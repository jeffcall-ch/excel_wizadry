Approved BOM Extractor Improvements
Summary of Approved Enhancements
Based on analysis of the 48-core AMD EPYC system processing 2500 engineering drawings, the following improvements have been approved to enhance performance, reliability, and production readiness.

1. CPU Core Utilization Optimization
Problem Statement
The current script defaults to cpu_count() - 1 workers, which on a 48-core system creates 47 workers but only reserves 1 core for the database writer. This leaves cores underutilized and doesn't account for system overhead or optimal resource distribution.

Strategy
Implement intelligent worker count calculation that considers:

Total available cores (48)
Reserved cores for database writer, progress monitor, and OS overhead
Dynamic scaling based on workload size
Memory constraints per worker
Implementation
class SystemResourceManager:
    def __init__(self, total_cores: int = 48, total_ram_gb: float = 32.0):
        self.total_cores = total_cores
        self.total_ram = total_ram_gb * 1024  # MB
        self.safe_ram = self.total_ram * 0.8  # 80% utilization
        self.os_overhead = 4 * 1024  # 4GB for OS
        
    def get_optimal_worker_count(self) -> int:
        """Calculate optimal worker count for 48-core EPYC system"""
        # Reserve cores: 1 DB writer, 1 progress monitor, 2 OS overhead  
        available_cores = self.total_cores - 4  # 44 workers
        
        # Ensure we don't exceed memory constraints
        available_worker_ram = self.safe_ram - self.os_overhead
        max_workers_by_memory = int(available_worker_ram // 512)  # 512MB per worker
        
        return min(available_cores, max_workers_by_memory)

# Integration into BOMExtractionConfig
def __post_init__(self):
    if self.max_workers is None:
        resource_manager = SystemResourceManager()
        self.max_workers = resource_manager.get_optimal_worker_count()
        logging.info(f"Optimized for {self.max_workers} workers on 48-core system")
2. Dynamic Memory Management
Problem Statement
Current script sets a fixed 2GB memory limit per worker. With 44+ workers, this could require 88GB+ RAM, but the system only has 32GB. This will cause unpredictable worker crashes and system instability.

Strategy
Implement dynamic memory allocation that:

Calculates safe per-worker memory limits based on available system RAM
Monitors actual memory usage during processing
Prevents system-wide memory exhaustion
Provides early warning for memory pressure
Implementation
class DynamicMemoryManager:
    def __init__(self, total_ram_gb: float = 32.0, num_workers: int = 44):
        self.total_ram = total_ram_gb * 1024  # MB
        self.safe_utilization = 0.8  # 80% of total RAM
        self.os_overhead = 4 * 1024  # 4GB for OS
        self.db_writer_allocation = 2 * 1024  # 2GB for DB operations
        self.num_workers = num_workers
        
        # Calculate per-worker allocation
        available_ram = (self.total_ram * self.safe_utilization) - self.os_overhead - self.db_writer_allocation
        self.per_worker_limit = max(256, int(available_ram / num_workers))  # Minimum 256MB
        
    def get_worker_memory_limit(self) -> int:
        """Get safe memory limit per worker"""
        return min(self.per_worker_limit, 1024)  # Cap at 1GB per worker

class EnhancedMemoryMonitor(SafeMemoryMonitor):
    def __init__(self, worker_id: int, memory_limit_mb: int):
        super().__init__()
        self.worker_id = worker_id
        self.memory_limit_mb = memory_limit_mb
        self.warning_threshold = memory_limit_mb * 0.8
        
    def check_memory_pressure(self) -> bool:
        """Check if approaching memory limit"""
        if self.process is None:
            return False
            
        try:
            current_mb = self.process.memory_info().rss / 1024 / 1024
            
            if current_mb > self.warning_threshold:
                logging.warning(f"Worker {self.worker_id} memory usage: {current_mb:.1f}MB (limit: {self.memory_limit_mb}MB)")
            
            if current_mb > self.memory_limit_mb:
                raise PDFProcessingError(
                    f"Worker {self.worker_id} exceeded memory limit: {current_mb:.1f}MB > {self.memory_limit_mb}MB",
                    ErrorCode.MEMORY_ERROR
                )
            
            return current_mb > self.warning_threshold
            
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return False

# Integration into worker processes
memory_manager = DynamicMemoryManager(32.0, config.max_workers)
worker_memory_limit = memory_manager.get_worker_memory_limit()
enhanced_monitor = EnhancedMemoryMonitor(worker_id, worker_memory_limit)
3. SQLite Database Writer Pool
Problem Statement
Single database writer becomes a bottleneck when 44 workers simultaneously send processing results. The queue fills up, workers timeout, and processing rate is limited by SQLite write performance rather than PDF processing capability.

Strategy
Implement multiple database writers with:

WAL mode optimization for concurrent writes
Load balancing across writer processes
Enhanced batching and transaction management
Improved error handling and recovery
Implementation
class DatabaseWriterPool:
    def __init__(self, db_path: str, num_writers: int = 3, results_queue_size: int = 2000):
        self.db_path = db_path
        self.num_writers = num_writers
        
        # Create separate queues for each DB writer to avoid contention
        self.writer_queues = [
            mp.Queue(maxsize=results_queue_size // num_writers) 
            for _ in range(num_writers)
        ]
        self.writer_processes = []
        self._setup_optimized_database()
        
    def _setup_optimized_database(self):
        """Configure SQLite for high-concurrency workloads"""
        with sqlite3.connect(self.db_path, timeout=30.0) as conn:
            # Optimize for concurrent writes
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL") 
            conn.execute("PRAGMA cache_size=-64000")  # 64MB cache
            conn.execute("PRAGMA temp_store=MEMORY")
            conn.execute("PRAGMA wal_autocheckpoint=10000")
            conn.execute("PRAGMA busy_timeout=30000")  # 30s timeout
            conn.execute("PRAGMA mmap_size=268435456")  # 256MB mmap
            
            # Create tables if they don't exist
            self._create_tables(conn)
            
    def start_writers(self):
        """Start multiple database writer processes"""
        for writer_id in range(self.num_writers):
            writer = OptimizedDatabaseWriter(
                db_path=self.db_path,
                results_queue=self.writer_queues[writer_id],
                writer_id=writer_id,
                batch_size=1000  # Larger batches for better performance
            )
            
            process = mp.Process(
                target=writer.run, 
                name=f"DBWriter-{writer_id}"
            )
            process.start()
            self.writer_processes.append(process)
            logging.info(f"Started database writer {writer_id} (PID: {process.pid})")
    
    def get_queue_for_worker(self, worker_id: int) -> mp.Queue:
        """Load balance workers across database writers"""
        return self.writer_queues[worker_id % self.num_writers]
    
    def shutdown(self):
        """Gracefully shutdown all database writers"""
        # Send shutdown signals
        for queue in self.writer_queues:
            try:
                shutdown_msg = WorkerMessage(
                    msg_type=MessageType.SHUTDOWN.value,
                    worker_id=-1,
                    timestamp=time.time()
                )
                queue.put_nowait(shutdown_msg.serialize())
            except:
                pass
        
        # Wait for processes to finish
        for process in self.writer_processes:
            process.join(timeout=30)
            if process.is_alive():
                process.terminate()

class OptimizedDatabaseWriter(DatabaseWriter):
    def __init__(self, db_path: str, results_queue: mp.Queue, writer_id: int, batch_size: int = 1000):
        self.writer_id = writer_id
        self.batch_size = batch_size
        super().__init__(db_path, results_queue, mp.Queue())  # No progress queue for pool writers
        
    def _flush_batch_optimized(self, bom_batch: List[BOMRow]):
        """Optimized batch insert with better error handling"""
        if not bom_batch:
            return
            
        try:
            with sqlite3.connect(self.db_path, timeout=45.0) as conn:
                # Use executemany with prepared statement
                conn.executemany("""
                    INSERT OR IGNORE INTO bom_data 
                    (id, filename, page, position, article_number, quantity, 
                     description, weight, total_weight, kks_codes, su_codes, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, [
                    (str(uuid.uuid4()), row.filename, row.page, row.position,
                     row.article_number, row.quantity, row.description,
                     row.weight, row.total_weight, row.kks_codes, row.su_codes,
                     datetime.now().isoformat())
                    for row in bom_batch
                ])
                conn.commit()
                
                self.logger.debug(f"Writer {self.writer_id} flushed {len(bom_batch)} rows")
                bom_batch.clear()
                
        except sqlite3.Error as e:
            self.logger.error(f"Writer {self.writer_id} database error: {e}")
            # Could implement retry logic here
            raise
7. Enhanced Progress Reporting
Problem Statement
Processing 2500 files can take several hours. Current progress reporting is basic and doesn't provide enough visibility into processing status, time estimates, or performance metrics needed for production monitoring.

Strategy
Implement comprehensive progress monitoring with:

Accurate time estimation based on rolling averages
Performance metrics and throughput tracking
Worker health monitoring
Detailed logging for production debugging
Implementation
class ProductionProgressMonitor:
    def __init__(self, total_files: int, progress_queue: mp.Queue):
        self.total_files = total_files
        self.progress_queue = SecureQueue(progress_queue, "progress")
        self.start_time = time.time()
        
        # Progress tracking
        self.processed_files = 0
        self.successful_files = 0
        self.failed_files = 0
        self.processing_rates = deque(maxlen=20)  # Rolling window
        
        # Worker monitoring
        self.active_workers = set()
        self.worker_last_seen = {}
        self.db_writer_health = DatabaseWriterHealthMonitor()
        
        # Performance metrics
        self.total_bom_rows = 0
        self.total_pages_processed = 0
        
        self.logger = logging.getLogger(f"{__name__}.ProductionProgressMonitor")
    
    def run(self):
        """Enhanced monitoring loop with detailed reporting"""
        last_report = 0
        last_rate_calculation = time.time()
        
        try:
            while True:
                msg_data = self.progress_queue.get_with_timeout(timeout=5.0)
                
                if msg_data is None:
                    # Timeout - perform health checks
                    self._check_system_health()
                    continue
                
                msg = WorkerMessage.deserialize(msg_data)
                
                if msg.msg_type == MessageType.SHUTDOWN.value:
                    break
                
                elif msg.msg_type == MessageType.PROGRESS.value:
                    self._process_progress_update(msg)
                
                elif msg.msg_type == MessageType.DB_WRITER_STATUS.value:
                    self.db_writer_health.update_heartbeat()
                    self._update_db_metrics(msg.data)
                
                elif msg.msg_type == MessageType.HEARTBEAT.value:
                    self._update_worker_activity(msg.worker_id)
                
                # Report progress periodically
                current_time = time.time()
                if (current_time - last_report) >= 10.0:  # Every 10 seconds
                    self._report_detailed_progress()
                    last_report = current_time
                
                # Calculate processing rate
                if (current_time - last_rate_calculation) >= 30.0:  # Every 30 seconds
                    self._update_processing_rate()
                    last_rate_calculation = current_time
        
        except Exception as e:
            self.logger.error(f"Progress monitor error: {e}")
        finally:
            self._report_final_statistics()
    
    def _report_detailed_progress(self):
        """Detailed progress report for production monitoring"""
        current_time = time.time()
        elapsed = current_time - self.start_time
        
        # Calculate progress percentage
        progress_pct = (self.processed_files / self.total_files) * 100 if self.total_files > 0 else 0
        
        # Calculate ETA using rolling average
        if self.processing_rates and len(self.processing_rates) >= 3:
            avg_rate = sum(self.processing_rates) / len(self.processing_rates)
            remaining_files = self.total_files - self.processed_files
            eta_seconds = remaining_files / avg_rate if avg_rate > 0 else 0
            eta_str = str(timedelta(seconds=int(eta_seconds)))
        else:
            eta_str = "Calculating..."
        
        # System health status
        db_status = "OK" if self.db_writer_health.check_health() else "⚠️ ISSUES"
        active_worker_count = len(self.active_workers)
        
        # Performance metrics
        success_rate = (self.successful_files / max(1, self.processed_files)) * 100
        
        self.logger.info(
            f"\n" + "="*80 +
            f"\n📊 PROGRESS REPORT - {datetime.now().strftime('%H:%M:%S')}" +
            f"\n📁 Files: {self.processed_files}/{self.total_files} ({progress_pct:.1f}%)" +
            f"\n✅ Success: {self.successful_files} ({success_rate:.1f}%) | ❌ Failed: {self.failed_files}" +
            f"\n⚡ Rate: {self.processing_rates[-1]:.1f} files/sec (avg)" +
            f"\n🕐 Elapsed: {str(timedelta(seconds=int(elapsed)))} | ETA: {eta_str}" +
            f"\n👥 Workers: {active_worker_count} active | 💾 Database: {db_status}" +
            f"\n📄 Total BOM rows: {self.total_bom_rows} | Pages: {self.total_pages_processed}" +
            f"\n" + "="*80
        )
    
    def _update_processing_rate(self):
        """Calculate current processing rate"""
        current_time = time.time()
        elapsed = current_time - self.start_time
        
        if elapsed > 0:
            current_rate = self.processed_files / elapsed
            self.processing_rates.append(current_rate)
    
    def _check_system_health(self):
        """Monitor system health and log warnings"""
        current_time = time.time()
        
        # Check for dead workers
        dead_workers = []
        for worker_id, last_seen in self.worker_last_seen.items():
            if current_time - last_seen > 60:  # No heartbeat for 60 seconds
                dead_workers.append(worker_id)
        
        if dead_workers:
            self.logger.warning(f"Workers not responding: {dead_workers}")
        
        # Check database writer health
        if not self.db_writer_health.check_health():
            self.logger.critical("Database writer appears unresponsive!")
    
    def _report_final_statistics(self):
        """Comprehensive final report"""
        total_time = time.time() - self.start_time
        
        self.logger.info(
            f"\n" + "="*80 +
            f"\n🏁 PROCESSING COMPLETE - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}" +
            f"\n📁 Total Files: {self.total_files}" +
            f"\n✅ Successful: {self.successful_files}" +
            f"\n❌ Failed: {self.failed_files}" +
            f"\n📊 Success Rate: {(self.successful_files/max(1,self.total_files)*100):.1f}%" +
            f"\n⏱️  Total Time: {str(timedelta(seconds=int(total_time)))}" +
            f"\n⚡ Average Rate: {(self.total_files/total_time):.1f} files/sec" +
            f"\n📄 Total BOM Rows Extracted: {self.total_bom_rows}" +
            f"\n📃 Total Pages Processed: {self.total_pages_processed}" +
            f"\n" + "="*80
        )
9. Production Error Recovery System
Problem Statement
Processing 2500 files over several hours requires robust error recovery. Current implementation cannot resume interrupted jobs, retry failed files, or provide detailed failure analysis for production debugging.

Strategy
Implement checkpoint-based recovery system with:

Persistent state tracking across runs
Selective retry capability for failed files
Detailed error classification and reporting
Resume capability for interrupted jobs
Implementation
class ProductionCheckpointManager:
    def __init__(self, checkpoint_file: str = "bom_processing_checkpoint.json"):
        self.checkpoint_file = Path(checkpoint_file)
        self.lock_file = Path(f"{checkpoint_file}.lock")
        
        # Processing state
        self.processed_files: Set[str] = set()
        self.failed_files: Dict[str, Dict] = {}  # filename -> error info
        self.processing_stats = {
            'start_time': None,
            'last_checkpoint': None,
            'total_files': 0,
            'successful_files': 0,
            'failed_files': 0,
            'bom_rows_extracted': 0
        }
        
        self.logger = logging.getLogger(f"{__name__}.CheckpointManager")
        self._load_checkpoint()
    
    def _load_checkpoint(self):
        """Load existing checkpoint with file locking"""
        if not self.checkpoint_file.exists():
            return
        
        try:
            with self._file_lock():
                with open(self.checkpoint_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.processed_files = set(data.get('processed_files', []))
                self.failed_files = data.get('failed_files', {})
                self.processing_stats.update(data.get('stats', {}))
                
                self.logger.info(
                    f"Loaded checkpoint: {len(self.processed_files)} processed, "
                    f"{len(self.failed_files)} failed"
                )
        
        except Exception as e:
            self.logger.error(f"Failed to load checkpoint: {e}")
            # Continue with empty state
    
    def save_checkpoint(self):
        """Save current state to checkpoint file"""
        try:
            checkpoint_data = {
                'processed_files': list(self.processed_files),
                'failed_files': self.failed_files,
                'stats': self.processing_stats,
                'timestamp': datetime.now().isoformat(),
                'version': '1.0'
            }
            
            # Atomic write using temporary file
            temp_file = self.checkpoint_file.with_suffix('.tmp')
            
            with self._file_lock():
                with open(temp_file, 'w', encoding='utf-8') as f:
                    json.dump(checkpoint_data, f, indent=2, ensure_ascii=False)
                
                # Atomic rename
                temp_file.replace(self.checkpoint_file)
                
            self.processing_stats['last_checkpoint'] = time.time()
            
        except Exception as e:
            self.logger.error(f"Failed to save checkpoint: {e}")
    
    @contextmanager
    def _file_lock(self):
        """Simple file locking for checkpoint access"""
        lock_acquired = False
        try:
            # Try to acquire lock
            for attempt in range(10):
                try:
                    with open(self.lock_file, 'x') as f:
                        f.write(str(os.getpid()))
                    lock_acquired = True
                    break
                except FileExistsError:
                    time.sleep(0.1)
            
            if not lock_acquired:
                raise RuntimeError("Could not acquire checkpoint lock")
            
            yield
            
        finally:
            if lock_acquired:
                try:
                    self.lock_file.unlink()
                except FileNotFoundError:
                    pass
    
    def should_process_file(self, file_path: str) -> bool:
        """Check if file should be processed"""
        return file_path not in self.processed_files
    
    def mark_file_completed(self, file_path: str, result: ProcessingResult):
        """Mark file as completed and update statistics"""
        if result.success:
            self.processed_files.add(file_path)
            self.failed_files.pop(file_path, None)  # Remove from failed if retrying
            self.processing_stats['successful_files'] += 1
            self.processing_stats['bom_rows_extracted'] += len(result.bom_rows)
        else:
            # Record detailed failure information
            self.failed_files[file_path] = {
                'error_code': result.error_code.value,
                'error_message': result.error_message,
                'file_size_mb': result.file_size_bytes / 1024 / 1024,
                'total_pages': result.total_pages,
                'timestamp': datetime.now().isoformat(),
                'retry_count': result.retry_count
            }
            self.processing_stats['failed_files'] += 1
        
        # Save checkpoint every 10 files
        if (self.processing_stats['successful_files'] + self.processing_stats['failed_files']) % 10 == 0:
            self.save_checkpoint()
    
    def get_failed_files_for_retry(self, max_retries: int = 3) -> List[str]:
        """Get list of failed files that should be retried"""
        retry_files = []
        
        for file_path, error_info in self.failed_files.items():
            retry_count = error_info.get('retry_count', 0)
            error_code = error_info.get('error_code', '')
            
            # Don't retry certain permanent failures
            permanent_failures = [
                ErrorCode.FILE_NOT_FOUND.value,
                ErrorCode.FILE_TOO_LARGE.value,
                ErrorCode.INVALID_PDF.value,
                ErrorCode.FILE_EMPTY.value
            ]
            
            if error_code not in permanent_failures and retry_count < max_retries:
                retry_files.append(file_path)
        
        return retry_files
    
    def generate_failure_report(self) -> str:
        """Generate detailed failure analysis report"""
        if not self.failed_files:
            return "No failed files to report."
        
        # Group failures by error code
        error_groups = defaultdict(list)
        for file_path, error_info in self.failed_files.items():
            error_code = error_info.get('error_code', 'UNKNOWN')
            error_groups[error_code].append((file_path, error_info))
        
        report_lines = [
            "="*80,
            "BOM EXTRACTION FAILURE REPORT",
            "="*80,
            f"Total Failed Files: {len(self.failed_files)}",
            ""
        ]
        
        for error_code, files in error_groups.items():
            report_lines.extend([
                f"Error Type: {error_code} ({len(files)} files)",
                "-" * 40
            ])
            
            for file_path, error_info in files[:5]:  # Show first 5 files per error type
                report_lines.append(
                    f"  {Path(file_path).name}: {error_info.get('error_message', 'No message')}"
                )
            
            if len(files) > 5:
                report_lines.append(f"  ... and {len(files) - 5} more files")
            
            report_lines.append("")
        
        return "\n".join(report_lines)

class ProductionJobManager:
    def __init__(self, config: BOMExtractionConfig):
        self.config = config
        self.checkpoint_manager = ProductionCheckpointManager()
        self.logger = logging.getLogger(f"{__name__}.JobManager")
    
    def run_with_recovery(self, resume: bool = False, retry_failed: bool = False):
        """Run extraction with recovery capabilities"""
        
        # Discover all files
        all_files = find_pdf_files(self.config.input_directory, self.config.max_files_to_process)
        
        if resume:
            # Filter out already processed files
            files_to_process = [f for f in all_files if self.checkpoint_manager.should_process_file(f)]
            self.logger.info(f"Resuming: {len(files_to_process)} files remaining out of {len(all_files)}")
        
        elif retry_failed:
            # Only process previously failed files
            files_to_process = self.checkpoint_manager.get_failed_files_for_retry()
            self.logger.info(f"Retrying {len(files_to_process)} failed files")
        
        else:
            # Fresh start
            files_to_process = all_files
            self.checkpoint_manager.processing_stats['start_time'] = time.time()
            self.checkpoint_manager.processing_stats['total_files'] = len(all_files)
        
        if not files_to_process:
            self.logger.info("No files to process")
            return
        
        try:
            # Run the extraction
            self._process_files_with_checkpoint(files_to_process)
            
        finally:
            # Always save final checkpoint and generate report
            self.checkpoint_manager.save_checkpoint()
            
            if self.checkpoint_manager.failed_files:
                failure_report = self.checkpoint_manager.generate_failure_report()
                self.logger.warning(f"\n{failure_report}")
                
                # Save failure report to file
                report_file = Path("bom_extraction_failures.txt")
                with open(report_file, 'w', encoding='utf-8') as f:
                    f.write(failure_report)
                self.logger.info(f"Detailed failure report saved to: {report_file}")

# Integration with main extractor
def main() -> int:
    parser.add_argument("--resume", action="store_true", 
                       help="Resume interrupted processing from checkpoint")
    parser.add_argument("--retry-failed", action="store_true",
                       help="Retry only previously failed files")
    parser.add_argument("--checkpoint", default="bom_processing_checkpoint.json",
                       help="Checkpoint file path")
    
    # In main execution
    if args.resume or args.retry_failed:
        job_manager = ProductionJobManager(config)
        job_manager.run_with_recovery(resume=args.resume, retry_failed=args.retry_failed)
    else:
        # Standard processing
        extractor = QueueBasedBOMExtractor(config)
        extractor.run()
Implementation Priority
Memory Management (Priority 1) - Prevents system crashes
CPU Optimization (Priority 2) - Maximizes performance on 48-core system
Database Writer Pool (Priority 2) - Removes processing bottleneck
Error Recovery (Priority 3) - Essential for production reliability
Progress Reporting (Priority 3) - Important for monitoring long-running jobs
These improvements will transform the script from a development tool into a production-ready system capable of reliably processing 2500 engineering drawings on your high-performance hardware.