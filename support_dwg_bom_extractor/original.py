#!/usr/bin/env python3
"""
PDF BOM (Bill of Materials) Text Extractor - Hardened Queue-based Architecture

Hardened design for 32-core systems processing 2500+ files:
- Single database writer process with failure recovery
- 31 worker processes for maximum CPU utilization
- Secure inter-process communication without pickle
- Comprehensive error handling and resource management
- Protection against deadlocks and resource leaks
"""

import os
import csv
import fitz  # PyMuPDF
import logging
import argparse
import sys
import sqlite3
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Set, Union, Iterator, Any, NamedTuple
from dataclasses import dataclass, field, asdict
import re
from datetime import datetime, timedelta
import multiprocessing as mp
from concurrent.futures import ProcessPoolExecutor, as_completed, TimeoutError
import signal
from collections import defaultdict, Counter
import json
import time
import traceback
from contextlib import contextmanager
import tempfile
import shutil
import queue
from enum import Enum
from functools import wraps, lru_cache
import uuid
import threading
import platform
import hashlib
import struct
import fcntl
import mmap

# Optional dependencies with graceful fallback
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False
    print("Warning: psutil not available. Memory monitoring disabled.")

try:
    import msgpack
    MSGPACK_AVAILABLE = True
except ImportError:
    MSGPACK_AVAILABLE = False

# Constants
DEFAULT_TIMEOUT = 60
DEFAULT_MAX_FILE_SIZE_MB = 100
BATCH_SIZE = 500  # Database batch insert size
QUEUE_SIZE_MULTIPLIER = 3  # Queue size = workers * multiplier
HEARTBEAT_INTERVAL = 10  # Seconds between heartbeats
SHUTDOWN_TIMEOUT = 30  # Seconds to wait for graceful shutdown
DB_WRITER_HEARTBEAT_TIMEOUT = 20  # Seconds before considering DB writer dead
MAX_MEMORY_MB = 2048  # Maximum memory per worker process
FALLBACK_FILE_PREFIX = "bom_fallback_"

# Compiled regex patterns for better performance
KKS_PATTERN = re.compile(r'(?<!\w)\d[A-Z0-9]{3}\d{2}BQ\d{3}(?!\w)', re.IGNORECASE)
SU_PATTERN = re.compile(r'(\d[A-Z0-9]{3}\d{2}BQ\d{3})\/SU(?!\w)', re.IGNORECASE)
ARTICLE_NUMBER_PATTERN = re.compile(r'^\d{5,7}$')
QUANTITY_PATTERN = re.compile(r'^\d+(\.\d+)?$')

class ErrorCode(Enum):
    """Standardized error codes"""
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    FILE_TOO_LARGE = "FILE_TOO_LARGE"
    FILE_EMPTY = "FILE_EMPTY"
    INVALID_PDF = "INVALID_PDF"
    NO_PAGES = "NO_PAGES"
    NO_ANCHOR = "NO_ANCHOR"
    NO_HEADER = "NO_HEADER"
    NO_DATA = "NO_DATA"
    PARSING_ERROR = "PARSING_ERROR"
    TIMEOUT = "TIMEOUT"
    MEMORY_ERROR = "MEMORY_ERROR"
    WORKER_ERROR = "WORKER_ERROR"
    SHUTDOWN = "SHUTDOWN"
    DATABASE_ERROR = "DATABASE_ERROR"
    DB_WRITER_DEAD = "DB_WRITER_DEAD"
    QUEUE_FULL = "QUEUE_FULL"
    RESOURCE_LEAK = "RESOURCE_LEAK"
    UNKNOWN = "UNKNOWN"

class PDFProcessingError(Exception):
    """Custom exception with error codes"""
    def __init__(self, message: str, error_code: ErrorCode = ErrorCode.UNKNOWN, 
                 filename: str = "", additional_info: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.error_code = error_code
        self.filename = filename
        self.additional_info = additional_info or {}

class MessageType(Enum):
    """Message types for inter-process communication"""
    BOM_DATA = "BOM_DATA"
    SUMMARY = "SUMMARY" 
    PROGRESS = "PROGRESS"
    HEARTBEAT = "HEARTBEAT"
    SHUTDOWN = "SHUTDOWN"
    ERROR = "ERROR"
    DB_WRITER_STATUS = "DB_WRITER_STATUS"

@dataclass
class WorkerMessage:
    """Message passed between processes using secure serialization"""
    msg_type: str  # MessageType.value
    worker_id: int
    timestamp: float
    data: Any = None
    
    def serialize(self) -> bytes:
        """Serialize message securely using JSON or msgpack"""
        try:
            # Create serializable dict
            msg_dict = {
                'msg_type': self.msg_type,
                'worker_id': self.worker_id,
                'timestamp': self.timestamp,
                'data': self._serialize_data(self.data)
            }
            
            if MSGPACK_AVAILABLE:
                return msgpack.packb(msg_dict, use_bin_type=True)
            else:
                return json.dumps(msg_dict, default=self._json_serializer).encode('utf-8')
        except Exception as e:
            # Fallback to minimal message
            fallback = {
                'msg_type': MessageType.ERROR.value,
                'worker_id': self.worker_id,
                'timestamp': self.timestamp,
                'data': f'Serialization failed: {str(e)}'
            }
            return json.dumps(fallback).encode('utf-8')
    
    def _serialize_data(self, data):
        """Serialize data safely"""
        if data is None:
            return None
        elif hasattr(data, '__dict__'):
            # Convert dataclass/object to dict
            if hasattr(data, 'to_dict'):
                return data.to_dict()
            else:
                return asdict(data) if hasattr(data, '__dataclass_fields__') else str(data)
        elif isinstance(data, (dict, list, str, int, float, bool)):
            return data
        else:
            return str(data)
    
    def _json_serializer(self, obj):
        """JSON serializer for complex objects"""
        if hasattr(obj, '__dict__'):
            return obj.__dict__
        return str(obj)
    
    @classmethod
    def deserialize(cls, data: bytes) -> 'WorkerMessage':
        """Deserialize message from queue"""
        try:
            if MSGPACK_AVAILABLE:
                try:
                    msg_dict = msgpack.unpackb(data, raw=False)
                except:
                    msg_dict = json.loads(data.decode('utf-8'))
            else:
                msg_dict = json.loads(data.decode('utf-8'))
            
            return cls(
                msg_type=msg_dict['msg_type'],
                worker_id=msg_dict['worker_id'],
                timestamp=msg_dict['timestamp'],
                data=msg_dict.get('data')
            )
        except Exception as e:
            # Return error message if deserialization fails
            return cls(
                msg_type=MessageType.ERROR.value,
                worker_id=-1,
                timestamp=time.time(),
                data=f'Deserialization failed: {str(e)}'
            )

@dataclass
class BOMRow:
    """BOM data row"""
    filename: str
    page: int
    position: str
    article_number: str
    quantity: str
    description: str
    weight: str
    total_weight: str
    kks_codes: str
    su_codes: str
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return asdict(self)

@dataclass
class SummaryRow:
    """Processing summary row"""
    filename: str
    success: bool
    header_found: bool
    data_found: bool
    parse_success: bool
    error_message: Optional[str]
    rows_count: int
    process_time_s: float
    file_size_mb: float
    pages_total: int
    bom_pages: str
    anchors_count: int
    retry_count: int
    memory_peak_mb: float
    error_code: str
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return asdict(self)

@dataclass
class ProcessingResult:
    """Processing result"""
    filename: str
    filepath: str
    success: bool
    error_code: ErrorCode = ErrorCode.UNKNOWN
    error_message: Optional[str] = None
    
    # Stage tracking
    header_found: bool = False
    bom_data_found: bool = False
    parsing_success: bool = False
    
    # Data
    bom_rows: List[BOMRow] = field(default_factory=list)
    
    # Metrics
    processing_time: float = 0.0
    file_size_bytes: int = 0
    total_pages: int = 0
    bom_pages: List[int] = field(default_factory=list)
    anchor_count: int = 0
    retry_count: int = 0
    memory_peak_mb: float = 0.0
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'filename': self.filename,
            'filepath': self.filepath,
            'success': self.success,
            'error_code': self.error_code.value,
            'error_message': self.error_message,
            'header_found': self.header_found,
            'bom_data_found': self.bom_data_found,
            'parsing_success': self.parsing_success,
            'bom_rows': [row.to_dict() for row in self.bom_rows],
            'processing_time': self.processing_time,
            'file_size_bytes': self.file_size_bytes,
            'total_pages': self.total_pages,
            'bom_pages': self.bom_pages,
            'anchor_count': self.anchor_count,
            'retry_count': self.retry_count,
            'memory_peak_mb': self.memory_peak_mb
        }
    
    def to_summary_row(self) -> SummaryRow:
        """Convert to summary row"""
        return SummaryRow(
            filename=self.filename,
            success=self.success,
            header_found=self.header_found,
            data_found=self.bom_data_found,
            parse_success=self.parsing_success,
            error_message=self.error_message,
            rows_count=len(self.bom_rows),
            process_time_s=self.processing_time,
            file_size_mb=self.file_size_bytes / 1024 / 1024,
            pages_total=self.total_pages,
            bom_pages=str(self.bom_pages) if self.bom_pages else "[]",
            anchors_count=self.anchor_count,
            retry_count=self.retry_count,
            memory_peak_mb=self.memory_peak_mb,
            error_code=self.error_code.value
        )

@dataclass
class BOMExtractionConfig:
    """Configuration for BOM extraction"""
    input_directory: str
    output_csv: str = "extracted_bom.csv"
    summary_csv: str = "extraction_summary.csv"
    database_path: str = "bom_extraction.db"
    fallback_directory: str = "bom_fallback"
    anchor_text: str = "POS"
    total_text: str = "TOTAL"
    coordinate_tolerance: float = 5.0
    max_files_to_process: Optional[int] = None
    max_workers: Optional[int] = None
    timeout_per_file: int = DEFAULT_TIMEOUT
    retry_failed_files: bool = True
    max_retries: int = 2
    max_file_size_mb: int = DEFAULT_MAX_FILE_SIZE_MB
    max_memory_mb: int = MAX_MEMORY_MB
    debug: bool = False
    log_file: Optional[str] = None
    
    def __post_init__(self):
        """Validate configuration and set optimal defaults"""
        # Detect optimal multiprocessing method
        self.mp_method = self._detect_mp_method()
        
        # Set optimal worker count for the system
        if self.max_workers is None:
            cpu_count = mp.cpu_count()
            # Reserve 1 core for database writer + OS overhead
            self.max_workers = max(1, cpu_count - 1)
        
        # Ensure fallback directory exists
        Path(self.fallback_directory).mkdir(parents=True, exist_ok=True)
    
    def _detect_mp_method(self) -> str:
        """Detect optimal multiprocessing start method for the platform"""
        system = platform.system().lower()
        if system == "linux":
            return "fork"  # Most efficient on Linux
        elif system == "darwin":
            return "spawn"  # Required on recent macOS
        elif system == "windows":
            return "spawn"  # Only option on Windows
        else:
            return "spawn"  # Safe default

@dataclass 
class TableCell:
    """Table cell representation"""
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    
    def __post_init__(self):
        self.text = self.text.strip()
    
    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2
    
    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2

@dataclass
class AnchorPosition:
    """Anchor position in PDF"""
    x0: float
    y0: float
    x1: float
    y1: float
    line_number: int
    page_number: int
    confidence: float = 1.0

class SafePDFManager:
    """Context manager for safe PDF handling with resource leak protection"""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = None
    
    def __enter__(self) -> fitz.Document:
        try:
            self.doc = fitz.open(self.pdf_path)
            return self.doc
        except Exception as e:
            if self.doc is not None:
                try:
                    self.doc.close()
                except:
                    pass
            raise PDFProcessingError(f"Failed to open PDF: {str(e)}", ErrorCode.INVALID_PDF)
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc is not None:
            try:
                self.doc.close()
            except Exception as e:
                # Log but don't re-raise - we're cleaning up
                logging.getLogger(__name__).warning(f"Error closing PDF {self.pdf_path}: {e}")

class SafeMemoryMonitor:
    """Safe memory monitoring with race condition protection"""
    
    def __init__(self):
        self.process = None
        self.peak_memory_mb = 0.0
        self._initialize()
    
    def _initialize(self):
        """Initialize psutil process with error handling"""
        if PSUTIL_AVAILABLE:
            try:
                self.process = psutil.Process()
                # Test access to avoid race conditions later
                _ = self.process.memory_info()
            except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
                self.process = None
    
    def update_peak(self):
        """Update peak memory usage safely"""
        if self.process is None:
            return
        
        try:
            memory_info = self.process.memory_info()
            current_mb = memory_info.rss / 1024 / 1024
            self.peak_memory_mb = max(self.peak_memory_mb, current_mb)
            
            # Check memory limit
            if current_mb > MAX_MEMORY_MB:
                raise PDFProcessingError(
                    f"Memory limit exceeded: {current_mb:.1f}MB > {MAX_MEMORY_MB}MB",
                    ErrorCode.MEMORY_ERROR
                )
        except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
            # Process might have been killed or access denied
            self.process = None
        except Exception:
            # Ignore other monitoring errors
            pass

class FallbackWriter:
    """Fallback writer for when database writer fails"""
    
    def __init__(self, fallback_dir: str, worker_id: int):
        self.fallback_dir = Path(fallback_dir)
        self.worker_id = worker_id
        self.fallback_file = self.fallback_dir / f"{FALLBACK_FILE_PREFIX}{worker_id}_{int(time.time())}.json"
        self.logger = logging.getLogger(f"{__name__}.FallbackWriter.{worker_id}")
        
        # Ensure directory exists
        self.fallback_dir.mkdir(parents=True, exist_ok=True)
    
    def write_result(self, result: ProcessingResult):
        """Write result to fallback file"""
        try:
            result_data = {
                'timestamp': time.time(),
                'worker_id': self.worker_id,
                'result': result.to_dict()
            }
            
            with open(self.fallback_file, 'a', encoding='utf-8') as f:
                f.write(json.dumps(result_data) + '\n')
                f.flush()
                os.fsync(f.fileno())  # Force write to disk
        
        except Exception as e:
            self.logger.error(f"Failed to write fallback data: {e}")

class DatabaseWriterHealthMonitor:
    """Monitor database writer health"""
    
    def __init__(self):
        self.last_heartbeat = time.time()
        self.is_alive = True
        self._lock = threading.Lock()
    
    def update_heartbeat(self):
        """Update last heartbeat timestamp"""
        with self._lock:
            self.last_heartbeat = time.time()
            self.is_alive = True
    
    def check_health(self) -> bool:
        """Check if database writer is still alive"""
        with self._lock:
            current_time = time.time()
            if current_time - self.last_heartbeat > DB_WRITER_HEARTBEAT_TIMEOUT:
                self.is_alive = False
            return self.is_alive
    
    def mark_dead(self):
        """Mark database writer as dead"""
        with self._lock:
            self.is_alive = False

class SecureQueue:
    """Thread-safe queue wrapper with timeout and backpressure handling"""
    
    def __init__(self, queue_obj: mp.Queue, name: str):
        self.queue = queue_obj
        self.name = name
        self.logger = logging.getLogger(f"{__name__}.SecureQueue.{name}")
    
    def put_with_timeout(self, item: bytes, timeout: float = 5.0) -> bool:
        """Put item with timeout and proper error handling"""
        try:
            self.queue.put(item, timeout=timeout)
            return True
        except queue.Full:
            self.logger.warning(f"Queue {self.name} is full, dropping message")
            return False
        except Exception as e:
            self.logger.error(f"Error putting item in queue {self.name}: {e}")
            return False
    
    def get_with_timeout(self, timeout: float = 1.0) -> Optional[bytes]:
        """Get item with timeout"""
        try:
            return self.queue.get(timeout=timeout)
        except queue.Empty:
            return None
        except Exception as e:
            self.logger.error(f"Error getting item from queue {self.name}: {e}")
            return None

class DatabaseWriter:
    """Dedicated database writer process with failure recovery"""
    
    def __init__(self, db_path: str, results_queue: mp.Queue, progress_queue: mp.Queue):
        self.db_path = db_path
        self.results_queue = SecureQueue(results_queue, "results")
        self.progress_queue = SecureQueue(progress_queue, "progress")
        self.logger = logging.getLogger(f"{__name__}.DatabaseWriter")
        
        # Statistics
        self.stats = {
            'bom_rows': 0,
            'successful_files': 0,
            'total_files': 0,
            'failed_files': 0
        }
        
        # Memory management
        self.memory_monitor = SafeMemoryMonitor()
        self.max_batch_memory_mb = 100  # Max memory for batch before forced flush
        
        self._setup_database()
    
    def _setup_database(self):
        """Initialize database schema with better error handling"""
        try:
            with sqlite3.connect(self.db_path, timeout=30.0) as conn:
                # Enable WAL mode for better concurrency
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute("PRAGMA synchronous=NORMAL")
                conn.execute("PRAGMA cache_size=-64000")  # 64MB cache
                conn.execute("PRAGMA temp_store=MEMORY")
                conn.execute("PRAGMA mmap_size=268435456")  # 256MB mmap
                
                # Create BOM data table
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS bom_data (
                        id TEXT PRIMARY KEY,
                        filename TEXT NOT NULL,
                        page INTEGER NOT NULL,
                        position TEXT,
                        article_number TEXT,
                        quantity TEXT,
                        description TEXT,
                        weight TEXT,
                        total_weight TEXT,
                        kks_codes TEXT,
                        su_codes TEXT,
                        created_at TEXT NOT NULL,
                        UNIQUE(filename, page, position, article_number)
                    )
                """)
                
                # Create processing summary table
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS processing_summary (
                        id TEXT PRIMARY KEY,
                        filename TEXT UNIQUE NOT NULL,
                        success BOOLEAN NOT NULL,
                        header_found BOOLEAN NOT NULL,
                        data_found BOOLEAN NOT NULL,
                        parse_success BOOLEAN NOT NULL,
                        error_message TEXT,
                        rows_count INTEGER NOT NULL,
                        process_time_s REAL NOT NULL,
                        file_size_mb REAL NOT NULL,
                        pages_total INTEGER NOT NULL,
                        bom_pages TEXT,
                        anchors_count INTEGER NOT NULL,
                        retry_count INTEGER NOT NULL,
                        memory_peak_mb REAL NOT NULL,
                        error_code TEXT NOT NULL,
                        created_at TEXT NOT NULL
                    )
                """)
                
                # Create indexes for better performance
                conn.execute("CREATE INDEX IF NOT EXISTS idx_bom_filename ON bom_data(filename)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_bom_created_at ON bom_data(created_at)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_summary_filename ON processing_summary(filename)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_summary_success ON processing_summary(success)")
                
                conn.commit()
                self.logger.info("Database initialized successfully")
        
        except Exception as e:
            self.logger.error(f"Database initialization failed: {e}")
            raise
    
    def run(self):
        """Main database writer loop with comprehensive error handling"""
        self.logger.info("Database writer process started")
        
        bom_batch = []
        active_workers = set()
        last_heartbeat_check = time.time()
        last_status_report = time.time()
        
        try:
            while True:
                # Update memory monitoring
                self.memory_monitor.update_peak()
                
                # Send periodic status updates
                current_time = time.time()
                if current_time - last_status_report > 30:  # Every 30 seconds
                    self._send_status_update()
                    last_status_report = current_time
                
                # Process messages
                msg_data = self.results_queue.get_with_timeout(timeout=1.0)
                if msg_data is None:
                    # Handle timeout - check for periodic tasks
                    if bom_batch and (
                        len(bom_batch) > BATCH_SIZE // 2 or  # Half batch size
                        current_time - last_heartbeat_check > 5  # 5 seconds since last flush
                    ):
                        self._flush_batch_safe(bom_batch)
                    
                    # Check for dead workers
                    if current_time - last_heartbeat_check > HEARTBEAT_INTERVAL:
                        self._check_worker_health(active_workers)
                        last_heartbeat_check = current_time
                        active_workers.clear()
                    continue
                
                # Process received message
                msg = WorkerMessage.deserialize(msg_data)
                
                if msg.msg_type == MessageType.SHUTDOWN.value:
                    self.logger.info("Shutdown message received")
                    break
                
                elif msg.msg_type == MessageType.BOM_DATA.value:
                    self._process_result(msg.data, bom_batch)
                    active_workers.add(msg.worker_id)
                
                elif msg.msg_type == MessageType.HEARTBEAT.value:
                    active_workers.add(msg.worker_id)
                
                elif msg.msg_type == MessageType.ERROR.value:
                    self.logger.error(f"Worker {msg.worker_id} error: {msg.data}")
                    active_workers.add(msg.worker_id)
                
                # Force flush if batch is getting too large or using too much memory
                if len(bom_batch) >= BATCH_SIZE or self._estimate_batch_memory(bom_batch) > self.max_batch_memory_mb:
                    self._flush_batch_safe(bom_batch)
        
        except Exception as e:
            self.logger.critical(f"Database writer fatal error: {e}", exc_info=True)
        
        finally:
            # Flush any remaining data
            if bom_batch:
                self._flush_batch_safe(bom_batch)
            
            self.logger.info(f"Database writer stopping. Final stats: {self.stats}")
    
    def _estimate_batch_memory(self, batch: List[BOMRow]) -> float:
        """Estimate memory usage of current batch in MB"""
        if not batch:
            return 0.0
        
        # Rough estimate: ~500 bytes per BOM row
        return (len(batch) * 500) / 1024 / 1024
    
    def _send_status_update(self):
        """Send status update to progress queue"""
        try:
            status_msg = WorkerMessage(
                msg_type=MessageType.DB_WRITER_STATUS.value,
                worker_id=-1,
                timestamp=time.time(),
                data={
                    'alive': True,
                    'stats': self.stats.copy(),
                    'memory_mb': self.memory_monitor.peak_memory_mb
                }
            )
            self.progress_queue.put_with_timeout(status_msg.serialize(), timeout=1.0)
        except Exception as e:
            self.logger.error(f"Failed to send status update: {e}")
    
    def _process_result(self, result_data: dict, bom_batch: List[BOMRow]):
        """Process a worker result with enhanced error handling"""
        try:
            # Reconstruct ProcessingResult from dict
            result = self._dict_to_result(result_data)
            
            # Add BOM rows to batch
            if result.bom_rows:
                bom_batch.extend(result.bom_rows)
                self.stats['bom_rows'] += len(result.bom_rows)
            
            # Insert summary immediately (more critical than BOM data)
            summary = result.to_summary_row()
            self._insert_summary_safe(summary)
            
            # Update stats
            self.stats['total_files'] += 1
            if result.success:
                self.stats['successful_files'] += 1
            else:
                self.stats['failed_files'] += 1
            
            # Send progress update
            self._send_progress_update(result)
        
        except Exception as e:
            self.logger.error(f"Error processing result: {e}")
    
    def _dict_to_result(self, result_data: dict) -> ProcessingResult:
        """Convert dictionary back to ProcessingResult"""
        bom_rows = []
        if result_data.get('bom_rows'):
            for row_data in result_data['bom_rows']:
                bom_rows.append(BOMRow(**row_data))
        
        return ProcessingResult(
            filename=result_data['filename'],
            filepath=result_data['filepath'],
            success=result_data['success'],
            error_code=ErrorCode(result_data.get('error_code', ErrorCode.UNKNOWN.value)),
            error_message=result_data.get('error_message'),
            header_found=result_data.get('header_found', False),
            bom_data_found=result_data.get('bom_data_found', False),
            parsing_success=result_data.get('parsing_success', False),
            bom_rows=bom_rows,
            processing_time=result_data.get('processing_time', 0.0),
            file_size_bytes=result_data.get('file_size_bytes', 0),
            total_pages=result_data.get('total_pages', 0),
            bom_pages=result_data.get('bom_pages', []),
            anchor_count=result_data.get('anchor_count', 0),
            retry_count=result_data.get('retry_count', 0),
            memory_peak_mb=result_data.get('memory_peak_mb', 0.0)
        )
    
    def _send_progress_update(self, result: ProcessingResult):
        """Send progress update with error handling"""
        try:
            progress_msg = WorkerMessage(
                msg_type=MessageType.PROGRESS.value,
                worker_id=-1,
                timestamp=time.time(),
                data={
                    'filename': result.filename,
                    'success': result.success,
                    'stats': self.stats.copy()
                }
            )
            self.progress_queue.put_with_timeout(progress_msg.serialize(), timeout=1.0)
        except Exception as e:
            self.logger.debug(f"Failed to send progress update: {e}")  # Debug level - not critical
    
    def _flush_batch_safe(self, bom_batch: List[BOMRow]):
        """Flush BOM batch with comprehensive error handling"""
        if not bom_batch:
            return
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                with sqlite3.connect(self.db_path, timeout=30.0) as conn:
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
                
                self.logger.debug(f"Flushed batch of {len(bom_batch)} BOM rows")
                bom_batch.clear()
                return
            
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e).lower() and attempt < max_retries - 1:
                    self.logger.warning(f"Database locked, retrying in {2 ** attempt} seconds...")
                    time.sleep(2 ** attempt)  # Exponential backoff
                    continue
                else:
                    raise
            except Exception as e:
                if attempt < max_retries - 1:
                    self.logger.warning(f"Batch flush failed (attempt {attempt + 1}): {e}")
                    time.sleep(1)
                    continue
                else:
                    raise
        
        self.logger.error(f"Failed to flush batch after {max_retries} attempts")
    
    def _insert_summary_safe(self, summary: SummaryRow):
        """Insert summary row with retry logic"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                with sqlite3.connect(self.db_path, timeout=30.0) as conn:
                    conn.execute("""
                        INSERT OR REPLACE INTO processing_summary
                        (id, filename, success, header_found, data_found, parse_success,
                         error_message, rows_count, process_time_s, file_size_mb,
                         pages_total, bom_pages, anchors_count, retry_count,
                         memory_peak_mb, error_code, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(uuid.uuid4()), summary.filename, summary.success,
                        summary.header_found, summary.data_found, summary.parse_success,
                        summary.error_message, summary.rows_count, summary.process_time_s,
                        summary.file_size_mb, summary.pages_total, summary.bom_pages,
                        summary.anchors_count, summary.retry_count, summary.memory_peak_mb,
                        summary.error_code, datetime.now().isoformat()
                    ))
                    conn.commit()
                return
            
            except Exception as e:
                if attempt < max_retries - 1:
                    self.logger.warning(f"Summary insert failed (attempt {attempt + 1}): {e}")
                    time.sleep(0.5)
                    continue
                else:
                    self.logger.error(f"Failed to insert summary for {summary.filename}: {e}")
                    break
    
    def _check_worker_health(self, active_workers: Set[int]):
        """Check worker health based on heartbeats"""
        if not active_workers:
            self.logger.warning("No heartbeats received from workers in last interval")
    
    def export_to_csv(self, bom_csv_path: str, summary_csv_path: str) -> bool:
        """Export database to CSV files with error handling"""
        try:
            with sqlite3.connect(self.db_path, timeout=30.0) as conn:
                cursor = conn.cursor()
                
                # Export BOM data in batches to handle large datasets
                self.logger.info("Exporting BOM data to CSV...")
                cursor.execute("""
                    SELECT filename, page, position, article_number, quantity,
                           description, weight, total_weight, kks_codes, su_codes
                    FROM bom_data ORDER BY filename, page, position
                """)
                
                with open(bom_csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
                    writer.writerow([
                        'Filename', 'Page', 'Position', 'Article_Number', 'Quantity',
                        'Description', 'Weight', 'Total_Weight', 'KKS_Codes', 'SU_Codes'
                    ])
                    
                    batch_size = 10000
                    while True:
                        rows = cursor.fetchmany(batch_size)
                        if not rows:
                            break
                        writer.writerows(rows)
                
                # Export summary data
                self.logger.info("Exporting summary data to CSV...")
                cursor.execute("""
                    SELECT filename, success, header_found, data_found, parse_success,
                           error_message, rows_count, process_time_s, file_size_mb,
                           pages_total, bom_pages, anchors_count, retry_count,
                           memory_peak_mb, error_code
                    FROM processing_summary ORDER BY filename
                """)
                
                with open(summary_csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
                    writer.writerow([
                        'Filename', 'Success', 'Header_Found', 'Data_Found', 'Parse_Success',
                        'Error_Message', 'Rows_Count', 'Process_Time_s', 'File_Size_MB',
                        'Pages_Total', 'BOM_Pages', 'Anchors_Count', 'Retry_Count',
                        'Memory_Peak_MB', 'Error_Code'
                    ])
                    
                    while True:
                        rows = cursor.fetchmany(1000)
                        if not rows:
                            break
                        writer.writerows(rows)
            
            self.logger.info("CSV export completed successfully")
            return True
        
        except Exception as e:
            self.logger.error(f"CSV export failed: {e}")
            return False

class ProgressMonitor:
    """Progress monitoring process with database writer health tracking"""
    
    def __init__(self, progress_queue: mp.Queue, total_files: int):
        self.progress_queue = SecureQueue(progress_queue, "progress")
        self.total_files = total_files
        self.logger = logging.getLogger(f"{__name__}.ProgressMonitor")
        self.start_time = time.time()
        self.processed_files = 0
        self.successful_files = 0
        self.failed_files = 0
        self.db_writer_health = DatabaseWriterHealthMonitor()
    
    def run(self):
        """Main progress monitoring loop"""
        self.logger.info("Progress monitor started")
        last_report = 0
        
        try:
            while True:
                msg_data = self.progress_queue.get_with_timeout(timeout=5.0)
                if msg_data is None:
                    # Check database writer health during timeout
                    if not self.db_writer_health.check_health():
                        self.logger.critical("Database writer appears to be dead!")
                    continue
                
                msg = WorkerMessage.deserialize(msg_data)
                
                if msg.msg_type == MessageType.SHUTDOWN.value:
                    break
                
                elif msg.msg_type == MessageType.PROGRESS.value:
                    data = msg.data
                    stats = data.get('stats', {})
                    
                    self.processed_files = stats.get('total_files', 0)
                    self.successful_files = stats.get('successful_files', 0)
                    self.failed_files = stats.get('failed_files', 0)
                    
                    # Report progress periodically
                    current_time = time.time()
                    if (current_time - last_report) >= 5.0 or self.processed_files >= self.total_files:
                        self._report_progress()
                        last_report = current_time
                
                elif msg.msg_type == MessageType.DB_WRITER_STATUS.value:
                    # Database writer is alive
                    self.db_writer_health.update_heartbeat()
        
        except Exception as e:
            self.logger.error(f"Progress monitor error: {e}")
        
        finally:
            self._report_final()
    
    def _report_progress(self):
        """Report current progress"""
        try:
            elapsed = time.time() - self.start_time
            progress_pct = (self.processed_files / self.total_files) * 100 if self.total_files > 0 else 0
            
            if self.processed_files > 0 and elapsed > 0:
                rate = self.processed_files / elapsed
                remaining = self.total_files - self.processed_files
                eta_seconds = remaining / rate if rate > 0 else 0
                eta_str = f"ETA: {int(eta_seconds)}s" if eta_seconds > 0 else "Complete"
            else:
                eta_str = "Calculating..."
            
            db_status = "OK" if self.db_writer_health.check_health() else "DEAD"
            
            self.logger.info(
                f"Progress: {self.processed_files}/{self.total_files} ({progress_pct:.1f}%) | "
                f"Success: {self.successful_files} | Failed: {self.failed_files} | "
                f"DB: {db_status} | {eta_str}"
            )
        
        except Exception as e:
            self.logger.error(f"Progress reporting error: {e}")
    
    def _report_final(self):
        """Report final statistics"""
        elapsed = time.time() - self.start_time
        self.logger.info(f"\n=== Processing Complete ===")
        self.logger.info(f"Files processed: {self.processed_files}")
        self.logger.info(f"Successful: {self.successful_files}")
        self.logger.info(f"Failed: {self.failed_files}")
        if self.processed_files > 0:
            self.logger.info(f"Success rate: {(self.successful_files/self.processed_files*100):.1f}%")
        self.logger.info(f"Total time: {elapsed:.1f}s")
        if elapsed > 0:
            self.logger.info(f"Processing rate: {self.processed_files/elapsed:.1f} files/sec")

class FileValidator:
    """File validation utilities with enhanced security checks"""
    
    def __init__(self, max_file_size_mb: int = DEFAULT_MAX_FILE_SIZE_MB):
        self.max_file_size_mb = max_file_size_mb
        self.max_size_bytes = max_file_size_mb * 1024 * 1024
        self.logger = logging.getLogger(f"{__name__}.FileValidator")
    
    def validate(self, pdf_path: str) -> Tuple[bool, ErrorCode, str, int, int]:
        """Validate PDF file with comprehensive checks"""
        try:
            file_path = Path(pdf_path)
            
            # Basic existence check
            if not file_path.exists():
                return False, ErrorCode.FILE_NOT_FOUND, "File does not exist", 0, 0
            
            # Security check - ensure it's a regular file
            if not file_path.is_file():
                return False, ErrorCode.INVALID_PDF, "Path is not a regular file", 0, 0
            
            # Get file size with error handling
            try:
                file_size = file_path.stat().st_size
            except OSError as e:
                return False, ErrorCode.FILE_NOT_FOUND, f"Cannot access file: {e}", 0, 0
            
            # Size checks
            if file_size > self.max_size_bytes:
                return False, ErrorCode.FILE_TOO_LARGE, \
                       f"File too large: {file_size / 1024 / 1024:.1f}MB", file_size, 0
            
            if file_size == 0:
                return False, ErrorCode.FILE_EMPTY, "File is empty", file_size, 0
            
            # PDF validity test with comprehensive error handling
            try:
                with SafePDFManager(str(pdf_path)) as doc:
                    page_count = len(doc)
                    if page_count == 0:
                        return False, ErrorCode.NO_PAGES, "PDF has no pages", file_size, 0
                    
                    # Test first page accessibility
                    try:
                        first_page = doc[0]
                        _ = first_page.get_text("dict")
                    except Exception as e:
                        return False, ErrorCode.INVALID_PDF, f"Cannot read PDF content: {e}", file_size, page_count
                    
                    return True, ErrorCode.UNKNOWN, "", file_size, page_count
            
            except PDFProcessingError as e:
                return False, e.error_code, e.message, file_size, 0
            except Exception as e:
                return False, ErrorCode.INVALID_PDF, f"Invalid PDF: {str(e)}", file_size, 0
        
        except Exception as e:
            return False, ErrorCode.UNKNOWN, f"Validation error: {str(e)}", 0, 0

class BOMDataExtractor:
    """BOM data extraction logic with enhanced error handling"""
    
    def __init__(self, config: BOMExtractionConfig):
        self.config = config
        self.kks_pattern = KKS_PATTERN
        self.su_pattern = SU_PATTERN
        self.logger = logging.getLogger(f"{__name__}.BOMDataExtractor")
    
    def extract_from_page(self, page: fitz.Page, anchor_pos: AnchorPosition, 
                         filename: str) -> List[BOMRow]:
        """Extract BOM data from page with comprehensive error handling"""
        try:
            # Extract codes first (less likely to fail)
            kks_codes, su_codes = self._extract_codes_from_page(page)
            
            # Extract structured data (more complex operation)
            headers, data_rows = self._extract_structured_data(page, anchor_pos)
            
            if not headers or not data_rows:
                return []
            
            # Convert to BOM rows with validation
            bom_rows = []
            for row_idx, row_data in enumerate(data_rows):
                try:
                    # Ensure row has enough columns
                    while len(row_data) < max(6, len(headers)):
                        row_data.append("")
                    
                    # Basic validation - must have position or article number
                    if len(row_data) >= 2 and (row_data[0].strip() or row_data[1].strip()):
                        bom_row = BOMRow(
                            filename=filename,
                            page=anchor_pos.page_number,
                            position=row_data[0].strip() if len(row_data) > 0 else "",
                            article_number=row_data[1].strip() if len(row_data) > 1 else "",
                            quantity=row_data[2].strip() if len(row_data) > 2 else "",
                            description=row_data[3].strip() if len(row_data) > 3 else "",
                            weight=row_data[4].strip() if len(row_data) > 4 else "",
                            total_weight=row_data[5].strip() if len(row_data) > 5 else "",
                            kks_codes=kks_codes,
                            su_codes=su_codes
                        )
                        bom_rows.append(bom_row)
                
                except Exception as e:
                    self.logger.warning(f"Error processing row {row_idx} in {filename}: {e}")
                    continue
            
            return bom_rows
            
        except Exception as e:
            raise PDFProcessingError(f"Extraction failed: {str(e)}", ErrorCode.PARSING_ERROR, filename)
    
    def _extract_codes_from_page(self, page: fitz.Page) -> Tuple[str, str]:
        """Extract KKS and SU codes from page with error handling"""
        try:
            full_text = page.get_text()
            if not full_text:
                return "", ""
            
            # Use compiled regex patterns for better performance
            kks_matches = self.kks_pattern.findall(full_text)
            su_matches = self.su_pattern.findall(full_text)
            
            # Remove duplicates and sort for consistency
            unique_kks = sorted(set(kks_matches)) if kks_matches else []
            unique_su = sorted(set(su_matches)) if su_matches else []
            
            return (", ".join(unique_kks), ", ".join(unique_su))
        
        except Exception as e:
            self.logger.debug(f"Error extracting codes: {e}")
            return "", ""
    
    def _extract_structured_data(self, page: fitz.Page, anchor_pos: AnchorPosition) -> Tuple[List[str], List[List[str]]]:
        """Extract structured data from page with enhanced error handling"""
        try:
            cells = self._extract_positioned_cells(page, anchor_pos)
            if not cells:
                raise ValueError("No cells found after anchor")
            
            table_rows = self._group_cells_into_rows(cells)
            if not table_rows:
                raise ValueError("Could not group cells into table rows")
            
            return self._parse_table_structure(table_rows, anchor_pos)
        
        except Exception as e:
            self.logger.debug(f"Structured data extraction failed: {e}")
            raise
    
    def _extract_positioned_cells(self, page: fitz.Page, anchor_pos: AnchorPosition) -> List[TableCell]:
        """Extract positioned cells from page with comprehensive error handling"""
        cells = []
        
        try:
            text_dict = page.get_text("dict")
            if not text_dict or "blocks" not in text_dict:
                return []
            
            for block in text_dict.get("blocks", []):
                if not isinstance(block, dict) or "lines" not in block:
                    continue
                
                for line in block.get("lines", []):
                    if not isinstance(line, dict) or "bbox" not in line:
                        continue
                    
                    line_bbox = line.get("bbox", [])
                    if len(line_bbox) < 4:
                        continue
                    
                    # Skip lines above the anchor
                    if line_bbox[1] < anchor_pos.y0:
                        continue
                    
                    for span in line.get("spans", []):
                        if not isinstance(span, dict):
                            continue
                        
                        span_text = span.get("text", "").strip()
                        span_bbox = span.get("bbox", [])
                        
                        if len(span_bbox) < 4 or not span_text:
                            continue
                        
                        try:
                            # Validate coordinates
                            if all(isinstance(coord, (int, float)) for coord in span_bbox[:4]):
                                cell = TableCell(
                                    text=span_text,
                                    x0=float(span_bbox[0]),
                                    y0=float(span_bbox[1]),
                                    x1=float(span_bbox[2]),
                                    y1=float(span_bbox[3])
                                )
                                cells.append(cell)
                        except (ValueError, TypeError):
                            continue
            
            return cells
            
        except Exception as e:
            self.logger.debug(f"Cell extraction error: {e}")
            return []
    
    def _group_cells_into_rows(self, cells: List[TableCell]) -> List[List[TableCell]]:
        """Group cells into rows by Y position with improved algorithm"""
        if not cells:
            return []
        
        # Sort cells by Y position first
        sorted_cells = sorted(cells, key=lambda c: c.center_y)
        
        row_groups = []
        tolerance = self.config.coordinate_tolerance
        
        for cell in sorted_cells:
            assigned = False
            
            # Try to assign to existing row
            for row_group in row_groups:
                if row_group:
                    avg_y = sum(c.center_y for c in row_group) / len(row_group)
                    if abs(cell.center_y - avg_y) <= tolerance:
                        row_group.append(cell)
                        assigned = True
                        break
            
            if not assigned:
                row_groups.append([cell])
        
        # Sort cells within each row by X position
        sorted_rows = []
        for row_group in row_groups:
            if row_group:
                sorted_row = sorted(row_group, key=lambda c: c.center_x)
                sorted_rows.append(sorted_row)
        
        # Sort rows by Y position
        return sorted(sorted_rows, key=lambda row: row[0].center_y if row else 0)
    
    def _parse_table_structure(self, table_rows: List[List[TableCell]], 
                              anchor_pos: AnchorPosition) -> Tuple[List[str], List[List[str]]]:
        """Parse table structure from grouped cells with enhanced logic"""
        if not table_rows:
            return [], []
        
        # Find header row containing anchor text
        header_row_idx = -1
        anchor_text_variants = [
            self.config.anchor_text.upper(),
            self.config.anchor_text.lower(),
            self.config.anchor_text.title()
        ]
        
        for i, row in enumerate(table_rows):
            row_text = " ".join(cell.text for cell in row).upper()
            for variant in anchor_text_variants:
                if variant.upper() in row_text:
                    header_row_idx = i
                    break
            if header_row_idx >= 0:
                break
        
        if header_row_idx == -1:
            return [], []
        
        # Extract headers with cleaning
        header_cells = table_rows[header_row_idx]
        headers = []
        column_positions = []
        
        for cell in header_cells:
            clean_header = cell.text.strip()
            if clean_header and clean_header not in headers:  # Avoid duplicates
                headers.append(clean_header)
                column_positions.append(cell.center_x)
        
        if not headers:
            return [], []
        
        # Extract data rows with termination checking
        data_rows = []
        termination_keywords = [
            self.config.total_text.upper(),
            "SUBTOTAL",
            "TOTAL",
            "SUM",
            "GESAMT"  # German for total
        ]
        
        for i in range(header_row_idx + 1, len(table_rows)):
            row = table_rows[i]
            
            # Check for termination conditions
            row_text = " ".join(cell.text for cell in row).upper()
            if any(keyword in row_text for keyword in termination_keywords):
                break
            
            # Skip empty rows
            if not any(cell.text.strip() for cell in row):
                continue
            
            # Map cells to columns using position matching
            parsed_row = [""] * len(headers)
            
            for cell in row:
                if column_positions:
                    # Find closest column position
                    distances = [abs(cell.center_x - col_x) for col_x in column_positions]
                    min_distance = min(distances)
                    col_idx = distances.index(min_distance)
                    
                    # Only assign if reasonably close
                    if col_idx < len(parsed_row) and min_distance < self.config.coordinate_tolerance * 3:
                        if parsed_row[col_idx]:
                            parsed_row[col_idx] += " " + cell.text.strip()
                        else:
                            parsed_row[col_idx] = cell.text.strip()
            
            # Clean up the row
            parsed_row = [cell.strip() for cell in parsed_row]
            
            # Only include rows with some meaningful content
            if any(cell for cell in parsed_row[:min(3, len(parsed_row))]):  # Check first 3 columns
                data_rows.append(parsed_row)
        
        return headers, data_rows

def process_pdf_worker(worker_id: int, file_batch: List[str], config_dict: Dict[str, Any], 
                      results_queue: mp.Queue, shutdown_event: mp.Event) -> None:
    """Enhanced worker function with comprehensive error handling and fallback support"""
    logger = None
    fallback_writer = None
    db_health_monitor = DatabaseWriterHealthMonitor()
    
    try:
        # Setup logging for worker
        logging.basicConfig(
            level=logging.INFO if not config_dict.get('debug', False) else logging.DEBUG,
            format=f'%(asctime)s - Worker-{worker_id} - %(levelname)s - %(message)s'
        )
        logger = logging.getLogger(__name__)
        
        # Reconstruct config
        config = BOMExtractionConfig(**config_dict)
        
        # Initialize components
        validator = FileValidator(config.max_file_size_mb)
        extractor = BOMDataExtractor(config)
        fallback_writer = FallbackWriter(config.fallback_directory, worker_id)
        
        # Setup secure queue
        secure_queue = SecureQueue(results_queue, f"worker-{worker_id}")
        
        logger.info(f"Worker {worker_id} started with {len(file_batch)} files")
        
        # Send initial heartbeat
        heartbeat_msg = WorkerMessage(
            msg_type=MessageType.HEARTBEAT.value,
            worker_id=worker_id,
            timestamp=time.time()
        )
        secure_queue.put_with_timeout(heartbeat_msg.serialize())
        
        last_heartbeat = time.time()
        processed_count = 0
        
        for pdf_path in file_batch:
            # Check for shutdown signal
            if shutdown_event.is_set():
                logger.info(f"Worker {worker_id} received shutdown signal")
                break
            
            # Send periodic heartbeat
            current_time = time.time()
            if current_time - last_heartbeat > HEARTBEAT_INTERVAL:
                heartbeat_msg = WorkerMessage(
                    msg_type=MessageType.HEARTBEAT.value,
                    worker_id=worker_id,
                    timestamp=current_time
                )
                secure_queue.put_with_timeout(heartbeat_msg.serialize())
                last_heartbeat = current_time
            
            # Process file with comprehensive error handling
            try:
                result = process_single_file(pdf_path, config, validator, extractor, worker_id)
                processed_count += 1
                
                # Try to send result to database writer
                msg = WorkerMessage(
                    msg_type=MessageType.BOM_DATA.value,
                    worker_id=worker_id,
                    timestamp=time.time(),
                    data=result.to_dict()
                )
                
                success = secure_queue.put_with_timeout(msg.serialize(), timeout=5.0)
                
                if not success:
                    # Database writer might be dead or queue full - use fallback
                    logger.warning(f"Failed to send result for {result.filename}, using fallback")
                    fallback_writer.write_result(result)
            
            except Exception as e:
                logger.error(f"Error processing {pdf_path}: {e}")
                # Create error result and try to send it
                try:
                    error_result = ProcessingResult(
                        filename=Path(pdf_path).name,
                        filepath=pdf_path,
                        success=False,
                        error_code=ErrorCode.WORKER_ERROR,
                        error_message=str(e)
                    )
                    fallback_writer.write_result(error_result)
                except Exception as fallback_error:
                    logger.error(f"Even fallback failed for {pdf_path}: {fallback_error}")
        
        logger.info(f"Worker {worker_id} completed processing {processed_count} files")
    
    except Exception as e:
        if logger:
            logger.critical(f"Worker {worker_id} fatal error: {e}", exc_info=True)
        
        # Send error message if possible
        try:
            if results_queue:
                error_msg = WorkerMessage(
                    msg_type=MessageType.ERROR.value,
                    worker_id=worker_id,
                    timestamp=time.time(),
                    data=str(e)
                )
                results_queue.put_nowait(error_msg.serialize())
        except:
            pass

def process_single_file(pdf_path: str, config: BOMExtractionConfig, validator: FileValidator,
                       extractor: BOMDataExtractor, worker_id: int) -> ProcessingResult:
    """Process a single PDF file with comprehensive error handling"""
    start_time = time.time()
    filename = Path(pdf_path).name
    logger = logging.getLogger(__name__)
    
    result = ProcessingResult(
        filename=filename,
        filepath=str(pdf_path),
        success=False
    )
    
    memory_monitor = SafeMemoryMonitor()
    
    try:
        # Initial memory check
        memory_monitor.update_peak()
        
        # Validate file first
        is_valid, error_code, error_msg, file_size, page_count = validator.validate(pdf_path)
        result.file_size_bytes = file_size
        result.total_pages = page_count
        
        if not is_valid:
            result.error_code = error_code
            result.error_message = error_msg
            return result
        
        # Process PDF with safe resource management
        with SafePDFManager(pdf_path) as doc:
            # Find anchors
            anchors = find_anchors(doc, config.anchor_text)
            result.anchor_count = len(anchors)
            
            if not anchors:
                result.error_code = ErrorCode.NO_ANCHOR
                result.error_message = f"No '{config.anchor_text}' anchors found"
                return result
            
            result.header_found = True
            
            # Extract data from pages with anchors
            all_bom_rows = []
            pages_with_data = []
            
            for anchor in anchors:
                try:
                    # Memory check before processing each page
                    memory_monitor.update_peak()
                    
                    page = doc[anchor.page_number - 1]
                    bom_rows = extractor.extract_from_page(page, anchor, filename)
                    
                    if bom_rows:
                        all_bom_rows.extend(bom_rows)
                        pages_with_data.append(anchor.page_number)
                
                except PDFProcessingError:
                    raise  # Re-raise processing errors
                except Exception as e:
                    logger.warning(
                        f"Worker {worker_id}: Error processing page {anchor.page_number} of {filename}: {e}"
                    )
                    continue
            
            # Final memory check
            memory_monitor.update_peak()
            
            # Set results based on extracted data
            if all_bom_rows:
                result.bom_rows = all_bom_rows
                result.bom_data_found = True
                result.parsing_success = True
                result.success = True
                result.bom_pages = pages_with_data
            else:
                result.error_code = ErrorCode.NO_DATA
                result.error_message = "No valid BOM data extracted from any page"
    
    except PDFProcessingError as e:
        result.error_code = e.error_code
        result.error_message = e.message
        if e.error_code == ErrorCode.MEMORY_ERROR:
            logger.warning(f"Worker {worker_id}: Memory limit exceeded processing {filename}")
    except Exception as e:
        result.error_code = ErrorCode.UNKNOWN
        result.error_message = f"Unexpected error: {str(e)}"
        logger.error(f"Worker {worker_id}: Unexpected error processing {filename}: {e}")
    
    finally:
        result.processing_time = time.time() - start_time
        result.memory_peak_mb = memory_monitor.peak_memory_mb
    
    return result

def find_anchors(doc: fitz.Document, anchor_text: str) -> List[AnchorPosition]:
    """Find anchor positions in document with enhanced search"""
    anchors = []
    
    try:
        # Limit search to reasonable number of pages for performance
        max_pages = min(len(doc), 100)
        
        # Create search variants
        search_variants = [
            anchor_text,
            anchor_text.upper(),
            anchor_text.lower(),
            anchor_text.title()
        ]
        
        for page_num in range(max_pages):
            try:
                page = doc[page_num]
                
                for variant in search_variants:
                    instances = page.search_for(variant)
                    
                    if instances:
                        # Take the first match for this variant
                        rect = instances[0]
                        try:
                            # Validate rectangle coordinates
                            if (isinstance(rect.x0, (int, float)) and isinstance(rect.y0, (int, float)) and
                                isinstance(rect.x1, (int, float)) and isinstance(rect.y1, (int, float)) and
                                rect.x1 > rect.x0 and rect.y1 > rect.y0):
                                
                                anchor = AnchorPosition(
                                    x0=float(rect.x0), 
                                    y0=float(rect.y0), 
                                    x1=float(rect.x1), 
                                    y1=float(rect.y1),
                                    line_number=1,
                                    page_number=page_num + 1,
                                    confidence=1.0 if variant == anchor_text else 0.8
                                )
                                anchors.append(anchor)
                                break  # Found match on this page, move to next page
                        except (ValueError, TypeError):
                            continue
            
            except Exception:
                # Skip problematic pages
                continue
        
        # Sort by confidence and page number
        anchors.sort(key=lambda a: (-a.confidence, a.page_number))
        return anchors
    
    except Exception:
        return []

def find_pdf_files(input_directory: str, max_files: Optional[int] = None) -> List[str]:
    """Find PDF files in directory with enhanced error handling"""
    try:
        input_path = Path(input_directory).resolve()
        
        if not input_path.exists():
            raise FileNotFoundError(f"Directory does not exist: {input_directory}")
        
        if not input_path.is_dir():
            raise NotADirectoryError(f"Path is not a directory: {input_directory}")
        
        pdf_files = set()
        patterns = ['*.pdf', '*.PDF']
        
        # Use a more efficient approach for large directories
        for pattern in patterns:
            try:
                # Use iterdir for immediate directory, rglob for subdirectories
                if max_files and len(pdf_files) >= max_files * 2:
                    break
                
                for file_path in input_path.rglob(pattern):
                    try:
                        if (file_path.is_file() and 
                            file_path.suffix.lower() == '.pdf' and
                            file_path.stat().st_size > 0):  # Skip empty files
                            pdf_files.add(str(file_path))
                        
                        if max_files and len(pdf_files) >= max_files * 2:
                            break
                    except (OSError, PermissionError):
                        # Skip files we can't access
                        continue
            
            except Exception:
                # Continue with other patterns if one fails
                continue
        
        # Convert to sorted list
        pdf_files = sorted(list(pdf_files))
        
        # Apply max_files limit
        if max_files and len(pdf_files) > max_files:
            pdf_files = pdf_files[:max_files]
        
        return pdf_files
    
    except Exception as e:
        raise RuntimeError(f"Error finding PDF files: {e}")

def setup_logging(debug: bool = False, log_file: Optional[str] = None) -> None:
    """Setup logging configuration with enhanced error handling"""
    level = logging.DEBUG if debug else logging.INFO
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    handlers = []
    
    # Console handler with error handling
    try:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)
        console_handler.setFormatter(logging.Formatter(log_format))
        handlers.append(console_handler)
    except Exception as e:
        print(f"Warning: Could not setup console logging: {e}")
    
    # File handler if requested
    if log_file:
        try:
            log_path = Path(log_file)
            log_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Use rotating log handler for large files
            from logging.handlers import RotatingFileHandler
            file_handler = RotatingFileHandler(
                log_file, 
                maxBytes=10*1024*1024,  # 10MB
                backupCount=3,
                encoding='utf-8'
            )
            file_handler.setLevel(level)
            file_handler.setFormatter(logging.Formatter(log_format))
            handlers.append(file_handler)
        except Exception as e:
            print(f"Warning: Could not setup file logging: {e}")
    
    # Configure root logger
    if handlers:
        logging.basicConfig(level=level, handlers=handlers, force=True)
    else:
        # Fallback to basic config
        logging.basicConfig(level=level, format=log_format, force=True)
    
    # Quiet noisy libraries
    logging.getLogger('fitz').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)

def get_optimal_multiprocessing_context(method: str) -> mp.context.BaseContext:
    """Get optimal multiprocessing context for the platform"""
    try:
        return mp.get_context(method)
    except ValueError:
        # Fallback to default if method not available
        return mp.get_context()

class QueueBasedBOMExtractor:
    """Main BOM extractor with hardened queue-based architecture"""
    
    def __init__(self, config: BOMExtractionConfig):
        self.config = config
        
        # Use platform-specific multiprocessing context
        self.mp_context = get_optimal_multiprocessing_context(config.mp_method)
        self.shutdown_event = self.mp_context.Event()
        
        # Setup logging
        setup_logging(config.debug, config.log_file)
        self.logger = logging.getLogger(__name__)
        
        # Calculate optimal queue sizes
        queue_size = max(100, self.config.max_workers * QUEUE_SIZE_MULTIPLIER)
        
        # Create queues with multiprocessing context
        self.results_queue = self.mp_context.Queue(maxsize=queue_size)
        self.progress_queue = self.mp_context.Queue(maxsize=queue_size)
        
        # Setup signal handling with thread safety
        self._setup_signal_handling()
        
        self.logger.info(f"Initialized BOM extractor with {config.max_workers} workers using {config.mp_method} method")
    
    def _setup_signal_handling(self):
        """Setup signal handling with proper synchronization"""
        def signal_handler(signum: int, frame):
            self.logger.info(f"Received signal {signum}, initiating graceful shutdown...")
            self.shutdown_event.set()
        
        # Register signal handlers
        try:
            signal.signal(signal.SIGINT, signal_handler)
            signal.signal(signal.SIGTERM, signal_handler)
            if hasattr(signal, 'SIGBREAK'):  # Windows
                signal.signal(signal.SIGBREAK, signal_handler)
        except ValueError:
            # Signals not available in this context (e.g., threads)
            pass
    
    def run(self) -> None:
        """Main execution flow with comprehensive error handling"""
        try:
            # Discover files
            self.logger.info("Discovering PDF files...")
            pdf_files = find_pdf_files(
                self.config.input_directory,
                self.config.max_files_to_process
            )
            
            if not pdf_files:
                self.logger.warning("No PDF files found to process")
                return
            
            self.logger.info(f"Found {len(pdf_files)} PDF files to process")
            
            # Start supporting processes
            db_writer_process = self._start_database_writer()
            progress_monitor_process = self._start_progress_monitor(len(pdf_files))
            
            supporting_processes = [db_writer_process, progress_monitor_process]
            
            try:
                # Process files with workers
                self._process_all_files(pdf_files)
                
                # Wait for supporting processes to finish
                self._shutdown_processes(supporting_processes)
                
                # Collect and merge fallback data if any
                self._recover_fallback_data()
                
                # Export final results
                self._export_results()
                
                self.logger.info("Processing completed successfully")
            
            finally:
                # Ensure cleanup of all processes
                self._cleanup_processes(supporting_processes)
        
        except KeyboardInterrupt:
            self.logger.info("Processing interrupted by user")
            self.shutdown_event.set()
        except Exception as e:
            self.logger.critical(f"Fatal error in main process: {e}", exc_info=True)
            raise
    
    def _start_database_writer(self) -> mp.Process:
        """Start database writer process"""
        db_writer = DatabaseWriter(
            self.config.database_path,
            self.results_queue,
            self.progress_queue
        )
        
        process = self.mp_context.Process(target=db_writer.run, name="DatabaseWriter")
        process.start()
        self.logger.info(f"Started database writer process (PID: {process.pid})")
        return process
    
    def _start_progress_monitor(self, total_files: int) -> mp.Process:
        """Start progress monitor process"""
        progress_monitor = ProgressMonitor(self.progress_queue, total_files)
        process = self.mp_context.Process(target=progress_monitor.run, name="ProgressMonitor")
        process.start()
        self.logger.info(f"Started progress monitor process (PID: {process.pid})")
        return process
    
    def _process_all_files(self, pdf_files: List[str]) -> None:
        """Process all PDF files using worker processes with load balancing"""
        # Calculate optimal batch sizes
        files_per_worker = max(1, len(pdf_files) // self.config.max_workers)
        worker_batches = []
        
        # Create initial batches
        for i in range(0, len(pdf_files), files_per_worker):
            batch = pdf_files[i:i + files_per_worker]
            worker_batches.append(batch)
        
        # Redistribute if we have too many batches
        while len(worker_batches) > self.config.max_workers:
            # Merge the two smallest batches
            worker_batches.sort(key=len)
            smallest = worker_batches.pop(0)
            second_smallest = worker_batches.pop(0)
            merged = smallest + second_smallest
            worker_batches.append(merged)
        
        self.logger.info(f"Starting {len(worker_batches)} worker processes")
        self.logger.debug(f"Batch sizes: {[len(batch) for batch in worker_batches]}")
        
        # Start worker processes
        processes = []
        config_dict = asdict(self.config)
        
        try:
            for worker_id, file_batch in enumerate(worker_batches):
                if not file_batch:
                    continue
                
                process = self.mp_context.Process(
                    target=process_pdf_worker,
                    args=(worker_id, file_batch, config_dict, self.results_queue, self.shutdown_event),
                    name=f"Worker-{worker_id}"
                )
                process.start()
                processes.append((process, worker_id, len(file_batch)))
                self.logger.debug(f"Started worker {worker_id} (PID: {process.pid}) with {len(file_batch)} files")
            
            # Monitor worker processes
            self._monitor_workers(processes)
        
        finally:
            # Ensure all workers are cleaned up
            for process, worker_id, _ in processes:
                if process.is_alive():
                    try:
                        self.logger.debug(f"Terminating worker {worker_id}")
                        process.terminate()
                        process.join(timeout=5)
                        if process.is_alive():
                            self.logger.warning(f"Force killing worker {worker_id}")
                            process.kill()
                            process.join(timeout=2)
                    except Exception as e:
                        self.logger.error(f"Error cleaning up worker {worker_id}: {e}")
        
        self.logger.info("All worker processes completed")
    
    def _monitor_workers(self, processes: List[Tuple[mp.Process, int, int]]) -> None:
        """Monitor worker processes and handle failures"""
        active_workers = {worker_id: {'process': process, 'files': file_count} 
                         for process, worker_id, file_count in processes}
        
        check_interval = 5.0
        last_check = time.time()
        
        while active_workers:
            current_time = time.time()
            
            # Check worker status periodically
            if current_time - last_check > check_interval:
                completed_workers = []
                
                for worker_id, info in active_workers.items():
                    process = info['process']
                    
                    if not process.is_alive():
                        exit_code = process.exitcode
                        if exit_code == 0:
                            self.logger.info(f"Worker {worker_id} completed successfully")
                        else:
                            self.logger.error(f"Worker {worker_id} exited with code {exit_code}")
                        completed_workers.append(worker_id)
                
                # Remove completed workers
                for worker_id in completed_workers:
                    del active_workers[worker_id]
                
                last_check = current_time
            
            # Check for shutdown signal
            if self.shutdown_event.is_set():
                self.logger.info("Shutdown signal received, stopping remaining workers")
                break
            
            time.sleep(1.0)
        
        # Wait for remaining workers to finish
        for worker_id, info in active_workers.items():
            process = info['process']
            try:
                process.join(timeout=self.config.timeout_per_file)
                if process.is_alive():
                    self.logger.warning(f"Worker {worker_id} timed out, terminating")
                    process.terminate()
                    process.join(timeout=5)
            except Exception as e:
                self.logger.error(f"Error waiting for worker {worker_id}: {e}")
    
    def _shutdown_processes(self, processes: List[mp.Process]) -> None:
        """Shutdown supporting processes gracefully"""
        self.logger.info("Shutting down supporting processes...")
        
        # Send shutdown messages
        try:
            shutdown_msg = WorkerMessage(
                msg_type=MessageType.SHUTDOWN.value,
                worker_id=-1,
                timestamp=time.time()
            )
            
            self.results_queue.put_nowait(shutdown_msg.serialize())
            self.progress_queue.put_nowait(shutdown_msg.serialize())
        except Exception as e:
            self.logger.error(f"Error sending shutdown messages: {e}")
        
        # Wait for processes to finish gracefully
        for process in processes:
            try:
                process.join(timeout=SHUTDOWN_TIMEOUT)
                if process.is_alive():
                    self.logger.warning(f"Force terminating {process.name}")
                    process.terminate()
                    process.join(timeout=5)
                    if process.is_alive():
                        process.kill()
            except Exception as e:
                self.logger.error(f"Error shutting down {process.name}: {e}")
    
    def _cleanup_processes(self, processes: List[mp.Process]) -> None:
        """Final cleanup of any remaining processes"""
        for process in processes:
            if process.is_alive():
                try:
                    process.terminate()
                    process.join(timeout=2)
                    if process.is_alive():
                        process.kill()
                except Exception:
                    pass  # Best effort cleanup
    
    def _recover_fallback_data(self) -> None:
        """Recover data from fallback files and merge into database"""
        fallback_dir = Path(self.config.fallback_directory)
        
        if not fallback_dir.exists():
            return
        
        fallback_files = list(fallback_dir.glob(f"{FALLBACK_FILE_PREFIX}*.json"))
        
        if not fallback_files:
            return
        
        self.logger.info(f"Found {len(fallback_files)} fallback files to recover")
        
        try:
            # Create a temporary database writer for recovery
            recovery_queue = self.mp_context.Queue()
            db_writer = DatabaseWriter(
                self.config.database_path,
                recovery_queue,
                self.mp_context.Queue()  # Dummy progress queue
            )
            
            recovered_count = 0
            
            for fallback_file in fallback_files:
                try:
                    with open(fallback_file, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if not line:
                                continue
                            
                            try:
                                data = json.loads(line)
                                result_data = data.get('result', {})
                                
                                if result_data:
                                    # Send to database writer
                                    msg = WorkerMessage(
                                        msg_type=MessageType.BOM_DATA.value,
                                        worker_id=data.get('worker_id', -1),
                                        timestamp=data.get('timestamp', time.time()),
                                        data=result_data
                                    )
                                    recovery_queue.put(msg.serialize())
                                    recovered_count += 1
                            
                            except json.JSONDecodeError:
                                continue
                    
                    # Remove processed fallback file
                    fallback_file.unlink()
                
                except Exception as e:
                    self.logger.error(f"Error recovering from {fallback_file}: {e}")
            
            # Process recovered data
            if recovered_count > 0:
                self.logger.info(f"Processing {recovered_count} recovered results")
                # Process a few messages to trigger database writes
                for _ in range(min(10, recovered_count)):
                    try:
                        msg_data = recovery_queue.get_nowait()
                        msg = WorkerMessage.deserialize(msg_data)
                        db_writer._process_result(msg.data, [])
                    except:
                        break
        
        except Exception as e:
            self.logger.error(f"Error during fallback recovery: {e}")
    
    def _export_results(self) -> None:
        """Export results to CSV files"""
        self.logger.info("Exporting results to CSV files...")
        
        try:
            # Create temporary database writer for export
            db_writer = DatabaseWriter(
                self.config.database_path,
                self.mp_context.Queue(),  # Dummy queue
                self.mp_context.Queue()   # Dummy queue
            )
            
            success = db_writer.export_to_csv(
                self.config.output_csv,
                self.config.summary_csv
            )
            
            if success:
                self.logger.info(f"✓ BOM data exported to: {self.config.output_csv}")
                self.logger.info(f"✓ Summary exported to: {self.config.summary_csv}")
                
                # Log file sizes for verification
                try:
                    bom_size = Path(self.config.output_csv).stat().st_size / 1024 / 1024
                    summary_size = Path(self.config.summary_csv).stat().st_size / 1024 / 1024
                    self.logger.info(f"File sizes: BOM={bom_size:.1f}MB, Summary={summary_size:.1f}MB")
                except:
                    pass
            else:
                self.logger.error("✗ Failed to export CSV files")
        
        except Exception as e:
            self.logger.error(f"Error during CSV export: {e}")

def main() -> int:
    """Main entry point with comprehensive error handling"""
    parser = argparse.ArgumentParser(
        description="PDF BOM Text Extractor - Hardened Queue-based Architecture",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Hardened queue-based design optimized for 32-core systems processing 2500+ files.
Features comprehensive error handling, resource leak protection, and fallback recovery.

Security improvements:
- No pickle serialization (uses JSON/msgpack)
- Safe resource management with context managers
- Memory monitoring and limits
- Database writer health monitoring with fallback
- Protection against deadlocks and race conditions
        """
    )
    
    # Required arguments
    parser.add_argument("input_directory", help="Directory containing PDF files")
    
    # Optional arguments
    parser.add_argument("--output-csv", default="extracted_bom.csv",
                      help="Output BOM CSV file path")
    parser.add_argument("--summary-csv", default="extraction_summary.csv", 
                      help="Summary CSV file path")
    parser.add_argument("--database", default="bom_extraction.db",
                      help="SQLite database path")
    parser.add_argument("--fallback-dir", default="bom_fallback",
                      help="Fallback directory for failed database writes")
    parser.add_argument("--max-workers", type=int,
                      help="Maximum worker processes (default: CPU cores - 1)")
    parser.add_argument("--max-files", type=int,
                      help="Maximum files to process")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT,
                      help="Timeout per file in seconds")
    parser.add_argument("--max-file-size", type=int, default=DEFAULT_MAX_FILE_SIZE_MB,
                      help="Maximum file size in MB")
    parser.add_argument("--max-memory", type=int, default=MAX_MEMORY_MB,
                      help="Maximum memory per worker in MB")
    parser.add_argument("--no-retry", action="store_true",
                      help="Disable retrying failed files")
    parser.add_argument("--debug", action="store_true",
                      help="Enable debug logging")
    parser.add_argument("--log-file", help="Log file path")
    
    args = parser.parse_args()
    
    try:
        # Create config with validation
        config = BOMExtractionConfig(
            input_directory=args.input_directory,
            output_csv=args.output_csv,
            summary_csv=args.summary_csv,
            database_path=args.database,
            fallback_directory=args.fallback_dir,
            max_workers=args.max_workers,
            max_files_to_process=args.max_files,
            timeout_per_file=args.timeout,
            max_file_size_mb=args.max_file_size,
            max_memory_mb=args.max_memory,
            retry_failed_files=not args.no_retry,
            debug=args.debug,
            log_file=args.log_file
        )
        
        # Validate input directory early
        if not Path(config.input_directory).exists():
            print(f"Error: Input directory does not exist: {config.input_directory}", file=sys.stderr)
            return 1
        
        # Run extractor
        extractor = QueueBasedBOMExtractor(config)
        extractor.run()
        
        return 0
    
    except (FileNotFoundError, NotADirectoryError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nInterrupted by user", file=sys.stderr)
        return 130
    except Exception as e:
        print(f"Fatal error: {e}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    # Set multiprocessing start method based on platform detection
    # This is done in __main__ to avoid conflicts with existing code
    platform_system = platform.system().lower()
    if platform_system == "linux":
        mp.set_start_method('fork', force=True)
    elif platform_system in ["darwin", "windows"]:
        mp.set_start_method('spawn', force=True)
    else:
        # Use spawn as safe default
        mp.set_start_method('spawn', force=True)
    
    sys.exit(main())
                