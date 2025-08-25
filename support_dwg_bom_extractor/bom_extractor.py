#!/usr/bin/env python3
"""
PDF BOM (Bill of Materials) Text Extractor - Production-Ready Enhanced Architecture

Enhanced for 48-core systems processing 2500+ files:
- Intelligent CPU utilization with dynamic worker scaling
- Multi-database writer pool to eliminate bottlenecks
- Dynamic memory management with per-worker limits
- Production checkpoint/recovery system
- Comprehensive progress monitoring and health checks
- Enhanced error handling and resource management
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
from collections import defaultdict, Counter, deque
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
BATCH_SIZE = 1000  # Increased batch size for better performance
QUEUE_SIZE_MULTIPLIER = 4  # Increased for higher throughput
HEARTBEAT_INTERVAL = 10
SHUTDOWN_TIMEOUT = 30
DB_WRITER_HEARTBEAT_TIMEOUT = 20
MAX_MEMORY_MB = 1024  # Will be dynamically adjusted
FALLBACK_FILE_PREFIX = "bom_fallback_"
NUM_DB_WRITERS = 3  # Multiple database writers

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

class SystemResourceManager:
    """Intelligent resource management for high-performance systems"""
    
    def __init__(self, total_cores: Optional[int] = None, total_ram_gb: Optional[float] = None):
        self.total_cores = total_cores or mp.cpu_count()
        self.total_ram = self._detect_system_ram_gb() if total_ram_gb is None else total_ram_gb
        self.safe_ram = self.total_ram * 0.8  # 80% utilization
        self.os_overhead = 4.0  # 4GB for OS
        
        # Log system configuration
        logging.getLogger(__name__).info(
            f"System Resources: {self.total_cores} cores, {self.total_ram:.1f}GB RAM"
        )

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
                      db_writer_pool_queues: List[mp.Queue], shutdown_event: mp.Event) -> None:
    """Enhanced worker function with comprehensive error handling and fallback support"""
    logger = None
    fallback_writer = None
    memory_monitor = None
    
    try:
        # Setup logging for worker
        logging.basicConfig(
            level=logging.INFO if not config_dict.get('debug', False) else logging.DEBUG,
            format=f'%(asctime)s - Worker-{worker_id} - %(levelname)s - %(message)s'
        )
        logger = logging.getLogger(__name__)
        
        # Reconstruct config
        config = BOMExtractionConfig(**config_dict)
        
        # Initialize enhanced memory monitoring
        memory_manager = DynamicMemoryManager(32.0, config.max_workers)  # Assume 32GB system
        memory_limit = memory_manager.get_worker_memory_limit()
        memory_monitor = EnhancedMemoryMonitor(worker_id, memory_limit)
        
        # Initialize components
        validator = FileValidator(config.max_file_size_mb)
        extractor = BOMDataExtractor(config)
        fallback_writer = FallbackWriter(config.fallback_directory, worker_id)
        
        # Setup secure queue (load balance across DB writers)
        assigned_queue = db_writer_pool_queues[worker_id % len(db_writer_pool_queues)]
        secure_queue = SecureQueue(assigned_queue, f"worker-{worker_id}")
        
        logger.info(f"Worker {worker_id} started with {len(file_batch)} files (Memory limit: {memory_limit}MB)")
        
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
            
            # Update memory monitoring
            if memory_monitor:
                memory_monitor.update_peak()
            
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
                result = process_single_file(pdf_path, config, validator, extractor, worker_id, memory_monitor)
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
                        error_message=str(e),
                        memory_peak_mb=memory_monitor.peak_memory_mb if memory_monitor else 0.0
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
            if db_writer_pool_queues:
                error_msg = WorkerMessage(
                    msg_type=MessageType.ERROR.value,
                    worker_id=worker_id,
                    timestamp=time.time(),
                    data=str(e)
                )
                db_writer_pool_queues[0].put_nowait(error_msg.serialize())
        except:
            pass

def process_single_file(pdf_path: str, config: BOMExtractionConfig, validator: FileValidator,
                       extractor: BOMDataExtractor, worker_id: int, 
                       memory_monitor: Optional[EnhancedMemoryMonitor] = None) -> ProcessingResult:
    """Process a single PDF file with comprehensive error handling"""
    start_time = time.time()
    filename = Path(pdf_path).name
    logger = logging.getLogger(__name__)
    
    result = ProcessingResult(
        filename=filename,
        filepath=str(pdf_path),
        success=False
    )
    
    try:
        # Initial memory check
        if memory_monitor:
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
                    if memory_monitor:
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
            if memory_monitor:
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
        if memory_monitor:
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

class ProductionJobManager:
    """Production job manager with checkpoint/recovery system"""
    
    def __init__(self, config: BOMExtractionConfig):
        self.config = config
        self.checkpoint_manager = ProductionCheckpointManager(config.checkpoint_file)
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
        
    
    def _detect_system_ram_gb(self) -> float:
        """Detect total system RAM"""
        if PSUTIL_AVAILABLE:
            try:
                return psutil.virtual_memory().total / (1024**3)
            except:
                pass
        
        # Fallback estimates based on core count
        if self.total_cores >= 32:
            return 32.0  # Assume server-class machine
        elif self.total_cores >= 16:
            return 16.0
        else:
            return 8.0
    
    def get_optimal_worker_count(self) -> int:
        """Calculate optimal worker count for the system"""
        # Reserve cores: 3 DB writers, 1 progress monitor, 2 OS overhead  
        available_cores = max(1, self.total_cores - 6)
        
        # Calculate max workers by memory constraints
        available_worker_ram = (self.safe_ram - self.os_overhead) * 1024  # MB
        reserved_db_ram = 2 * 1024  # 2GB for database operations
        worker_ram = available_worker_ram - reserved_db_ram
        
        # Minimum 256MB per worker, but cap at reasonable limit
        max_workers_by_memory = int(worker_ram // 256)
        
        optimal_workers = min(available_cores, max_workers_by_memory, 44)  # Cap at 44 for stability
        
        logging.getLogger(__name__).info(
            f"Optimal configuration: {optimal_workers} workers "
            f"(cores: {available_cores}, memory limit: {max_workers_by_memory})"
        )
        
        return optimal_workers
    
    def get_worker_memory_limit_mb(self, num_workers: int) -> int:
        """Calculate safe memory limit per worker"""
        available_worker_ram = (self.safe_ram - self.os_overhead) * 1024  # MB
        reserved_db_ram = 2 * 1024  # 2GB for database operations
        worker_ram = available_worker_ram - reserved_db_ram
        
        per_worker_limit = max(256, int(worker_ram / num_workers))
        return min(per_worker_limit, 1024)  # Cap at 1GB per worker

class DynamicMemoryManager:
    """Dynamic memory allocation and monitoring"""
    
    def __init__(self, total_ram_gb: float, num_workers: int):
        self.total_ram = total_ram_gb
        self.num_workers = num_workers
        self.safe_utilization = 0.8
        self.os_overhead = 4.0  # GB
        self.db_writer_allocation = 2.0  # GB
        
        # Calculate per-worker allocation
        available_ram_gb = (self.total_ram * self.safe_utilization) - self.os_overhead - self.db_writer_allocation
        self.per_worker_limit_mb = max(256, int((available_ram_gb * 1024) / num_workers))
        
        logging.getLogger(__name__).info(
            f"Memory allocation: {self.per_worker_limit_mb}MB per worker "
            f"({self.per_worker_limit_mb * num_workers / 1024:.1f}GB total worker memory)"
        )
    
    def get_worker_memory_limit(self) -> int:
        """Get safe memory limit per worker"""
        return min(self.per_worker_limit_mb, 1024)  # Cap at 1GB

class EnhancedMemoryMonitor:
    """Enhanced memory monitoring with pressure detection"""
    
    def __init__(self, worker_id: int, memory_limit_mb: int):
        self.worker_id = worker_id
        self.memory_limit_mb = memory_limit_mb
        self.warning_threshold = memory_limit_mb * 0.8
        self.process = None
        self.peak_memory_mb = 0.0
        self.logger = logging.getLogger(f"{__name__}.MemoryMonitor.{worker_id}")
        
        if PSUTIL_AVAILABLE:
            try:
                self.process = psutil.Process()
                # Test access
                _ = self.process.memory_info()
            except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
                self.process = None
    
    def update_peak(self):
        """Update peak memory usage with pressure detection"""
        if self.process is None:
            return
        
        try:
            memory_info = self.process.memory_info()
            current_mb = memory_info.rss / 1024 / 1024
            self.peak_memory_mb = max(self.peak_memory_mb, current_mb)
            
            # Warning threshold
            if current_mb > self.warning_threshold:
                self.logger.warning(
                    f"Worker {self.worker_id} memory usage: {current_mb:.1f}MB "
                    f"(limit: {self.memory_limit_mb}MB)"
                )
            
            # Hard limit
            if current_mb > self.memory_limit_mb:
                raise PDFProcessingError(
                    f"Worker {self.worker_id} exceeded memory limit: "
                    f"{current_mb:.1f}MB > {self.memory_limit_mb}MB",
                    ErrorCode.MEMORY_ERROR
                )
                
        except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
            self.process = None
        except Exception:
            pass  # Ignore other monitoring errors
    
    def check_memory_pressure(self) -> bool:
        """Check if approaching memory limit"""
        if self.process is None:
            return False
            
        try:
            current_mb = self.process.memory_info().rss / 1024 / 1024
            return current_mb > self.warning_threshold
        except:
            return False

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

class ProductionCheckpointManager:
    """Production-ready checkpoint and recovery system"""
    
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
    
    def save_checkpoint(self):
        """Save current state to checkpoint file"""
        try:
            checkpoint_data = {
                'processed_files': list(self.processed_files),
                'failed_files': self.failed_files,
                'stats': self.processing_stats,
                'timestamp': datetime.now().isoformat(),
                'version': '2.0'
            }
            
            temp_file = self.checkpoint_file.with_suffix('.tmp')
            
            with self._file_lock():
                with open(temp_file, 'w', encoding='utf-8') as f:
                    json.dump(checkpoint_data, f, indent=2, ensure_ascii=False)
                
                temp_file.replace(self.checkpoint_file)
                
            self.processing_stats['last_checkpoint'] = time.time()
            
        except Exception as e:
            self.logger.error(f"Failed to save checkpoint: {e}")
    
    @contextmanager
    def _file_lock(self):
        """Simple file locking for checkpoint access"""
        lock_acquired = False
        try:
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
            self.failed_files.pop(file_path, None)
            self.processing_stats['successful_files'] += 1
            self.processing_stats['bom_rows_extracted'] += len(result.bom_rows)
        else:
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
            
            for file_path, error_info in files[:5]:
                report_lines.append(
                    f"  {Path(file_path).name}: {error_info.get('error_message', 'No message')}"
                )
            
            if len(files) > 5:
                report_lines.append(f"  ... and {len(files) - 5} more files")
            
            report_lines.append("")
        
        return "\n".join(report_lines)

@dataclass
class BOMExtractionConfig:
    """Enhanced configuration for BOM extraction"""
    input_directory: str
    output_csv: str = "extracted_bom.csv"
    summary_csv: str = "extraction_summary.csv"
    database_path: str = "bom_extraction.db"
    fallback_directory: str = "bom_fallback"
    checkpoint_file: str = "bom_processing_checkpoint.json"
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
    num_db_writers: int = NUM_DB_WRITERS
    
    def __post_init__(self):
        """Validate configuration and set optimal defaults"""
        # Initialize resource manager for intelligent configuration
        resource_manager = SystemResourceManager()
        
        # Set optimal worker count if not specified
        if self.max_workers is None:
            self.max_workers = resource_manager.get_optimal_worker_count()
        
        # Set dynamic memory limit per worker
        self.max_memory_mb = resource_manager.get_worker_memory_limit_mb(self.max_workers)
        
        # Detect optimal multiprocessing method
        self.mp_method = self._detect_mp_method()
        
        # Ensure directories exist
        Path(self.fallback_directory).mkdir(parents=True, exist_ok=True)
        
        # Log configuration
        logger = logging.getLogger(__name__)
        logger.info(f"Configuration: {self.max_workers} workers, {self.max_memory_mb}MB per worker")
        logger.info(f"Database writers: {self.num_db_writers}, MP method: {self.mp_method}")
    
    def _detect_mp_method(self) -> str:
        """Detect optimal multiprocessing start method for the platform"""
        system = platform.system().lower()
        if system == "linux":
            return "fork"
        elif system == "darwin":
            return "spawn"
        elif system == "windows":
            return "spawn"
        else:
            return "spawn"

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
                logging.getLogger(__name__).warning(f"Error closing PDF {self.pdf_path}: {e}")

class FallbackWriter:
    """Fallback writer for when database writer fails"""
    
    def __init__(self, fallback_dir: str, worker_id: int):
        self.fallback_dir = Path(fallback_dir)
        self.worker_id = worker_id
        self.fallback_file = self.fallback_dir / f"{FALLBACK_FILE_PREFIX}{worker_id}_{int(time.time())}.json"
        self.logger = logging.getLogger(f"{__name__}.FallbackWriter.{worker_id}")
        
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
                os.fsync(f.fileno())
        
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

class DatabaseWriterPool:
    """Multi-database writer pool to eliminate bottlenecks"""
    
    def __init__(self, db_path: str, num_writers: int = NUM_DB_WRITERS, 
                 results_queue_size: int = 2000, mp_context=None):
        self.db_path = db_path
        self.num_writers = num_writers
        self.mp_context = mp_context or mp
        
        # Create separate queues for each DB writer to avoid contention
        self.writer_queues = [
            self.mp_context.Queue(maxsize=results_queue_size // num_writers) 
            for _ in range(num_writers)
        ]
        self.writer_processes = []
        self.logger = logging.getLogger(f"{__name__}.DatabaseWriterPool")
        
        self._setup_optimized_database()
        
    def _setup_optimized_database(self):
        """Configure SQLite for high-concurrency workloads"""
        try:
            with sqlite3.connect(self.db_path, timeout=30.0) as conn:
                # Optimize for concurrent writes
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute("PRAGMA synchronous=NORMAL") 
                conn.execute("PRAGMA cache_size=-64000")  # 64MB cache
                conn.execute("PRAGMA temp_store=MEMORY")
                conn.execute("PRAGMA wal_autocheckpoint=10000")
                conn.execute("PRAGMA busy_timeout=30000")  # 30s timeout
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
    
    def start_writers(self, progress_queue: mp.Queue):
        """Start multiple database writer processes"""
        for writer_id in range(self.num_writers):
            writer = OptimizedDatabaseWriter(
                db_path=self.db_path,
                results_queue=self.writer_queues[writer_id],
                progress_queue=progress_queue,
                writer_id=writer_id,
                batch_size=BATCH_SIZE
            )
            
            process = self.mp_context.Process(
                target=writer.run, 
                name=f"DBWriter-{writer_id}"
            )
            process.start()
            self.writer_processes.append(process)
            self.logger.info(f"Started database writer {writer_id} (PID: {process.pid})")
    
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
            try:
                process.join(timeout=SHUTDOWN_TIMEOUT)
                if process.is_alive():
                    process.terminate()
                    process.join(timeout=5)
                    if process.is_alive():
                        process.kill()
            except Exception as e:
                self.logger.error(f"Error shutting down database writer: {e}")
    
    def export_to_csv(self, bom_csv_path: str, summary_csv_path: str) -> bool:
        """Export database to CSV files with error handling"""
        try:
            with sqlite3.connect(self.db_path, timeout=30.0) as conn:
                cursor = conn.cursor()
                
                # Export BOM data in batches
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

class OptimizedDatabaseWriter:
    """Optimized database writer with enhanced performance"""
    
    def __init__(self, db_path: str, results_queue: mp.Queue, progress_queue: mp.Queue, 
                 writer_id: int, batch_size: int = BATCH_SIZE):
        self.db_path = db_path
        self.results_queue = SecureQueue(results_queue, f"results-{writer_id}")
        self.progress_queue = SecureQueue(progress_queue, f"progress-{writer_id}")
        self.writer_id = writer_id
        self.batch_size = batch_size
        self.logger = logging.getLogger(f"{__name__}.DatabaseWriter.{writer_id}")
        
        # Statistics
        self.stats = {
            'bom_rows': 0,
            'successful_files': 0,
            'total_files': 0,
            'failed_files': 0
        }
        
        # Memory management
        memory_limit = 512  # 512MB per DB writer
        self.memory_monitor = EnhancedMemoryMonitor(f"DB-{writer_id}", memory_limit)
        self.max_batch_memory_mb = 100
    
    def run(self):
        """Main database writer loop with comprehensive error handling"""
        self.logger.info(f"Database writer {self.writer_id} started")
        
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
                if current_time - last_status_report > 30:
                    self._send_status_update()
                    last_status_report = current_time
                
                # Process messages
                msg_data = self.results_queue.get_with_timeout(timeout=1.0)
                if msg_data is None:
                    # Handle timeout
                    if bom_batch and (
                        len(bom_batch) > self.batch_size // 2 or
                        current_time - last_heartbeat_check > 5
                    ):
                        self._flush_batch_optimized(bom_batch)
                    
                    if current_time - last_heartbeat_check > HEARTBEAT_INTERVAL:
                        last_heartbeat_check = current_time
                        active_workers.clear()
                    continue
                
                # Process received message
                msg = WorkerMessage.deserialize(msg_data)
                
                if msg.msg_type == MessageType.SHUTDOWN.value:
                    self.logger.info(f"Writer {self.writer_id} received shutdown")
                    break
                
                elif msg.msg_type == MessageType.BOM_DATA.value:
                    self._process_result(msg.data, bom_batch)
                    active_workers.add(msg.worker_id)
                
                elif msg.msg_type == MessageType.HEARTBEAT.value:
                    active_workers.add(msg.worker_id)
                
                elif msg.msg_type == MessageType.ERROR.value:
                    self.logger.error(f"Worker {msg.worker_id} error: {msg.data}")
                    active_workers.add(msg.worker_id)
                
                # Force flush if batch is getting too large
                if len(bom_batch) >= self.batch_size or self._estimate_batch_memory(bom_batch) > self.max_batch_memory_mb:
                    self._flush_batch_optimized(bom_batch)
        
        except Exception as e:
            self.logger.critical(f"Database writer {self.writer_id} fatal error: {e}", exc_info=True)
        
        finally:
            if bom_batch:
                self._flush_batch_optimized(bom_batch)
            
            self.logger.info(f"Database writer {self.writer_id} stopping. Stats: {self.stats}")
    
    def _estimate_batch_memory(self, batch: List[BOMRow]) -> float:
        """Estimate memory usage of current batch in MB"""
        if not batch:
            return 0.0
        return (len(batch) * 500) / 1024 / 1024  # ~500 bytes per row
    
    def _send_status_update(self):
        """Send status update to progress queue"""
        try:
            status_msg = WorkerMessage(
                msg_type=MessageType.DB_WRITER_STATUS.value,
                worker_id=self.writer_id,
                timestamp=time.time(),
                data={
                    'alive': True,
                    'stats': self.stats.copy(),
                    'memory_mb': self.memory_monitor.peak_memory_mb,
                    'writer_id': self.writer_id
                }
            )
            self.progress_queue.put_with_timeout(status_msg.serialize(), timeout=1.0)
        except Exception as e:
            self.logger.error(f"Failed to send status update: {e}")
    
    def _process_result(self, result_data: dict, bom_batch: List[BOMRow]):
        """Process a worker result with enhanced error handling"""
        try:
            result = self._dict_to_result(result_data)
            
            # Add BOM rows to batch
            if result.bom_rows:
                bom_batch.extend(result.bom_rows)
                self.stats['bom_rows'] += len(result.bom_rows)
            
            # Insert summary immediately
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
                worker_id=self.writer_id,
                timestamp=time.time(),
                data={
                    'filename': result.filename,
                    'success': result.success,
                    'stats': self.stats.copy(),
                    'bom_rows_count': len(result.bom_rows),
                    'pages_processed': result.total_pages
                }
            )
            self.progress_queue.put_with_timeout(progress_msg.serialize(), timeout=1.0)
        except Exception as e:
            self.logger.debug(f"Failed to send progress update: {e}")
    
    def _flush_batch_optimized(self, bom_batch: List[BOMRow]):
        """Optimized batch insert with better error handling"""
        if not bom_batch:
            return
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                with sqlite3.connect(self.db_path, timeout=45.0) as conn:
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
                return
            
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e).lower() and attempt < max_retries - 1:
                    self.logger.warning(f"Database locked, retrying in {2 ** attempt} seconds...")
                    time.sleep(2 ** attempt)
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

class ProductionProgressMonitor:
    """Enhanced progress monitoring with detailed reporting"""
    
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
    
    def _process_progress_update(self, msg: WorkerMessage):
        """Process progress update message"""
        data = msg.data
        if not data:
            return
        
        stats = data.get('stats', {})
        
        self.processed_files = stats.get('total_files', 0)
        self.successful_files = stats.get('successful_files', 0)
        self.failed_files = stats.get('failed_files', 0)
        
        # Update totals
        if 'bom_rows_count' in data:
            self.total_bom_rows += data['bom_rows_count']
        
        if 'pages_processed' in data:
            self.total_pages_processed += data['pages_processed']
    
    def _update_db_metrics(self, data: dict):
        """Update database metrics"""
        if data and 'stats' in data:
            db_stats = data['stats']
            # Could aggregate stats from multiple DB writers here
    
    def _update_worker_activity(self, worker_id: int):
        """Update worker activity tracking"""
        current_time = time.time()
        self.active_workers.add(worker_id)
        self.worker_last_seen[worker_id] = current_time
    
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
            f"\n⚡ Rate: {self.processing_rates[-1]:.1f} files/sec" if self.processing_rates else f"\n⚡ Rate: Calculating..." +
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
            self.logger.critical("Database writers appear unresponsive!")
    
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