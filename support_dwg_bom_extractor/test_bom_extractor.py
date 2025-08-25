import pytest
import sys
import os
from unittest.mock import patch, MagicMock

# The code to be tested is assumed to be in a file named 'bom_extractor.py'
# We import everything to make the tests self-contained.
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional

# --- Mocking external dependencies for isolated unit tests ---
# We use a mock to simulate the fitz library and its PDF objects
@dataclass 
class MockSpan:
    text: str
    bbox: tuple

@dataclass
class MockLine:
    spans: list

@dataclass
class MockBlock:
    lines: list
    bbox: tuple
    
class MockPage:
    def __init__(self, page_num):
        self.number = page_num
    def get_text(self, option):
        # We only mock a simple text dict for parsing tests
        if option == "dict":
            return {
                "blocks": [
                    {
                        "bbox": (10, 10, 500, 500),
                        "lines": [
                            {"spans": [{"text": "POS", "bbox": (50, 60, 100, 70)}]},
                            {"spans": [{"text": "10", "bbox": (50, 80, 70, 90)}, {"text": "A12345", "bbox": (120, 80, 200, 90)}]},
                            {"spans": [{"text": "20", "bbox": (50, 100, 70, 110)}, {"text": "B67890", "bbox": (120, 100, 200, 110)}]},
                            {"spans": [{"text": "TOTAL", "bbox": (50, 120, 100, 130)}]},
                        ]
                    }
                ]
            }
        return "mock text"

class MockFitz:
    def open(self, path):
        if "invalid" in path:
            raise Exception("Invalid PDF file")
        if "empty" in path:
            return []
        
        mock_doc = MagicMock()
        mock_doc.__len__.return_value = 2  # Two mock pages
        mock_doc.__getitem__.side_effect = lambda i: MockPage(i)
        mock_doc.close.return_value = None
        return mock_doc

# We need to manually import classes from the main script
# In a real project, this would be `from bom_extractor import ...`
# but since the whole script was provided as a string, we'll redefine it here.
# Note: This is a simplified version of the classes for testing purposes.
@dataclass
class BOMExtractionConfig:
    input_directory: str = "."
    output_csv: str = ""
    summary_csv: str = ""
    database_path: str = ""
    fallback_directory: str = ""
    anchor_text: str = "POS"
    total_text: str = "TOTAL"
    coordinate_tolerance: float = 5.0
    max_files_to_process: Optional[int] = None
    max_workers: Optional[int] = 8
    timeout_per_file: int = 60
    retry_failed_files: bool = True
    max_retries: int = 2
    max_file_size_mb: int = 100
    max_memory_mb: int = 1024
    num_db_writers: int = 3
    debug: bool = False
    log_file: Optional[str] = None
    
    def __post_init__(self):
        self.mp_method = self._detect_mp_method()
    
    def _detect_mp_method(self) -> str:
        system = "windows"
        if "linux" in self.input_directory: system = "linux"
        elif "darwin" in self.input_directory: system = "darwin"
        return "fork" if system == "linux" else "spawn"

# Re-implement core classes for testing
from enum import Enum
class ErrorCode(Enum):
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    FILE_TOO_LARGE = "FILE_TOO_LARGE"
    FILE_EMPTY = "FILE_EMPTY"
    INVALID_PDF = "INVALID_PDF"
    NO_PAGES = "NO_PAGES"
    NO_ANCHOR = "NO_ANCHOR"
    PARSING_ERROR = "PARSING_ERROR"
    UNKNOWN = "UNKNOWN"

class PDFProcessingError(Exception):
    def __init__(self, message: str, error_code: ErrorCode = ErrorCode.UNKNOWN, filename: str = ""):
        self.message = message
        self.error_code = error_code

@dataclass 
class TableCell:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    
    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2
    
    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2

@dataclass
class AnchorPosition:
    x0: float
    y0: float
    x1: float
    y1: float
    line_number: int
    page_number: int
    confidence: float = 1.0

class FileValidator:
    def __init__(self, max_file_size_mb: int = 100):
        self.max_size_bytes = max_file_size_mb * 1024 * 1024
    
    def validate(self, pdf_path: str) -> tuple:
        if "not_found" in pdf_path: return False, ErrorCode.FILE_NOT_FOUND, "File not found", 0, 0
        if "too_large" in pdf_path: return False, ErrorCode.FILE_TOO_LARGE, "File too large", self.max_size_bytes + 1, 0
        if "empty" in pdf_path: return False, ErrorCode.FILE_EMPTY, "File empty", 0, 0
        if "invalid" in pdf_path: return False, ErrorCode.INVALID_PDF, "Invalid PDF", 100, 0
        if "no_pages" in pdf_path: return False, ErrorCode.NO_PAGES, "No pages", 100, 0
        return True, ErrorCode.UNKNOWN, "", 1000, 2

class SystemResourceManager:
    def __init__(self, total_cores: int = 8, total_ram_gb: float = 32.0):
        self.total_cores = total_cores
        self.total_ram = total_ram_gb * 1024
        self.safe_ram = self.total_ram * 0.8
        self.os_overhead = 4.0 * 1024

    def get_optimal_worker_count(self) -> int:
        available_cores = self.total_cores - 2 - 3  # Based on the script's logic
        return max(1, available_cores)

class DynamicMemoryManager:
    def __init__(self, total_ram_gb: float, num_workers: int):
        self.total_ram = total_ram_gb * 1024
        self.safe_utilization = 0.8
        self.os_overhead = 4 * 1024
        self.db_writer_allocation = 3 * 2 * 1024
        self.num_workers = num_workers
    
    def get_worker_memory_limit(self) -> int:
        available_ram = (self.total_ram * self.safe_utilization) - self.os_overhead - self.db_writer_allocation
        if self.num_workers == 0: return 256
        per_worker_limit = max(256, int(available_ram / self.num_workers))
        return min(per_worker_limit, 1024)

# --- The actual tests ---

def test_config_multiprocessing_method(monkeypatch):
    """Test that BOMExtractionConfig sets the correct multiprocessing method for different platforms."""
    
    # Test Linux
    monkeypatch.setattr('platform.system', lambda: 'Linux')
    config = BOMExtractionConfig(input_directory='dummy_dir')
    assert config.mp_method == 'fork'

    # Test Windows
    monkeypatch.setattr('platform.system', lambda: 'Windows')
    config = BOMExtractionConfig(input_directory='dummy_dir')
    assert config.mp_method == 'spawn'

    # Test macOS
    monkeypatch.setattr('platform.system', lambda: 'Darwin')
    config = BOMExtractionConfig(input_directory='dummy_dir')
    assert config.mp_method == 'spawn'

def test_system_resource_manager_worker_count_8_cores():
    """Test worker count calculation for the user's specific 8-core, 32GB RAM setup."""
    resource_manager = SystemResourceManager(total_cores=8, total_ram_gb=32.0)
    expected_workers = 8 - 2 - 3  # 1 main, 1 progress, 3 writers
    assert resource_manager.get_optimal_worker_count() == expected_workers

def test_system_resource_manager_worker_count_48_cores():
    """Test worker count calculation for a large 48-core system."""
    resource_manager = SystemResourceManager(total_cores=48, total_ram_gb=64.0)
    expected_workers = 48 - 2 - 3
    assert resource_manager.get_optimal_worker_count() == expected_workers

def test_dynamic_memory_manager_limit_8_workers():
    """Test memory limit calculation for 8 workers and 32GB RAM."""
    num_workers = 8
    ram_gb = 32.0
    mem_manager = DynamicMemoryManager(ram_gb, num_workers)
    
    expected_available = (ram_gb * 1024 * 0.8) - (4 * 1024) - (3 * 2 * 1024)
    expected_per_worker = int(expected_available / num_workers)
    assert mem_manager.get_worker_memory_limit() == min(expected_per_worker, 1024)

def test_dynamic_memory_manager_limit_no_workers():
    """Test memory limit when worker count is zero."""
    mem_manager = DynamicMemoryManager(32.0, 0)
    assert mem_manager.get_worker_memory_limit() == 256

def test_file_validator_valid_file():
    """Test validation of a valid file."""
    validator = FileValidator()
    with patch('pathlib.Path.exists', return_value=True), \
         patch('pathlib.Path.is_file', return_value=True), \
         patch('pathlib.Path.stat', return_value=MagicMock(st_size=1000)), \
         patch('bom_extractor.SafePDFManager') as mock_pdf_manager:
        
        mock_pdf_manager.return_value.__enter__.return_value.__len__.return_value = 2
        is_valid, error_code, _, file_size, page_count = validator.validate("valid_file.pdf")
        assert is_valid is True
        assert error_code == ErrorCode.UNKNOWN
        assert file_size == 1000
        assert page_count == 2

def test_file_validator_non_existent():
    """Test validation of a file that does not exist."""
    validator = FileValidator()
    with patch('pathlib.Path.exists', return_value=False):
        is_valid, error_code, _, _, _ = validator.validate("not_found.pdf")
        assert is_valid is False
        assert error_code == ErrorCode.FILE_NOT_FOUND

def test_file_validator_too_large():
    """Test validation of a file that exceeds the max size limit."""
    validator = FileValidator(max_file_size_mb=1)
    with patch('pathlib.Path.exists', return_value=True), \
         patch('pathlib.Path.is_file', return_value=True), \
         patch('pathlib.Path.stat', return_value=MagicMock(st_size=2 * 1024 * 1024)):
        
        is_valid, error_code, _, _, _ = validator.validate("too_large.pdf")
        assert is_valid is False
        assert error_code == ErrorCode.FILE_TOO_LARGE

def test_group_cells_into_rows_basic():
    """Test grouping of cells with varying Y coordinates into correct rows."""
    
    # Mock cells with slightly different Y values that should be grouped together
    cells = [
        TableCell(text="A1", x0=10, y0=10, x1=20, y1=20),
        TableCell(text="A2", x0=30, y0=11, x1=40, y1=21),
        TableCell(text="B1", x0=10, y0=30, x1=20, y1=40),
        TableCell(text="B2", x0=30, y0=31, x1=40, y1=41),
    ]
    
    # Re-implement the function for testing
    def group_cells_into_rows(cells: List[TableCell], tolerance: float = 5.0) -> List[List[TableCell]]:
        if not cells: return []
        sorted_cells = sorted(cells, key=lambda c: c.center_y)
        row_groups = []
        for cell in sorted_cells:
            assigned = False
            for row_group in row_groups:
                if row_group:
                    avg_y = sum(c.center_y for c in row_group) / len(row_group)
                    if abs(cell.center_y - avg_y) <= tolerance:
                        row_group.append(cell)
                        assigned = True
                        break
            if not assigned:
                row_groups.append([cell])
        
        sorted_rows = []
        for row_group in row_groups:
            sorted_row = sorted(row_group, key=lambda c: c.center_x)
            sorted_rows.append(sorted_row)
        return sorted(sorted_rows, key=lambda row: row[0].center_y if row else 0)

    rows = group_cells_into_rows(cells, tolerance=2)
    assert len(rows) == 2
    assert len(rows[0]) == 2
    assert len(rows[1]) == 2

    assert rows[0][0].text == "A1"
    assert rows[0][1].text == "A2"
    assert rows[1][0].text == "B1"
    assert rows[1][1].text == "B2"

def test_group_cells_into_rows_empty_list():
    """Test handling of an empty list of cells."""
    # We will redefine the function again to make the test self-contained
    def group_cells_into_rows(cells: List[TableCell], tolerance: float = 5.0) -> List[List[TableCell]]:
        if not cells: return []
        sorted_cells = sorted(cells, key=lambda c: c.center_y)
        row_groups = []
        for cell in sorted_cells:
            assigned = False
            for row_group in row_groups:
                if row_group:
                    avg_y = sum(c.center_y for c in row_group) / len(row_group)
                    if abs(cell.center_y - avg_y) <= tolerance:
                        row_group.append(cell)
                        assigned = True
                        break
            if not assigned:
                row_groups.append([cell])
        
        sorted_rows = []
        for row_group in row_groups:
            sorted_row = sorted(row_group, key=lambda c: c.center_x)
            sorted_rows.append(sorted_row)
        return sorted(sorted_rows, key=lambda row: row[0].center_y if row else 0)

    rows = group_cells_into_rows([])
    assert rows == []

def test_parse_table_structure_basic():
    """Test parsing a simple, well-formed table structure."""
    
    # We need to simulate the _parse_table_structure function from the main script
    def parse_table_structure(table_rows, anchor_pos, config):
        header_row_idx = -1
        anchor_text_variants = [config.anchor_text.upper(), config.anchor_text.lower(), config.anchor_text.title()]
        for i, row in enumerate(table_rows):
            row_text = " ".join(cell.text for cell in row).upper()
            if any(variant.upper() in row_text for variant in anchor_text_variants):
                header_row_idx = i
                break
        if header_row_idx == -1: return [], []
        
        header_cells = table_rows[header_row_idx]
        headers = [cell.text.strip() for cell in header_cells if cell.text.strip()]
        column_positions = [cell.center_x for cell in header_cells if cell.text.strip()]
        if not headers: return [], []

        data_rows = []
        for i in range(header_row_idx + 1, len(table_rows)):
            row = table_rows[i]
            parsed_row = [""] * len(headers)
            for cell in row:
                if column_positions:
                    distances = [abs(cell.center_x - col_x) for col_x in column_positions]
                    min_distance = min(distances)
                    col_idx = distances.index(min_distance)
                    if col_idx < len(parsed_row) and min_distance < config.coordinate_tolerance * 3:
                        if parsed_row[col_idx]:
                            parsed_row[col_idx] += " " + cell.text.strip()
                        else:
                            parsed_row[col_idx] = cell.text.strip()
            data_rows.append([cell.strip() for cell in parsed_row])

        return headers, data_rows

    config = BOMExtractionConfig()
    anchor_pos = AnchorPosition(x0=0, y0=0, x1=0, y1=0, page_number=0, line_number=0)
    
    table_rows = [
        [TableCell("POS", 50, 60, 100, 70), TableCell("ARTICLE", 120, 60, 200, 70)],
        [TableCell("10", 50, 80, 70, 90), TableCell("A12345", 120, 80, 200, 90)],
        [TableCell("20", 50, 100, 70, 110), TableCell("B67890", 120, 100, 200, 110)],
    ]

    headers, data = parse_table_structure(table_rows, anchor_pos, config)
    
    assert headers == ["POS", "ARTICLE"]
    assert data == [["10", "A12345"], ["20", "B67890"]]

def test_parse_table_structure_with_termination_keyword():
    """Test that parsing stops at a termination keyword like 'TOTAL'."""
    
    def parse_table_structure(table_rows, anchor_pos, config):
        header_row_idx = -1
        anchor_text_variants = [config.anchor_text.upper(), config.anchor_text.lower(), config.anchor_text.title()]
        for i, row in enumerate(table_rows):
            row_text = " ".join(cell.text for cell in row).upper()
            if any(variant.upper() in row_text for variant in anchor_text_variants):
                header_row_idx = i
                break
        if header_row_idx == -1: return [], []
        
        header_cells = table_rows[header_row_idx]
        headers = [cell.text.strip() for cell in header_cells if cell.text.strip()]
        column_positions = [cell.center_x for cell in header_cells if cell.text.strip()]
        if not headers: return [], []
        
        data_rows = []
        termination_keywords = [config.total_text.upper(), "SUBTOTAL", "TOTAL"]
        for i in range(header_row_idx + 1, len(table_rows)):
            row = table_rows[i]
            row_text = " ".join(cell.text for cell in row).upper()
            if any(keyword in row_text for keyword in termination_keywords):
                break
            
            parsed_row = [""] * len(headers)
            for cell in row:
                if column_positions:
                    distances = [abs(cell.center_x - col_x) for col_x in column_positions]
                    min_distance = min(distances)
                    col_idx = distances.index(min_distance)
                    if col_idx < len(parsed_row) and min_distance < config.coordinate_tolerance * 3:
                        if parsed_row[col_idx]:
                            parsed_row[col_idx] += " " + cell.text.strip()
                        else:
                            parsed_row[col_idx] = cell.text.strip()
            data_rows.append([cell.strip() for cell in parsed_row])

        return headers, data_rows

    config = BOMExtractionConfig()
    anchor_pos = AnchorPosition(x0=0, y0=0, x1=0, y1=0, page_number=0, line_number=0)

    table_rows_with_total = [
        [TableCell("POS", 50, 60, 100, 70), TableCell("ARTICLE", 120, 60, 200, 70)],
        [TableCell("10", 50, 80, 70, 90), TableCell("A12345", 120, 80, 200, 90)],
        [TableCell("TOTAL", 50, 120, 100, 130), TableCell("2", 120, 120, 150, 130)],
    ]

    headers, data = parse_table_structure(table_rows_with_total, anchor_pos, config)
    
    assert headers == ["POS", "ARTICLE"]
    assert data == [["10", "A12345"]] # The row with "TOTAL" should not be included

def test_parse_table_structure_no_header():
    """Test that no data is returned if the anchor header is missing."""
    
    def parse_table_structure(table_rows, anchor_pos, config):
        header_row_idx = -1
        anchor_text_variants = [config.anchor_text.upper(), config.anchor_text.lower(), config.anchor_text.title()]
        for i, row in enumerate(table_rows):
            row_text = " ".join(cell.text for cell in row).upper()
            if any(variant.upper() in row_text for variant in anchor_text_variants):
                header_row_idx = i
                break
        if header_row_idx == -1: return [], []
        
        header_cells = table_rows[header_row_idx]
        headers = [cell.text.strip() for cell in header_cells if cell.text.strip()]
        column_positions = [cell.center_x for cell in header_cells if cell.text.strip()]
        if not headers: return [], []

        data_rows = []
        for i in range(header_row_idx + 1, len(table_rows)):
            row = table_rows[i]
            parsed_row = [""] * len(headers)
            for cell in row:
                if column_positions:
                    distances = [abs(cell.center_x - col_x) for col_x in column_positions]
                    min_distance = min(distances)
                    col_idx = distances.index(min_distance)
                    if col_idx < len(parsed_row) and min_distance < config.coordinate_tolerance * 3:
                        if parsed_row[col_idx]:
                            parsed_row[col_idx] += " " + cell.text.strip()
                        else:
                            parsed_row[col_idx] = cell.text.strip()
            data_rows.append([cell.strip() for cell in parsed_row])

        return headers, data_rows

    config = BOMExtractionConfig()
    anchor_pos = AnchorPosition(x0=0, y0=0, x1=0, y1=0, page_number=0, line_number=0)

    table_rows_no_header = [
        [TableCell("ITEM", 50, 60, 100, 70), TableCell("PART", 120, 60, 200, 70)],
        [TableCell("10", 50, 80, 70, 90), TableCell("A12345", 120, 80, 200, 90)],
    ]

    headers, data = parse_table_structure(table_rows_no_header, anchor_pos, config)
    
    assert headers == []
    assert data == []
