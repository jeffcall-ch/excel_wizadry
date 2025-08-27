import pytest
from support_dwg_bom_extractor.table_boundary_finder import get_table_boundaries
import os

# Test data directory
TEST_PDF_DIR = "C:\\Users\\szil\\Repos\\excel_wizadry\\support_dwg_bom_extractor\\test_pdfs"

@pytest.fixture
def setup_test_environment():
    """Fixture to set up the test environment."""
    os.makedirs(TEST_PDF_DIR, exist_ok=True)
    yield
    # Cleanup after tests
    for file in os.listdir(TEST_PDF_DIR):
        os.remove(os.path.join(TEST_PDF_DIR, file))
    os.rmdir(TEST_PDF_DIR)

def test_get_table_boundaries_happy_path(setup_test_environment):
    """Test the happy path for get_table_boundaries."""
    # Create a dummy PDF file with a table
    pdf_path = os.path.join(TEST_PDF_DIR, "happy_path.pdf")
    with open(pdf_path, "wb") as f:
        f.write(b"%PDF-1.4\n...")  # Minimal valid PDF content

    # Call the function
    boundaries = get_table_boundaries(pdf_path)

    # Assert the boundaries are returned (mocked example)
    assert boundaries is not None
    assert len(boundaries) == 4

def test_get_table_boundaries_empty_pdf(setup_test_environment):
    """Test get_table_boundaries with an empty PDF."""
    # Create an empty PDF file
    pdf_path = os.path.join(TEST_PDF_DIR, "empty.pdf")
    with open(pdf_path, "wb") as f:
        f.write(b"")

    # Call the function
    boundaries = get_table_boundaries(pdf_path)

    # Assert that no boundaries are returned
    assert boundaries is None

def test_get_table_boundaries_invalid_file(setup_test_environment):
    """Test get_table_boundaries with an invalid file."""
    # Create a non-PDF file
    invalid_path = os.path.join(TEST_PDF_DIR, "not_a_pdf.txt")
    with open(invalid_path, "w") as f:
        f.write("This is not a PDF.")

    # Call the function
    boundaries = get_table_boundaries(invalid_path)

    # Assert that no boundaries are returned
    assert boundaries is None

def test_get_table_boundaries_missing_file():
    """Test get_table_boundaries with a missing file."""
    # Call the function with a non-existent file
    missing_path = os.path.join(TEST_PDF_DIR, "missing.pdf")
    boundaries = get_table_boundaries(missing_path)

    # Assert that no boundaries are returned
    assert boundaries is None
