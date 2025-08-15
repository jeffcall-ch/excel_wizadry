import os
import tempfile
import pytest
import pandas as pd
from pdf_search import multi_pdf_full_text_search as mpfts

# Helper to create a dummy PDF file with text
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def create_pdf_with_text(path, texts):
    c = canvas.Canvas(path, pagesize=letter)
    width, height = letter
    for text in texts:
        c.drawString(72, height - 72, text)
        c.showPage()
    c.save()

@pytest.fixture
def temp_pdf_dir():
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf1 = os.path.join(tmpdir, 'doc1.pdf')
        pdf2 = os.path.join(tmpdir, 'doc2.pdf')
        create_pdf_with_text(pdf1, ["This is a test page with histogram and history.", "Another page with nothing."])
        create_pdf_with_text(pdf2, ["This page mentions histograms.", "No relevant terms here."])
        yield tmpdir

def test_parse_boolean_query():
    query = "term1 AND (term2 OR term3) AND NOT term4"
    tokens = mpfts.parse_boolean_query(query)
    assert any(t[0] == 'AND' for t in tokens)
    assert any(t[0] == 'OR' for t in tokens)
    assert any(t[0] == 'NOT' for t in tokens)
    assert any(t[0] == 'LPAREN' for t in tokens)
    assert any(t[0] == 'RPAREN' for t in tokens)
    assert any(t[0] == 'TERM' for t in tokens)

def test_evaluate_boolean_expression_simple():
    tokens = mpfts.parse_boolean_query("test")
    assert mpfts.evaluate_boolean_expression(tokens, "this is a test")
    assert not mpfts.evaluate_boolean_expression(tokens, "no match here")

def test_evaluate_boolean_expression_complex():
    tokens = mpfts.parse_boolean_query("test AND histogram")
    assert mpfts.evaluate_boolean_expression(tokens, "test histogram")
    assert not mpfts.evaluate_boolean_expression(tokens, "test only")
    tokens = mpfts.parse_boolean_query("test OR histogram")
    assert mpfts.evaluate_boolean_expression(tokens, "test only")
    assert mpfts.evaluate_boolean_expression(tokens, "histogram only")
    tokens = mpfts.parse_boolean_query("test AND NOT histogram")
    assert mpfts.evaluate_boolean_expression(tokens, "test only")
    assert not mpfts.evaluate_boolean_expression(tokens, "test histogram")

def test_partial_match():
    tokens = mpfts.parse_boolean_query("hist*")
    assert mpfts.evaluate_boolean_expression(tokens, "histogram history")
    assert mpfts.evaluate_boolean_expression(tokens, "this is a histogram")
    assert not mpfts.evaluate_boolean_expression(tokens, "no match here")

def test_search_pdf_for_boolean_text(temp_pdf_dir):
    pdf_path = os.path.join(temp_pdf_dir, 'doc1.pdf')
    matches = mpfts.search_pdf_for_boolean_text(pdf_path, "hist*")
    assert any('hist' in m['matched_term'] for m in matches)
    assert all(m['page_number'] == 1 for m in matches)

def test_search_directory_for_pdfs(temp_pdf_dir):
    matches = mpfts.search_directory_for_pdfs(temp_pdf_dir, "hist*")
    assert len(matches) > 0
    assert all('matched_term' in m for m in matches)

def test_create_excel_report(temp_pdf_dir):
    matches = mpfts.search_directory_for_pdfs(temp_pdf_dir, "hist*")
    output_file = os.path.join(temp_pdf_dir, "report.xlsx")
    mpfts.create_excel_report(matches, output_file, "hist*")
    assert os.path.exists(output_file)
    df = pd.read_excel(output_file, sheet_name='Search Results')
    assert 'file_name' in df.columns
    assert len(df) == len(matches)
