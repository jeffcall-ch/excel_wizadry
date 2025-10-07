"""
Contract Screening Script for Piping Keywords
============================================

This script processes PDF contracts to identify piping-related content by:
1. Extracting text from PDF files
2. Searching for keywords from a CSV file
3. Capturing context (5 words before/after each match)
4. Scoring pages based on keyword frequency
5. Generating comprehensive Excel reports

Author: GitHub Copilot
Date: September 2025
"""

import re
import csv
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
from typing import List, Dict, Tuple, Any
from dataclasses import dataclass
from collections import defaultdict, Counter
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@dataclass
class KeywordMatch:
    """Data class to store keyword match information"""
    page_number: int
    keyword: str
    category: str
    context: str
    position: int

@dataclass
class PageStats:
    """Data class to store page statistics"""
    page_number: int
    total_matches: int
    unique_keywords: int
    categories: set
    score: float

class ContractScreener:
    """Main class for contract screening functionality"""
    
    def __init__(self, keywords_file: str):
        """Initialize with keywords CSV file"""
        self.keywords_data = self._load_keywords(keywords_file)
        self.matches: List[KeywordMatch] = []
        self.page_stats: Dict[int, PageStats] = {}
        
    def _load_keywords(self, keywords_file: str) -> Dict[str, str]:
        """Load keywords from CSV file and return dict mapping keyword to category"""
        keywords = {}
        try:
            with open(keywords_file, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    keyword = row['Keyword'].strip().lower()
                    category = row['Category'].strip()
                    keywords[keyword] = category
            logger.info(f"Loaded {len(keywords)} keywords from {keywords_file}")
            return keywords
        except Exception as e:
            logger.error(f"Error loading keywords file: {e}")
            raise
    
    def _extract_text_from_pdf(self, pdf_path: str) -> Dict[int, str]:
        """Extract text from PDF file, returning dict of page_number: text"""
        page_texts = {}
        try:
            doc = fitz.open(pdf_path)
            logger.info(f"Processing PDF with {len(doc)} pages")
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text = page.get_text()
                page_texts[page_num + 1] = text  # 1-based page numbering
                
                if (page_num + 1) % 50 == 0:  # Progress logging
                    logger.info(f"Processed {page_num + 1} pages")
            
            doc.close()
            logger.info(f"Successfully extracted text from {len(page_texts)} pages")
            return page_texts
            
        except Exception as e:
            logger.error(f"Error processing PDF: {e}")
            raise
    
    def _get_context(self, text: str, match_start: int, match_end: int, words_before: int = 10, words_after: int = 10) -> str:
        """Extract context around a keyword match"""
        # Split text into words while preserving positions
        words = re.findall(r'\S+', text)
        word_positions = [(m.start(), m.end()) for m in re.finditer(r'\S+', text)]
        
        # Find which word contains our match
        match_word_idx = None
        for idx, (start, end) in enumerate(word_positions):
            if start <= match_start < end:
                match_word_idx = idx
                break
        
        if match_word_idx is None:
            return "Context not found"
        
        # Get surrounding words
        start_idx = max(0, match_word_idx - words_before)
        end_idx = min(len(words), match_word_idx + words_after + 1)
        
        context_words = words[start_idx:end_idx]
        return ' '.join(context_words)
    
    def _search_keywords_in_text(self, page_number: int, text: str) -> List[KeywordMatch]:
        """Search for all keywords in text and return matches with context"""
        matches = []
        text_lower = text.lower()
        
        for keyword, category in self.keywords_data.items():
            # Find all occurrences of keyword (case-insensitive substring search)
            for match in re.finditer(re.escape(keyword), text_lower):
                context = self._get_context(text, match.start(), match.end())
                
                keyword_match = KeywordMatch(
                    page_number=page_number,
                    keyword=keyword,
                    category=category,
                    context=context,
                    position=match.start()
                )
                matches.append(keyword_match)
        
        return matches
    
    def _calculate_page_stats(self) -> None:
        """Calculate statistics for each page"""
        page_data = defaultdict(lambda: {'matches': [], 'keywords': set(), 'categories': set()})
        
        # Group matches by page
        for match in self.matches:
            page_data[match.page_number]['matches'].append(match)
            page_data[match.page_number]['keywords'].add(match.keyword)
            page_data[match.page_number]['categories'].add(match.category)
        
        # Calculate stats for each page
        for page_num, data in page_data.items():
            total_matches = len(data['matches'])
            unique_keywords = len(data['keywords'])
            categories = data['categories']
            
            # Scoring: 1 point per match + bonus for keyword diversity
            score = total_matches + (unique_keywords * 0.5)
            
            self.page_stats[page_num] = PageStats(
                page_number=page_num,
                total_matches=total_matches,
                unique_keywords=unique_keywords,
                categories=categories,
                score=score
            )
    
    def process_pdf(self, pdf_path: str) -> None:
        """Main method to process PDF and find keyword matches"""
        logger.info(f"Starting contract screening for: {pdf_path}")
        
        # Extract text from PDF
        page_texts = self._extract_text_from_pdf(pdf_path)
        
        # Search for keywords in each page
        all_matches = []
        for page_num, text in page_texts.items():
            page_matches = self._search_keywords_in_text(page_num, text)
            all_matches.extend(page_matches)
            
            if page_num % 100 == 0:  # Progress logging
                logger.info(f"Searched keywords in {page_num} pages")
        
        self.matches = all_matches
        self._calculate_page_stats()
        
        logger.info(f"Found {len(self.matches)} total keyword matches across {len(self.page_stats)} pages")
    
    def generate_excel_report(self, output_path: str) -> None:
        """Generate comprehensive Excel report with multiple sheets"""
        logger.info(f"Generating Excel report: {output_path}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Sheet 1: All Results
            self._create_main_results_sheet(writer)
            
            # Sheet 2: Page Summary with Keywords
            self._create_page_summary_sheet(writer)
            
            # Sheet 3: Summary by Category
            self._create_category_summary_sheet(writer)
            
            # Sheet 4: High-Scoring Pages
            self._create_high_scoring_pages_sheet(writer)
            
            # Sheet 5: Page Statistics
            self._create_page_statistics_sheet(writer)
        
        logger.info(f"Excel report generated successfully: {output_path}")
    
    def _create_main_results_sheet(self, writer) -> None:
        """Create main results sheet with all matches"""
        if not self.matches:
            logger.warning("No matches found to write to Excel")
            return
        
        # Group matches by page to get all keywords per page
        page_keywords = defaultdict(set)
        for match in self.matches:
            page_keywords[match.page_number].add(match.keyword)
        
        data = []
        for match in self.matches:
            keywords_on_page = ', '.join(sorted(page_keywords[match.page_number]))
            data.append({
                'Page': match.page_number,
                'Keyword': match.keyword,
                'Category': match.category,
                'Context': match.context,
                'All Keywords on Page': keywords_on_page
            })
        
        df = pd.DataFrame(data)
        df.to_excel(writer, sheet_name='All Results', index=False)
        
        # Format the sheet
        worksheet = writer.sheets['All Results']
        worksheet.freeze_panes = 'A2'  # Freeze header row
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def _create_page_summary_sheet(self, writer) -> None:
        """Create sheet with page-by-page summary of all found keywords"""
        if not self.matches:
            return
        
        # Group matches by page
        page_data = defaultdict(lambda: {'keywords': set(), 'categories': set(), 'total_matches': 0})
        
        for match in self.matches:
            page_data[match.page_number]['keywords'].add(match.keyword)
            page_data[match.page_number]['categories'].add(match.category)
            page_data[match.page_number]['total_matches'] += 1
        
        data = []
        for page_num in sorted(page_data.keys()):
            page_info = page_data[page_num]
            keywords_str = ', '.join(sorted(page_info['keywords']))
            categories_str = ', '.join(sorted(page_info['categories']))
            
            data.append({
                'Page': page_num,
                'Total Matches': page_info['total_matches'],
                'Unique Keywords': len(page_info['keywords']),
                'Keywords Found': keywords_str,
                'Categories': categories_str,
                'Score': self.page_stats[page_num].score if page_num in self.page_stats else 0
            })
        
        df = pd.DataFrame(data)
        df.to_excel(writer, sheet_name='Page Summary', index=False)
        
        # Format the sheet
        worksheet = writer.sheets['Page Summary']
        worksheet.freeze_panes = 'A2'
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 60)  # Wider for keywords list
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def _create_category_summary_sheet(self, writer) -> None:
        """Create summary sheet by category"""
        category_stats = defaultdict(lambda: {'matches': 0, 'pages': set()})
        
        for match in self.matches:
            category_stats[match.category]['matches'] += 1
            category_stats[match.category]['pages'].add(match.page_number)
        
        data = []
        for category, stats in category_stats.items():
            data.append({
                'Category': category,
                'Total Matches': stats['matches'],
                'Pages with Matches': len(stats['pages']),
                'Avg Matches per Page': round(stats['matches'] / len(stats['pages']), 2) if stats['pages'] else 0
            })
        
        df = pd.DataFrame(data).sort_values('Total Matches', ascending=False)
        df.to_excel(writer, sheet_name='Category Summary', index=False)
        
        # Format the sheet
        worksheet = writer.sheets['Category Summary']
        worksheet.freeze_panes = 'A2'
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def _create_high_scoring_pages_sheet(self, writer) -> None:
        """Create sheet with high-scoring pages (top 20% or minimum score of 5)"""
        if not self.page_stats:
            return
        
        # Get high-scoring pages (top 20% or score >= 5)
        sorted_pages = sorted(self.page_stats.values(), key=lambda x: x.score, reverse=True)
        top_20_percent = max(1, len(sorted_pages) // 5)
        high_scoring = [p for p in sorted_pages if p.score >= 5][:top_20_percent]
        
        data = []
        for page_stat in high_scoring:
            categories_str = ', '.join(sorted(page_stat.categories))
            # Get all keywords found on this page
            page_keywords = [match.keyword for match in self.matches if match.page_number == page_stat.page_number]
            keywords_str = ', '.join(sorted(set(page_keywords)))
            data.append({
                'Page': page_stat.page_number,
                'Score': round(page_stat.score, 2),
                'Total Matches': page_stat.total_matches,
                'Unique Keywords': page_stat.unique_keywords,
                'Categories Found': categories_str,
                'Keywords Found': keywords_str
            })
        
        df = pd.DataFrame(data)
        df.to_excel(writer, sheet_name='High-Scoring Pages', index=False)
        
        # Format the sheet
        worksheet = writer.sheets['High-Scoring Pages']
        worksheet.freeze_panes = 'A2'
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 40)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def _create_page_statistics_sheet(self, writer) -> None:
        """Create sheet with overall statistics"""
        if not self.matches or not self.page_stats:
            return
        
        # Calculate overall statistics
        total_pages_with_matches = len(self.page_stats)
        total_matches = len(self.matches)
        unique_keywords_found = len(set(match.keyword for match in self.matches))
        categories_found = len(set(match.category for match in self.matches))
        avg_matches_per_page = round(total_matches / total_pages_with_matches, 2) if total_pages_with_matches else 0
        
        # Top keywords
        keyword_counts = Counter(match.keyword for match in self.matches)
        top_keywords = keyword_counts.most_common(10)
        
        # Create summary data
        summary_data = [
            ['Metric', 'Value'],
            ['Total Pages with Matches', total_pages_with_matches],
            ['Total Keyword Matches', total_matches],
            ['Unique Keywords Found', unique_keywords_found],
            ['Categories Found', categories_found],
            ['Average Matches per Page', avg_matches_per_page],
            ['', ''],  # Empty row
            ['Top 10 Keywords', 'Count']
        ]
        
        for keyword, count in top_keywords:
            summary_data.append([keyword, count])
        
        df = pd.DataFrame(summary_data)
        df.to_excel(writer, sheet_name='Statistics', index=False, header=False)
        
        # Format the sheet
        worksheet = writer.sheets['Statistics']
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            worksheet.column_dimensions[column_letter].width = adjusted_width


def main():
    """Main function to run the contract screening"""
    # Hardcoded paths
    script_dir = Path(__file__).parent
    pdf_path = r"C:\Users\szil\Repos\excel_wizadry\Contract_screening\00_Tees Valley contract full.pdf"
    keywords_csv = script_dir / "piping_contract_search_phrases_for_piping_and_supports.txt"
    output_path = script_dir / "Tees_Valley_contract_screening_results.xlsx"
    
    logger.info(f"PDF to process: {pdf_path}")
    logger.info(f"Keywords file: {keywords_csv}")
    logger.info(f"Output will be saved to: {output_path}")
    
    # Validate input files
    if not Path(pdf_path).exists():
        logger.error(f"PDF file not found: {pdf_path}")
        return
    
    if not Path(keywords_csv).exists():
        logger.error(f"Keywords CSV file not found: {keywords_csv}")
        return
    
    try:
        # Initialize screener and process PDF
        screener = ContractScreener(str(keywords_csv))
        screener.process_pdf(pdf_path)
        
        # Generate Excel report
        screener.generate_excel_report(str(output_path))
        
        logger.info("Contract screening completed successfully!")
        logger.info(f"Results saved to: {output_path}")
        
    except Exception as e:
        logger.error(f"Error during contract screening: {e}")
        raise


if __name__ == "__main__":
    main()