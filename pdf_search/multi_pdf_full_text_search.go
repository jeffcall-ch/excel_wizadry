package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/ledongthuc/pdf"
	"github.com/xuri/excelize/v2"
)

// MatchResult represents a single search match
type MatchResult struct {
	FilePath     string
	PageNumber   int
	MatchContext string
	FileName     string
}

// SearchConfig holds the search configuration
type SearchConfig struct {
	SearchText    string
	Directory     string
	OutputFile    string
	CaseSensitive bool
	PartialMatch  bool
}

// searchPDFForText searches a PDF file for the specified text and returns matches with page numbers
func searchPDFForText(pdfPath, searchText string, caseSensitive, partialMatch bool) ([]MatchResult, error) {
	var matches []MatchResult

	// Open the PDF file
	f, r, err := pdf.Open(pdfPath)
	if err != nil {
		return nil, fmt.Errorf("error opening PDF %s: %v", pdfPath, err)
	}
	defer f.Close()

	// Get the number of pages
	numPages := r.NumPage()

	for pageNum := 1; pageNum <= numPages; pageNum++ {
		// Get the page
		page := r.Page(pageNum)
		if page.V.IsNull() {
			continue
		}

		// Extract text from the page
		pageText, err := page.GetPlainText(nil)
		if err != nil {
			log.Printf("Error extracting text from page %d of %s: %v", pageNum, pdfPath, err)
			continue
		}

		// Clean up the text (remove excessive whitespace)
		pageText = strings.ReplaceAll(pageText, "\n", " ")
		pageText = regexp.MustCompile(`\s+`).ReplaceAllString(pageText, " ")

		// Prepare search pattern and text for comparison
		var searchPattern, textToSearch string
		if caseSensitive {
			searchPattern = searchText
			textToSearch = pageText
		} else {
			searchPattern = strings.ToLower(searchText)
			textToSearch = strings.ToLower(pageText)
		}

		// Check if the text exists on this page
		var found bool
		var matchPositions []int
		
		if partialMatch {
			// For partial matching, use regex to find the search text as part of words
			var pattern string
			if caseSensitive {
				pattern = regexp.QuoteMeta(searchText)
			} else {
				pattern = "(?i)" + regexp.QuoteMeta(searchText)
			}
			
			re, err := regexp.Compile(pattern)
			if err != nil {
				log.Printf("Error compiling regex pattern for %s: %v", searchText, err)
				continue
			}
			
			// Find all matches
			allMatches := re.FindAllStringIndex(textToSearch, -1)
			if len(allMatches) > 0 {
				found = true
				for _, match := range allMatches {
					matchPositions = append(matchPositions, match[0])
				}
			}
		} else {
			// Exact matching
			if strings.Contains(textToSearch, searchPattern) {
				found = true
				// Find all occurrences
				index := 0
				for {
					pos := strings.Index(textToSearch[index:], searchPattern)
					if pos == -1 {
						break
					}
					matchPositions = append(matchPositions, index+pos)
					index = index + pos + len(searchPattern)
				}
			}
		}

		if found {
			// Process all matches and get surrounding context
			for _, matchPos := range matchPositions {
				// Get context (50 characters before and after)
				startPos := maxInt(0, matchPos-50)
				endPos := minInt(len(pageText), matchPos+len(searchText)+50)
				
				context := "..." + pageText[startPos:endPos] + "..."
				
				// Store match information
				matches = append(matches, MatchResult{
					FilePath:     pdfPath,
					PageNumber:   pageNum,
					MatchContext: context,
					FileName:     filepath.Base(pdfPath),
				})
			}
		}
	}

	return matches, nil
}

// searchDirectoryForPDFs searches all PDF files in the specified directory and its subdirectories
func searchDirectoryForPDFs(directory, searchText string, caseSensitive, partialMatch bool) ([]MatchResult, error) {
	var allMatches []MatchResult
	pdfCount := 0
	matchCount := 0

	err := filepath.Walk(directory, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			log.Printf("Error accessing path %s: %v", path, err)
			return nil
		}

		// Check if it's a PDF file
		if !info.IsDir() && strings.ToLower(filepath.Ext(path)) == ".pdf" {
			pdfCount++
			fmt.Printf("Searching %s...\n", path)

			// Search this PDF
			matches, err := searchPDFForText(path, searchText, caseSensitive, partialMatch)
			if err != nil {
				log.Printf("Error searching PDF %s: %v", path, err)
				return nil
			}

			if len(matches) > 0 {
				matchCount += len(matches)
				allMatches = append(allMatches, matches...)
			}
		}

		return nil
	})

	if err != nil {
		return nil, fmt.Errorf("error walking directory: %v", err)
	}

	fmt.Printf("\nSearch completed. Found %d matches across %d PDF files.\n", matchCount, pdfCount)
	return allMatches, nil
}

// createExcelReport creates an Excel report with match results and hyperlinks to the original files
func createExcelReport(matches []MatchResult, outputFile, searchText string) error {
	if len(matches) == 0 {
		fmt.Println("No matches found. No Excel report generated.")
		return nil
	}

	// Create a new Excel file
	f := excelize.NewFile()

	// Create the main results sheet
	sheetName := "Search Results"
	f.NewSheet(sheetName)
	f.DeleteSheet("Sheet1") // Remove default sheet

	// Set headers
	headers := []string{"File Path", "File Name", "Page Number", "Match Context"}
	for i, header := range headers {
		cell := fmt.Sprintf("%c1", 'A'+i)
		f.SetCellValue(sheetName, cell, header)
	}

	// Add data rows
	for i, match := range matches {
		row := i + 2 // Start from row 2 (after header)
		
		// File Path (with hyperlink)
		filePathCell := fmt.Sprintf("A%d", row)
		f.SetCellValue(sheetName, filePathCell, match.FilePath)
		f.SetCellHyperLink(sheetName, filePathCell, match.FilePath, "External")
		
		// File Name
		fileNameCell := fmt.Sprintf("B%d", row)
		f.SetCellValue(sheetName, fileNameCell, match.FileName)
		
		// Page Number
		pageCell := fmt.Sprintf("C%d", row)
		f.SetCellValue(sheetName, pageCell, match.PageNumber)
		
		// Match Context
		contextCell := fmt.Sprintf("D%d", row)
		f.SetCellValue(sheetName, contextCell, match.MatchContext)
	}

	// Set column widths
	f.SetColWidth(sheetName, "A", "A", 50)
	f.SetColWidth(sheetName, "B", "B", 30)
	f.SetColWidth(sheetName, "C", "C", 15)
	f.SetColWidth(sheetName, "D", "D", 80)

	// Create summary sheet
	summarySheet := "Summary"
	f.NewSheet(summarySheet)

	summaryData := [][]interface{}{
		{"Item", "Value"},
		{"Date of Search", time.Now().Format("2006-01-02 15:04:05")},
		{"Search Term", searchText},
		{"Total PDF Files Searched", len(getUniqueFilePaths(matches))},
		{"Total Matches Found", len(matches)},
		{"Report Generated By", "PDF Search Tool (Go)"},
	}

	for i, row := range summaryData {
		for j, value := range row {
			cell := fmt.Sprintf("%c%d", 'A'+j, i+1)
			f.SetCellValue(summarySheet, cell, value)
		}
	}

	// Set column widths for summary
	f.SetColWidth(summarySheet, "A", "A", 25)
	f.SetColWidth(summarySheet, "B", "B", 40)

	// Ensure output directory exists
	outputDir := filepath.Dir(outputFile)
	if outputDir != "" && outputDir != "." {
		err := os.MkdirAll(outputDir, 0755)
		if err != nil {
			return fmt.Errorf("error creating output directory: %v", err)
		}
	}

	// Save the file
	if err := f.SaveAs(outputFile); err != nil {
		return fmt.Errorf("error saving Excel file: %v", err)
	}

	return nil
}

// getUniqueFilePaths returns unique file paths from matches
func getUniqueFilePaths(matches []MatchResult) []string {
	uniquePaths := make(map[string]bool)
	for _, match := range matches {
		uniquePaths[match.FilePath] = true
	}
	
	var paths []string
	for path := range uniquePaths {
		paths = append(paths, path)
	}
	return paths
}

// Helper functions
func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func main() {
	// Define command line flags
	var config SearchConfig
	var showHelp bool

	flag.StringVar(&config.SearchText, "search", "", "Text to search for in PDF files (required)")
	flag.StringVar(&config.Directory, "dir", "", "Directory containing PDF files to search (required)")
	flag.StringVar(&config.OutputFile, "output", "", "Output file path for Excel report")
	flag.BoolVar(&config.CaseSensitive, "case-sensitive", false, "Enable case-sensitive search")
	flag.BoolVar(&config.PartialMatch, "partial", false, "Enable partial matching (e.g., 'fail' matches 'failed', 'failure')")
	flag.BoolVar(&showHelp, "help", false, "Show help message")

	flag.Parse()

	// Show help if requested or if required arguments are missing
	if showHelp || config.SearchText == "" || config.Directory == "" {
		fmt.Println("PDF Search Tool - Search PDF files for specific text and generate Excel report")
		fmt.Println()
		fmt.Println("Usage:")
		fmt.Println("  go run multi_pdf_full_text_search.go -search=\"text to find\" -dir=\"/path/to/pdfs\" [options]")
		fmt.Println()
		fmt.Println("Required flags:")
		fmt.Println("  -search     Text to search for in PDF files")
		fmt.Println("  -dir        Directory containing PDF files to search")
		fmt.Println()
		fmt.Println("Optional flags:")
		fmt.Println("  -output     Output file path for Excel report (default: auto-generated)")
		fmt.Println("  -case-sensitive Enable case-sensitive search")
		fmt.Println("  -partial    Enable partial matching (e.g., 'fail' matches 'failed', 'failure')")
		fmt.Println("  -help       Show this help message")
		fmt.Println()
		fmt.Println("Examples:")
		fmt.Println("  go run multi_pdf_full_text_search.go -search=\"error\" -dir=\"./documents\"")
		fmt.Println("  go run multi_pdf_full_text_search.go -search=\"Error\" -dir=\"./docs\" -case-sensitive -output=\"results.xlsx\"")
		fmt.Println("  go run multi_pdf_full_text_search.go -search=\"fail\" -dir=\"./pdfs\" -partial")
		return
	}

	// Set default output file if not specified
	if config.OutputFile == "" {
		timestamp := time.Now().Format("20060102_150405")
		config.OutputFile = fmt.Sprintf("pdf_search_results_%s.xlsx", timestamp)
	}

	fmt.Printf("Searching for '%s' in %s...\n", config.SearchText, config.Directory)

	// Perform search
	matches, err := searchDirectoryForPDFs(config.Directory, config.SearchText, config.CaseSensitive, config.PartialMatch)
	if err != nil {
		log.Fatalf("Error during search: %v", err)
	}

	// Create Excel report
	err = createExcelReport(matches, config.OutputFile, config.SearchText)
	if err != nil {
		log.Fatalf("Error creating Excel report: %v", err)
	}

	if len(matches) > 0 {
		fmt.Printf("Excel report created: %s\n", config.OutputFile)
	}
}
