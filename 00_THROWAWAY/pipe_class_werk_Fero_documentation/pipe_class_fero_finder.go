package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"sync"

	"github.com/ledongthuc/pdf"
)

type MatchResult struct {
	FilePath        string
	FoundCodes      string
	DatesInPath     string
	DatesInFilename string
	DatesInContent  string
}

var pipeClassCodes = []string{
	"EFDX", "EFDJ", "EHDX", "EEDX", "ECDE", "ECDM", "EHDQ", "EEDQ",
	"EHGN", "EHFD", "EEFD", "AHDX", "AEDX", "ACDE", "ACDM", "AHDQ",
	"AHFD", "AHGN",
}

func extractDates(text string) []string {
	var dates []string
	seenDates := make(map[string]bool)

	// Pattern for years (1900-2099)
	yearPattern := regexp.MustCompile(`\b(19\d{2}|20\d{2})\b`)
	years := yearPattern.FindAllString(text, -1)
	for _, year := range years {
		if !seenDates[year] {
			dates = append(dates, year)
			seenDates[year] = true
		}
	}

	// Pattern for dates like: 2021-05-15, 2021/05/15, 2021.05.15
	datePattern1 := regexp.MustCompile(`\b(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\b`)
	dates1 := datePattern1.FindAllString(text, -1)
	for _, date := range dates1 {
		if !seenDates[date] {
			dates = append(dates, date)
			seenDates[date] = true
		}
	}

	// Pattern for dates like: 15-05-2021, 15/05/2021, 15.05.2021
	datePattern2 := regexp.MustCompile(`\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})\b`)
	dates2 := datePattern2.FindAllString(text, -1)
	for _, date := range dates2 {
		if !seenDates[date] {
			dates = append(dates, date)
			seenDates[date] = true
		}
	}

	// Pattern for dates like: May 15, 2021 or 15 May 2021
	datePattern3 := regexp.MustCompile(`(?i)\b((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})\b`)
	dates3 := datePattern3.FindAllString(text, -1)
	for _, date := range dates3 {
		if !seenDates[date] {
			dates = append(dates, date)
			seenDates[date] = true
		}
	}

	// Pattern for dates like: 2021 May 15
	datePattern4 := regexp.MustCompile(`(?i)\b(\d{4}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2})\b`)
	dates4 := datePattern4.FindAllString(text, -1)
	for _, date := range dates4 {
		if !seenDates[date] {
			dates = append(dates, date)
			seenDates[date] = true
		}
	}

	return dates
}

func extractTextFromPDF(path string) (string, error) {
	f, r, err := pdf.Open(path)
	if err != nil {
		return "", err
	}
	defer f.Close()

	var text strings.Builder
	totalPages := r.NumPage()

	for pageNum := 1; pageNum <= totalPages; pageNum++ {
		page := r.Page(pageNum)
		if page.V.IsNull() {
			continue
		}

		content, err := page.GetPlainText(nil)
		if err != nil {
			continue
		}
		text.WriteString(content)
	}

	return text.String(), nil
}

func processPDF(path string, index int, total int) (result *MatchResult) {
	// Recover from panics in PDF processing
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("  ✗ ERROR: panic recovered - %v\n", r)
			result = nil
		}
	}()

	fmt.Printf("[%d/%d] Processing: %s\n", index, total, filepath.Base(path))

	text, err := extractTextFromPDF(path)
	if err != nil {
		fmt.Printf("  ✗ ERROR reading file: %v\n", err)
		return nil
	}

	textUpper := strings.ToUpper(text)

	// Check if FERO is present
	if !strings.Contains(textUpper, "FERO") {
		return nil
	}

	// Check if any pipe class code is present
	var foundCodes []string
	for _, code := range pipeClassCodes {
		if strings.Contains(textUpper, code) {
			foundCodes = append(foundCodes, code)
		}
	}

	if len(foundCodes) == 0 {
		return nil
	}

	// Extract dates
	datesInPath := extractDates(filepath.Dir(path))
	datesInFilename := extractDates(filepath.Base(path))
	datesInContent := extractDates(text)

	fmt.Printf("  ✓ MATCH FOUND - Contains: %s\n", strings.Join(foundCodes, ", "))

	return &MatchResult{
		FilePath:        path,
		FoundCodes:      strings.Join(foundCodes, ", "),
		DatesInPath:     strings.Join(datesInPath, ", "),
		DatesInFilename: strings.Join(datesInFilename, ", "),
		DatesInContent:  strings.Join(datesInContent, ", "),
	}
}

func main() {
	searchDir := `C:\temp\WERK`
	outputCSV := "fero_pipe_class_files.csv"

	fmt.Printf("Searching for PDFs in: %s\n", searchDir)
	fmt.Printf("Looking for 'FERO' AND one of: %s\n", strings.Join(pipeClassCodes, ", "))
	fmt.Println(strings.Repeat("-", 80))

	// Check if directory exists
	if _, err := os.Stat(searchDir); os.IsNotExist(err) {
		fmt.Printf("ERROR: Directory does not exist: %s\n", searchDir)
		return
	}

	// Find all PDF files
	var pdfFiles []string
	err := filepath.Walk(searchDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return nil
		}
		if !info.IsDir() && strings.HasSuffix(strings.ToLower(info.Name()), ".pdf") {
			pdfFiles = append(pdfFiles, path)
		}
		return nil
	})

	if err != nil {
		fmt.Printf("ERROR walking directory: %v\n", err)
		return
	}

	fmt.Printf("Found %d PDF files to process\n\n", len(pdfFiles))

	// Process PDFs concurrently
	results := make([]*MatchResult, 0)
	var mu sync.Mutex
	var wg sync.WaitGroup
	semaphore := make(chan struct{}, 8) // Limit to 8 concurrent workers

	for i, pdfPath := range pdfFiles {
		wg.Add(1)
		go func(path string, index int) {
			defer wg.Done()
			semaphore <- struct{}{}        // Acquire
			defer func() { <-semaphore }() // Release

			result := processPDF(path, index+1, len(pdfFiles))
			if result != nil {
				mu.Lock()
				results = append(results, result)
				mu.Unlock()
			}
		}(pdfPath, i)
	}

	wg.Wait()

	// Write results to CSV
	fmt.Println("\n" + strings.Repeat("=", 80))
	fmt.Printf("Found %d matching files\n", len(results))

	// Get current working directory for output
	scriptDir, err := os.Getwd()
	if err != nil {
		scriptDir = "."
	}
	outputPath := filepath.Join(scriptDir, outputCSV)

	fmt.Printf("Writing results to: %s\n", outputPath)

	file, err := os.Create(outputPath)
	if err != nil {
		fmt.Printf("ERROR creating CSV: %v\n", err)
		return
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	// Write header
	writer.Write([]string{"file_path", "found_codes", "dates_in_path", "dates_in_filename", "dates_in_content"})

	// Write data
	for _, result := range results {
		writer.Write([]string{
			result.FilePath,
			result.FoundCodes,
			result.DatesInPath,
			result.DatesInFilename,
			result.DatesInContent,
		})
	}

	fmt.Println("✓ CSV file created successfully!")
	fmt.Println("\nMatching files:")
	for _, result := range results {
		fmt.Printf("  - %s\n", result.FilePath)
		fmt.Printf("    Codes: %s\n", result.FoundCodes)
		if result.DatesInPath != "" {
			fmt.Printf("    Dates in path: %s\n", result.DatesInPath)
		}
		if result.DatesInFilename != "" {
			fmt.Printf("    Dates in filename: %s\n", result.DatesInFilename)
		}
		if result.DatesInContent != "" {
			fmt.Printf("    Dates in content: %s\n", result.DatesInContent)
		}
	}
}
