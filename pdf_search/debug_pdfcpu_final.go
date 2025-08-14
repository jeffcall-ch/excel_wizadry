package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/pdfcpu/pdfcpu/pkg/api"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run debug_pdfcpu_final.go <pdf_file>")
		return
	}

	pdfPath := os.Args[1]
	
	fmt.Printf("PDF: %s\n", pdfPath)

	// Try to extract text using pdfcpu's text extraction
	// Use current directory as output
	outputDir := "."
	err := api.ExtractContentFile(pdfPath, outputDir, nil, nil)
	if err != nil {
		log.Printf("Error extracting text: %v", err)
		fmt.Printf("\n✗ pdfcpu failed to extract text from this PDF.\n")
		return
	}

	fmt.Printf("Text extraction completed, looking for output files...\n")

	// Look for generated text files
	baseName := strings.TrimSuffix(filepath.Base(pdfPath), filepath.Ext(pdfPath))
	pattern := baseName + "*Content*.txt"
	
	matches, err := filepath.Glob(pattern)
	if err != nil {
		log.Printf("Error searching for text files: %v", err)
		return
	}

	if len(matches) == 0 {
		fmt.Printf("\n✗ No text files found with pattern: %s\n", pattern)
		fmt.Printf("✗ PDF might not contain extractable text.\n")
		return
	}

	fmt.Printf("Found %d text file(s): %v\n", len(matches), matches)

	// Read the first text file
	textFile := matches[0]
	textBytes, err := os.ReadFile(textFile)
	if err != nil {
		log.Printf("Error reading text file %s: %v", textFile, err)
		return
	}
	
	text := string(textBytes)
	
	fmt.Printf("\nRaw text length: %d characters\n", len(text))
	if len(text) > 0 {
		fmt.Printf("First 500 characters:\n%s\n", text[:min(500, len(text))])
		
		// Check for "pip" (case insensitive)
		if strings.Contains(strings.ToLower(text), "pip") {
			fmt.Printf("\n✓ Found 'pip' in the text!\n")
		} else {
			fmt.Printf("\n✗ 'pip' not found in the text.\n")
		}
	} else {
		fmt.Printf("\n✗ No text extracted from the PDF.\n")
	}
	
	// Clean up generated files
	for _, file := range matches {
		os.Remove(file)
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
