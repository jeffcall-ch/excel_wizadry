package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/pdfcpu/pdfcpu/pkg/api"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run debug_pdfcpu.go <pdf_file>")
		return
	}

	pdfPath := os.Args[1]
	
	// Create a configuration for pdfcpu
	config := pdfcpu.NewDefaultConfiguration()
	
	fmt.Printf("PDF: %s\n", pdfPath)

	// Try to extract text using pdfcpu's text extraction
	// Note: pdfcpu may extract to a separate text file
	outputPath := "extracted_text.txt"
	err := api.ExtractContentFile(pdfPath, outputPath, nil, config)
	if err != nil {
		log.Fatalf("Error extracting text: %v", err)
	}

	// Read the extracted text file
	textBytes, err := os.ReadFile(outputPath)
	if err != nil {
		log.Fatalf("Error reading extracted text file: %v", err)
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
	
	// Clean up
	os.Remove(outputPath)
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
