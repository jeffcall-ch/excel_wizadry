package main

import (
	"fmt"
	"log"
	"os"

	"github.com/unidoc/unipdf/v3/extractor"
	"github.com/unidoc/unipdf/v3/model"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run debug_unipdf.go <pdf_file>")
		return
	}

	pdfPath := os.Args[1]
	
	// Open the PDF file
	file, err := os.Open(pdfPath)
	if err != nil {
		log.Fatalf("Error opening file: %v", err)
	}
	defer file.Close()

	// Create a PDF reader
	pdfReader, err := model.NewPdfReader(file)
	if err != nil {
		log.Fatalf("Error creating PDF reader: %v", err)
	}

	fmt.Printf("PDF: %s\n", pdfPath)
	
	numPages, err := pdfReader.GetNumPages()
	if err != nil {
		log.Fatalf("Error getting number of pages: %v", err)
	}
	fmt.Printf("Number of pages: %d\n", numPages)

	// Process first page only
	if numPages > 0 {
		page, err := pdfReader.GetPage(1)
		if err != nil {
			log.Fatalf("Error getting page 1: %v", err)
		}

		// Extract text from the page
		ex, err := extractor.New(page)
		if err != nil {
			log.Fatalf("Error creating extractor: %v", err)
		}

		text, err := ex.ExtractText()
		if err != nil {
			log.Fatalf("Error extracting text: %v", err)
		}

		fmt.Printf("\nRaw text length: %d characters\n", len(text))
		if len(text) > 0 {
			fmt.Printf("First 500 characters:\n%s\n", text[:min(500, len(text))])
			
			// Check for "pip" (case insensitive)
			if contains(text, "pip") {
				fmt.Printf("\n✓ Found 'pip' in the text!\n")
			} else {
				fmt.Printf("\n✗ 'pip' not found in the text.\n")
			}
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func contains(text, substr string) bool {
	// Case insensitive search
	for i := 0; i <= len(text)-len(substr); i++ {
		if len(text[i:i+len(substr)]) == len(substr) {
			match := true
			for j := 0; j < len(substr); j++ {
				c1 := text[i+j]
				c2 := substr[j]
				if c1 >= 'A' && c1 <= 'Z' {
					c1 = c1 + 32 // convert to lowercase
				}
				if c2 >= 'A' && c2 <= 'Z' {
					c2 = c2 + 32 // convert to lowercase
				}
				if c1 != c2 {
					match = false
					break
				}
			}
			if match {
				return true
			}
		}
	}
	return false
}
