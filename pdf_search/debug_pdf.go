package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/ledongthuc/pdf"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run debug_pdf.go <pdf_file>")
		return
	}

	pdfPath := os.Args[1]
	
	f, r, err := pdf.Open(pdfPath)
	if err != nil {
		log.Fatalf("Error opening PDF %s: %v", pdfPath, err)
	}
	defer f.Close()

	fmt.Printf("PDF: %s\n", pdfPath)
	fmt.Printf("Number of pages: %d\n", r.NumPage())

	// Process first page only
	if r.NumPage() > 0 {
		page := r.Page(1)
		if !page.V.IsNull() {
			pageText, err := page.GetPlainText(nil)
			if err != nil {
				log.Printf("Error extracting text from page 1: %v", err)
				return
			}

			// Clean up the text
			pageText = strings.ReplaceAll(pageText, "\n", " ")
			
			fmt.Printf("\nRaw text length: %d characters\n", len(pageText))
			fmt.Printf("First 500 characters:\n%s\n", pageText[:min(500, len(pageText))])
			
			// Check for "pip" (case insensitive)
			lowerText := strings.ToLower(pageText)
			if strings.Contains(lowerText, "pip") {
				fmt.Printf("\n✓ Found 'pip' in the text!\n")
				
				// Show context around first occurrence
				index := strings.Index(lowerText, "pip")
				start := max(0, index-50)
				end := min(len(pageText), index+53)
				fmt.Printf("Context: ...%s...\n", pageText[start:end])
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

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}
