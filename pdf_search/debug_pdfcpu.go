paimport (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/pdfcpu/pdfcpu/pkg/api"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu"
)n

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/pdfcpu/pdfcpu/pkg/api"
	"github.com/pdfcpu/pdfcpu/pkg/pdfcpu/model"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run debug_pdfcpu.go <pdf_file>")
		return
	}

	pdfPath := os.Args[1]
	
	// Create a configuration for pdfcpu
	config := model.NewDefaultConfiguration()
	
	fmt.Printf("PDF: %s\n", pdfPath)

	// Extract text from the PDF
	err := api.ExtractContentFile(pdfPath, "", nil, config)
	if err != nil {
		log.Fatalf("Error extracting text: %v", err)
	}

	fmt.Printf("Text extraction completed\n")
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
