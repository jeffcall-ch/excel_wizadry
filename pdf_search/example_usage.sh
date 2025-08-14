#!/bin/bash

# Example usage script for PDF Search Tool (Go version)
# This script demonstrates different ways to use the tool

echo "PDF Search Tool (Go) - Example Usage"
echo "====================================="
echo

# Example 1: Basic search
echo "Example 1: Basic search for 'error' in current directory"
echo "Command: go run multi_pdf_full_text_search.go -search=\"error\" -dir=\".\""
echo

# Example 2: Case-sensitive search
echo "Example 2: Case-sensitive search for 'ERROR' with custom output"
echo "Command: go run multi_pdf_full_text_search.go -search=\"ERROR\" -dir=\"./pdfs\" -case-sensitive -output=\"error_results.xlsx\""
echo

# Example 3: Search in specific directory (Windows path)
echo "Example 3: Search in Windows directory"
echo "Command: go run multi_pdf_full_text_search.go -search=\"fail\" -dir=\"C:\\temp\\rev1_isos\""
echo

# Example 4: Build and run executable
echo "Example 4: Build executable and run"
echo "Commands:"
echo "  go build -o pdf_search.exe multi_pdf_full_text_search.go"
echo "  .\\pdf_search.exe -search=\"warning\" -dir=\"./documents\""
echo

# Example 5: Show help
echo "Example 5: Show help"
echo "Command: go run multi_pdf_full_text_search.go -help"
echo

echo "Note: Replace paths and search terms with your actual values"
echo "Make sure to run 'go mod tidy' first to download dependencies"
