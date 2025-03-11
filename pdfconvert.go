package pdfconvert

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"

	"github.com/jung-kurt/gofpdf"
)

// WordXML represents the structure of a DOCX document
type WordXML struct {
	Body struct {
		Paragraphs []struct {
			TextRuns []struct {
				Text string `xml:"t"`
			} `xml:"r"`
		} `xml:"p"`
	} `xml:"body"`
}

// ExtractText extracts text from a DOCX file
func ExtractText(docxPath string) (string, error) {
	zipReader, err := zip.OpenReader(docxPath)
	if err != nil {
		return "", fmt.Errorf("failed to open DOCX: %v", err)
	}
	defer zipReader.Close()

	var contentXML string
	for _, file := range zipReader.File {
		if file.Name == "word/document.xml" {
			rc, err := file.Open()
			if err != nil {
				return "", fmt.Errorf("failed to read document.xml: %v", err)
			}
			defer rc.Close()

			buf := new(bytes.Buffer)
			_, err = io.Copy(buf, rc)
			if err != nil {
				return "", fmt.Errorf("failed to read XML content: %v", err)
			}

			contentXML = buf.String()
			break
		}
	}

	var doc WordXML
	err = xml.Unmarshal([]byte(contentXML), &doc)
	if err != nil {
		return "", fmt.Errorf("failed to parse document.xml: %v", err)
	}

	var extractedText []string
	for _, p := range doc.Body.Paragraphs {
		for _, r := range p.TextRuns {
			extractedText = append(extractedText, r.Text)
		}
	}

	return strings.Join(extractedText, "\n"), nil
}

// GeneratePDF creates a PDF file from extracted text
func GeneratePDF(text, pdfPath string) error {
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.SetFont("Arial", "", 12)
	pdf.AddPage()
	pdf.MultiCell(190, 10, text, "", "L", false)

	err := pdf.OutputFileAndClose(pdfPath)
	if err != nil {
		return fmt.Errorf("failed to save PDF: %v", err)
	}

	return nil
}

// Convert handles the full process of converting DOCX to PDF
func Convert(docxPath, pdfPath string) error {
	text, err := ExtractText(docxPath)
	if err != nil {
		return err
	}

	err = GeneratePDF(text, pdfPath)
	if err != nil {
		return err
	}

	return nil
}
