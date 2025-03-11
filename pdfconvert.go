package pdfconvert

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"github.com/jung-kurt/gofpdf"
	"io"
	"strings"
)

// WordXML represents the structure of a DOCX document
type WordXML struct {
	Body struct {
		Paragraphs []struct {
			TextRuns []struct {
				Text      string `xml:"t"`
				Bold      bool   `xml:"rPr>b"`
				Italic    bool   `xml:"rPr<i"`
				Underline bool   `xml:"rPr<u"`
			} `xml:"r"`
		} `xml:"p"`
	} `xml:"body"`
}

// ExtractTextWithFormatting extracts text with basic formatting from a DOCX file
func ExtractTextWithFormatting(docxPath string) (string, []string, error) {
	zipReader, err := zip.OpenReader(docxPath)
	if err != nil {
		return "", nil, fmt.Errorf("failed to open DOCX: %v", err)
	}
	defer zipReader.Close()

	var contentXML string
	var images []string
	for _, file := range zipReader.File {
		if file.Name == "word/document.xml" {
			rc, err := file.Open()
			if err != nil {
				return "", nil, fmt.Errorf("failed to read document.xml: %v", err)
			}
			defer rc.Close()

			buf := new(bytes.Buffer)
			_, err = io.Copy(buf, rc)
			if err != nil {
				return "", nil, fmt.Errorf("failed to read XML content: %v", err)
			}

			contentXML = buf.String()
		}
		if strings.HasPrefix(file.Name, "word/media/") {
			images = append(images, file.Name) // Store image names for later use
		}
	}

	var doc WordXML
	err = xml.Unmarshal([]byte(contentXML), &doc)
	if err != nil {
		return "", nil, fmt.Errorf("failed to parse document.xml: %v", err)
	}

	var extractedText []string
	for _, p := range doc.Body.Paragraphs {
		for _, r := range p.TextRuns {
			// Handle text formatting: Bold, Italic, Underline
			style := ""
			if r.Bold {
				style += "B"
			}
			if r.Italic {
				style += "I"
			}
			if r.Underline {
				style += "U"
			}
			extractedText = append(extractedText, fmt.Sprintf("%s%s", style, r.Text))
		}
	}

	return strings.Join(extractedText, "\n"), images, nil
}

// GeneratePDF creates a more advanced PDF with images and formatting
func GeneratePDF(text string, images []string, pdfPath string) error {
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.SetFont("Arial", "", 12)
	pdf.AddPage()

	// Add text with formatting
	lines := strings.Split(text, "\n")
	for _, line := range lines {
		// Apply formatting if necessary (simple bold, italic, underline handling)
		if strings.Contains(line, "B") {
			pdf.SetFont("Arial", "B", 12)
		} else if strings.Contains(line, "I") {
			pdf.SetFont("Arial", "I", 12)
		} else {
			pdf.SetFont("Arial", "", 12)
		}
		pdf.MultiCell(190, 10, line, "", "L", false)
	}

	// Handle embedded images (basic image inclusion)
	for _, imgPath := range images {
		// We would extract image content here and add it to the PDF
		// Assuming the images are extracted into a directory
		imageFilePath := fmt.Sprintf("word/media/%s", imgPath)
		pdf.Image(imageFilePath, 10, 40, 100, 100, false, "", 0, "")
	}

	// Save PDF to file
	err := pdf.OutputFileAndClose(pdfPath)
	if err != nil {
		return fmt.Errorf("failed to save PDF: %v", err)
	}

	return nil
}
