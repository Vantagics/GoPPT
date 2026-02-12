package gopresentation

import (
	"fmt"
	"strings"
)

// Validate checks the presentation for structural issues and returns an error
// describing all problems found, or nil if the presentation is valid.
// This is analogous to unioffice's Validate() method.
func (p *Presentation) Validate() error {
	var errs []string

	if p.properties == nil {
		errs = append(errs, "document properties are nil")
	}
	if p.presentationProperties == nil {
		errs = append(errs, "presentation properties are nil")
	}
	if p.layout == nil {
		errs = append(errs, "document layout is nil")
	} else {
		if p.layout.CX <= 0 {
			errs = append(errs, "layout width (CX) must be positive")
		}
		if p.layout.CY <= 0 {
			errs = append(errs, "layout height (CY) must be positive")
		}
	}
	if len(p.slides) == 0 {
		errs = append(errs, "presentation must have at least one slide")
	}

	for i, slide := range p.slides {
		prefix := fmt.Sprintf("slide %d", i+1)
		if slideErrs := validateSlide(slide); len(slideErrs) > 0 {
			for _, e := range slideErrs {
				errs = append(errs, prefix+": "+e)
			}
		}
	}

	if len(errs) == 0 {
		return nil
	}
	return fmt.Errorf("validation failed:\n  %s", strings.Join(errs, "\n  "))
}

func validateSlide(s *Slide) []string {
	var errs []string
	for j, shape := range s.shapes {
		prefix := fmt.Sprintf("shape %d", j+1)
		if shape == nil {
			errs = append(errs, prefix+": shape is nil")
			continue
		}
		if shape.GetWidth() < 0 {
			errs = append(errs, prefix+": width is negative")
		}
		if shape.GetHeight() < 0 {
			errs = append(errs, prefix+": height is negative")
		}

		switch sh := shape.(type) {
		case *DrawingShape:
			if sh.data == nil && sh.path == "" {
				errs = append(errs, prefix+": drawing shape has no image data or path")
			}
			if sh.mimeType != "" && !isValidImageMime(sh.mimeType) {
				errs = append(errs, prefix+": unsupported image MIME type: "+sh.mimeType)
			}
		case *TableShape:
			if sh.numRows <= 0 || sh.numCols <= 0 {
				errs = append(errs, prefix+": table must have at least 1 row and 1 column")
			}
			if sh.numRows > 0 && sh.numCols > 0 && len(sh.rows) != sh.numRows {
				errs = append(errs, prefix+": table row count mismatch")
			}
		case *ChartShape:
			if sh.plotArea.chartType == nil {
				errs = append(errs, prefix+": chart shape has no chart type set")
			}
		case *RichTextShape:
			if len(sh.paragraphs) == 0 {
				errs = append(errs, prefix+": rich text shape has no paragraphs")
			}
			if sh.columns < 1 {
				errs = append(errs, prefix+": text columns must be >= 1")
			}
			errs = append(errs, validateParagraphs(sh.paragraphs, prefix)...)
		case *PlaceholderShape:
			if len(sh.paragraphs) == 0 {
				errs = append(errs, prefix+": placeholder shape has no paragraphs")
			}
			if sh.phType == "" {
				errs = append(errs, prefix+": placeholder type is empty")
			}
		case *LineShape:
			if !isValidARGB(sh.lineColor.ARGB) {
				errs = append(errs, prefix+": line color is invalid ARGB")
			}
		case *GroupShape:
			for k, gs := range sh.shapes {
				if gs == nil {
					errs = append(errs, fmt.Sprintf("%s: group child %d is nil", prefix, k+1))
				}
			}
		}
	}

	for j, c := range s.comments {
		if c == nil {
			errs = append(errs, fmt.Sprintf("comment %d: is nil", j+1))
			continue
		}
		if c.Author == nil {
			errs = append(errs, fmt.Sprintf("comment %d: missing author", j+1))
		}
		if c.Text == "" {
			errs = append(errs, fmt.Sprintf("comment %d: empty text", j+1))
		}
	}

	return errs
}

// validateParagraphs checks paragraph elements for common issues.
func validateParagraphs(paragraphs []*Paragraph, prefix string) []string {
	var errs []string
	for i, para := range paragraphs {
		if para == nil {
			errs = append(errs, fmt.Sprintf("%s: paragraph %d is nil", prefix, i+1))
			continue
		}
		if para.alignment == nil {
			errs = append(errs, fmt.Sprintf("%s: paragraph %d has nil alignment", prefix, i+1))
		}
		for k, elem := range para.elements {
			if elem == nil {
				errs = append(errs, fmt.Sprintf("%s: paragraph %d element %d is nil", prefix, i+1, k+1))
				continue
			}
			if tr, ok := elem.(*TextRun); ok {
				if tr.font == nil {
					errs = append(errs, fmt.Sprintf("%s: paragraph %d text run %d has nil font", prefix, i+1, k+1))
				}
			}
		}
	}
	return errs
}

// isValidImageMime checks if a MIME type is a supported image format.
func isValidImageMime(mime string) bool {
	switch mime {
	case "image/png", "image/jpeg", "image/gif", "image/bmp", "image/svg+xml":
		return true
	}
	return false
}
