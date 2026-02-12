package gopresentation

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
)

// Writer is the interface for presentation writers.
type Writer interface {
	Save(path string) error
	WriteTo(w io.Writer) error
}

// WriterType represents the output format.
type WriterType string

const (
	WriterPowerPoint2007 WriterType = "PowerPoint2007"
)

// NewWriter creates a writer for the given format.
func NewWriter(p *Presentation, format WriterType) (Writer, error) {
	switch format {
	case WriterPowerPoint2007:
		return &PPTXWriter{presentation: p}, nil
	default:
		return nil, fmt.Errorf("unsupported writer format: %s", format)
	}
}

// PPTXWriter writes presentations in PPTX format.
type PPTXWriter struct {
	presentation *Presentation
	relID        int
}

func (w *PPTXWriter) nextRelID() string {
	w.relID++
	return fmt.Sprintf("rId%d", w.relID)
}

// Save writes the presentation to a file.
func (w *PPTXWriter) Save(path string) error {
	dir := filepath.Dir(path)
	if dir != "" && dir != "." {
		if err := os.MkdirAll(dir, 0750); err != nil {
			return fmt.Errorf("failed to create directory: %w", err)
		}
	}
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}

	writeErr := w.WriteTo(f)
	closeErr := f.Close()

	if writeErr != nil {
		// Attempt cleanup on write failure
		os.Remove(path)
		return writeErr
	}
	return closeErr
}

// WriteTo writes the presentation to a writer.
func (w *PPTXWriter) WriteTo(writer io.Writer) error {
	if w.presentation == nil {
		return fmt.Errorf("presentation is nil")
	}

	zw := zip.NewWriter(writer)

	w.relID = 0

	// Write [Content_Types].xml
	if err := w.writeContentTypes(zw); err != nil {
		return err
	}

	// Write _rels/.rels
	if err := w.writeRootRels(zw); err != nil {
		return err
	}

	// Write docProps/app.xml
	if err := w.writeAppProperties(zw); err != nil {
		return err
	}

	// Write docProps/core.xml
	if err := w.writeCoreProperties(zw); err != nil {
		return err
	}

	// Write ppt/presentation.xml
	if err := w.writePresentation(zw); err != nil {
		return err
	}

	// Write ppt/_rels/presentation.xml.rels
	if err := w.writePresentationRels(zw); err != nil {
		return err
	}

	// Write ppt/presProps.xml
	if err := w.writePresProps(zw); err != nil {
		return err
	}

	// Write ppt/viewProps.xml
	if err := w.writeViewProps(zw); err != nil {
		return err
	}

	// Write ppt/tableStyles.xml
	if err := w.writeTableStyles(zw); err != nil {
		return err
	}

	// Write slide master and layout
	if err := w.writeSlideMaster(zw); err != nil {
		return err
	}

	if err := w.writeSlideLayout(zw); err != nil {
		return err
	}

	// Write theme
	if err := w.writeTheme(zw); err != nil {
		return err
	}

	// Write slides
	for i, slide := range w.presentation.slides {
		hlinkRelMap := w.buildHyperlinkRelMap(slide)
		if err := w.writeSlide(zw, slide, i+1, hlinkRelMap); err != nil {
			return err
		}
		if err := w.writeSlideRels(zw, slide, i+1, hlinkRelMap); err != nil {
			return err
		}
	}

	// Write images
	if err := w.writeMedia(zw); err != nil {
		return err
	}

	// Write charts
	chartIdx := 1
	for _, slide := range w.presentation.slides {
		for _, shape := range slide.shapes {
			if cs, ok := shape.(*ChartShape); ok {
				if err := w.writeChartPart(zw, cs, chartIdx); err != nil {
					return err
				}
				chartIdx++
			}
		}
	}

	// Write comments
	if w.hasComments() {
		if err := w.writeCommentAuthors(zw); err != nil {
			return err
		}
		for i, slide := range w.presentation.slides {
			if err := w.writeSlideComments(zw, slide, i+1); err != nil {
				return err
			}
		}
	}

	// Write notes slides
	for i, slide := range w.presentation.slides {
		if slide.notes != "" {
			if err := w.writeNotesSlide(zw, slide, i+1); err != nil {
				return err
			}
		}
	}

	return zw.Close()
}
