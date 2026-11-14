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
	// layoutPlan groups the slides by their layout part for the emission:
	// one layout part per distinct layout path, each carrying its own deco.
	layoutPlan []layoutPlanEntry
}

// layoutPlanEntry is one layout part: the slides grouped under it, its part
// index (1-based), and its decoration shapes.
type layoutPlanEntry struct {
	layoutPath string
	index      int
	decoShapes []Shape
}

// buildLayoutPlan groups the slides by their layout part path and assigns
// the layout part indexes. A single-layout deck yields one entry.
func (w *PPTXWriter) buildLayoutPlan() []layoutPlanEntry {
	var plan []layoutPlanEntry
	byLayout := map[string]int{}
	for _, slide := range w.presentation.slides {
		if slide.layoutPath == "" {
			continue
		}
		li, ok := byLayout[slide.layoutPath]
		if !ok {
			li = len(plan) + 1
			byLayout[slide.layoutPath] = li
			entry := layoutPlanEntry{layoutPath: slide.layoutPath, index: li}
			if cached, ok2 := w.presentation.layoutDecoCache[slide.layoutPath]; ok2 {
				entry.decoShapes = cached.shapes
			}
			plan = append(plan, entry)
		}
	}
	if len(plan) == 0 {
		plan = append(plan, layoutPlanEntry{index: 1})
	}
	return plan
}

// slideLayoutIndex resolves the layout part index a slide's rels point at —
// 1 when the deck carries no plan.
func (w *PPTXWriter) slideLayoutIndex(slide *Slide) int {
	for _, e := range w.layoutPlan {
		if e.layoutPath == slide.layoutPath {
			return e.index
		}
	}
	return 1
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
	w.layoutPlan = w.buildLayoutPlan()

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

	// Write docProps/custom.xml when the presentation defines any
	if err := w.writeCustomProperties(zw); err != nil {
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

	// Write slide master and the layout parts the plan names
	if err := w.writeSlideMaster(zw); err != nil {
		return err
	}

	for _, entry := range w.layoutPlan {
		if err := w.writeSlideLayoutN(zw, entry); err != nil {
			return err
		}
	}

	// Write theme
	if err := w.writeTheme(zw); err != nil {
		return err
	}

	// The reader merged the master/layout decoration into each slide's shape
	// list for rendering. Serialising them as slide content re-parents the
	// layout's own canvas geometry into the slide — the slide parts keep
	// only the slide's own shapes; the deco goes with its layout part.
	for _, slide := range w.presentation.slides {
		kept := make([]Shape, 0, len(slide.shapes))
		for _, sh := range slide.shapes {
			if !sh.base().deco {
				kept = append(kept, sh)
			}
		}
		slide.shapes = kept
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

	// Write SmartArt parts, verbatim from what was read
	if err := w.writeDiagramParts(zw); err != nil {
		return err
	}

	// Write charts, including those nested in groups, in the same document order
	// getChartIndex numbers them in.
	chartIdx := 1
	for _, slide := range w.presentation.slides {
		for _, shape := range flattenShapes(slide.shapes) {
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
