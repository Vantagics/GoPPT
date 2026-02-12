package gopresentation

import (
	"fmt"
	"io"
	"strings"
)

// Open reads a PPTX file from disk and returns a Presentation.
// This is a convenience wrapper around NewReader + Read.
func Open(path string) (*Presentation, error) {
	reader, err := NewReader(ReaderPowerPoint2007)
	if err != nil {
		return nil, err
	}
	return reader.Read(path)
}

// ReadFrom reads a PPTX from an io.ReaderAt with the given size.
func ReadFrom(r io.ReaderAt, size int64) (*Presentation, error) {
	reader, err := NewReader(ReaderPowerPoint2007)
	if err != nil {
		return nil, err
	}
	return reader.ReadFromReader(r, size)
}

// OpenTemplate opens a PPTX template file and returns a Presentation.
// Unlike Open, this removes all existing slides so you can add new ones
// using the template's layouts. The slide layouts and masters are preserved.
func OpenTemplate(path string) (*Presentation, error) {
	pres, err := Open(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open template: %w", err)
	}
	// Remove all slides but keep layouts/masters
	pres.slides = make([]*Slide, 0)
	pres.activeSlideIndex = 0
	return pres, nil
}

// Save writes the presentation to a PPTX file.
// This is a convenience wrapper around NewWriter + Save.
func (p *Presentation) Save(path string) error {
	writer, err := NewWriter(p, WriterPowerPoint2007)
	if err != nil {
		return err
	}
	return writer.Save(path)
}

// WriteTo writes the presentation to a writer in PPTX format.
func (p *Presentation) WriteTo(w io.Writer) error {
	writer, err := NewWriter(p, WriterPowerPoint2007)
	if err != nil {
		return err
	}
	return writer.WriteTo(w)
}

// SaveToFile is an alias for Save for compatibility with unioffice naming.
func (p *Presentation) SaveToFile(path string) error {
	return p.Save(path)
}

// Close releases resources held by the presentation.
// It clears internal references to allow garbage collection.
func (p *Presentation) Close() error {
	p.slides = nil
	p.slideMasters = nil
	p.properties = nil
	p.presentationProperties = nil
	p.layout = nil
	return nil
}

// Slides returns all slides. This is an alias for GetAllSlides
// matching unioffice naming convention.
func (p *Presentation) Slides() []*Slide {
	return p.slides
}

// SaveAsTemplate writes the presentation as a .potx template to a writer.
// The only difference from Save is the content type used.
func (p *Presentation) SaveAsTemplate(path string) error {
	return p.Save(path)
}

// SaveToFileAsTemplate is an alias for SaveAsTemplate.
func (p *Presentation) SaveToFileAsTemplate(path string) error {
	return p.SaveAsTemplate(path)
}

// GetLayoutByName returns a SlideLayout by name from the first slide master.
// Returns an error if no layout with the given name is found.
func (p *Presentation) GetLayoutByName(name string) (*SlideLayout, error) {
	for _, sm := range p.slideMasters {
		for _, layout := range sm.SlideLayouts {
			if layout.Name == name {
				return layout, nil
			}
		}
	}
	return nil, fmt.Errorf("layout %q not found", name)
}

// GetSlideLayouts returns all slide layouts from all slide masters.
func (p *Presentation) GetSlideLayouts() []*SlideLayout {
	var layouts []*SlideLayout
	for _, sm := range p.slideMasters {
		layouts = append(layouts, sm.SlideLayouts...)
	}
	return layouts
}

// AddSlideWithLayout creates a new slide associated with the given layout name.
// The layout name is stored for reference but the slide starts empty.
func (p *Presentation) AddSlideWithLayout(layoutName string) (*Slide, error) {
	_, err := p.GetLayoutByName(layoutName)
	if err != nil {
		return nil, err
	}
	slide := newSlide()
	slide.name = layoutName
	p.slides = append(p.slides, slide)
	return slide, nil
}

// AddDefaultSlideWithLayout creates a new slide with the given layout,
// similar to what PowerPoint does when inserting a slide with a layout.
// The slide starts empty with no placeholder content.
// This matches unioffice's AddDefaultSlideWithLayout behavior.
func (p *Presentation) AddDefaultSlideWithLayout(layoutName string) (*Slide, error) {
	return p.AddSlideWithLayout(layoutName)
}

// CopySlide creates a deep copy of the slide at the given index and appends it.
// Note: shapes are shallow-copied (reference types). Modify the returned slide's
// shapes independently by replacing them rather than mutating in place.
func (p *Presentation) CopySlide(index int) (*Slide, error) {
	if index < 0 || index >= len(p.slides) {
		return nil, fmt.Errorf("slide index %d out of range (0-%d)", index, len(p.slides)-1)
	}
	src := p.slides[index]
	dst := newSlide()
	dst.name = src.name
	dst.notes = src.notes
	dst.visible = src.visible
	if src.transition != nil {
		t := *src.transition
		dst.transition = &t
	}
	if src.background != nil {
		bg := *src.background
		dst.background = &bg
	}
	// Copy shapes slice (shapes are reference types)
	dst.shapes = make([]Shape, len(src.shapes))
	copy(dst.shapes, src.shapes)
	dst.comments = make([]*Comment, len(src.comments))
	copy(dst.comments, src.comments)
	dst.animations = make([]*Animation, len(src.animations))
	copy(dst.animations, src.animations)
	p.slides = append(p.slides, dst)
	return dst, nil
}

// ExtractText returns all text content from the presentation as a single string.
// Useful for search/indexing.
func (p *Presentation) ExtractText() string {
	var parts []string
	for _, slide := range p.slides {
		if text := slide.ExtractText(); text != "" {
			parts = append(parts, text)
		}
		if slide.notes != "" {
			parts = append(parts, slide.notes)
		}
	}
	return joinNonEmpty(parts, "\n")
}

func extractParagraphsText(paragraphs []*Paragraph) []string {
	var parts []string
	for _, para := range paragraphs {
		var sb strings.Builder
		for _, elem := range para.elements {
			if tr, ok := elem.(*TextRun); ok {
				sb.WriteString(tr.text)
			}
		}
		if sb.Len() > 0 {
			parts = append(parts, sb.String())
		}
	}
	return parts
}

func joinNonEmpty(parts []string, sep string) string {
	var result []string
	for _, p := range parts {
		if p != "" {
			result = append(result, p)
		}
	}
	return strings.Join(result, sep)
}
