// Package gopresentation provides a pure Go library for reading and writing
// PowerPoint presentation files (.pptx) following the Office Open XML (OOXML) standard.
//
// It is inspired by PHPOffice/PHPPresentation and provides an idiomatic Go API
// for creating, manipulating, and saving presentation documents.
//
// See the Version variable for the current library version.
package gopresentation

import (
	"errors"
	"time"
)

// Presentation represents an in-memory PowerPoint presentation.
type Presentation struct {
	properties             *DocumentProperties
	presentationProperties *PresentationProperties
	slides                 []*Slide
	slideMasters           []*SlideMaster
	activeSlideIndex       int
	layout                 *DocumentLayout
}

// New creates a new Presentation with one default blank slide.
func New() *Presentation {
	p := &Presentation{
		properties:             NewDocumentProperties(),
		presentationProperties: NewPresentationProperties(),
		slides:                 make([]*Slide, 0),
		slideMasters:           make([]*SlideMaster, 0),
		activeSlideIndex:       0,
		layout:                 NewDocumentLayout(),
	}
	// Add a default slide
	p.CreateSlide()
	return p
}

// GetDocumentProperties returns the document properties.
func (p *Presentation) GetDocumentProperties() *DocumentProperties {
	return p.properties
}

// SetDocumentProperties sets the document properties.
func (p *Presentation) SetDocumentProperties(props *DocumentProperties) {
	p.properties = props
}

// GetPresentationProperties returns the presentation properties.
func (p *Presentation) GetPresentationProperties() *PresentationProperties {
	return p.presentationProperties
}

// GetLayout returns the document layout.
func (p *Presentation) GetLayout() *DocumentLayout {
	return p.layout
}

// SetLayout sets the document layout.
func (p *Presentation) SetLayout(layout *DocumentLayout) {
	p.layout = layout
}

// CreateSlide creates a new slide and adds it to the presentation.
func (p *Presentation) CreateSlide() *Slide {
	slide := newSlide()
	p.slides = append(p.slides, slide)
	return slide
}

// AddSlide adds an existing slide to the presentation.
func (p *Presentation) AddSlide(slide *Slide) *Slide {
	p.slides = append(p.slides, slide)
	return slide
}

// GetActiveSlide returns the currently active slide.
func (p *Presentation) GetActiveSlide() *Slide {
	if len(p.slides) == 0 {
		return nil
	}
	if p.activeSlideIndex >= len(p.slides) {
		p.activeSlideIndex = 0
	}
	return p.slides[p.activeSlideIndex]
}

// SetActiveSlideIndex sets the active slide by index.
func (p *Presentation) SetActiveSlideIndex(index int) error {
	if index < 0 || index >= len(p.slides) {
		return errors.New("slide index out of range")
	}
	p.activeSlideIndex = index
	return nil
}

// GetActiveSlideIndex returns the active slide index.
func (p *Presentation) GetActiveSlideIndex() int {
	return p.activeSlideIndex
}

// GetSlide returns a slide by index.
func (p *Presentation) GetSlide(index int) (*Slide, error) {
	if index < 0 || index >= len(p.slides) {
		return nil, errors.New("slide index out of range")
	}
	return p.slides[index], nil
}

// GetAllSlides returns all slides.
func (p *Presentation) GetAllSlides() []*Slide {
	return p.slides
}

// GetSlideCount returns the number of slides.
func (p *Presentation) GetSlideCount() int {
	return len(p.slides)
}

// RemoveSlideByIndex removes a slide by index.
// Returns an error if the index is out of range or if it would remove the last slide.
func (p *Presentation) RemoveSlideByIndex(index int) error {
	if index < 0 || index >= len(p.slides) {
		return errors.New("slide index out of range")
	}
	if len(p.slides) <= 1 {
		return errors.New("cannot remove the last slide")
	}
	p.slides = append(p.slides[:index], p.slides[index+1:]...)
	if p.activeSlideIndex >= len(p.slides) && len(p.slides) > 0 {
		p.activeSlideIndex = len(p.slides) - 1
	}
	return nil
}

// MoveSlide moves a slide from one index to another.
func (p *Presentation) MoveSlide(fromIndex, toIndex int) error {
	if fromIndex < 0 || fromIndex >= len(p.slides) {
		return errors.New("fromIndex out of range")
	}
	if toIndex < 0 || toIndex >= len(p.slides) {
		return errors.New("toIndex out of range")
	}
	if fromIndex == toIndex {
		return nil
	}
	slide := p.slides[fromIndex]
	// Remove from old position
	p.slides = append(p.slides[:fromIndex], p.slides[fromIndex+1:]...)
	// Insert at new position using copy to avoid intermediate allocation
	p.slides = append(p.slides, nil) // grow by one
	copy(p.slides[toIndex+1:], p.slides[toIndex:])
	p.slides[toIndex] = slide
	return nil
}

// GetSlideMasters returns all slide masters.
func (p *Presentation) GetSlideMasters() []*SlideMaster {
	return p.slideMasters
}

// CreateSlideMaster creates a new slide master and adds it.
func (p *Presentation) CreateSlideMaster() *SlideMaster {
	sm := &SlideMaster{}
	p.slideMasters = append(p.slideMasters, sm)
	return sm
}

// DocumentProperties holds standard and custom document properties.
type DocumentProperties struct {
	Creator        string
	LastModifiedBy string
	Created        time.Time
	Modified       time.Time
	Title          string
	Description    string
	Subject        string
	Keywords       string
	Category       string
	Company        string
	Status         string
	Revision       string
	customProps    map[string]*CustomProperty
}

// CustomProperty represents a custom document property.
type CustomProperty struct {
	Name  string
	Value interface{}
	Type  PropertyType
}

// PropertyType represents the type of a custom property.
type PropertyType int

const (
	PropertyTypeString  PropertyType = iota
	PropertyTypeBoolean
	PropertyTypeInteger
	PropertyTypeFloat
	PropertyTypeDate
	PropertyTypeUnknown
)

// NewDocumentProperties creates new document properties with defaults.
func NewDocumentProperties() *DocumentProperties {
	now := time.Now()
	return &DocumentProperties{
		Creator:        "GoPresentation",
		LastModifiedBy: "GoPresentation",
		Created:        now,
		Modified:       now,
		customProps:    make(map[string]*CustomProperty),
	}
}

// SetCustomProperty sets a custom property.
func (dp *DocumentProperties) SetCustomProperty(name string, value interface{}, propType PropertyType) {
	dp.customProps[name] = &CustomProperty{
		Name:  name,
		Value: value,
		Type:  propType,
	}
}

// IsCustomPropertySet checks if a custom property exists.
func (dp *DocumentProperties) IsCustomPropertySet(name string) bool {
	_, ok := dp.customProps[name]
	return ok
}

// GetCustomProperties returns all custom property names.
func (dp *DocumentProperties) GetCustomProperties() []string {
	names := make([]string, 0, len(dp.customProps))
	for name := range dp.customProps {
		names = append(names, name)
	}
	return names
}

// GetCustomPropertyValue returns the value of a custom property.
func (dp *DocumentProperties) GetCustomPropertyValue(name string) interface{} {
	if prop, ok := dp.customProps[name]; ok {
		return prop.Value
	}
	return nil
}

// GetCustomPropertyType returns the type of a custom property.
func (dp *DocumentProperties) GetCustomPropertyType(name string) PropertyType {
	if prop, ok := dp.customProps[name]; ok {
		return prop.Type
	}
	return PropertyTypeUnknown
}
