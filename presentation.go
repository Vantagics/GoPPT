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
	// themeColors maps scheme color names (dk1, dk2, lt1, lt2, accent1..accent6,
	// hlink, folHlink) to ARGB hex strings (e.g. "FF000000").
	themeColors map[string]string
	// themeFonts maps the font references a theme declares to real typeface
	// names: "+mj-lt" (major latin), "+mj-ea", "+mj-cs", "+mn-lt", "+mn-ea",
	// "+mn-cs". They are references into theme1.xml's <a:fontScheme>, not names,
	// and a reference stored as a name is a font no file matches.
	themeFonts map[string]string
	// masterTextStyles is the slide master's <p:txStyles>: the font, alignment
	// and indentation a placeholder inherits when neither the slide nor the
	// layout says anything. Read once, because every slide shares it.
	masterTextStyles *masterTextStyles
	// tableStyles holds ppt/tableStyles.xml keyed by style GUID, and
	// defaultTableStyle is the list's def — the fallback for a table that
	// names no style. A cell that declares no fill of its own takes the
	// style's fill for its band.
	tableStyles       map[string]*tableStyle
	defaultTableStyle string
	// masterPlaceholders is the geometry the slide master gives each
	// placeholder. It is the last rung before the renderer's own default and
	// is frequently the only rung that answers at all.
	masterPlaceholders []layoutPlaceholder
	// themeFillStyles / themeLnStyles are the theme's <a:fillStyleLst> and
	// <a:lnStyleLst> bodies. A <p:style> fillRef idx="2" does not name a
	// colour — it names the SECOND of these bodies, whose stops are phClr
	// (whatever colour the reference carries) passed through transforms.
	// Resolving a reference therefore means substituting the colour into the
	// body, not reading a solid from it; treating idx>1 as a solid fill was
	// painting theme-gradient text boxes flat black.
	themeFillStyles []themeFillStyle
	// themeBgFillStyles is the theme's <a:bgFillStyleLst>: a bgRef's
	// 1001-based idx indexes THIS list.
	themeBgFillStyles []themeFillStyle
	themeLnStyles     []themeLnStyle
	// themeEffectStyles is the theme's <a:effectStyleLst>: the outer shadow
	// (if any) an <a:effectRef idx="N"> resolves to. Only the shadow is
	// modelled — the third Office style's camera and bevel have no renderer
	// counterpart, and a reference without a modelled effect falls back to
	// no shadow, which is what shapes without effectRef draw anyway.
	themeEffectStyles []themeEffectStyle
	// masterBackground / masterBgImage are the slide master's <p:bg>: the
	// rung below the layout in the background ladder. A deck that paints its
	// background only on the master (PowerPoint's own themed templates do)
	// used to render white here, because only the layout rung was read.
	masterBackground *Fill
	masterBgImage    *DrawingShape
	// masterDecoShapes are the master's non-placeholder shapes — the colour
	// bands, logos and footers PowerPoint draws on every slide of the deck,
	// beneath anything the layout or the slide adds. Layouts are cached per
	// part path in layoutDecoCache, since a deck's slides share a handful of
	// layouts and each would otherwise reparse the same decoration XML. The
	// cache also keeps the layout's own showMasterSp flag, which decides
	// whether the master's decoration shows through it.
	masterDecoShapes []Shape
	layoutDecoCache  map[string]layoutDeco
	// masterStates holds one masterState per master part path, and
	// defaultTheme is the fallback theme parsed from theme1.xml/theme2.xml
	// for slides that reach no master with a readable theme. A deck carries
	// one theme per master — Office writes seven of each without blinking —
	// so the flat fields above are the ACTIVE rung's values, re-pointed as
	// each slide is read; by the time the renderer runs, every reference has
	// already resolved to concrete colours and fonts.
	masterStates map[string]*masterState
	defaultTheme *themeState
	// masterRead records that at least one master was looked for, so a deck
	// whose masters carry no placeholders is not treated as master-less.
	masterRead bool
}

// themeState is one theme part's parsed content: the colour scheme, the font
// scheme and the fill/line/effect style lists a shape's <p:style> references
// resolve against.
type themeState struct {
	colors       map[string]string
	fonts        map[string]string
	fillStyles   []themeFillStyle
	bgFillStyles []themeFillStyle
	lnStyles     []themeLnStyle
	effectStyles []themeEffectStyle
}

// layoutDeco is one slide layout's cached decoration: its non-placeholder
// shapes and whether its root sets showMasterSp="0" — a layout's flag keeps
// only the master's graphics off, while a slide's silences its layout's too.
type layoutDeco struct {
	shapes       []Shape
	hideMasterSp bool
}

// masterState is one slide master's parsed contribution: its placeholder
// geometry, its text styles, its background rung, its decoration shapes and
// the theme part its rels name.
type masterState struct {
	placeholders []layoutPlaceholder
	textStyles   *masterTextStyles
	background   *Fill
	bgImage      *DrawingShape
	decoShapes   []Shape
	theme        *themeState
}

// applyThemeState re-points the active theme fields. All consumers read them
// while a slide is being parsed, never at render time, so swapping the
// pointers between slides is safe.
func (p *Presentation) applyThemeState(ts *themeState) {
	if ts == nil {
		return
	}
	p.themeColors = ts.colors
	p.themeFonts = ts.fonts
	p.themeFillStyles = ts.fillStyles
	p.themeBgFillStyles = ts.bgFillStyles
	p.themeLnStyles = ts.lnStyles
	p.themeEffectStyles = ts.effectStyles
}

// activeTheme returns the theme state the presentation resolves against —
// the flat fields are the live copy, re-pointed as slides are read. The
// colour table is nil when nothing has been read yet; callers check.
func (p *Presentation) activeTheme() *themeState {
	return &themeState{
		colors:       p.themeColors,
		fonts:        p.themeFonts,
		fillStyles:   p.themeFillStyles,
		bgFillStyles: p.themeBgFillStyles,
		lnStyles:     p.themeLnStyles,
		effectStyles: p.themeEffectStyles,
	}
}

// applyMasterState re-points the active master rung, and with it the theme
// that master names. A master without a readable theme falls back to the
// default one: carrying the previous slide's theme across would tint the
// deck in whichever master happened to parse last.
func (p *Presentation) applyMasterState(st *masterState) {
	if st == nil {
		return
	}
	p.masterPlaceholders = st.placeholders
	p.masterTextStyles = st.textStyles
	p.masterBackground = st.background
	p.masterBgImage = st.bgImage
	p.masterDecoShapes = st.decoShapes
	if st.theme != nil {
		p.applyThemeState(st.theme)
	} else {
		p.applyThemeState(p.defaultTheme)
	}
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
//
// GetSlide never panics, including when called on a nil *Presentation (which
// can result from a failed Open whose error the caller ignored).
func (p *Presentation) GetSlide(index int) (slide *Slide, err error) {
	defer recoverToError(&err, "Presentation.GetSlide")
	if p == nil {
		return nil, errors.New("presentation is nil")
	}
	if index < 0 || index >= len(p.slides) {
		return nil, errors.New("slide index out of range")
	}
	return p.slides[index], nil
}

// GetAllSlides returns all slides.
func (p *Presentation) GetAllSlides() []*Slide {
	if p == nil {
		return nil
	}
	return p.slides
}

// GetSlideCount returns the number of slides.
func (p *Presentation) GetSlideCount() int {
	if p == nil {
		return 0
	}
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
	PropertyTypeString PropertyType = iota
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
//
// A zero-valued DocumentProperties — one a caller built rather than got from
// NewDocumentProperties — has a nil map, and assigning into a nil map panics.
func (dp *DocumentProperties) SetCustomProperty(name string, value interface{}, propType PropertyType) {
	if dp.customProps == nil {
		dp.customProps = make(map[string]*CustomProperty)
	}
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

// MasterDecoShapes exposes the master's decoration shapes for diagnostics.
func (p *Presentation) MasterDecoShapes() []Shape { return p.masterDecoShapes }
