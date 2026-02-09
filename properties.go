package gopresentation

// PresentationProperties holds presentation-level properties.
type PresentationProperties struct {
	zoom           float64
	lastView       ViewType
	slideshowType  SlideshowType
	commentVisible bool
	markedAsFinal  bool
	thumbnailPath  string
	thumbnailData  []byte
}

// ViewType represents the last view type.
type ViewType int

const (
	ViewSlide         ViewType = iota
	ViewNotes
	ViewHandout
	ViewOutline
	ViewSlideMaster
	ViewSlideSorter
)

// SlideshowType represents the slideshow type.
type SlideshowType int

const (
	SlideshowTypePresent SlideshowType = iota
	SlideshowTypeBrowse
	SlideshowTypeKiosk
)

// NewPresentationProperties creates new presentation properties with defaults.
func NewPresentationProperties() *PresentationProperties {
	return &PresentationProperties{
		zoom:           1.0,
		lastView:       ViewSlide,
		slideshowType:  SlideshowTypePresent,
		commentVisible: false,
		markedAsFinal:  false,
	}
}

// GetZoom returns the zoom level.
func (pp *PresentationProperties) GetZoom() float64 {
	return pp.zoom
}

// SetZoom sets the zoom level (clamped to 0.1–4.0).
func (pp *PresentationProperties) SetZoom(zoom float64) {
	if zoom < 0.1 {
		zoom = 0.1
	}
	if zoom > 4.0 {
		zoom = 4.0
	}
	pp.zoom = zoom
}

// GetLastView returns the last view type.
func (pp *PresentationProperties) GetLastView() ViewType {
	return pp.lastView
}

// SetLastView sets the last view type.
func (pp *PresentationProperties) SetLastView(view ViewType) {
	pp.lastView = view
}

// GetSlideshowType returns the slideshow type.
func (pp *PresentationProperties) GetSlideshowType() SlideshowType {
	return pp.slideshowType
}

// SetSlideshowType sets the slideshow type.
func (pp *PresentationProperties) SetSlideshowType(t SlideshowType) {
	pp.slideshowType = t
}

// IsCommentVisible returns whether comments are visible.
func (pp *PresentationProperties) IsCommentVisible() bool {
	return pp.commentVisible
}

// SetCommentVisible sets comment visibility.
func (pp *PresentationProperties) SetCommentVisible(visible bool) {
	pp.commentVisible = visible
}

// IsMarkedAsFinal returns whether the presentation is marked as final.
func (pp *PresentationProperties) IsMarkedAsFinal() bool {
	return pp.markedAsFinal
}

// MarkAsFinal marks the presentation as final.
func (pp *PresentationProperties) MarkAsFinal(final ...bool) {
	if len(final) == 0 {
		pp.markedAsFinal = true
		return
	}
	pp.markedAsFinal = final[0]
}

// SetThumbnailPath sets the thumbnail from a file path.
func (pp *PresentationProperties) SetThumbnailPath(path string) {
	pp.thumbnailPath = path
}

// GetThumbnailPath returns the thumbnail path.
func (pp *PresentationProperties) GetThumbnailPath() string {
	return pp.thumbnailPath
}

// SetThumbnailData sets the thumbnail from raw data.
func (pp *PresentationProperties) SetThumbnailData(data []byte) {
	pp.thumbnailData = data
}

// GetThumbnailData returns the thumbnail data.
func (pp *PresentationProperties) GetThumbnailData() []byte {
	return pp.thumbnailData
}

// DocumentLayout represents the slide dimensions.
type DocumentLayout struct {
	CX   int64 // width in EMU (English Metric Units)
	CY   int64 // height in EMU
	Name string
}

// Standard layout constants (in EMU: 1 inch = 914400 EMU).
const (
	LayoutScreen4x3  = "screen4x3"
	LayoutScreen16x9 = "screen16x9"
	LayoutScreen16x10 = "screen16x10"
	LayoutA4          = "A4"
	LayoutLetter      = "letter"
	LayoutCustom      = "custom"
)

// NewDocumentLayout creates a default 4:3 layout.
func NewDocumentLayout() *DocumentLayout {
	return &DocumentLayout{
		CX:   9144000,  // 10 inches
		CY:   6858000,  // 7.5 inches
		Name: LayoutScreen4x3,
	}
}

// SetLayout sets a predefined layout.
func (dl *DocumentLayout) SetLayout(name string) {
	dl.Name = name
	switch name {
	case LayoutScreen4x3:
		dl.CX = 9144000
		dl.CY = 6858000
	case LayoutScreen16x9:
		dl.CX = 12192000
		dl.CY = 6858000
	case LayoutScreen16x10:
		dl.CX = 10972800
		dl.CY = 6858000
	case LayoutA4:
		dl.CX = 9906000
		dl.CY = 6858000
	case LayoutLetter:
		dl.CX = 9144000
		dl.CY = 6858000
	}
}

// SetCustomLayout sets custom dimensions in EMU. Both values must be positive.
func (dl *DocumentLayout) SetCustomLayout(cx, cy int64) {
	if cx <= 0 {
		cx = 9144000 // default 10 inches
	}
	if cy <= 0 {
		cy = 6858000 // default 7.5 inches
	}
	dl.CX = cx
	dl.CY = cy
	dl.Name = LayoutCustom
}

// SlideMaster represents a slide master.
type SlideMaster struct {
	Name         string
	SlideLayouts []*SlideLayout
}

// SlideLayout represents a slide layout.
type SlideLayout struct {
	Name string
	Type string
}
