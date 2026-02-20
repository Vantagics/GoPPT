package gopresentation

import "errors"

// GroupShape represents a group of shapes.
type GroupShape struct {
	BaseShape
	shapes []Shape
	// Child coordinate space (from grpSpPr xfrm chOff/chExt)
	childOffX int64
	childOffY int64
	childExtX int64
	childExtY int64
	// groupFill is the fill defined on the group's grpSpPr, inherited by
	// child shapes that use <a:grpFill/>.
	groupFill *Fill
}

// ShapeTypeGroup is the shape type for groups.
const ShapeTypeGroup ShapeType = 10

func (g *GroupShape) GetType() ShapeType { return ShapeTypeGroup }

// NewGroupShape creates a new group shape.
func NewGroupShape() *GroupShape {
	return &GroupShape{
		shapes: make([]Shape, 0),
	}
}

// AddShape adds a shape to the group.
func (g *GroupShape) AddShape(s Shape) *GroupShape {
	g.shapes = append(g.shapes, s)
	return g
}

// GetShapes returns all shapes in the group.
func (g *GroupShape) GetShapes() []Shape {
	return g.shapes
}
// GetGroupFill returns the group-level fill (from grpSpPr), if any.
func (g *GroupShape) GetGroupFill() *Fill {
	return g.groupFill
}

// GetShapeCount returns the number of shapes in the group.
func (g *GroupShape) GetShapeCount() int {
	return len(g.shapes)
}

// RemoveShape removes a shape by index.
func (g *GroupShape) RemoveShape(index int) error {
	if index < 0 || index >= len(g.shapes) {
		return errOutOfRange
	}
	g.shapes = append(g.shapes[:index], g.shapes[index+1:]...)
	return nil
}

// PlaceholderShape represents a placeholder shape (title, body, etc.).
type PlaceholderShape struct {
	RichTextShape
	phType PlaceholderType
	phIdx  int
}

// ShapeTypePlaceholder is the shape type for placeholders.
const ShapeTypePlaceholder ShapeType = 11

func (p *PlaceholderShape) GetType() ShapeType { return ShapeTypePlaceholder }

// PlaceholderType represents the type of placeholder.
type PlaceholderType string

const (
	PlaceholderTitle    PlaceholderType = "title"
	PlaceholderBody     PlaceholderType = "body"
	PlaceholderCtrTitle PlaceholderType = "ctrTitle"
	PlaceholderSubTitle PlaceholderType = "subTitle"
	PlaceholderDate     PlaceholderType = "dt"
	PlaceholderFooter   PlaceholderType = "ftr"
	PlaceholderSlideNum PlaceholderType = "sldNum"
)

// NewPlaceholderShape creates a new placeholder shape.
func NewPlaceholderShape(phType PlaceholderType) *PlaceholderShape {
	return &PlaceholderShape{
		RichTextShape: *NewRichTextShape(),
		phType:        phType,
	}
}

// GetPlaceholderType returns the placeholder type.
func (p *PlaceholderShape) GetPlaceholderType() PlaceholderType {
	return p.phType
}

// SetPlaceholderIndex sets the placeholder index.
func (p *PlaceholderShape) SetPlaceholderIndex(idx int) {
	p.phIdx = idx
}

// GetPlaceholderIndex returns the placeholder index.
func (p *PlaceholderShape) GetPlaceholderIndex() int {
	return p.phIdx
}

// SetText sets the placeholder text, replacing all existing content with a single paragraph.
func (p *PlaceholderShape) SetText(text string) {
	p.paragraphs = []*Paragraph{NewParagraph()}
	p.paragraphs[0].CreateTextRun(text)
	p.activeParagraph = 0
}

// Clear clears the placeholder content and adds a single empty paragraph.
// An empty paragraph is required by PowerPoint for the file to be valid.
func (p *PlaceholderShape) Clear() {
	p.paragraphs = []*Paragraph{NewParagraph()}
	p.activeParagraph = 0
}

// ClearAll completely removes all paragraphs from the placeholder.
// You must add at least one paragraph via CreateParagraph before saving.
func (p *PlaceholderShape) ClearAll() {
	p.paragraphs = nil
	p.activeParagraph = 0
}

// Remove removes this placeholder from the given slide.
// Returns true if the placeholder was found and removed.
func (p *PlaceholderShape) Remove(slide *Slide) bool {
	return slide.RemoveShapeByPointer(p)
}

// errors
var errOutOfRange = errors.New("index out of range")
