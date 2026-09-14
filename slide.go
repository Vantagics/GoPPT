package gopresentation

import "errors"

// Transition represents a slide transition.
//
// It maps onto <p:transition>, the child of <p:sld> that says how the
// presentation moves into this slide. Only the effect elements of
// pml-animationInfo.xsd are modelled; the PowerPoint 2010 transitions that live
// in the p14 namespace (p14:flash, p14:ripple, p14:vortex and the rest) are a
// separate content model and are not represented here.
type Transition struct {
	Type     TransitionType
	Speed    TransitionSpeed
	Duration int // milliseconds; see the field's note under Transition

	// Direction is the dir attribute of the effects that take one. Which
	// tokens are legal depends on Type: push and wipe take l/r/u/d; cover, pull
	// and strips take ld/lu/rd/ru/l/r/u/d; blinds, checker, comb and randomBar
	// take horz/vert; split and zoom take in/out. The writer emits the token as
	// given and omits the attribute for the effects that have no dir at all.
	Direction TransitionDirection

	// Orientation is split's orient attribute: horz or vert. Other effects
	// state their orientation through Direction instead, because their
	// attribute happens to be called dir.
	Orientation TransitionOrientation

	// ThroughBlack is thruBlk on cut and fade, which holds the previous slide
	// on screen while the new one comes in through black. False is the schema
	// default and is therefore not written.
	ThroughBlack bool

	// Spokes is wheel's spoke count. Zero means the schema default of four.
	Spokes int

	// AdvanceOnClick is advClick. Nil omits the attribute, and omitting it
	// means the same as true, so only a false is worth stating.
	AdvanceOnClick *bool

	// AdvanceAfterTime is advTm: the delay in milliseconds before the
	// transition starts. Zero omits it, which means no auto-advance.
	AdvanceAfterTime int
}

// Duration is written as p14:dur, which has no home in the base schema's
// CT_SlideTransition — so a transition that carries one is wrapped in
// mc:AlternateContent, with the plain element repeated in mc:Fallback for
// consumers that do not know the p14 namespace.

// TransitionType represents the type of slide transition.
//
// The constants below TransitionDissolve were added when the writer learned to
// emit <p:transition>; the earlier eight keep the numeric values they have
// always had, so nothing that stored one as an integer changes meaning.
type TransitionType int

const (
	TransitionNone TransitionType = iota
	TransitionFade
	TransitionPush
	TransitionWipe
	TransitionSplit
	TransitionCover
	TransitionUncover
	TransitionDissolve

	TransitionBlinds
	TransitionChecker
	TransitionCircle
	TransitionComb
	TransitionCut
	TransitionDiamond
	TransitionNewsflash
	TransitionPlus
	TransitionRandom
	TransitionRandomBar
	TransitionStrips
	TransitionWedge
	TransitionWheel
	TransitionZoom
)

// TransitionPull is the name the format uses for what TransitionUncover was
// named after. The schema has no <p:uncover>: the element that moves the
// previous slide off-screen, continually revealing more of the new one, is
// <p:pull>, and it is the eight-direction counterpart of <p:cover> — same
// CT_EightDirectionTransition, same l default. The two constants are therefore
// one value, and either name round-trips.
const TransitionPull = TransitionUncover

// TransitionDirection is a transition effect's dir attribute. The token is the
// XML token, so the type doubles as the string the attribute is written with.
type TransitionDirection string

const (
	TransitionDirectionLeft       TransitionDirection = "l"
	TransitionDirectionRight      TransitionDirection = "r"
	TransitionDirectionUp         TransitionDirection = "u"
	TransitionDirectionDown       TransitionDirection = "d"
	TransitionDirectionLeftDown   TransitionDirection = "ld"
	TransitionDirectionLeftUp     TransitionDirection = "lu"
	TransitionDirectionRightDown  TransitionDirection = "rd"
	TransitionDirectionRightUp    TransitionDirection = "ru"
	TransitionDirectionHorizontal TransitionDirection = "horz"
	TransitionDirectionVertical   TransitionDirection = "vert"
	TransitionDirectionIn         TransitionDirection = "in"
	TransitionDirectionOut        TransitionDirection = "out"
)

// TransitionOrientation is split's orient attribute.
type TransitionOrientation string

const (
	TransitionOrientationHorizontal TransitionOrientation = "horz"
	TransitionOrientationVertical   TransitionOrientation = "vert"
)

// TransitionSpeed represents the speed of a transition.
type TransitionSpeed string

const (
	TransitionSpeedSlow   TransitionSpeed = "slow"
	TransitionSpeedMedium TransitionSpeed = "med"
	TransitionSpeedFast   TransitionSpeed = "fast"
)

// transitionDirSet is the value set an effect's dir attribute accepts. The sets
// differ per effect, which is why the reader keeps the raw token: a document can
// hold a combination this writer would not have chosen, and dropping it would
// lose the transition.
type transitionDirSet int

const (
	dirNone        transitionDirSet = iota
	dirOrientation                  // horz, vert
	dirSide                         // l, r, u, d
	dirEight                        // ld, lu, rd, ru, l, r, u, d
	dirCorner                       // ld, lu, rd, ru
	dirInOut                        // in, out
)

// transitionEffect is one effect element of CT_SlideTransition: its name, and
// the value set its dir attribute accepts.
type transitionEffect struct {
	name string
	dir  transitionDirSet
}

// transitionEffects maps each type to its element. The names and types are the
// ones in pml-animationInfo.xsd's CT_SlideTransition. TransitionUncover is absent
// because it and TransitionPull are one value.
var transitionEffects = map[TransitionType]transitionEffect{
	TransitionBlinds:    {"blinds", dirOrientation},
	TransitionChecker:   {"checker", dirOrientation},
	TransitionCircle:    {"circle", dirNone},
	TransitionComb:      {"comb", dirOrientation},
	TransitionCover:     {"cover", dirEight},
	TransitionCut:       {"cut", dirNone},
	TransitionDiamond:   {"diamond", dirNone},
	TransitionDissolve:  {"dissolve", dirNone},
	TransitionFade:      {"fade", dirNone},
	TransitionNewsflash: {"newsflash", dirNone},
	TransitionPlus:      {"plus", dirNone},
	TransitionPull:      {"pull", dirEight},
	TransitionPush:      {"push", dirSide},
	TransitionRandom:    {"random", dirNone},
	TransitionRandomBar: {"randomBar", dirOrientation},
	TransitionSplit:     {"split", dirInOut},
	TransitionStrips:    {"strips", dirCorner},
	TransitionWedge:     {"wedge", dirNone},
	TransitionWheel:     {"wheel", dirNone},
	TransitionWipe:      {"wipe", dirSide},
	TransitionZoom:      {"zoom", dirInOut},
}

// transitionElement returns the effect element for a type, and whether the type
// is one this package models. TransitionNone is not.
func transitionElement(t TransitionType) (transitionEffect, bool) {
	e, ok := transitionEffects[t]
	return e, ok
}

// transitionTypeForElement is the inverse of transitionEffects, for the reader.
func transitionTypeForElement(name string) (TransitionType, bool) {
	for t, e := range transitionEffects {
		if e.name == name {
			return t, true
		}
	}
	return TransitionNone, false
}

// Slide represents a single slide in a presentation.
type Slide struct {
	shapes     []Shape
	name       string
	notes      string
	transition *Transition
	visible    bool
	comments   []*Comment
	animations []*Animation
	background *Fill
}

// newSlide creates a new empty slide.
func newSlide() *Slide {
	return &Slide{
		shapes:     make([]Shape, 0),
		visible:    true,
		comments:   make([]*Comment, 0),
		animations: make([]*Animation, 0),
	}
}

// GetName returns the slide name.
func (s *Slide) GetName() string {
	return s.name
}

// SetName sets the slide name.
func (s *Slide) SetName(name string) {
	s.name = name
}

// GetNotes returns the slide notes.
func (s *Slide) GetNotes() string {
	return s.notes
}

// SetNotes sets the slide notes.
func (s *Slide) SetNotes(notes string) {
	s.notes = notes
}

// IsVisible returns whether the slide is visible.
func (s *Slide) IsVisible() bool {
	return s.visible
}

// SetVisible sets the slide visibility.
func (s *Slide) SetVisible(visible bool) {
	s.visible = visible
}

// GetTransition returns the slide transition.
func (s *Slide) GetTransition() *Transition {
	return s.transition
}

// SetTransition sets the slide transition.
func (s *Slide) SetTransition(t *Transition) {
	s.transition = t
}

// GetShapes returns all shapes on the slide.
func (s *Slide) GetShapes() []Shape {
	return s.shapes
}

// AddShape adds a shape to the slide.
func (s *Slide) AddShape(shape Shape) {
	s.shapes = append(s.shapes, shape)
}

// RemoveShape removes a shape by index.
func (s *Slide) RemoveShape(index int) error {
	if index < 0 || index >= len(s.shapes) {
		return errors.New("shape index out of range")
	}
	s.shapes = append(s.shapes[:index], s.shapes[index+1:]...)
	return nil
}

// CreateRichTextShape creates a new rich text shape and adds it to the slide.
func (s *Slide) CreateRichTextShape() *RichTextShape {
	shape := NewRichTextShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateDrawingShape creates a new drawing (image) shape and adds it to the slide.
func (s *Slide) CreateDrawingShape() *DrawingShape {
	shape := NewDrawingShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// AddImage creates a drawing shape from a file path and adds it to the slide.
// This is a convenience method matching unioffice's Slide.AddImage pattern.
func (s *Slide) AddImage(path string) (*DrawingShape, error) {
	shape := NewDrawingShape()
	if err := shape.SetImageFromFile(path); err != nil {
		return nil, err
	}
	s.shapes = append(s.shapes, shape)
	return shape, nil
}

// AddImageData creates a drawing shape from raw image data and adds it to the slide.
func (s *Slide) AddImageData(data []byte, mimeType string) *DrawingShape {
	shape := NewDrawingShape()
	shape.SetImageData(data, mimeType)
	s.shapes = append(s.shapes, shape)
	return shape
}

// AddTextBox creates a new rich text shape and adds it to the slide.
// This is an alias for CreateRichTextShape matching unioffice naming.
func (s *Slide) AddTextBox() *RichTextShape {
	return s.CreateRichTextShape()
}

// AddAutoShape creates a new auto shape and adds it to the slide.
// This matches unioffice's Slide.AddShape pattern.
func (s *Slide) AddAutoShape() *AutoShape {
	return s.CreateAutoShape()
}

// AddTable creates a new table shape and adds it to the slide.
// This matches unioffice's Slide.AddTable pattern.
func (s *Slide) AddTable(rows, cols int) *TableShape {
	return s.CreateTableShape(rows, cols)
}

// CreateTableShape creates a new table shape and adds it to the slide.
func (s *Slide) CreateTableShape(rows, cols int) *TableShape {
	shape := NewTableShape(rows, cols)
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateAutoShape creates a new auto shape and adds it to the slide.
func (s *Slide) CreateAutoShape() *AutoShape {
	shape := NewAutoShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateLineShape creates a new line shape and adds it to the slide.
func (s *Slide) CreateLineShape() *LineShape {
	shape := NewLineShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateChartShape creates a new chart shape and adds it to the slide.
func (s *Slide) CreateChartShape() *ChartShape {
	shape := NewChartShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreateGroupShape creates a new group shape and adds it to the slide.
func (s *Slide) CreateGroupShape() *GroupShape {
	shape := NewGroupShape()
	s.shapes = append(s.shapes, shape)
	return shape
}

// CreatePlaceholderShape creates a new placeholder shape and adds it to the slide.
func (s *Slide) CreatePlaceholderShape(phType PlaceholderType) *PlaceholderShape {
	shape := NewPlaceholderShape(phType)
	s.shapes = append(s.shapes, shape)
	return shape
}

// --- Comments ---

// AddComment adds a comment to the slide.
func (s *Slide) AddComment(c *Comment) {
	s.comments = append(s.comments, c)
}

// GetComments returns all comments on the slide.
func (s *Slide) GetComments() []*Comment {
	return s.comments
}

// GetCommentCount returns the number of comments.
func (s *Slide) GetCommentCount() int {
	return len(s.comments)
}

// --- Animations ---

// AddAnimation adds an animation to the slide.
func (s *Slide) AddAnimation(a *Animation) {
	s.animations = append(s.animations, a)
}

// GetAnimations returns all animations on the slide.
func (s *Slide) GetAnimations() []*Animation {
	return s.animations
}

// --- Background ---

// SetBackground sets the slide background fill.
func (s *Slide) SetBackground(f *Fill) {
	s.background = f
}

// GetBackground returns the slide background fill.
func (s *Slide) GetBackground() *Fill {
	return s.background
}

// --- Placeholder access ---

// GetPlaceholder returns the first placeholder of the given type.
// Returns nil if no placeholder of that type exists.
func (s *Slide) GetPlaceholder(phType PlaceholderType) *PlaceholderShape {
	for _, shape := range s.shapes {
		if ph, ok := shape.(*PlaceholderShape); ok {
			if ph.phType == phType {
				return ph
			}
		}
	}
	return nil
}

// GetPlaceholderByIndex returns a placeholder by its index.
// Returns nil if no placeholder with that index exists.
func (s *Slide) GetPlaceholderByIndex(idx int) *PlaceholderShape {
	for _, shape := range s.shapes {
		if ph, ok := shape.(*PlaceholderShape); ok {
			if ph.phIdx == idx {
				return ph
			}
		}
	}
	return nil
}

// GetPlaceholders returns all placeholder shapes on the slide.
func (s *Slide) GetPlaceholders() []*PlaceholderShape {
	var phs []*PlaceholderShape
	for _, shape := range s.shapes {
		if ph, ok := shape.(*PlaceholderShape); ok {
			phs = append(phs, ph)
		}
	}
	return phs
}

// GetShapeCount returns the number of shapes on the slide.
func (s *Slide) GetShapeCount() int {
	return len(s.shapes)
}

// ExtractText returns all text content from this slide as a single string.
func (s *Slide) ExtractText() string {
	var parts []string
	for _, shape := range s.shapes {
		switch sh := shape.(type) {
		case *RichTextShape:
			parts = append(parts, extractParagraphsText(sh.paragraphs)...)
		case *PlaceholderShape:
			parts = append(parts, extractParagraphsText(sh.paragraphs)...)
		case *AutoShape:
			if sh.text != "" {
				parts = append(parts, sh.text)
			}
		case *TableShape:
			for _, row := range sh.rows {
				for _, cell := range row {
					parts = append(parts, extractParagraphsText(cell.paragraphs)...)
				}
			}
		case *GroupShape:
			for _, gs := range sh.shapes {
				switch gsh := gs.(type) {
				case *RichTextShape:
					parts = append(parts, extractParagraphsText(gsh.paragraphs)...)
				case *PlaceholderShape:
					parts = append(parts, extractParagraphsText(gsh.paragraphs)...)
				}
			}
		}
	}
	return joinNonEmpty(parts, "\n")
}

// GetTextBoxes returns all RichTextShape (text box) shapes on the slide.
func (s *Slide) GetTextBoxes() []*RichTextShape {
	var boxes []*RichTextShape
	for _, shape := range s.shapes {
		if tb, ok := shape.(*RichTextShape); ok {
			boxes = append(boxes, tb)
		}
	}
	return boxes
}

// RemoveShapeByPointer removes a specific shape from the slide.
func (s *Slide) RemoveShapeByPointer(target Shape) bool {
	for i, shape := range s.shapes {
		if shape == target {
			s.shapes = append(s.shapes[:i], s.shapes[i+1:]...)
			return true
		}
	}
	return false
}
