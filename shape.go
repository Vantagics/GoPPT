package gopresentation

import (
	"fmt"
	"os"
	"strings"
)

// Shape is the interface that all shapes implement.
type Shape interface {
	GetType() ShapeType
	GetOffsetX() int64
	GetOffsetY() int64
	GetWidth() int64
	GetHeight() int64
	GetName() string
	GetRotation() int
	// base returns the underlying BaseShape (unexported, internal use only).
	base() *BaseShape
}

// ShapeType represents the type of shape.
type ShapeType int

const (
	ShapeTypeRichText ShapeType = iota
	ShapeTypeDrawing
	ShapeTypeTable
	ShapeTypeAutoShape
	ShapeTypeLine
	ShapeTypeChart
)

// BaseShape contains common shape properties.
type BaseShape struct {
	name           string
	description    string
	offsetX        int64 // in EMU
	offsetY        int64 // in EMU
	width          int64 // in EMU
	height         int64 // in EMU
	rotation       int   // in degrees
	flipHorizontal bool
	flipVertical   bool
	fill           *Fill
	border         *Border
	shadow         *Shadow
	hyperlink      *Hyperlink
}

func (b *BaseShape) GetOffsetX() int64 { return b.offsetX }
func (b *BaseShape) GetOffsetY() int64 { return b.offsetY }
func (b *BaseShape) GetWidth() int64   { return b.width }
func (b *BaseShape) GetHeight() int64  { return b.height }
func (b *BaseShape) GetName() string   { return b.name }
func (b *BaseShape) GetRotation() int  { return b.rotation }
func (b *BaseShape) base() *BaseShape  { return b }

func (b *BaseShape) SetOffsetX(x int64) *BaseShape { b.offsetX = x; return b }
func (b *BaseShape) SetOffsetY(y int64) *BaseShape { b.offsetY = y; return b }
func (b *BaseShape) SetWidth(w int64) *BaseShape   { b.width = w; return b }
func (b *BaseShape) SetHeight(h int64) *BaseShape  { b.height = h; return b }
func (b *BaseShape) SetName(n string) *BaseShape   { b.name = n; return b }
func (b *BaseShape) SetRotation(r int) *BaseShape  { b.rotation = ((r % 360) + 360) % 360; return b }

// SetPosition sets both offset X and Y in EMU.
func (b *BaseShape) SetPosition(x, y int64) *BaseShape {
	b.offsetX = x
	b.offsetY = y
	return b
}

// SetSize sets both width and height in EMU.
func (b *BaseShape) SetSize(w, h int64) *BaseShape {
	b.width = w
	b.height = h
	return b
}

// SetFlipHorizontal controls horizontal flipping.
func (b *BaseShape) SetFlipHorizontal(flip bool) *BaseShape {
	b.flipHorizontal = flip
	return b
}

// GetFlipHorizontal returns whether the shape is flipped horizontally.
func (b *BaseShape) GetFlipHorizontal() bool { return b.flipHorizontal }

// SetFlipVertical controls vertical flipping.
func (b *BaseShape) SetFlipVertical(flip bool) *BaseShape {
	b.flipVertical = flip
	return b
}

// GetFlipVertical returns whether the shape is flipped vertically.
func (b *BaseShape) GetFlipVertical() bool { return b.flipVertical }

func (b *BaseShape) GetDescription() string  { return b.description }
func (b *BaseShape) SetDescription(d string) { b.description = d }

func (b *BaseShape) GetFill() *Fill {
	if b.fill == nil {
		b.fill = NewFill()
	}
	return b.fill
}

func (b *BaseShape) SetFill(f *Fill) { b.fill = f }

func (b *BaseShape) GetBorder() *Border {
	if b.border == nil {
		b.border = NewBorder()
	}
	return b.border
}

func (b *BaseShape) SetBorder(border *Border) { b.border = border }

func (b *BaseShape) GetShadow() *Shadow {
	if b.shadow == nil {
		b.shadow = NewShadow()
	}
	return b.shadow
}

func (b *BaseShape) SetShadow(s *Shadow) { b.shadow = s }

func (b *BaseShape) GetHyperlink() *Hyperlink  { return b.hyperlink }
func (b *BaseShape) SetHyperlink(h *Hyperlink) { b.hyperlink = h }

// CustomGeomPath represents a custom geometry path for freeform shapes.
type CustomGeomPath struct {
	Width    int64         // path coordinate space width
	Height   int64         // path coordinate space height
	Commands []PathCommand // path commands (moveTo, lineTo, close, etc.)
}

// PathCommand represents a single path command.
type PathCommand struct {
	Type string // "moveTo", "lnTo", "close", "cubicBezTo", "quadBezTo", "arcTo"
	Pts  []PathPoint
	// Arc parameters (only for arcTo): radii and angles in OOXML 60000ths of a degree
	WR, HR         int64 // ellipse radii in path coordinate units
	StAng, SwAng   int64 // start angle and sweep angle (60000ths of a degree)
}

// PathPoint represents a point in path coordinates.
type PathPoint struct {
	X, Y int64
}

// RichTextShape represents a rich text shape.
type RichTextShape struct {
	BaseShape
	paragraphs      []*Paragraph
	activeParagraph int
	autoFit         AutoFitType
	fontScale       int // normAutofit fontScale in thousandths of a percent (e.g. 62500 = 62.5%), 0 means 100%
	wordWrap        bool
	verticalAlign   VerticalAlignment
	textAnchor      TextAnchorType
	textDirection   string // "horz", "vert", "vert270", "eaVert", etc.
	columns         int
	columnSpacing   int64
	// Text insets (padding) in EMU. Defaults: lIns=91440, rIns=91440, tIns=45720, bIns=45720
	insetLeft   int64
	insetRight  int64
	insetTop    int64
	insetBottom int64
	insetsSet   bool            // true if insets were explicitly parsed from XML
	customPath  *CustomGeomPath // non-nil for freeform/custGeom shapes
	headEnd     *LineEnd        // arrow at start of custom path (from <a:ln><a:headEnd>)
	tailEnd     *LineEnd        // arrow at end of custom path (from <a:ln><a:tailEnd>)
}

// TextAnchorType represents the text anchoring type within a shape.
type TextAnchorType string

const (
	TextAnchorTop    TextAnchorType = "t"
	TextAnchorMiddle TextAnchorType = "ctr"
	TextAnchorBottom TextAnchorType = "b"
	TextAnchorNone   TextAnchorType = ""
)

// AutoFitType represents the auto-fit behavior.
type AutoFitType int

const (
	AutoFitNone AutoFitType = iota
	AutoFitNormal
	AutoFitShape
)

func (r *RichTextShape) GetType() ShapeType { return ShapeTypeRichText }

// NewRichTextShape creates a new rich text shape.
func NewRichTextShape() *RichTextShape {
	rt := &RichTextShape{
		paragraphs: []*Paragraph{NewParagraph()},
		wordWrap:   true,
		columns:    1,
	}
	return rt
}

// SetHeight sets the height and returns the shape for chaining.
func (r *RichTextShape) SetHeight(h int64) *RichTextShape {
	r.height = h
	return r
}

// SetWidth sets the width and returns the shape for chaining.
func (r *RichTextShape) SetWidth(w int64) *RichTextShape {
	r.width = w
	return r
}

// SetOffsetX sets the X offset and returns the shape for chaining.
func (r *RichTextShape) SetOffsetX(x int64) *RichTextShape {
	r.offsetX = x
	return r
}

// SetOffsetY sets the Y offset and returns the shape for chaining.
func (r *RichTextShape) SetOffsetY(y int64) *RichTextShape {
	r.offsetY = y
	return r
}

// GetActiveParagraph returns the active paragraph.
func (r *RichTextShape) GetActiveParagraph() *Paragraph {
	if len(r.paragraphs) == 0 {
		r.paragraphs = append(r.paragraphs, NewParagraph())
	}
	return r.paragraphs[r.activeParagraph]
}

// CreateParagraph creates a new paragraph and makes it active.
func (r *RichTextShape) CreateParagraph() *Paragraph {
	p := NewParagraph()
	r.paragraphs = append(r.paragraphs, p)
	r.activeParagraph = len(r.paragraphs) - 1
	return p
}

// GetParagraphs returns all paragraphs.
func (r *RichTextShape) GetParagraphs() []*Paragraph {
	return r.paragraphs
}

// CreateTextRun creates a text run in the active paragraph.
func (r *RichTextShape) CreateTextRun(text string) *TextRun {
	return r.GetActiveParagraph().CreateTextRun(text)
}

// CreateBreak creates a line break in the active paragraph.
func (r *RichTextShape) CreateBreak() *BreakElement {
	return r.GetActiveParagraph().CreateBreak()
}

// SetAutoFit sets the auto-fit type.
func (r *RichTextShape) SetAutoFit(fit AutoFitType) {
	r.autoFit = fit
}

// GetAutoFit returns the auto-fit type.
func (r *RichTextShape) GetAutoFit() AutoFitType {
	return r.autoFit
}

// SetWordWrap sets word wrap.
func (r *RichTextShape) SetWordWrap(wrap bool) {
	r.wordWrap = wrap
}

// GetWordWrap returns word wrap setting.
func (r *RichTextShape) GetWordWrap() bool {
	return r.wordWrap
}

// SetColumns sets the number of text columns.
func (r *RichTextShape) SetColumns(cols int) {
	r.columns = cols
}

// GetColumns returns the number of text columns.
func (r *RichTextShape) GetColumns() int {
	return r.columns
}

// SetTextAnchor sets the text anchoring type (vertical position of text within the shape).
func (r *RichTextShape) SetTextAnchor(anchor TextAnchorType) {
	r.textAnchor = anchor
}

// GetTextAnchor returns the text anchoring type.
func (r *RichTextShape) GetTextAnchor() TextAnchorType {
	return r.textAnchor
}

// GetCustomPath returns the custom geometry path, if any.
func (r *RichTextShape) GetCustomPath() *CustomGeomPath {
	return r.customPath
}

// Paragraph represents a text paragraph.
type Paragraph struct {
	elements    []ParagraphElement
	alignment   *Alignment
	bullet      *Bullet
	lineSpacing int // in points * 100
	spaceBefore int
	spaceAfter  int
}

// ParagraphElement is the interface for paragraph content.
type ParagraphElement interface {
	GetElementType() string
}

// NewParagraph creates a new paragraph.
func NewParagraph() *Paragraph {
	return &Paragraph{
		elements:  make([]ParagraphElement, 0),
		alignment: NewAlignment(),
	}
}

// GetAlignment returns the paragraph alignment.
func (p *Paragraph) GetAlignment() *Alignment {
	return p.alignment
}

// SetAlignment sets the paragraph alignment.
func (p *Paragraph) SetAlignment(a *Alignment) {
	p.alignment = a
}

// GetBullet returns the paragraph bullet.
func (p *Paragraph) GetBullet() *Bullet {
	return p.bullet
}

// SetBullet sets the paragraph bullet.
func (p *Paragraph) SetBullet(b *Bullet) {
	p.bullet = b
}

// GetLineSpacing returns the line spacing.
func (p *Paragraph) GetLineSpacing() int {
	return p.lineSpacing
}

// SetLineSpacing sets the line spacing.
func (p *Paragraph) SetLineSpacing(spacing int) {
	p.lineSpacing = spacing
}

// GetElements returns all paragraph elements.
func (p *Paragraph) GetElements() []ParagraphElement {
	return p.elements
}

// GetSpaceBefore returns the space before the paragraph.
func (p *Paragraph) GetSpaceBefore() int { return p.spaceBefore }

// SetSpaceBefore sets the space before the paragraph.
func (p *Paragraph) SetSpaceBefore(v int) { p.spaceBefore = v }

// GetSpaceAfter returns the space after the paragraph.
func (p *Paragraph) GetSpaceAfter() int { return p.spaceAfter }

// SetSpaceAfter sets the space after the paragraph.
func (p *Paragraph) SetSpaceAfter(v int) { p.spaceAfter = v }

// CreateTextRun creates a new text run.
func (p *Paragraph) CreateTextRun(text string) *TextRun {
	tr := &TextRun{
		text: text,
		font: NewFont(),
	}
	p.elements = append(p.elements, tr)
	return tr
}

// CreateBreak creates a line break element.
func (p *Paragraph) CreateBreak() *BreakElement {
	br := &BreakElement{}
	p.elements = append(p.elements, br)
	return br
}

// TextRun represents a run of text with formatting.
type TextRun struct {
	text      string
	font      *Font
	hyperlink *Hyperlink
}

func (tr *TextRun) GetElementType() string { return "textrun" }

// GetText returns the text content.
func (tr *TextRun) GetText() string { return tr.text }

// SetText sets the text content.
func (tr *TextRun) SetText(text string) { tr.text = text }

// GetFont returns the font properties.
func (tr *TextRun) GetFont() *Font { return tr.font }

// SetFont sets the font properties.
func (tr *TextRun) SetFont(f *Font) { tr.font = f }

// GetHyperlink returns the hyperlink.
func (tr *TextRun) GetHyperlink() *Hyperlink { return tr.hyperlink }

// SetHyperlink sets the hyperlink.
func (tr *TextRun) SetHyperlink(h *Hyperlink) { tr.hyperlink = h }

// BreakElement represents a line break.
type BreakElement struct{}

func (br *BreakElement) GetElementType() string { return "break" }

// DrawingShape represents an image/drawing shape.
type DrawingShape struct {
	BaseShape
	path               string // file path
	data               []byte // raw image data
	mimeType           string
	resizeProportional bool
	alpha              int // alphaModFix amount (0-100000); 0 means fully opaque (default)
	// srcRect crop percentages in 1/1000 of a percent (e.g. 56333 = 56.333%)
	cropLeft   int
	cropTop    int
	cropRight  int
	cropBottom int
}

func (d *DrawingShape) GetType() ShapeType { return ShapeTypeDrawing }

// NewDrawingShape creates a new drawing shape.
func NewDrawingShape() *DrawingShape {
	return &DrawingShape{
		resizeProportional: true,
	}
}

// SetPath sets the image file path.
func (d *DrawingShape) SetPath(path string) *DrawingShape {
	d.path = path
	return d
}

// GetPath returns the image file path.
func (d *DrawingShape) GetPath() string { return d.path }

// SetImageData sets the raw image data.
func (d *DrawingShape) SetImageData(data []byte, mimeType string) *DrawingShape {
	d.data = data
	d.mimeType = mimeType
	return d
}

// GetImageData returns the raw image data.
func (d *DrawingShape) GetImageData() []byte { return d.data }

// GetMimeType returns the image MIME type.
func (d *DrawingShape) GetMimeType() string { return d.mimeType }

// maxImageFileSize is the maximum allowed size for an image file loaded from disk.
const maxImageFileSize = 50 << 20 // 50 MB

// SetImageFromFile loads an image from a file path and sets the data and MIME type.
// Returns an error if the file exceeds maxImageFileSize or cannot be read.
func (d *DrawingShape) SetImageFromFile(path string) error {
	info, err := os.Stat(path)
	if err != nil {
		return fmt.Errorf("failed to stat image file: %w", err)
	}
	if info.Size() > maxImageFileSize {
		return fmt.Errorf("image file too large: %d bytes (max %d)", info.Size(), maxImageFileSize)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return fmt.Errorf("failed to read image file: %w", err)
	}
	mime := guessMimeFromPath(path)
	d.data = data
	d.mimeType = mime
	return nil
}

// guessMimeFromPath guesses the MIME type from a file extension.
func guessMimeFromPath(path string) string {
	lower := strings.ToLower(path)
	switch {
	case strings.HasSuffix(lower, ".png"):
		return "image/png"
	case strings.HasSuffix(lower, ".jpg"), strings.HasSuffix(lower, ".jpeg"):
		return "image/jpeg"
	case strings.HasSuffix(lower, ".gif"):
		return "image/gif"
	case strings.HasSuffix(lower, ".bmp"):
		return "image/bmp"
	case strings.HasSuffix(lower, ".svg"):
		return "image/svg+xml"
	default:
		return "image/png"
	}
}

// SetHeight sets the height and returns for chaining.
func (d *DrawingShape) SetHeight(h int64) *DrawingShape {
	d.height = h
	return d
}

// SetWidth sets the width and returns for chaining.
func (d *DrawingShape) SetWidth(w int64) *DrawingShape {
	d.width = w
	return d
}

// SetOffsetX sets the X offset and returns for chaining.
func (d *DrawingShape) SetOffsetX(x int64) *DrawingShape {
	d.offsetX = x
	return d
}

// SetOffsetY sets the Y offset and returns for chaining.
func (d *DrawingShape) SetOffsetY(y int64) *DrawingShape {
	d.offsetY = y
	return d
}

// GetCropLeft returns the left crop percentage (in 1/1000 of a percent).
func (d *DrawingShape) GetCropLeft() int { return d.cropLeft }

// GetCropTop returns the top crop percentage (in 1/1000 of a percent).
func (d *DrawingShape) GetCropTop() int { return d.cropTop }

// GetCropRight returns the right crop percentage (in 1/1000 of a percent).
func (d *DrawingShape) GetCropRight() int { return d.cropRight }

// GetCropBottom returns the bottom crop percentage (in 1/1000 of a percent).
func (d *DrawingShape) GetCropBottom() int { return d.cropBottom }

// GetAlphaValue returns the alphaModFix amount (0-100000).
func (d *DrawingShape) GetAlphaValue() int { return d.alpha }

// AutoShape represents a predefined shape (rectangle, ellipse, etc.).
type AutoShape struct {
	BaseShape
	shapeType     AutoShapeType
	text          string
	paragraphs    []*Paragraph
	textAnchor    TextAnchorType
	textDirection string
	adjustValues  map[string]int // avLst adjustment values (e.g. "adj1" -> 10690)
	fontScale     int            // normAutofit fontScale in thousandths of a percent (e.g. 62500 = 62.5%), 0 means 100%
	// Text insets (padding) in EMU.
	insetLeft   int64
	insetRight  int64
	insetTop    int64
	insetBottom int64
	insetsSet   bool
	headEnd     *LineEnd // arrow at start of arc
	tailEnd     *LineEnd // arrow at end of arc
}

// AutoShapeType represents the type of auto shape.
type AutoShapeType string

const (
	AutoShapeRectangle            AutoShapeType = "rect"
	AutoShapeRoundedRect          AutoShapeType = "roundRect"
	AutoShapeEllipse              AutoShapeType = "ellipse"
	AutoShapeTriangle             AutoShapeType = "triangle"
	AutoShapeDiamond              AutoShapeType = "diamond"
	AutoShapeParallelogram        AutoShapeType = "parallelogram"
	AutoShapeTrapezoid            AutoShapeType = "trapezoid"
	AutoShapePentagon             AutoShapeType = "pentagon"
	AutoShapeHexagon              AutoShapeType = "hexagon"
	AutoShapeArrowRight           AutoShapeType = "rightArrow"
	AutoShapeArrowLeft            AutoShapeType = "leftArrow"
	AutoShapeArrowUp              AutoShapeType = "upArrow"
	AutoShapeArrowDown            AutoShapeType = "downArrow"
	AutoShapeStar4                AutoShapeType = "star4"
	AutoShapeStar5                AutoShapeType = "star5"
	AutoShapeStar10               AutoShapeType = "star10"
	AutoShapeStar12               AutoShapeType = "star12"
	AutoShapeStar16               AutoShapeType = "star16"
	AutoShapeStar24               AutoShapeType = "star24"
	AutoShapeStar32               AutoShapeType = "star32"
	AutoShapeHeart                AutoShapeType = "heart"
	AutoShapeLightningBolt        AutoShapeType = "lightningBolt"
	AutoShapeChevron              AutoShapeType = "chevron"
	AutoShapeCloud                AutoShapeType = "cloud"
	AutoShapePlus                 AutoShapeType = "mathPlus"
	AutoShapeMinus                AutoShapeType = "mathMinus"
	AutoShapeFlowchartProcess     AutoShapeType = "flowChartProcess"
	AutoShapeFlowchartDecision    AutoShapeType = "flowChartDecision"
	AutoShapeFlowchartPreparation AutoShapeType = "flowChartPreparation"
	AutoShapeCallout1             AutoShapeType = "wedgeRoundRectCallout"
	AutoShapeCallout2             AutoShapeType = "wedgeEllipseCallout"
	AutoShapeRibbon               AutoShapeType = "ribbon2"
	AutoShapeSmileyFace           AutoShapeType = "smileyFace"
	AutoShapeDonut                AutoShapeType = "donut"
	AutoShapeNoSmoking            AutoShapeType = "noSmoking"
	AutoShapeBlockArc             AutoShapeType = "blockArc"
	AutoShapeCube                 AutoShapeType = "cube"
	AutoShapeCan                  AutoShapeType = "can"
	AutoShapeBevel                AutoShapeType = "bevel"
	AutoShapeFoldedCorner         AutoShapeType = "foldedCorner"
	AutoShapeFrame                AutoShapeType = "frame"
	AutoShapePlaque               AutoShapeType = "plaque"
	AutoShapeLeftRightArrow       AutoShapeType = "leftRightArrow"
	AutoShapeRtTriangle           AutoShapeType = "rtTriangle"
	AutoShapeHomePlate            AutoShapeType = "homePlate"
	AutoShapeSnip2SameRect        AutoShapeType = "snip2SameRect"
	AutoShapePie                  AutoShapeType = "pie"
	AutoShapeArc                  AutoShapeType = "arc"
	AutoShapeBentArrow            AutoShapeType = "bentArrow"
	AutoShapeUturnArrow           AutoShapeType = "uturnArrow"
	AutoShapeMathEqual            AutoShapeType = "mathEqual"
	AutoShapeCurvedRightArrow     AutoShapeType = "curvedRightArrow"
	AutoShapeCurvedLeftArrow      AutoShapeType = "curvedLeftArrow"
	AutoShapeCurvedUpArrow        AutoShapeType = "curvedUpArrow"
	AutoShapeCurvedDownArrow      AutoShapeType = "curvedDownArrow"
)

func (a *AutoShape) GetType() ShapeType { return ShapeTypeAutoShape }

// NewAutoShape creates a new auto shape.
func NewAutoShape() *AutoShape {
	return &AutoShape{
		shapeType: AutoShapeRectangle,
	}
}

// SetAutoShapeType sets the auto shape type.
func (a *AutoShape) SetAutoShapeType(t AutoShapeType) *AutoShape {
	a.shapeType = t
	return a
}

// SetGeometry sets the shape geometry type (alias for SetAutoShapeType).
// This matches the unioffice naming convention.
func (a *AutoShape) SetGeometry(t AutoShapeType) *AutoShape {
	a.shapeType = t
	return a
}

// GetAutoShapeType returns the auto shape type.
func (a *AutoShape) GetAutoShapeType() AutoShapeType {
	return a.shapeType
}

// SetSolidFill sets a solid fill on the auto shape.
func (a *AutoShape) SetSolidFill(c Color) *AutoShape {
	a.GetFill().SetSolid(c)
	return a
}

// SetText sets the text content.
func (a *AutoShape) SetText(text string) *AutoShape {
	a.text = text
	return a
}

// GetText returns the text content.
func (a *AutoShape) GetText() string {
	return a.text
}

// GetParagraphs returns the rich text paragraphs (if any).
func (a *AutoShape) GetParagraphs() []*Paragraph {
	return a.paragraphs
}

// GetAdjustValues returns the adjustment values map.
// GetHeadEnd returns the head end arrow.
func (a *AutoShape) GetHeadEnd() *LineEnd { return a.headEnd }

// GetTailEnd returns the tail end arrow.
func (a *AutoShape) GetTailEnd() *LineEnd { return a.tailEnd }
func (a *AutoShape) GetAdjustValues() map[string]int {
	return a.adjustValues
}

// LineShape represents a line shape.
type LineShape struct {
	BaseShape
	lineStyle     BorderStyle
	lineWidth     int
	lineWidthEMU  int             // raw line width in EMU for precision; 0 means use lineWidth*12700
	lineColor     Color
	headEnd       *LineEnd
	tailEnd       *LineEnd
	connectorType string          // prstGeom value: "line", "straightConnector1", "bentConnector3", etc.
	adjustValues  map[string]int  // adjustment values for connector geometry
	customPath    *CustomGeomPath // non-nil for custGeom connectors (freeform curved arrows)
}

func (l *LineShape) GetType() ShapeType { return ShapeTypeLine }

// NewLineShape creates a new line shape.
func NewLineShape() *LineShape {
	return &LineShape{
		lineStyle: BorderSolid,
		lineWidth: 1,
		lineColor: ColorBlack,
	}
}

// SetLineStyle sets the line style.
func (l *LineShape) SetLineStyle(s BorderStyle) *LineShape {
	l.lineStyle = s
	return l
}

// GetLineStyle returns the line style.
func (l *LineShape) GetLineStyle() BorderStyle { return l.lineStyle }

// SetLineWidth sets the line width.
func (l *LineShape) SetLineWidth(w int) *LineShape {
	l.lineWidth = w
	return l
}

// GetLineWidth returns the line width.
func (l *LineShape) GetLineWidth() int { return l.lineWidth }

// GetLineWidthEMU returns the line width in EMU for precise rendering.
// If the raw EMU value was not set, it falls back to lineWidth * 12700.
func (l *LineShape) GetLineWidthEMU() int {
	if l.lineWidthEMU > 0 {
		return l.lineWidthEMU
	}
	return l.lineWidth * 12700
}

// SetLineColor sets the line color.
func (l *LineShape) SetLineColor(c Color) *LineShape {
	l.lineColor = c
	return l
}

// GetLineColor returns the line color.
func (l *LineShape) GetLineColor() Color { return l.lineColor }

// SetHeadEnd sets the head end (arrow at start of line).
func (l *LineShape) SetHeadEnd(e *LineEnd) *LineShape {
	l.headEnd = e
	return l
}

// GetHeadEnd returns the head end.
func (l *LineShape) GetHeadEnd() *LineEnd { return l.headEnd }

// SetTailEnd sets the tail end (arrow at end of line).
func (l *LineShape) SetTailEnd(e *LineEnd) *LineShape {
	l.tailEnd = e
	return l
}

// GetTailEnd returns the tail end.
func (l *LineShape) GetTailEnd() *LineEnd { return l.tailEnd }

// GetConnectorType returns the connector type (prstGeom value).
func (l *LineShape) GetConnectorType() string { return l.connectorType }

// GetAdjustValues returns the adjustment values for connector geometry.
func (l *LineShape) GetAdjustValues() map[string]int { return l.adjustValues }

// TableShape represents a table shape.
type TableShape struct {
	BaseShape
	rows       [][]*TableCell
	numRows    int
	numCols    int
	colWidths  []int64 // individual column widths in EMU (from gridCol)
	rowHeights []int64 // individual row heights in EMU (from tr)
}

func (t *TableShape) GetType() ShapeType { return ShapeTypeTable }

// NewTableShape creates a new table shape.
func NewTableShape(rows, cols int) *TableShape {
	table := &TableShape{
		numRows: rows,
		numCols: cols,
		rows:    make([][]*TableCell, rows),
	}
	for i := 0; i < rows; i++ {
		table.rows[i] = make([]*TableCell, cols)
		for j := 0; j < cols; j++ {
			table.rows[i][j] = NewTableCell()
		}
	}
	return table
}

// GetCell returns a cell at the given row and column.
func (t *TableShape) GetCell(row, col int) *TableCell {
	if row < 0 || row >= t.numRows || col < 0 || col >= t.numCols {
		return nil
	}
	return t.rows[row][col]
}

// GetRows returns all rows.
func (t *TableShape) GetRows() [][]*TableCell {
	return t.rows
}

// GetNumRows returns the number of rows.
func (t *TableShape) GetNumRows() int { return t.numRows }

// GetNumCols returns the number of columns.
func (t *TableShape) GetNumCols() int { return t.numCols }

// SetHeight sets the height and returns for chaining.
func (t *TableShape) SetHeight(h int64) *TableShape {
	t.height = h
	return t
}

// SetWidth sets the width and returns for chaining.
func (t *TableShape) SetWidth(w int64) *TableShape {
	t.width = w
	return t
}

// TableCell represents a table cell.
type TableCell struct {
	paragraphs []*Paragraph
	fill       *Fill
	border     *CellBorders
	colSpan    int
	rowSpan    int
	hMerge     bool // continuation of horizontal merge (skip rendering)
	vMerge     bool // continuation of vertical merge (skip rendering)
}

// CellBorders represents borders for a table cell.
type CellBorders struct {
	Top    *Border
	Bottom *Border
	Left   *Border
	Right  *Border
}

// NewTableCell creates a new table cell.
func NewTableCell() *TableCell {
	return &TableCell{
		paragraphs: []*Paragraph{NewParagraph()},
		fill:       NewFill(),
		border: &CellBorders{
			Top:    NewBorder(),
			Bottom: NewBorder(),
			Left:   NewBorder(),
			Right:  NewBorder(),
		},
		colSpan: 1,
		rowSpan: 1,
	}
}

// SetText sets the cell text (convenience method).
func (tc *TableCell) SetText(text string) *TableCell {
	if len(tc.paragraphs) == 0 {
		tc.paragraphs = append(tc.paragraphs, NewParagraph())
	}
	tc.paragraphs[0].CreateTextRun(text)
	return tc
}

// GetParagraphs returns the cell paragraphs.
func (tc *TableCell) GetParagraphs() []*Paragraph {
	return tc.paragraphs
}

// GetFill returns the cell fill.
func (tc *TableCell) GetFill() *Fill { return tc.fill }

// SetFill sets the cell fill.
func (tc *TableCell) SetFill(f *Fill) { tc.fill = f }

// GetBorders returns the cell borders.
func (tc *TableCell) GetBorders() *CellBorders { return tc.border }

// SetColSpan sets the column span.
func (tc *TableCell) SetColSpan(span int) { tc.colSpan = span }

// GetColSpan returns the column span.
func (tc *TableCell) GetColSpan() int { return tc.colSpan }

// SetRowSpan sets the row span.
func (tc *TableCell) SetRowSpan(span int) { tc.rowSpan = span }

// GetRowSpan returns the row span.
func (tc *TableCell) GetRowSpan() int { return tc.rowSpan }
