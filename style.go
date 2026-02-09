package gopresentation

import (
	"strings"
)

// Color represents an ARGB color.
type Color struct {
	ARGB string // 8-character hex string, e.g., "FF000000" for black
}

// Predefined colors.
var (
	ColorBlack   = Color{ARGB: "FF000000"}
	ColorWhite   = Color{ARGB: "FFFFFFFF"}
	ColorRed     = Color{ARGB: "FFFF0000"}
	ColorGreen   = Color{ARGB: "FF00FF00"}
	ColorBlue    = Color{ARGB: "FF0000FF"}
	ColorYellow  = Color{ARGB: "FFFFFF00"}
)

// NewColor creates a new Color from an ARGB hex string.
// Accepts 6-char RGB (e.g. "FF0000") or 8-char ARGB (e.g. "FFFF0000").
// A leading "#" is stripped automatically.
func NewColor(argb string) Color {
	argb = strings.TrimPrefix(argb, "#")
	if len(argb) == 6 {
		argb = "FF" + argb
	}
	argb = strings.ToUpper(argb)
	if !isValidARGB(argb) {
		return Color{ARGB: "FF000000"} // fallback to black
	}
	return Color{ARGB: argb}
}

// isValidARGB checks that s is exactly 8 hex characters.
func isValidARGB(s string) bool {
	if len(s) != 8 {
		return false
	}
	for _, c := range s {
		if !((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F')) {
			return false
		}
	}
	return true
}

// GetRed returns the red component (0-255).
func (c Color) GetRed() uint8 {
	return parseHexByte(c.ARGB, 2)
}

// GetGreen returns the green component (0-255).
func (c Color) GetGreen() uint8 {
	return parseHexByte(c.ARGB, 4)
}

// GetBlue returns the blue component (0-255).
func (c Color) GetBlue() uint8 {
	return parseHexByte(c.ARGB, 6)
}

// GetAlpha returns the alpha component (0-255).
func (c Color) GetAlpha() uint8 {
	return parseHexByte(c.ARGB, 0)
}

// parseHexByte parses two hex characters at offset into a uint8.
// Returns 0 on any error (out of range, invalid chars).
func parseHexByte(s string, offset int) uint8 {
	if offset+2 > len(s) {
		return 0
	}
	h := hexVal(s[offset])
	l := hexVal(s[offset+1])
	if h < 0 || l < 0 {
		return 0
	}
	return uint8(h<<4 | l)
}

func hexVal(c byte) int {
	switch {
	case c >= '0' && c <= '9':
		return int(c - '0')
	case c >= 'A' && c <= 'F':
		return int(c-'A') + 10
	case c >= 'a' && c <= 'f':
		return int(c-'a') + 10
	default:
		return -1
	}
}

// Font represents text font properties.
type Font struct {
	Name          string
	Size          int     // in points
	Bold          bool
	Italic        bool
	Underline     UnderlineType
	Strikethrough bool
	Color         Color
	Superscript   bool
	Subscript     bool
}

// UnderlineType represents the underline style.
type UnderlineType string

const (
	UnderlineNone   UnderlineType = "none"
	UnderlineSingle UnderlineType = "sng"
	UnderlineDouble UnderlineType = "dbl"
	UnderlineHeavy  UnderlineType = "heavy"
	UnderlineDash   UnderlineType = "dash"
	UnderlineWavy   UnderlineType = "wavy"
)

// NewFont creates a new Font with defaults.
func NewFont() *Font {
	return &Font{
		Name:      "Calibri",
		Size:      10,
		Bold:      false,
		Italic:    false,
		Underline: UnderlineNone,
		Color:     ColorBlack,
	}
}

// SetBold sets the bold property and returns the font for chaining.
func (f *Font) SetBold(bold bool) *Font {
	f.Bold = bold
	return f
}

// SetItalic sets the italic property.
func (f *Font) SetItalic(italic bool) *Font {
	f.Italic = italic
	return f
}

// SetSize sets the font size in points (clamped to 1–4000).
func (f *Font) SetSize(size int) *Font {
	if size < 1 {
		size = 1
	}
	if size > 4000 {
		size = 4000
	}
	f.Size = size
	return f
}

// SetColor sets the font color.
func (f *Font) SetColor(color Color) *Font {
	f.Color = color
	return f
}

// SetName sets the font name.
func (f *Font) SetName(name string) *Font {
	f.Name = name
	return f
}

// SetUnderline sets the underline type.
func (f *Font) SetUnderline(u UnderlineType) *Font {
	f.Underline = u
	return f
}

// SetStrikethrough sets the strikethrough property.
func (f *Font) SetStrikethrough(s bool) *Font {
	f.Strikethrough = s
	return f
}

// Alignment represents text alignment properties.
type Alignment struct {
	Horizontal HorizontalAlignment
	Vertical   VerticalAlignment
	MarginLeft int64 // in EMU
	MarginRight int64
	MarginTop  int64
	MarginBottom int64
	Indent     int64
	Level      int
}

// HorizontalAlignment represents horizontal text alignment.
type HorizontalAlignment string

const (
	HorizontalLeft      HorizontalAlignment = "l"
	HorizontalCenter    HorizontalAlignment = "ctr"
	HorizontalRight     HorizontalAlignment = "r"
	HorizontalJustify   HorizontalAlignment = "just"
	HorizontalDistributed HorizontalAlignment = "dist"
)

// VerticalAlignment represents vertical text alignment.
type VerticalAlignment string

const (
	VerticalTop    VerticalAlignment = "t"
	VerticalMiddle VerticalAlignment = "ctr"
	VerticalBottom VerticalAlignment = "b"
)

// NewAlignment creates a new Alignment with defaults.
func NewAlignment() *Alignment {
	return &Alignment{
		Horizontal: HorizontalLeft,
		Vertical:   VerticalTop,
	}
}

// SetHorizontal sets horizontal alignment.
func (a *Alignment) SetHorizontal(h HorizontalAlignment) *Alignment {
	a.Horizontal = h
	return a
}

// SetVertical sets vertical alignment.
func (a *Alignment) SetVertical(v VerticalAlignment) *Alignment {
	a.Vertical = v
	return a
}

// Fill represents a shape fill.
type Fill struct {
	Type      FillType
	Color     Color
	EndColor  Color // for gradient fills
	Rotation  int   // gradient rotation in degrees
}

// FillType represents the type of fill.
type FillType int

const (
	FillNone FillType = iota
	FillSolid
	FillGradientLinear
	FillGradientPath
)

// NewFill creates a new Fill with no fill.
func NewFill() *Fill {
	return &Fill{Type: FillNone}
}

// SetSolid sets a solid fill.
func (f *Fill) SetSolid(color Color) *Fill {
	f.Type = FillSolid
	f.Color = color
	return f
}

// SetGradientLinear sets a linear gradient fill. Rotation is normalized to 0–359.
func (f *Fill) SetGradientLinear(startColor, endColor Color, rotation int) *Fill {
	f.Type = FillGradientLinear
	f.Color = startColor
	f.EndColor = endColor
	f.Rotation = ((rotation % 360) + 360) % 360
	return f
}

// Border represents a shape border.
type Border struct {
	Style BorderStyle
	Width int // in EMU
	Color Color
}

// BorderStyle represents the border line style.
type BorderStyle string

const (
	BorderNone  BorderStyle = "none"
	BorderSolid BorderStyle = "solid"
	BorderDash  BorderStyle = "dash"
	BorderDot   BorderStyle = "dot"
)

// NewBorder creates a new Border with no border.
func NewBorder() *Border {
	return &Border{Style: BorderNone}
}

// Shadow represents a shape shadow.
type Shadow struct {
	Visible   bool
	Direction int // in degrees
	Distance  int // in points
	BlurRadius int
	Color     Color
	Alpha     int // 0-100
}

// NewShadow creates a new Shadow.
func NewShadow() *Shadow {
	return &Shadow{
		Visible:   false,
		Direction: 0,
		Distance:  0,
		Color:     Color{ARGB: "80000000"},
		Alpha:     50,
	}
}

// SetVisible sets shadow visibility.
func (s *Shadow) SetVisible(v bool) *Shadow {
	s.Visible = v
	return s
}

// SetDirection sets shadow direction in degrees (normalized to 0–359).
func (s *Shadow) SetDirection(d int) *Shadow {
	s.Direction = ((d % 360) + 360) % 360
	return s
}

// SetDistance sets shadow distance in points (clamped to >= 0).
func (s *Shadow) SetDistance(d int) *Shadow {
	if d < 0 {
		d = 0
	}
	s.Distance = d
	return s
}

// Hyperlink represents a hyperlink.
type Hyperlink struct {
	URL     string
	Tooltip string
	IsInternal bool
	SlideNumber int
}

// NewHyperlink creates a new external hyperlink.
func NewHyperlink(url string) *Hyperlink {
	return &Hyperlink{URL: url}
}

// NewInternalHyperlink creates a hyperlink to another slide.
func NewInternalHyperlink(slideNumber int) *Hyperlink {
	return &Hyperlink{
		IsInternal:  true,
		SlideNumber: slideNumber,
	}
}
