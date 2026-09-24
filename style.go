package gopresentation

import (
	"math"
	"strings"
)

// Color represents an ARGB color.
type Color struct {
	ARGB string // 8-character hex string, e.g., "FF000000" for black
}

// Predefined colors.
var (
	ColorBlack  = Color{ARGB: "FF000000"}
	ColorWhite  = Color{ARGB: "FFFFFFFF"}
	ColorRed    = Color{ARGB: "FFFF0000"}
	ColorGreen  = Color{ARGB: "FF00FF00"}
	ColorBlue   = Color{ARGB: "FF0000FF"}
	ColorYellow = Color{ARGB: "FFFFFF00"}
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
	NameEA        string // East Asian font name (from <a:ea> element)
	NameSym       string // symbol font name (from <a:sym> element) — carries the PUA U+F000-F0FF glyphs
	Size          int    // in points
	Bold          bool
	Italic        bool
	Underline     UnderlineType
	Strikethrough bool
	Color         Color
	Superscript   bool
	Subscript     bool
	Shadow        *Shadow // run-level text shadow, from <a:rPr><a:effectLst><a:outerShdw>
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
	Horizontal   HorizontalAlignment
	Vertical     VerticalAlignment
	MarginLeft   int64 // in EMU
	MarginRight  int64
	MarginTop    int64
	MarginBottom int64
	Indent       int64
	Level        int

	// marLSet/indentSet record that the paragraph's markup stated the
	// attribute, even as an explicit 0. The master's bodyStyle carries a
	// marL/indent of its own, and inheritance must not overwrite a value the
	// slide spelled out — "0" is a deliberate override there, not "unset".
	marLSet   bool
	indentSet bool
}

// HorizontalAlignment represents horizontal text alignment.
type HorizontalAlignment string

const (
	HorizontalLeft        HorizontalAlignment = "l"
	HorizontalCenter      HorizontalAlignment = "ctr"
	HorizontalRight       HorizontalAlignment = "r"
	HorizontalJustify     HorizontalAlignment = "just"
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
	Type     FillType
	Color    Color
	EndColor Color // for gradient fills
	Rotation int   // gradient rotation in degrees
	// MidColor/MidPos carry a middle gradient stop (MidPos in 0..100000
	// gradient-vector units, 0 = absent). The Office themes' fill styles are
	// three-stop gradients whose knee sits well off the straight line between
	// the ends — dropping it shows as a several-level band across every
	// fillRef-2 shape. The writer writes it back between the other stops.
	MidColor Color
	MidPos   int
	// Stops carries the full ordered stop list of an N-stop gradient (nil =
	// legacy fields above). Pos is in 0..100000 gradient-vector units.
	Stops []GradStop
	// Path is the gradient-path kind for FillGradientPath fills ("circle",
	// "rect", "shape"); empty for linear gradients. FillTo holds the
	// <a:fillToRect> insets and TileTo the <a:tileRect> insets, both as
	// (l, t, r, b) in 0..100000 gradient-box units — tileRect values may be
	// negative, which grows the tile beyond the box.
	Path   string
	FillTo [4]int
	TileTo [4]int
}

// GradStop is one stop of a multi-stop gradient.
type GradStop struct {
	Pos   int // 0..100000 along the gradient vector
	Color Color
}

// SetGradientStops installs a full N-stop gradient, linear (path == "") or
// path-shaped. The legacy two/three-stop fields are kept in sync so callers
// and the writer that predate Stops keep working.
func (f *Fill) SetGradientStops(stops []GradStop, angle int, path string, fillTo, tileTo [4]int) *Fill {
	f.Stops = stops
	f.Color = stops[0].Color
	f.EndColor = stops[len(stops)-1].Color
	if len(stops) == 3 {
		f.MidColor = stops[1].Color
		f.MidPos = stops[1].Pos
	} else {
		f.MidColor = Color{}
		f.MidPos = 0
	}
	if path != "" {
		f.Type = FillGradientPath
		f.Path = path
		f.FillTo = fillTo
		f.TileTo = tileTo
	} else {
		f.Type = FillGradientLinear
		f.Rotation = ((angle % 360) + 360) % 360
	}
	return f
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

// SetNoFill clears the fill.
func (f *Fill) SetNoFill() *Fill {
	f.Type = FillNone
	return f
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
	Width int // in points (1 pt = 12700 EMU)
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

// ArrowType represents the type of arrow head/tail on a line.
type ArrowType string

const (
	ArrowNone     ArrowType = "none"
	ArrowTriangle ArrowType = "triangle"
	ArrowStealth  ArrowType = "stealth"
	ArrowDiamond  ArrowType = "diamond"
	ArrowOval     ArrowType = "oval"
	ArrowArrow    ArrowType = "arrow"
)

// ArrowSize represents the size of an arrow head/tail.
type ArrowSize string

const (
	ArrowSizeSm  ArrowSize = "sm"
	ArrowSizeMed ArrowSize = "med"
	ArrowSizeLg  ArrowSize = "lg"
)

// LineEnd represents the arrow head or tail of a line.
type LineEnd struct {
	Type   ArrowType
	Width  ArrowSize
	Length ArrowSize
}

// NewBorder creates a new Border with no border.
func NewBorder() *Border {
	return &Border{Style: BorderNone}
}

// SetSolidFill sets a solid border with the given color.
func (b *Border) SetSolidFill(c Color) *Border {
	b.Style = BorderSolid
	b.Color = c
	return b
}

// SetWidth sets the border width in points. It is stored in points, not EMU:
// the reader converts the EMU value found in the file (v/12700) and the writer
// converts back, so a value of 1 means a 1pt line.
func (b *Border) SetWidth(w int) *Border {
	b.Width = w
	return b
}

// SetNoFill removes the border.
func (b *Border) SetNoFill() *Border {
	b.Style = BorderNone
	return b
}

// Shadow represents a shape shadow.
type Shadow struct {
	Visible    bool
	Direction  int // in degrees
	Distance   int // in points
	BlurRadius int
	Color      Color
	Alpha      int // 0-100
	// ScaleX/ScaleY scale the shadow silhouette relative to the shape
	// (percent; 100 = same size). PowerPoint writes e.g. sx="105000"
	// for a 5% enlarged silhouette.
	ScaleX int
	ScaleY int
	// Algn anchors the scaled silhouette to the shape box ("tl", "ctr",
	// ...). With tl the enlargement grows right/down; empty means the
	// writer/reader never saw one and the silhouette stays centred.
	Algn string
}

// NewShadow creates a new Shadow.
func NewShadow() *Shadow {
	return &Shadow{
		Visible:   false,
		Direction: 0,
		Distance:  0,
		Color:     Color{ARGB: "80000000"},
		Alpha:     50,
		ScaleX:    100,
		ScaleY:    100,
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
	URL         string
	Tooltip     string
	IsInternal  bool
	SlideNumber int
}

// allowedHyperlinkSchemes defines the URL schemes permitted in hyperlinks.
// This prevents injection of dangerous schemes like javascript: or data:.
var allowedHyperlinkSchemes = []string{"http://", "https://", "mailto:", "ftp://", "ftps://"}

// isValidHyperlinkURL checks that a URL uses an allowed scheme.
func isValidHyperlinkURL(url string) bool {
	lower := strings.ToLower(strings.TrimSpace(url))
	for _, scheme := range allowedHyperlinkSchemes {
		if strings.HasPrefix(lower, scheme) {
			return true
		}
	}
	return false
}

// NewHyperlink creates a new external hyperlink.
// The URL must use an allowed scheme (http, https, mailto, ftp, ftps).
// Returns nil if the URL scheme is not allowed.
func NewHyperlink(url string) *Hyperlink {
	if !isValidHyperlinkURL(url) {
		return nil
	}
	return &Hyperlink{URL: url}
}

// NewInternalHyperlink creates a hyperlink to another slide.
func NewInternalHyperlink(slideNumber int) *Hyperlink {
	return &Hyperlink{
		IsInternal:  true,
		SlideNumber: slideNumber,
	}
}

// --- Color modification helpers for OOXML color transforms ---

// rgbToHSL converts RGB (0-255) to HSL (h: 0-360, s: 0-1, l: 0-1).
func rgbToHSL(r, g, b uint8) (float64, float64, float64) {
	rf := float64(r) / 255.0
	gf := float64(g) / 255.0
	bf := float64(b) / 255.0

	max := rf
	if gf > max {
		max = gf
	}
	if bf > max {
		max = bf
	}
	min := rf
	if gf < min {
		min = gf
	}
	if bf < min {
		min = bf
	}

	l := (max + min) / 2.0
	if max == min {
		return 0, 0, l
	}

	d := max - min
	s := d / (1.0 - abs64(2.0*l-1.0))

	var h float64
	switch max {
	case rf:
		h = (gf - bf) / d
		if gf < bf {
			h += 6
		}
	case gf:
		h = (bf-rf)/d + 2
	case bf:
		h = (rf-gf)/d + 4
	}
	h *= 60

	return h, s, l
}

// hslToRGB converts HSL (h: 0-360, s: 0-1, l: 0-1) to RGB (0-255).
func hslToRGB(h, s, l float64) (uint8, uint8, uint8) {
	if s == 0 {
		v := clamp8(l * 255)
		return v, v, v
	}

	var q float64
	if l < 0.5 {
		q = l * (1 + s)
	} else {
		q = l + s - l*s
	}
	p := 2*l - q

	hk := h / 360.0
	tr := hk + 1.0/3.0
	tg := hk
	tb := hk - 1.0/3.0

	hueToRGB := func(t float64) float64 {
		if t < 0 {
			t += 1
		}
		if t > 1 {
			t -= 1
		}
		switch {
		case t < 1.0/6.0:
			return p + (q-p)*6*t
		case t < 1.0/2.0:
			return q
		case t < 2.0/3.0:
			return p + (q-p)*(2.0/3.0-t)*6
		default:
			return p
		}
	}

	return clamp8(hueToRGB(tr) * 255), clamp8(hueToRGB(tg) * 255), clamp8(hueToRGB(tb) * 255)
}

func abs64(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}

func clamp8(v float64) uint8 {
	if v < 0 {
		return 0
	}
	if v > 255 {
		return 255
	}
	return uint8(v + 0.5)
}

// applySatMod multiplies the saturation by factor (e.g. 1.6 = 160%) WITHOUT
// clamping S to 1 first: PowerPoint lets the multiplied S exceed 1 and clamps
// only the resulting RGB channels. ED7D31 with satMod 160% renders FF7200;
// clamping S first would give FF7A1F. Measured against PowerPoint's own
// output (r39 op-probe deck, 34 solid swatches).
func applySatMod(c *Color, factor float64) {
	h, s, l := rgbToHSL(c.GetRed(), c.GetGreen(), c.GetBlue())
	nr, ng, nb := hslToRGB(h, s*factor, l)
	c.ARGB = c.ARGB[:2] + colorHex(nr) + colorHex(ng) + colorHex(nb)
}

// applyLumMod multiplies the luminance by factor (e.g. 0.75 = 75%).
func applyLumMod(c *Color, factor float64) {
	r, g, b := c.GetRed(), c.GetGreen(), c.GetBlue()
	h, s, l := rgbToHSL(r, g, b)
	l *= factor
	if l > 1 {
		l = 1
	}
	if l < 0 {
		l = 0
	}
	nr, ng, nb := hslToRGB(h, s, l)
	c.ARGB = c.ARGB[:2] + colorHex(nr) + colorHex(ng) + colorHex(nb)
}

// applyLumOff adds offset to luminance (e.g. 0.25 = +25%).
func applyLumOff(c *Color, offset float64) {
	r, g, b := c.GetRed(), c.GetGreen(), c.GetBlue()
	h, s, l := rgbToHSL(r, g, b)
	l += offset
	if l > 1 {
		l = 1
	}
	if l < 0 {
		l = 0
	}
	nr, ng, nb := hslToRGB(h, s, l)
	c.ARGB = c.ARGB[:2] + colorHex(nr) + colorHex(ng) + colorHex(nb)
}

// srgbToLinear and linearToSRGB convert one colour channel between sRGB and
// the linear-light space PowerPoint mixes colours in. The transfer function is
// the piecewise sRGB EOTF, not a plain 2.2 power: the two disagree by enough to
// miss a measured channel by one.
func srgbToLinear(v float64) float64 {
	if v <= 0.04045 {
		return v / 12.92
	}
	return math.Pow((v+0.055)/1.055, 2.4)
}

func linearToSRGB(v float64) float64 {
	if v <= 0.0031308 {
		return v * 12.92
	}
	return 1.055*math.Pow(v, 1.0/2.4) - 0.055
}

// mixInLinearLight blends a channel toward `target` (1 = white, 0 = black) by
// `amount` in linear light.
func mixInLinearLight(v uint8, amount, target float64) uint8 {
	if amount < 0 {
		amount = 0
	}
	if amount > 1 {
		amount = 1
	}
	lin := amount*srgbToLinear(float64(v)/255.0) + (1-amount)*target
	if lin < 0 {
		lin = 0
	}
	if lin > 1 {
		lin = 1
	}
	return uint8(linearToSRGB(lin)*255.0 + 0.5)
}

// applyTint mixes the colour toward white by the given amount (0-1) — where
// `amount` is the weight of the *original* colour, so a tint of 1 is the colour
// and a tint of 0 is white.
//
// Two things here are not the obvious reading of the spec and were both
// measured off PowerPoint's own output (the r29 COM probe put a series of
// explicit tint/shade fills on a table and read the rendered pixels back):
//
//  1. The weight. <a:tint val="0"/> renders white and <a:tint val="100000"/>
//     renders the colour itself, so the percentage counts the colour that
//     survives, not the white that is added.
//  2. The space. The blend happens in linear light: accent1 #4F81BD with
//     tint 40000 is #D0D8E8 on screen, which the linear formula reproduces
//     exactly, where an sRGB blend at the same weight gives #B8CDE5 and the
//     old inverted-and-sRGB formula gave #95B3D7. Every channel of every
//     probe value (tint 0/20000/25000/40000/50000/75000/100000) lands within
//     one unit.
func applyTint(c *Color, amount float64) {
	r, g, b := c.GetRed(), c.GetGreen(), c.GetBlue()
	c.ARGB = c.ARGB[:2] + colorHex(mixInLinearLight(r, amount, 1.0)) +
		colorHex(mixInLinearLight(g, amount, 1.0)) +
		colorHex(mixInLinearLight(b, amount, 1.0))
}

// applyShade blends the colour toward black by the given amount (0-1), again in
// linear light. The weight here does count the colour (shade 50000 halves the
// linear channel: accent1 becomes #385D8A, which both spaces agree is "half",
// but the space still shows — #264264 for shade 25000 is the linear quarter).
func applyShade(c *Color, amount float64) {
	r, g, b := c.GetRed(), c.GetGreen(), c.GetBlue()
	c.ARGB = c.ARGB[:2] + colorHex(mixInLinearLight(r, amount, 0.0)) +
		colorHex(mixInLinearLight(g, amount, 0.0)) +
		colorHex(mixInLinearLight(b, amount, 0.0))
}

func colorHex(v uint8) string {
	const hex = "0123456789ABCDEF"
	return string([]byte{hex[v>>4], hex[v&0x0f]})
}

// presetColorToColor converts an OOXML preset color name to a Color.
// See ECMA-376 §20.1.10.47 for the full list.
func presetColorToColor(name string) Color {
	switch name {
	case "black":
		return ColorBlack
	case "white":
		return ColorWhite
	case "red":
		return ColorRed
	case "green":
		return ColorGreen
	case "blue":
		return ColorBlue
	case "yellow":
		return ColorYellow
	case "cyan", "aqua":
		return NewColor("FF00FFFF")
	case "magenta", "fuchsia":
		return NewColor("FFFF00FF")
	case "gray", "grey":
		return NewColor("FF808080")
	case "dkGray", "darkGray", "darkGrey":
		return NewColor("FFA9A9A9")
	case "ltGray", "lightGray", "lightGrey":
		return NewColor("FFD3D3D3")
	case "maroon":
		return NewColor("FF800000")
	case "olive":
		return NewColor("FF808000")
	case "navy":
		return NewColor("FF000080")
	case "purple":
		return NewColor("FF800080")
	case "teal":
		return NewColor("FF008080")
	case "silver":
		return NewColor("FFC0C0C0")
	case "orange":
		return NewColor("FFFFA500")
	default:
		return ColorBlack
	}
}
