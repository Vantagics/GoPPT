package gopresentation

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"image"
	"image/color"
	"io"
	"math"
	"sort"
	"strconv"
	"strings"
)

// --- SVG rasterisation -------------------------------------------------------
//
// The library stores picture bytes verbatim, so an SVG survives read and write
// unchanged. Drawing it is a separate problem: nothing in the dependency set
// rasterises SVG, and a picture that could not be drawn used to fall through to
// a bare grey frame — indistinguishable from a picture that legitimately has a
// grey border. A deck whose icons, arrows and connector lines are SVG therefore
// lost all of them from every preview, with nothing anywhere saying so.
//
// The subset below covers what design tooling actually emits for slide
// decoration: polygon, polyline, line, rect (including rounded), circle,
// ellipse, and path using M/L/H/V/C/S/Q/T/A/Z in absolute and relative form,
// with fill, stroke, both fill rules, opacity, and the common transforms.
// Anything outside the subset is refused with a reason rather than drawn wrong,
// so the caller can keep the picture visible as a labelled placeholder.
//
// Fill uses the nonzero winding rule for multi-subpath art. That matters for
// icon sets like Font Awesome, whose subpaths deliberately overlap: even-odd
// would punch holes where, say, a figure's torso crosses its neighbours, which
// is not what the icon looks like.

// svgSS is the supersampling factor. Geometry is filled with hard-edged
// scanline coverage, so anti-aliasing comes from rasterising large and
// box-averaging down; 3 gives 9 samples per output pixel, enough for icon and
// line art without the cost of a real coverage integrator.
const svgSS = 3

// maxSVGDimension caps the rasterisation buffer so a malformed presentation
// cannot ask for a multi-gigabyte allocation.
const maxSVGDimension = 8192

// maxSVGPathPoints caps how many flattened points one document may contribute,
// bounding both memory and the scanline sweep.
const maxSVGPathPoints = 200000

// svgPoint is a position in user units (or in device pixels, once transformed).
type svgPoint struct{ x, y float64 }

// svgMatrix is a 2x3 affine transform: x' = a*x + c*y + e, y' = b*x + d*y + f.
type svgMatrix struct {
	a, b, c, d, e, f float64
}

// mul composes two transforms so that the receiver is applied after n.
func (m svgMatrix) mul(n svgMatrix) svgMatrix {
	return svgMatrix{
		a: m.a*n.a + m.c*n.b,
		b: m.b*n.a + m.d*n.b,
		c: m.a*n.c + m.c*n.d,
		d: m.b*n.c + m.d*n.d,
		e: m.a*n.e + m.c*n.f + m.e,
		f: m.b*n.e + m.d*n.f + m.f,
	}
}

func (m svgMatrix) apply(p svgPoint) svgPoint {
	return svgPoint{m.a*p.x + m.c*p.y + m.e, m.b*p.x + m.d*p.y + m.f}
}

// uniformScale returns the factor by which the transform scales lengths, or 0
// when the transform is degenerate. Stroke widths are scaled by this, since a
// transform applies to the stroke as well as to the path.
func (m svgMatrix) uniformScale() float64 {
	return math.Sqrt(math.Abs(m.a*m.d - m.b*m.c))
}

// svgSubpath is one M...Z sequence, already flattened to line segments.
type svgSubpath struct {
	pts    []svgPoint
	closed bool
}

// svgCap is a stroke's line-cap style.
type svgCap int

const (
	svgCapButt svgCap = iota
	svgCapRound
	svgCapSquare
)

// svgStyle is a fully resolved paint state: presentation attributes inherit
// down the element tree and are resolved as the tree is walked, so no bitmask
// of "was this set here" is needed.
type svgStyle struct {
	fill        *svgPaint // nil means fill="none"
	stroke      *svgPaint // nil means no stroke
	strokeWidth float64   // user units
	lineCap     svgCap
	evenOdd     bool
}

// svgShape is one drawable element: its subpaths in user units, the element
// transform chain that applies to them, and its resolved paint state.
type svgShape struct {
	subpaths []svgSubpath
	xf       svgMatrix
	style    svgStyle
}

// svgDoc is a parsed document: the viewport mapping plus the shapes in
// document order, which is also paint order.
type svgDoc struct {
	viewBox    [4]float64
	hasViewBox bool
	width      float64
	height     float64
	noStretch  bool // preserveAspectRatio="none"
	shapes     []svgShape
}

// renderSVG rasterises an SVG document into a w x h image whose area outside the
// drawn geometry is transparent. It returns an error naming the construct it
// could not handle rather than approximating it, so the caller can decide
// whether to show a placeholder instead.
func renderSVG(data []byte, w, h int) (*image.RGBA, error) {
	if w <= 0 || h <= 0 {
		return nil, fmt.Errorf("svg: unusable target size %dx%d", w, h)
	}
	if w > maxSVGDimension || h > maxSVGDimension {
		return nil, fmt.Errorf("svg: target size %dx%d exceeds the %d px limit", w, h, maxSVGDimension)
	}
	doc, err := parseSVG(data, float64(w*svgSS), float64(h*svgSS))
	if err != nil {
		return nil, err
	}
	dw := float64(w * svgSS)
	dh := float64(h * svgSS)
	base := doc.deviceTransform(dw, dh)
	img := image.NewRGBA(image.Rect(0, 0, w*svgSS, h*svgSS))
	for i := range doc.shapes {
		doc.shapes[i].rasterise(img, base)
	}
	return downsampleBox(img, w, h), nil
}

// deviceTransform maps user units to the supersampled raster, honouring the
// viewBox and the default preserveAspectRatio="xMidYMid meet".
func (d *svgDoc) deviceTransform(dw, dh float64) svgMatrix {
	vbX, vbY, vbW, vbH := 0.0, 0.0, dw, dh
	switch {
	case d.hasViewBox:
		vbX, vbY, vbW, vbH = d.viewBox[0], d.viewBox[1], d.viewBox[2], d.viewBox[3]
	case d.width > 0 && d.height > 0:
		vbW, vbH = d.width, d.height
	}
	if vbW <= 0 || vbH <= 0 {
		return svgMatrix{a: 1, d: 1}
	}
	sx := dw / vbW
	sy := dh / vbH
	if d.noStretch {
		return svgMatrix{a: sx, d: sy, e: -vbX * sx, f: -vbY * sy}
	}
	s := math.Min(sx, sy)
	if !d.hasViewBox && d.width <= 0 {
		// No viewBox and no intrinsic size: user units already are device
		// pixels, so there is nothing to scale or centre.
		return svgMatrix{a: s, d: s}
	}
	return svgMatrix{
		a: s, d: s,
		e: (dw-vbW*s)/2 - vbX*s,
		f: (dh-vbH*s)/2 - vbY*s,
	}
}

// downsampleBox box-filters the supersampled raster down to the target size.
//
// Averaging the premultiplied samples image.RGBA stores is exactly right, but
// the 8-bit result must be written back premultiplied too, which is why the
// channels are summed before any rounding.
func downsampleBox(src *image.RGBA, w, h int) *image.RGBA {
	dst := image.NewRGBA(image.Rect(0, 0, w, h))
	inv := 1.0 / float64(svgSS*svgSS)
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			var sr, sg, sb, sa uint32
			for dy := 0; dy < svgSS; dy++ {
				row := src.Pix[(y*svgSS+dy)*src.Stride+(x*svgSS)*4:]
				for dx := 0; dx < svgSS; dx++ {
					sr += uint32(row[dx*4])
					sg += uint32(row[dx*4+1])
					sb += uint32(row[dx*4+2])
					sa += uint32(row[dx*4+3])
				}
			}
			o := y*dst.Stride + x*4
			dst.Pix[o+0] = round8(float64(sr) * inv)
			dst.Pix[o+1] = round8(float64(sg) * inv)
			dst.Pix[o+2] = round8(float64(sb) * inv)
			dst.Pix[o+3] = round8(float64(sa) * inv)
		}
	}
	return dst
}

func round8(v float64) uint8 {
	if v <= 0 {
		return 0
	}
	if v >= 255 {
		return 255
	}
	return uint8(v + 0.5)
}

// premultiply converts a straight-alpha colour into the premultiplied form an
// image.RGBA buffer stores.
func premultiply(r, g, b uint8, alpha float64) color.RGBA {
	if alpha < 0 {
		alpha = 0
	}
	if alpha > 1 {
		alpha = 1
	}
	a := round8(alpha * 255)
	f := alpha
	return color.RGBA{
		R: round8(float64(r) * f),
		G: round8(float64(g) * f),
		B: round8(float64(b) * f),
		A: a,
	}
}

// svgPaint is a straight-alpha colour. It is deliberately not a color.RGBA:
// that type is defined to hold premultiplied values, and premultiplication is
// the rasteriser's job, so keeping the two separate avoids the classic
// double-multiply that turns a 50% fill into a 25% one. `a` is 0..1.
type svgPaint struct {
	r, g, b uint8
	a       float64
}

// premultiply converts the paint into the form the raster buffer stores.
func (p svgPaint) premultiply() color.RGBA { return premultiply(p.r, p.g, p.b, p.a) }

// withAlpha returns a copy of the paint with its alpha scaled by f.
func (p svgPaint) withAlpha(f float64) svgPaint {
	p.a *= f
	if p.a < 0 {
		p.a = 0
	}
	if p.a > 1 {
		p.a = 1
	}
	return p
}

// --- parsing -----------------------------------------------------------------

// svgRootStyle is the initial paint state, matching SVG's initial values:
// solid black fill, no stroke, stroke-width 1.
var svgRootStyle = svgStyle{
	fill:        &svgPaint{a: 1},
	strokeWidth: 1,
}

// svgParser walks the document. Presentation attributes are resolved as the
// tree is descended, so each shape is emitted with a complete paint state.
type svgParser struct {
	doc       *svgDoc
	mats      []svgMatrix
	styles    []svgStyle
	pushed    []bool
	points    int
	reason    string // first construct outside the supported subset
	defsDepth int    // nesting depth of <defs>, whose content is not drawn
	deviceW   float64
	deviceH   float64
}

func (p *svgParser) top() svgMatrix  { return p.mats[len(p.mats)-1] }
func (p *svgParser) style() svgStyle { return p.styles[len(p.styles)-1] }

// refuse records the first unsupported construct. Reporting the first is
// enough: it already explains why the picture cannot be drawn.
func (p *svgParser) refuse(format string, args ...any) {
	if p.reason == "" {
		p.reason = fmt.Sprintf(format, args...)
	}
}

func (p *svgParser) ensureDoc() *svgDoc {
	if p.doc == nil {
		p.doc = &svgDoc{}
	}
	return p.doc
}

// flattenTolerance returns the curve-flattening tolerance in user units: half a
// supersampled pixel. Deriving it from the device scale is what stops a picture
// authored against a 640-unit viewBox but rendered at 55 px from being flattened
// with the step count appropriate to 640 px, and the reverse.
func (p *svgParser) flattenTolerance() float64 {
	dw, dh := p.deviceW, p.deviceH
	vbW, vbH := dw, dh
	if p.doc != nil {
		switch {
		case p.doc.hasViewBox:
			vbW, vbH = p.doc.viewBox[2], p.doc.viewBox[3]
		case p.doc.width > 0 && p.doc.height > 0:
			vbW, vbH = p.doc.width, p.doc.height
		}
	}
	scale := 1.0
	if vbW > 0 && vbH > 0 && dw > 0 && dh > 0 {
		scale = math.Min(dw/vbW, dh/vbH)
	}
	if scale <= 0 {
		scale = 1
	}
	tol := 0.5 / scale
	if tol < 1e-4 {
		tol = 1e-4
	}
	if tol > 1e3 {
		tol = 1e3
	}
	return tol
}

func parseSVG(data []byte, deviceW, deviceH float64) (*svgDoc, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	p := &svgParser{deviceW: deviceW, deviceH: deviceH}
	p.mats = append(p.mats, svgMatrix{a: 1, d: 1})
	p.styles = append(p.styles, svgRootStyle)

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("svg: cannot parse XML: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			p.pushed = append(p.pushed, p.start(t))
		case xml.EndElement:
			p.end(t.Name.Local)
			if n := len(p.pushed); n > 0 {
				if p.pushed[n-1] {
					p.mats = p.mats[:len(p.mats)-1]
					p.styles = p.styles[:len(p.styles)-1]
				}
				p.pushed = p.pushed[:n-1]
			}
		}
	}
	if p.reason != "" {
		return nil, fmt.Errorf("svg: %s", p.reason)
	}
	if p.doc == nil {
		return nil, fmt.Errorf("svg: no <svg> element found")
	}
	if len(p.doc.shapes) == 0 {
		return nil, fmt.Errorf("svg: nothing drawable in the supported subset")
	}
	return p.doc, nil
}

// start handles one element and reports whether it pushed a scope onto the
// transform and style stacks, which the matching end tag must pop.
func (p *svgParser) start(t xml.StartElement) bool {
	name := strings.ToLower(t.Name.Local)
	switch name {
	case "svg":
		if p.doc == nil {
			p.initRoot(t)
			return false
		}
		// A nested <svg> establishes a new viewport; composing its transform is
		// not supported, so treat it as a plain group rather than misplacing
		// its content.
		p.refuse("nested <svg> element")
		return p.pushGroup(t)
	case "g":
		return p.pushGroup(t)
	case "defs":
		// Content of <defs> is a definition, not content: it is drawn only
		// where something references it, and the referencing element, <use>, is
		// refused. Drawing it anyway would scatter artwork that the document
		// deliberately keeps off the canvas — a sprite sheet or a set of
		// reusable symbols — across the picture.
		p.defsDepth++
		return false
	case "polygon", "polyline", "line", "rect", "circle", "ellipse", "path":
		if p.defsDepth > 0 {
			return false
		}
		p.emitShape(name, t)
		return false
	default:
		p.unsupportedElement(name)
		return false
	}
}

// end closes the scope opened by start for the given element name.
func (p *svgParser) end(name string) {
	if strings.EqualFold(name, "defs") && p.defsDepth > 0 {
		p.defsDepth--
	}
}

func (p *svgParser) pushGroup(t xml.StartElement) bool {
	p.mats = append(p.mats, p.top().mul(p.parseTransform(attrOf(t, "transform"))))
	p.styles = append(p.styles, p.styleFor(t))
	return true
}

func (p *svgParser) initRoot(t xml.StartElement) {
	d := &svgDoc{}
	if vb, ok := parseNumberList(attrOf(t, "viewBox"), 4); ok && vb[2] > 0 && vb[3] > 0 {
		copy(d.viewBox[:], vb)
		d.hasViewBox = true
	}
	d.width = parseLength(attrOf(t, "width"))
	d.height = parseLength(attrOf(t, "height"))
	switch v := strings.TrimSpace(attrOf(t, "preserveAspectRatio")); v {
	case "", "xMidYMid", "xMidYMid meet":
	case "none":
		d.noStretch = true
	default:
		p.refuse("preserveAspectRatio=%q is not supported", v)
	}
	p.styles[0] = p.styleFor(t)
	p.doc = d
}

// unsupportedElement classifies an element the rasteriser does not draw.
//
// name is already lower-cased by the caller, so every label here is too. SVG
// spells several of these elements in camel case — clipPath, linearGradient,
// foreignObject — and matching the document's spelling against a lower-cased
// name silently sent them to the default branch below.
func (p *svgParser) unsupportedElement(name string) {
	switch name {
	case "", "title", "desc", "metadata", "defs":
		// No visual effect on their own. `defs` is inert unless something
		// references it, and the referencing element, <use>, is refused below.
	case "text", "tspan", "textpath":
		p.refuse("<%s> text is not supported", name)
	case "image":
		p.refuse("nested <image> is not supported")
	case "use":
		p.refuse("<use> references are not supported")
	case "style":
		p.refuse("embedded CSS <style> is not supported")
	case "lineargradient", "radialgradient", "stop":
		// Paint servers. Their definitions are inert, and a shape that actually
		// paints with one is refused by parsePaint, which is where the failure
		// is worth reporting.
	case "mask", "clippath", "filter", "marker", "pattern", "symbol", "switch", "foreignobject", "script":
		p.refuse("<%s> is not supported", name)
	default:
		// Unknown elements are refused rather than ignored: an unrecognised
		// element is far more likely to be a construct that would have drawn
		// something than a no-op.
		p.refuse("<%s> is not supported", name)
	}
}

// emitShape builds one shape from an element and appends it to the document.
func (p *svgParser) emitShape(name string, t xml.StartElement) {
	st := p.styleFor(t)
	// Geometry that nothing paints is invisible, not an error.
	if st.fill == nil && (st.stroke == nil || st.strokeWidth <= 0) {
		return
	}
	sub := p.shapeSubpaths(name, t)
	if len(sub) == 0 {
		return
	}
	for _, sp := range sub {
		p.points += len(sp.pts)
	}
	if p.points > maxSVGPathPoints {
		p.refuse("path data exceeds the %d point limit", maxSVGPathPoints)
		return
	}
	d := p.ensureDoc()
	xf := p.top().mul(p.parseTransform(attrOf(t, "transform")))
	d.shapes = append(d.shapes, svgShape{subpaths: sub, xf: xf, style: st})
}

// styleFor resolves an element's presentation attributes against the inherited
// style. The CSS form wins over the attribute form, which is what the cascade
// requires.
func (p *svgParser) styleFor(t xml.StartElement) svgStyle {
	st := p.style()
	props := elementProps(t)

	if v, ok := props["fill"]; ok {
		c, err := p.parsePaint(v)
		if err != nil {
			p.refuse("%v", err)
		} else {
			st.fill = c
		}
	}
	if v, ok := props["stroke"]; ok {
		c, err := p.parsePaint(v)
		if err != nil {
			p.refuse("%v", err)
		} else {
			st.stroke = c
		}
	}
	if v, ok := props["stroke-width"]; ok {
		if f, ok := parseUnitFloat(v); ok && f >= 0 {
			st.strokeWidth = f
		}
	}
	if v, ok := props["fill-rule"]; ok {
		switch strings.TrimSpace(v) {
		case "evenodd":
			st.evenOdd = true
		case "nonzero":
			st.evenOdd = false
		default:
			p.refuse("fill-rule=%q is not supported", v)
		}
	}
	if v, ok := props["stroke-linecap"]; ok {
		switch strings.TrimSpace(v) {
		case "butt":
			st.lineCap = svgCapButt
		case "round":
			st.lineCap = svgCapRound
		case "square":
			st.lineCap = svgCapSquare
		default:
			p.refuse("stroke-linecap=%q is not supported", v)
		}
	}
	// Opacity is folded into the paint's alpha. The group `opacity` property
	// really applies to a composited layer, but for the fill and stroke of a
	// single element the two are equivalent.
	if v, ok := props["fill-opacity"]; ok && st.fill != nil {
		if f, ok := parseUnitFloat(v); ok {
			c := st.fill.withAlpha(f)
			st.fill = &c
		}
	}
	if v, ok := props["stroke-opacity"]; ok && st.stroke != nil {
		if f, ok := parseUnitFloat(v); ok {
			c := st.stroke.withAlpha(f)
			st.stroke = &c
		}
	}
	if v, ok := props["opacity"]; ok {
		if f, ok := parseUnitFloat(v); ok {
			if st.fill != nil {
				c := st.fill.withAlpha(f)
				st.fill = &c
			}
			if st.stroke != nil {
				c := st.stroke.withAlpha(f)
				st.stroke = &c
			}
		}
	}
	return st
}

// --- attribute and value helpers ---------------------------------------------

// svgNS reports whether a namespace URI is SVG's. Attributes with no prefix
// carry an empty Space, which is the common case and must be accepted.
func svgNS(space string) bool {
	return space == "" || strings.Contains(strings.ToLower(space), "svg")
}

// elementProps collects an element's presentation properties from both the
// attribute form (fill="#fff") and the CSS form (style="fill:#fff"). The CSS
// form is applied second so that it wins regardless of attribute order, which
// is what the cascade requires.
func elementProps(t xml.StartElement) map[string]string {
	props := make(map[string]string, len(t.Attr)+4)
	for _, a := range t.Attr {
		if !svgNS(a.Name.Space) {
			continue
		}
		name := strings.ToLower(a.Name.Local)
		if name == "style" {
			continue
		}
		props[name] = strings.TrimSpace(a.Value)
	}
	for _, a := range t.Attr {
		if strings.ToLower(a.Name.Local) != "style" {
			continue
		}
		for k, v := range parseStyleAttr(a.Value) {
			props[k] = v
		}
	}
	return props
}

// parseStyleAttr splits a CSS style attribute into declarations. A style
// attribute holds only simple "name: value" pairs, so no full CSS parser is
// needed.
func parseStyleAttr(v string) map[string]string {
	out := make(map[string]string, 4)
	for _, decl := range strings.Split(v, ";") {
		i := strings.IndexByte(decl, ':')
		if i < 0 {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(decl[:i]))
		val := strings.TrimSpace(decl[i+1:])
		if name == "" || val == "" {
			continue
		}
		out[name] = val
	}
	return out
}

func attrOf(t xml.StartElement, name string) string {
	for _, a := range t.Attr {
		if strings.EqualFold(a.Name.Local, name) {
			return a.Value
		}
	}
	return ""
}

// parseLength reads a length attribute as a plain number, accepting a px unit.
// Percentages and font-relative units depend on context this rasteriser does not
// track, so they report 0 and the caller falls back to the viewBox.
func parseLength(v string) float64 {
	v = strings.TrimSpace(v)
	if i := strings.IndexFunc(v, func(r rune) bool {
		return r != '.' && (r < '0' || r > '9') && r != '-' && r != '+' && r != 'e' && r != 'E'
	}); i >= 0 {
		v = v[:i]
	}
	f, err := strconv.ParseFloat(strings.TrimSpace(v), 64)
	if err != nil {
		return 0
	}
	return f
}

// parseUnitFloat reads a number or a percentage, the latter returned as a
// fraction. That is the pair of forms opacity and stroke-width accept.
func parseUnitFloat(v string) (float64, bool) {
	v = strings.TrimSpace(v)
	if strings.HasSuffix(v, "%") {
		f, err := strconv.ParseFloat(strings.TrimSpace(strings.TrimSuffix(v, "%")), 64)
		if err != nil {
			return 0, false
		}
		return f / 100, true
	}
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return 0, false
	}
	return f, true
}

// parseNumberListRaw splits a whitespace/comma separated number list.
func parseNumberListRaw(s string) []float64 {
	fields := strings.FieldsFunc(s, func(r rune) bool {
		return r == ' ' || r == ',' || r == '\t' || r == '\n' || r == '\r'
	})
	out := make([]float64, 0, len(fields))
	for _, f := range fields {
		v, err := strconv.ParseFloat(f, 64)
		if err != nil {
			return out
		}
		out = append(out, v)
	}
	return out
}

// parseNumberList requires at least n numbers and fails otherwise.
func parseNumberList(s string, n int) ([]float64, bool) {
	out := parseNumberListRaw(s)
	if len(out) < n {
		return nil, false
	}
	return out, true
}

// numAt indexes a number list, returning def when the index is absent.
func numAt(v []float64, i int, def float64) float64 {
	if i >= 0 && i < len(v) {
		return v[i]
	}
	return def
}

// parseTransform composes an SVG transform list. The list applies left to right
// as nested coordinate systems, so each item is composed onto the running
// result rather than replacing it.
func (p *svgParser) parseTransform(s string) svgMatrix {
	m := svgMatrix{a: 1, d: 1}
	if strings.TrimSpace(s) == "" {
		return m
	}
	for i := 0; i < len(s); {
		for i < len(s) && !isASCIINameStart(s[i]) {
			i++
		}
		if i >= len(s) {
			break
		}
		start := i
		for i < len(s) && isASCIINameChar(s[i]) {
			i++
		}
		name := strings.ToLower(s[start:i])
		for i < len(s) && s[i] != '(' {
			i++
		}
		if i >= len(s) {
			p.refuse("malformed transform=%q", s)
			break
		}
		i++
		argStart := i
		for i < len(s) && s[i] != ')' {
			i++
		}
		args := parseNumberListRaw(s[argStart:i])
		if i < len(s) {
			i++
		}
		switch name {
		case "translate":
			m = m.mul(svgMatrix{a: 1, d: 1, e: numAt(args, 0, 0), f: numAt(args, 1, 0)})
		case "scale":
			sx := numAt(args, 0, 1)
			m = m.mul(svgMatrix{a: sx, d: numAt(args, 1, sx)})
		case "rotate":
			r := numAt(args, 0, 0) * math.Pi / 180
			rot := svgMatrix{a: math.Cos(r), b: math.Sin(r), c: -math.Sin(r), d: math.Cos(r)}
			if len(args) >= 3 {
				cx, cy := args[1], args[2]
				m = m.mul(svgMatrix{a: 1, d: 1, e: cx, f: cy}).mul(rot).mul(svgMatrix{a: 1, d: 1, e: -cx, f: -cy})
			} else {
				m = m.mul(rot)
			}
		case "matrix":
			m = m.mul(svgMatrix{
				a: numAt(args, 0, 1), b: numAt(args, 1, 0),
				c: numAt(args, 2, 0), d: numAt(args, 3, 1),
				e: numAt(args, 4, 0), f: numAt(args, 5, 0),
			})
		case "skewx", "skewy":
			// A skew cannot be represented as anything but itself, and
			// ignoring it would silently misplace the content.
			p.refuse("transform skew is not supported")
		default:
			p.refuse("transform %q is not supported", name)
		}
	}
	return m
}

func isASCIINameStart(c byte) bool {
	return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')
}

func isASCIINameChar(c byte) bool {
	return isASCIINameStart(c) || (c >= '0' && c <= '9')
}

// --- paint parsing -----------------------------------------------------------

// svgNamedColors covers SVG's basic colour keywords.
var svgNamedColors = map[string]svgPaint{
	"black": {a: 1}, "silver": {r: 192, g: 192, b: 192, a: 1},
	"gray": {r: 128, g: 128, b: 128, a: 1}, "grey": {r: 128, g: 128, b: 128, a: 1},
	"white": {r: 255, g: 255, b: 255, a: 1}, "maroon": {r: 128, a: 1},
	"red": {r: 255, a: 1}, "purple": {r: 128, b: 128, a: 1},
	"fuchsia": {r: 255, b: 255, a: 1}, "magenta": {r: 255, b: 255, a: 1},
	"green": {g: 128, a: 1}, "lime": {g: 255, a: 1},
	"olive": {r: 128, g: 128, a: 1}, "yellow": {r: 255, g: 255, a: 1},
	"navy": {b: 128, a: 1}, "blue": {b: 255, a: 1},
	"teal": {g: 128, b: 128, a: 1}, "aqua": {g: 255, b: 255, a: 1},
	"cyan": {g: 255, b: 255, a: 1}, "orange": {r: 255, g: 165, a: 1},
	"pink": {r: 255, g: 192, b: 203, a: 1}, "brown": {r: 165, g: 42, b: 42, a: 1},
	"gold": {r: 255, g: 215, a: 1}, "indigo": {r: 75, b: 130, a: 1},
	"violet": {r: 238, g: 130, b: 238, a: 1}, "beige": {r: 245, g: 245, b: 220, a: 1},
	"ivory": {r: 255, g: 255, b: 240, a: 1}, "khaki": {r: 240, g: 230, b: 140, a: 1},
	"crimson": {r: 220, g: 20, b: 60, a: 1}, "salmon": {r: 250, g: 128, b: 114, a: 1},
	"coral": {r: 255, g: 127, b: 80, a: 1}, "tomato": {r: 255, g: 99, b: 71, a: 1},
	"orchid": {r: 218, g: 112, b: 214, a: 1}, "plum": {r: 221, g: 160, b: 221, a: 1},
	"turquoise": {r: 64, g: 224, b: 208, a: 1}, "skyblue": {r: 135, g: 206, b: 235, a: 1},
	"steelblue": {r: 70, g: 130, b: 180, a: 1}, "royalblue": {r: 65, g: 105, b: 225, a: 1},
	"midnightblue": {r: 25, g: 25, b: 112, a: 1}, "slategray": {r: 112, g: 128, b: 144, a: 1},
	"slategrey": {r: 112, g: 128, b: 144, a: 1}, "whitesmoke": {r: 245, g: 245, b: 245, a: 1},
	"gainsboro": {r: 220, g: 220, b: 220, a: 1}, "lightgray": {r: 211, g: 211, b: 211, a: 1},
	"lightgrey": {r: 211, g: 211, b: 211, a: 1}, "darkgray": {r: 169, g: 169, b: 169, a: 1},
	"darkgrey": {r: 169, g: 169, b: 169, a: 1}, "dimgray": {r: 105, g: 105, b: 105, a: 1},
	"dimgrey": {r: 105, g: 105, b: 105, a: 1}, "darkblue": {b: 139, a: 1},
	"darkgreen": {g: 100, a: 1}, "darkred": {r: 139, a: 1},
	"lightblue": {r: 173, g: 216, b: 230, a: 1}, "lightgreen": {r: 144, g: 238, b: 144, a: 1},
	"transparent": {a: 0},
}

// parsePaint resolves a fill or stroke value. A nil colour with a nil error
// means "nothing is painted", which is what `none` means; a non-nil error means
// the value is outside the supported subset and the caller refuses the picture.
func (p *svgParser) parsePaint(v string) (*svgPaint, error) {
	v = strings.TrimSpace(v)
	if v == "" {
		return nil, fmt.Errorf("empty paint value")
	}
	low := strings.ToLower(v)
	switch low {
	case "none":
		return nil, nil
	case "currentcolor":
		// No inherited `color` property is tracked; black is SVG's initial
		// value, so this is the documented default rather than a guess.
		return &svgPaint{a: 1}, nil
	}
	if strings.HasPrefix(low, "url(") {
		return nil, fmt.Errorf("gradient or pattern paint %q is not supported", v)
	}
	if strings.HasPrefix(low, "#") {
		if c, ok := parseHexPaint(low); ok {
			return c, nil
		}
		return nil, fmt.Errorf("paint %q is not supported", v)
	}
	if strings.HasPrefix(low, "rgb") {
		if c, ok := parseRGBPaint(low); ok {
			return c, nil
		}
		return nil, fmt.Errorf("paint %q is not supported", v)
	}
	if c, ok := svgNamedColors[low]; ok {
		return &c, nil
	}
	return nil, fmt.Errorf("colour %q is not supported", v)
}

// parseHexPaint reads #rgb, #rgba, #rrggbb and #rrggbbaa.
func parseHexPaint(s string) (*svgPaint, bool) {
	h := strings.TrimPrefix(s, "#")
	hexVal := func(part string) (uint8, bool) {
		n, err := strconv.ParseUint(part, 16, 8)
		if err != nil {
			return 0, false
		}
		return uint8(n), true
	}
	switch len(h) {
	case 3, 4:
		var out [4]uint8
		out[3] = 255
		for i := 0; i < len(h); i++ {
			v, ok := hexVal(strings.Repeat(string(h[i]), 2))
			if !ok {
				return nil, false
			}
			out[i] = v
		}
		return &svgPaint{r: out[0], g: out[1], b: out[2], a: float64(out[3]) / 255}, true
	case 6, 8:
		var out [4]uint8
		out[3] = 255
		for i := 0; i < len(h)/2; i++ {
			v, ok := hexVal(h[i*2 : i*2+2])
			if !ok {
				return nil, false
			}
			out[i] = v
		}
		return &svgPaint{r: out[0], g: out[1], b: out[2], a: float64(out[3]) / 255}, true
	}
	return nil, false
}

// parseRGBPaint reads rgb(r,g,b), rgb(r%,g%,b%) and the rgba(...) forms.
func parseRGBPaint(s string) (*svgPaint, bool) {
	open := strings.IndexByte(s, '(')
	close := strings.LastIndexByte(s, ')')
	if open < 0 || close < open {
		return nil, false
	}
	name := strings.TrimSpace(s[:open])
	if name != "rgb" && name != "rgba" {
		return nil, false
	}
	args := strings.FieldsFunc(s[open+1:close], func(r rune) bool {
		return r == ',' || r == ' ' || r == '/' || r == '\t' || r == '\n'
	})
	if len(args) < 3 {
		return nil, false
	}
	comp := func(v string) (uint8, bool) {
		f, ok := parseUnitFloat(v)
		if !ok {
			return 0, false
		}
		if strings.HasSuffix(strings.TrimSpace(v), "%") {
			f *= 255
		}
		if f < 0 {
			f = 0
		}
		if f > 255 {
			f = 255
		}
		return uint8(f + 0.5), true
	}
	r, ok1 := comp(args[0])
	g, ok2 := comp(args[1])
	b, ok3 := comp(args[2])
	if !ok1 || !ok2 || !ok3 {
		return nil, false
	}
	alpha := 1.0
	if len(args) >= 4 {
		f, ok := parseUnitFloat(args[3])
		if !ok {
			return nil, false
		}
		alpha = f
	}
	return &svgPaint{r: r, g: g, b: b, a: alpha}, true
}

// --- geometry ----------------------------------------------------------------

// svgKappa is the circle-to-cubic-Bezier constant, 4/3*(sqrt(2)-1). One cubic
// with these control offsets approximates a quarter circle to within about
// 0.02% of the radius, which is far below one supersampled pixel.
const svgKappa = 0.5522847498307936

// appendCubic appends the cubic Bezier (p0 -> p3) to pts as n line segments,
// skipping p0 itself because the caller has already placed it.
func appendCubic(pts []svgPoint, p0, c1, c2, p3 svgPoint, n int) []svgPoint {
	if n < 2 {
		n = 2
	}
	for i := 1; i <= n; i++ {
		t := float64(i) / float64(n)
		u := 1 - t
		pts = append(pts, svgPoint{
			x: u*u*u*p0.x + 3*u*u*t*c1.x + 3*u*t*t*c2.x + t*t*t*p3.x,
			y: u*u*u*p0.y + 3*u*u*t*c1.y + 3*u*t*t*c2.y + t*t*t*p3.y,
		})
	}
	return pts
}

// cubicSteps picks a segment count for one cubic from the length of its control
// polygon and the flattening tolerance. The control polygon is always at least
// as long as the curve, so this errs towards over-sampling, which is the safe
// direction: the cost is a few hundred extra edges, the benefit is that a
// picture rendered larger than its author drew it does not facet.
func cubicSteps(p0, c1, c2, p3 svgPoint, tol float64) int {
	if tol <= 0 {
		tol = 0.5
	}
	approx := svgDist(p0, c1) + svgDist(c1, c2) + svgDist(c2, p3)
	n := int(approx/tol) + 2
	if n < 4 {
		n = 4
	}
	if n > 96 {
		n = 96
	}
	return n
}

func svgDist(a, b svgPoint) float64 {
	dx, dy := b.x-a.x, b.y-a.y
	return math.Sqrt(dx*dx + dy*dy)
}

// ellipseSubpath builds a closed subpath approximating an ellipse.
func ellipseSubpath(cx, cy, rx, ry float64, tol float64) svgSubpath {
	steps := func(a, b svgPoint) int {
		return cubicSteps(a, a, b, b, tol)
	}
	right := svgPoint{cx + rx, cy}
	bottom := svgPoint{cx, cy + ry}
	left := svgPoint{cx - rx, cy}
	top := svgPoint{cx, cy - ry}

	pts := []svgPoint{right}
	pts = appendCubic(pts, right, svgPoint{cx + rx, cy + svgKappa*ry}, svgPoint{cx + svgKappa*rx, cy + ry}, bottom, steps(right, bottom))
	pts = appendCubic(pts, bottom, svgPoint{cx - svgKappa*rx, cy + ry}, svgPoint{cx - rx, cy + svgKappa*ry}, left, steps(bottom, left))
	pts = appendCubic(pts, left, svgPoint{cx - rx, cy - svgKappa*ry}, svgPoint{cx - svgKappa*rx, cy - ry}, top, steps(left, top))
	pts = appendCubic(pts, top, svgPoint{cx + svgKappa*rx, cy - ry}, svgPoint{cx + rx, cy - svgKappa*ry}, right, steps(top, right))
	return svgSubpath{pts: pts, closed: true}
}

// shapeSubpaths builds the geometry for one shape element, already flattened to
// line segments in user units.
func (p *svgParser) shapeSubpaths(name string, t xml.StartElement) []svgSubpath {
	tol := p.flattenTolerance()
	switch name {
	case "path":
		sub, ok := svgPathSubpaths(attrOf(t, "d"), tol)
		if !ok {
			p.refuse("path data uses a command outside the supported subset")
			return nil
		}
		return sub
	case "polygon", "polyline":
		nums := parseNumberListRaw(attrOf(t, "points"))
		var pts []svgPoint
		for i := 0; i+1 < len(nums); i += 2 {
			pts = append(pts, svgPoint{nums[i], nums[i+1]})
		}
		if len(pts) < 2 {
			return nil
		}
		return []svgSubpath{{pts: pts, closed: name == "polygon"}}
	case "line":
		p0 := svgPoint{parseLength(attrOf(t, "x1")), parseLength(attrOf(t, "y1"))}
		p1 := svgPoint{parseLength(attrOf(t, "x2")), parseLength(attrOf(t, "y2"))}
		if p0 == p1 {
			return nil
		}
		return []svgSubpath{{pts: []svgPoint{p0, p1}}}
	case "rect":
		return svgRectSubpaths(t, tol)
	case "circle":
		r := parseLength(attrOf(t, "r"))
		if r <= 0 {
			return nil
		}
		return []svgSubpath{ellipseSubpath(parseLength(attrOf(t, "cx")), parseLength(attrOf(t, "cy")), r, r, tol)}
	case "ellipse":
		rx := parseLength(attrOf(t, "rx"))
		ry := parseLength(attrOf(t, "ry"))
		if rx <= 0 || ry <= 0 {
			return nil
		}
		return []svgSubpath{ellipseSubpath(parseLength(attrOf(t, "cx")), parseLength(attrOf(t, "cy")), rx, ry, tol)}
	}
	return nil
}

// svgRectSubpaths builds a rectangle, with rounded corners when rx/ry are set.
func svgRectSubpaths(t xml.StartElement, tol float64) []svgSubpath {
	x := parseLength(attrOf(t, "x"))
	y := parseLength(attrOf(t, "y"))
	w := parseLength(attrOf(t, "width"))
	h := parseLength(attrOf(t, "height"))
	if w <= 0 || h <= 0 {
		return nil
	}
	rx := parseLength(attrOf(t, "rx"))
	ry := parseLength(attrOf(t, "ry"))
	if rx == 0 {
		rx = ry
	}
	if ry == 0 {
		ry = rx
	}
	if rx > w/2 {
		rx = w / 2
	}
	if ry > h/2 {
		ry = h / 2
	}
	if rx <= 0 || ry <= 0 {
		return []svgSubpath{{pts: []svgPoint{
			{x, y}, {x + w, y}, {x + w, y + h}, {x, y + h},
		}, closed: true}}
	}
	// Clockwise from the top-left corner's end, using one cubic per corner.
	var pts []svgPoint
	pts = append(pts, svgPoint{x + rx, y})
	pts = append(pts, svgPoint{x + w - rx, y})
	n := cubicSteps(svgPoint{x + w - rx, y}, svgPoint{x + w, y}, svgPoint{x + w, y + ry}, svgPoint{x + w, y + ry}, tol)
	pts = appendCubic(pts, svgPoint{x + w - rx, y}, svgPoint{x + w - rx + svgKappa*rx, y}, svgPoint{x + w, y + ry - svgKappa*ry}, svgPoint{x + w, y + ry}, n)
	pts = append(pts, svgPoint{x + w, y + h - ry})
	pts = appendCubic(pts, svgPoint{x + w, y + h - ry}, svgPoint{x + w, y + h - ry + svgKappa*ry}, svgPoint{x + w - rx + svgKappa*rx, y + h}, svgPoint{x + w - rx, y + h}, n)
	pts = append(pts, svgPoint{x + rx, y + h})
	pts = appendCubic(pts, svgPoint{x + rx, y + h}, svgPoint{x + rx - svgKappa*rx, y + h}, svgPoint{x, y + h - ry + svgKappa*ry}, svgPoint{x, y + h - ry}, n)
	pts = append(pts, svgPoint{x, y + ry})
	pts = appendCubic(pts, svgPoint{x, y + ry}, svgPoint{x, y + ry - svgKappa*ry}, svgPoint{x + rx - svgKappa*rx, y}, svgPoint{x + rx, y}, n)
	return []svgSubpath{{pts: pts, closed: true}}
}

// --- path data ---------------------------------------------------------------

// svgPathSubpaths parses a path `d` attribute into flattened subpaths. The
// second result is false when the data uses a construct this rasteriser does not
// implement, so the caller can refuse the whole picture instead of drawing part
// of it.
func svgPathSubpaths(d string, tol float64) ([]svgSubpath, bool) {
	if strings.TrimSpace(d) == "" {
		return nil, true
	}
	b := &svgPathBuilder{tol: tol}
	b.run(d)
	if b.failed {
		return nil, false
	}
	return b.subs, true
}

// svgPathScanner walks a path `d` string. Numbers may be separated by
// whitespace, by commas, or by nothing at all when the next one starts with a
// sign or a dot — "10-20" and ".5.5" are two numbers each — so splitting on
// separators is not enough and the scanner is written by hand.
type svgPathScanner struct {
	s string
	i int
}

func (s *svgPathScanner) skipSep() {
	for s.i < len(s.s) {
		switch s.s[s.i] {
		case ' ', '\t', '\n', '\r', '\f', ',':
			s.i++
		default:
			return
		}
	}
}

// number reads one number, consuming nothing when there is not one here.
func (s *svgPathScanner) number() (float64, bool) {
	s.skipSep()
	start := s.i
	if s.i < len(s.s) && (s.s[s.i] == '-' || s.s[s.i] == '+') {
		s.i++
	}
	digits := false
	for s.i < len(s.s) && s.s[s.i] >= '0' && s.s[s.i] <= '9' {
		s.i++
		digits = true
	}
	if s.i < len(s.s) && s.s[s.i] == '.' {
		s.i++
		for s.i < len(s.s) && s.s[s.i] >= '0' && s.s[s.i] <= '9' {
			s.i++
			digits = true
		}
	}
	if !digits {
		s.i = start
		return 0, false
	}
	if s.i < len(s.s) && (s.s[s.i] == 'e' || s.s[s.i] == 'E') {
		j := s.i + 1
		if j < len(s.s) && (s.s[j] == '-' || s.s[j] == '+') {
			j++
		}
		if j < len(s.s) && s.s[j] >= '0' && s.s[j] <= '9' {
			for j < len(s.s) && s.s[j] >= '0' && s.s[j] <= '9' {
				j++
			}
			s.i = j
		}
	}
	v, err := strconv.ParseFloat(s.s[start:s.i], 64)
	if err != nil {
		s.i = start
		return 0, false
	}
	return v, true
}

// numbers reads exactly n numbers, or reports failure.
func (s *svgPathScanner) numbers(n int) ([]float64, bool) {
	out := make([]float64, 0, n)
	for i := 0; i < n; i++ {
		v, ok := s.number()
		if !ok {
			return nil, false
		}
		out = append(out, v)
	}
	return out, true
}

func isPathCommand(c byte) bool {
	switch c {
	case 'M', 'm', 'L', 'l', 'H', 'h', 'V', 'v', 'C', 'c', 'S', 's',
		'Q', 'q', 'T', 't', 'A', 'a', 'Z', 'z':
		return true
	}
	return false
}

func addPt(a, b svgPoint) svgPoint { return svgPoint{a.x + b.x, a.y + b.y} }

// reflectPt mirrors about through through, which is how an S or T command
// derives its implied control point from the previous one.
func reflectPt(about, through svgPoint) svgPoint {
	return svgPoint{2*through.x - about.x, 2*through.y - about.y}
}

type svgPathBuilder struct {
	subs  []svgSubpath
	cur   []svgPoint
	pos   svgPoint
	start svgPoint
	// lastCtl is the previous cubic's second control point and lastQuad the
	// previous quadratic's control point, kept so S and T can reflect them.
	lastCtl  svgPoint
	lastQuad svgPoint
	prev     byte
	tol      float64
	failed   bool
}

// run consumes the path data. An implicit repetition of the previous command is
// how SVG allows "L 1 2 3 4" to mean two line segments, and after a moveto the
// repeated command is a lineto — which is why the command byte is rewritten
// rather than remembered separately.
func (b *svgPathBuilder) run(d string) {
	sc := &svgPathScanner{s: d}
	var cmd byte
	for {
		sc.skipSep()
		if sc.i >= len(sc.s) {
			break
		}
		if c := sc.s[sc.i]; isPathCommand(c) {
			cmd = c
			sc.i++
			if cmd == 'Z' || cmd == 'z' {
				b.closePath()
				cmd = 0
			}
			continue
		}
		if cmd == 0 {
			b.failed = true
			return
		}
		if !b.exec(cmd, sc) {
			b.failed = true
			return
		}
		switch cmd {
		case 'M':
			cmd = 'L'
		case 'm':
			cmd = 'l'
		}
	}
	b.flush(false)
}

func (b *svgPathBuilder) flush(closed bool) {
	if len(b.cur) >= 2 {
		b.subs = append(b.subs, svgSubpath{pts: b.cur, closed: closed})
	}
	b.cur = nil
}

func (b *svgPathBuilder) closePath() {
	b.flush(true)
	b.pos = b.start
	b.prev = 'z'
}

// lineTo starts a subpath at the current point when one is not already open,
// which is what a drawing command after a closepath must do.
func (b *svgPathBuilder) lineTo(p svgPoint) {
	if len(b.cur) == 0 {
		b.cur = append(b.cur, b.pos)
	}
	b.cur = append(b.cur, p)
	b.pos = p
}

func (b *svgPathBuilder) cubicTo(c1, c2, p svgPoint) {
	if len(b.cur) == 0 {
		b.cur = append(b.cur, b.pos)
	}
	b.cur = appendCubic(b.cur, b.pos, c1, c2, p, cubicSteps(b.pos, c1, c2, p, b.tol))
	b.lastCtl = c2
	b.pos = p
}

// quadTo emits a quadratic by promoting it to the equivalent cubic, whose
// control points are the de Casteljau midpoints of the quadratic's.
func (b *svgPathBuilder) quadTo(q, p svgPoint) {
	c1 := svgPoint{b.pos.x + 2.0/3.0*(q.x-b.pos.x), b.pos.y + 2.0/3.0*(q.y-b.pos.y)}
	c2 := svgPoint{p.x + 2.0/3.0*(q.x-p.x), p.y + 2.0/3.0*(q.y-p.y)}
	b.cubicTo(c1, c2, p)
	b.lastQuad = q
}

// arcTo appends an elliptical arc. Endpoint parameterisation is converted to
// centre parameterisation (SVG 1.1 appendix F.6) and the result split into
// segments of at most a quarter turn, where the kappa approximation is
// accurate; a single cubic across a large sweep would visibly bulge.
func (b *svgPathBuilder) arcTo(rx, ry, rotationDeg float64, largeArc, sweep bool, p1 svgPoint) {
	p0 := b.pos
	rx, ry = math.Abs(rx), math.Abs(ry)
	if rx == 0 || ry == 0 || p0 == p1 {
		b.lineTo(p1)
		return
	}
	phi := rotationDeg * math.Pi / 180
	cosPhi, sinPhi := math.Cos(phi), math.Sin(phi)

	dx2 := (p0.x - p1.x) / 2
	dy2 := (p0.y - p1.y) / 2
	x1p := cosPhi*dx2 + sinPhi*dy2
	y1p := -sinPhi*dx2 + cosPhi*dy2

	// Radii too small to span the endpoints are scaled up (F.6.6).
	if lambda := (x1p*x1p)/(rx*rx) + (y1p*y1p)/(ry*ry); lambda > 1 {
		s := math.Sqrt(lambda)
		rx *= s
		ry *= s
	}

	num := rx*rx*ry*ry - rx*rx*y1p*y1p - ry*ry*x1p*x1p
	den := rx*rx*y1p*y1p + ry*ry*x1p*x1p
	coef := 0.0
	if den > 0 {
		coef = math.Sqrt(math.Max(0, num/den))
	}
	if largeArc == sweep {
		coef = -coef
	}
	cxp := coef * rx * y1p / ry
	cyp := -coef * ry * x1p / rx
	cx := cosPhi*cxp - sinPhi*cyp + (p0.x+p1.x)/2
	cy := sinPhi*cxp + cosPhi*cyp + (p0.y+p1.y)/2

	ux, uy := (x1p-cxp)/rx, (y1p-cyp)/ry
	vx, vy := (-x1p-cxp)/rx, (-y1p-cyp)/ry
	theta1 := svgAngle(1, 0, ux, uy)
	dTheta := svgAngle(ux, uy, vx, vy)
	if !sweep && dTheta > 0 {
		dTheta -= 2 * math.Pi
	} else if sweep && dTheta < 0 {
		dTheta += 2 * math.Pi
	}

	segments := int(math.Ceil(math.Abs(dTheta) / (math.Pi / 2)))
	if segments < 1 {
		segments = 1
	}
	if segments > 8 {
		segments = 8
	}
	delta := dTheta / float64(segments)
	for i := 0; i < segments; i++ {
		b.arcSegment(cx, cy, rx, ry, cosPhi, sinPhi, theta1+float64(i)*delta, theta1+float64(i+1)*delta)
	}
	// Land exactly on the requested endpoint rather than on the accumulated
	// approximation of it, so a following segment starts where the caller
	// intended.
	b.pos = p1
}

func (b *svgPathBuilder) arcSegment(cx, cy, rx, ry, cosPhi, sinPhi, t1, t2 float64) {
	half := (t2 - t1) / 2
	alpha := math.Sin(t2-t1) * (math.Sqrt(4+3*math.Pow(math.Tan(half), 2)) - 1) / 3
	if math.IsNaN(alpha) {
		alpha = math.Tan(half) * 4 / 3
	}
	point := func(t float64) svgPoint {
		x, y := rx*math.Cos(t), ry*math.Sin(t)
		return svgPoint{cosPhi*x - sinPhi*y + cx, sinPhi*x + cosPhi*y + cy}
	}
	deriv := func(t float64) svgPoint {
		x, y := -rx*math.Sin(t), ry*math.Cos(t)
		return svgPoint{cosPhi*x - sinPhi*y, sinPhi*x + cosPhi*y}
	}
	p0, p1 := point(t1), point(t2)
	d0, d1 := deriv(t1), deriv(t2)
	c1 := svgPoint{p0.x + alpha*d0.x, p0.y + alpha*d0.y}
	c2 := svgPoint{p1.x - alpha*d1.x, p1.y - alpha*d1.y}
	if len(b.cur) == 0 {
		b.cur = append(b.cur, p0)
	}
	b.cur = appendCubic(b.cur, p0, c1, c2, p1, cubicSteps(p0, c1, c2, p1, b.tol))
}

// svgAngle returns the signed angle from vector u to vector v.
func svgAngle(ux, uy, vx, vy float64) float64 {
	lenU, lenV := math.Hypot(ux, uy), math.Hypot(vx, vy)
	if lenU == 0 || lenV == 0 {
		return 0
	}
	c := (ux*vx + uy*vy) / (lenU * lenV)
	c = math.Max(-1, math.Min(1, c))
	ang := math.Acos(c)
	if ux*vy-uy*vx < 0 {
		ang = -ang
	}
	return ang
}

// exec reads the operands of one command instance.
func (b *svgPathBuilder) exec(cmd byte, sc *svgPathScanner) bool {
	rel := cmd >= 'a'
	switch cmd {
	case 'M', 'm':
		n, ok := sc.numbers(2)
		if !ok {
			return false
		}
		p := svgPoint{n[0], n[1]}
		if rel {
			p = addPt(b.pos, p)
		}
		b.flush(false)
		b.pos = p
		b.start = p
		b.cur = append(b.cur, p)
	case 'L', 'l':
		n, ok := sc.numbers(2)
		if !ok {
			return false
		}
		p := svgPoint{n[0], n[1]}
		if rel {
			p = addPt(b.pos, p)
		}
		b.lineTo(p)
	case 'H', 'h':
		n, ok := sc.numbers(1)
		if !ok {
			return false
		}
		x := n[0]
		if rel {
			x += b.pos.x
		}
		b.lineTo(svgPoint{x, b.pos.y})
	case 'V', 'v':
		n, ok := sc.numbers(1)
		if !ok {
			return false
		}
		y := n[0]
		if rel {
			y += b.pos.y
		}
		b.lineTo(svgPoint{b.pos.x, y})
	case 'C', 'c':
		n, ok := sc.numbers(6)
		if !ok {
			return false
		}
		c1, c2, p := svgPoint{n[0], n[1]}, svgPoint{n[2], n[3]}, svgPoint{n[4], n[5]}
		if rel {
			c1, c2, p = addPt(b.pos, c1), addPt(b.pos, c2), addPt(b.pos, p)
		}
		b.cubicTo(c1, c2, p)
	case 'S', 's':
		n, ok := sc.numbers(4)
		if !ok {
			return false
		}
		c2, p := svgPoint{n[0], n[1]}, svgPoint{n[2], n[3]}
		if rel {
			c2, p = addPt(b.pos, c2), addPt(b.pos, p)
		}
		c1 := b.pos
		if b.prev == 'C' || b.prev == 'c' || b.prev == 'S' || b.prev == 's' {
			c1 = reflectPt(b.lastCtl, b.pos)
		}
		b.cubicTo(c1, c2, p)
	case 'Q', 'q':
		n, ok := sc.numbers(4)
		if !ok {
			return false
		}
		q, p := svgPoint{n[0], n[1]}, svgPoint{n[2], n[3]}
		if rel {
			q, p = addPt(b.pos, q), addPt(b.pos, p)
		}
		b.quadTo(q, p)
	case 'T', 't':
		n, ok := sc.numbers(2)
		if !ok {
			return false
		}
		p := svgPoint{n[0], n[1]}
		if rel {
			p = addPt(b.pos, p)
		}
		q := b.pos
		if b.prev == 'Q' || b.prev == 'q' || b.prev == 'T' || b.prev == 't' {
			q = reflectPt(b.lastQuad, b.pos)
		}
		b.quadTo(q, p)
	case 'A', 'a':
		n, ok := sc.numbers(7)
		if !ok {
			return false
		}
		p := svgPoint{n[5], n[6]}
		if rel {
			p = addPt(b.pos, p)
		}
		b.arcTo(n[0], n[1], n[2], n[3] != 0, n[4] != 0, p)
	default:
		return false
	}
	b.prev = cmd
	return true
}

// --- rasterisation -----------------------------------------------------------

// svgEdge is one directed segment in device space, in the form the scanline
// sweep wants. dir is +1 or -1 by direction: the nonzero winding rule counts it,
// the even-odd rule ignores it.
type svgEdge struct {
	x0, y0, x1, y1 float64
	dir            int
}

// rasterise draws one shape into dst. base maps user units to the raster.
func (s *svgShape) rasterise(dst *image.RGBA, base svgMatrix) {
	xf := base.mul(s.xf)
	sub := make([][]svgPoint, len(s.subpaths))
	closed := make([]bool, len(s.subpaths))
	for i, sp := range s.subpaths {
		pts := make([]svgPoint, len(sp.pts))
		for j, p := range sp.pts {
			pts[j] = xf.apply(p)
		}
		sub[i] = pts
		closed[i] = sp.closed
	}
	if s.style.fill != nil {
		svgFill(dst, sub, s.style.evenOdd, s.style.fill.premultiply())
	}
	if s.style.stroke != nil && s.style.strokeWidth > 0 {
		svgStroke(dst, sub, closed, s.style.strokeWidth*xf.uniformScale(), s.style.lineCap, s.style.stroke.premultiply())
	}
}

// svgFill fills the union of the subpaths under the given fill rule.
//
// An open subpath is closed implicitly, which is what SVG specifies for fill
// and why the closing edge is added regardless of the `closed` flag: a polyline
// used as a filled shape must not leak.
func svgFill(dst *image.RGBA, sub [][]svgPoint, evenOdd bool, c color.RGBA) {
	edges := make([]svgEdge, 0, 16)
	for _, pts := range sub {
		edges = appendClosedEdges(edges, pts)
	}
	svgSweep(dst, edges, evenOdd, c)
}

func appendClosedEdges(edges []svgEdge, pts []svgPoint) []svgEdge {
	if len(pts) < 2 {
		return edges
	}
	for i := 0; i+1 < len(pts); i++ {
		edges = appendEdge(edges, pts[i], pts[i+1])
	}
	return appendEdge(edges, pts[len(pts)-1], pts[0])
}

func appendEdge(edges []svgEdge, a, b svgPoint) []svgEdge {
	if a.y == b.y {
		// A horizontal span contributes no crossing, and leaving it out also
		// avoids a division by zero in the sweep.
		return edges
	}
	dir := 1
	if b.y < a.y {
		dir = -1
	}
	return append(edges, svgEdge{x0: a.x, y0: a.y, x1: b.x, y1: b.y, dir: dir})
}

// svgSweep rasterises edges with a scanline sweep, one sample row per device
// row.
//
// Coverage is deliberately hard-edged rather than fractionally weighted: the
// raster is supersampled and box-filtered afterwards, which produces the same
// anti-aliasing for a fraction of the work and without the risk of blending a
// shared edge twice.
func svgSweep(dst *image.RGBA, edges []svgEdge, evenOdd bool, c color.RGBA) {
	if len(edges) == 0 || c.A == 0 {
		return
	}
	minY := math.Min(edges[0].y0, edges[0].y1)
	maxY := math.Max(edges[0].y0, edges[0].y1)
	for _, e := range edges {
		minY = math.Min(minY, math.Min(e.y0, e.y1))
		maxY = math.Max(maxY, math.Max(e.y0, e.y1))
	}
	b := dst.Bounds()
	y0 := maxInt(int(math.Floor(minY)), b.Min.Y)
	y1 := minInt(int(math.Ceil(maxY)), b.Max.Y)

	type crossing struct {
		x   float64
		dir int
	}
	crs := make([]crossing, 0, len(edges))
	for y := y0; y < y1; y++ {
		fy := float64(y) + 0.5
		crs = crs[:0]
		for _, e := range edges {
			lo, hi := e.y0, e.y1
			if lo > hi {
				lo, hi = hi, lo
			}
			// Half-open in y, so a vertex shared by two edges is counted once.
			if fy < lo || fy >= hi {
				continue
			}
			t := (fy - e.y0) / (e.y1 - e.y0)
			crs = append(crs, crossing{x: e.x0 + t*(e.x1-e.x0), dir: e.dir})
		}
		if len(crs) < 2 {
			continue
		}
		sort.Slice(crs, func(i, j int) bool { return crs[i].x < crs[j].x })

		wind := 0
		for i := range crs {
			wind += crs[i].dir
			if i+1 >= len(crs) {
				break
			}
			// The span between this crossing and the next is inside when the
			// winding so far says so; even-odd only counts how many edges have
			// been crossed.
			inside := wind != 0
			if evenOdd {
				inside = (i+1)%2 == 1
			}
			if inside {
				svgFillSpan(dst, y, crs[i].x, crs[i+1].x, c)
			}
		}
	}
}

// svgFillSpan paints the pixels whose centre lies inside (x1, x2).
func svgFillSpan(dst *image.RGBA, y int, x1, x2 float64, c color.RGBA) {
	if x2 <= x1 {
		return
	}
	b := dst.Bounds()
	start := maxInt(int(math.Ceil(x1-0.5)), b.Min.X)
	end := minInt(int(math.Floor(x2-0.5)), b.Max.X-1)
	if start > end {
		return
	}
	if c.A == 255 {
		row := dst.Pix[y*dst.Stride:]
		for x := start; x <= end; x++ {
			o := x * 4
			row[o+0], row[o+1], row[o+2], row[o+3] = c.R, c.G, c.B, 255
		}
		return
	}
	for x := start; x <= end; x++ {
		svgBlendPremul(dst, x, y, c)
	}
}

// svgBlendPremul composites a premultiplied colour over the destination.
func svgBlendPremul(dst *image.RGBA, x, y int, c color.RGBA) {
	if c.A == 0 {
		return
	}
	o := y*dst.Stride + x*4
	ia := uint32(255 - c.A)
	dst.Pix[o+0] = uint8(uint32(c.R) + uint32(dst.Pix[o+0])*ia/255)
	dst.Pix[o+1] = uint8(uint32(c.G) + uint32(dst.Pix[o+1])*ia/255)
	dst.Pix[o+2] = uint8(uint32(c.B) + uint32(dst.Pix[o+2])*ia/255)
	dst.Pix[o+3] = uint8(uint32(c.A) + uint32(dst.Pix[o+3])*ia/255)
}

// svgCirclePoly approximates a circle. Sixteen segments hold the error below
// 1% of the radius, which is invisible once the raster is box-filtered; caps and
// joins are the only users.
func svgCirclePoly(cx, cy, r float64) []svgPoint {
	const segments = 16
	pts := make([]svgPoint, 0, segments)
	for i := 0; i < segments; i++ {
		a := 2 * math.Pi * float64(i) / segments
		pts = append(pts, svgPoint{cx + r*math.Cos(a), cy + r*math.Sin(a)})
	}
	return pts
}

// svgStroke paints a stroke by filling one quadrilateral per segment, widened
// by half the stroke width. A quadrilateral covers the whole width in one pass,
// so the sweep never has to know about strokes, and butt line caps — SVG's
// default — come for free. Curves reach here already flattened finely enough
// that the notch left at an unfilled join is a fraction of a pixel; line joins
// are otherwise not modelled, which is noted in the package documentation.
func svgStroke(dst *image.RGBA, sub [][]svgPoint, closed []bool, width float64, cap svgCap, c color.RGBA) {
	if width < 1 {
		// A sub-pixel stroke still has to be visible; a preview that silently
		// drops hairline rules is worse than one that draws them a pixel wide,
		// which is also what the rest of the renderer does with minimum widths.
		width = 1
	}
	half := width / 2
	edges := make([]svgEdge, 0, 8)
	for i, pts := range sub {
		n := len(pts)
		if n < 2 {
			continue
		}
		for j := 0; j+1 < n || (closed[i] && j < n); j++ {
			k := j + 1
			if k >= n {
				k = 0
			}
			a, b := pts[j], pts[k]
			dx, dy := b.x-a.x, b.y-a.y
			l := math.Hypot(dx, dy)
			if l == 0 {
				continue
			}
			nx, ny := -dy/l*half, dx/l*half
			// Square caps extend the two end quadrilaterals along the path.
			extA, extB := 0.0, 0.0
			if !closed[i] && cap == svgCapSquare {
				if j == 0 {
					extA = half
				}
				if j+2 == n {
					extB = half
				}
			}
			ux, uy := dx/l, dy/l
			ax, ay := a.x-ux*extA, a.y-uy*extA
			bx, by := b.x+ux*extB, b.y+uy*extB
			edges = edges[:0]
			edges = appendClosedEdges(edges, []svgPoint{
				{ax + nx, ay + ny},
				{bx + nx, by + ny},
				{bx - nx, by - ny},
				{ax - nx, ay - ny},
			})
			svgSweep(dst, edges, false, c)
		}
		if cap != svgCapRound || closed[i] {
			// Discs model an open subpath's caps. Joins along a closed path are
			// a stroke-linejoin concern, which is not modelled.
			continue
		}
		for _, p := range pts {
			edges = edges[:0]
			edges = appendClosedEdges(edges, svgCirclePoly(p.x, p.y, half))
			svgSweep(dst, edges, false, c)
		}
	}
}
