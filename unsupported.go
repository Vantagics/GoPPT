package gopresentation

import (
	"image"
	"image/color"
	"strings"

	"golang.org/x/image/font"
	"golang.org/x/image/font/basicfont"
)

// ShapeTypeUnsupported is the type of the stand-in the reader creates for an
// OOXML construct it recognises but cannot rasterise: SmartArt diagrams, OLE
// objects, embedded 3-D models and the like.
//
// The reader never drops such a shape silently. A blank region in a preview is
// indistinguishable from a shape that rendered correctly but happens to be
// empty, which hides real generation problems; keeping a stand-in preserves the
// frame geometry and makes the condition visible and reportable. Use
// Presentation.UnsupportedShapes to enumerate them.
const ShapeTypeUnsupported ShapeType = 12

// UnsupportedShape stands in for a shape that exists in the file but that this
// library cannot draw.
//
// The original XML is not retained, so writing the presentation back out drops
// the shape: the writer skips it. That is a deliberate trade-off, since the
// alternative is failing to read the whole presentation because of one
// construct that does not affect anything else.
type UnsupportedShape struct {
	BaseShape

	// reason is a short human-readable description of what the original shape
	// was, e.g. "SmartArt diagram". It is drawn inside the placeholder box, so
	// someone looking at a preview can tell what is missing.
	reason string
	// contentType is the a:graphicData uri the shape was read from, kept
	// verbatim so callers can classify precisely rather than by label.
	contentType string
}

// NewUnsupportedShape creates a stand-in for a construct the library cannot
// draw. reason should be a short description, e.g. "SmartArt diagram".
func NewUnsupportedShape(reason string) *UnsupportedShape {
	return &UnsupportedShape{reason: reason}
}

func (u *UnsupportedShape) GetType() ShapeType { return ShapeTypeUnsupported }

// GetReason returns a short human-readable description of what the original
// shape was, e.g. "SmartArt diagram".
func (u *UnsupportedShape) GetReason() string { return u.reason }

// SetReason sets the human-readable description.
func (u *UnsupportedShape) SetReason(reason string) *UnsupportedShape {
	u.reason = reason
	return u
}

// GetContentType returns the OOXML content the shape was read from, normally
// the a:graphicData uri. It is empty for a shape created in code.
func (u *UnsupportedShape) GetContentType() string { return u.contentType }

// SetContentType records the OOXML content type the shape came from.
func (u *UnsupportedShape) SetContentType(ct string) *UnsupportedShape {
	u.contentType = ct
	return u
}

// Label returns the text drawn inside the placeholder box.
func (u *UnsupportedShape) Label() string {
	if u.reason == "" {
		return "Unsupported shape"
	}
	return "Unsupported: " + u.reason
}

// classifyGraphicData turns an a:graphicData uri into a short description of
// what the enclosing graphicFrame holds. It is used for both the placeholder
// label and diagnostics, so a construct with no special case still gets a name
// derived from the uri rather than an opaque one.
func classifyGraphicData(uri string) string {
	switch {
	case uri == "":
		return "unknown graphicFrame content"
	case strings.Contains(uri, "/diagram"):
		return "SmartArt diagram"
	case strings.Contains(uri, "/chart"):
		return "chart"
	case strings.Contains(uri, "/table"):
		return "table"
	case strings.Contains(uri, "/ole") || strings.Contains(uri, "oleObject"):
		return "OLE object"
	case strings.Contains(uri, "/ink") || strings.Contains(uri, "inkml"):
		return "ink annotation"
	case strings.Contains(uri, "model3d") || strings.Contains(uri, "/model"):
		return "3-D model"
	case strings.Contains(uri, "/picture"):
		return "picture"
	default:
		// Fall back to the last path segment so the label still says something
		// more useful than the full uri.
		if i := strings.LastIndex(uri, "/"); i >= 0 && i+1 < len(uri) {
			return uri[i+1:]
		}
		return uri
	}
}

// UnsupportedShapes returns every shape that the reader could not represent, in
// slide order and descending into groups. An empty result means every shape in
// the presentation is renderable.
//
// Preview and validation tooling can use this to fail a build, or to annotate a
// report, rather than showing a quietly incomplete image:
//
//	if bad := pres.UnsupportedShapes(); len(bad) > 0 {
//		for _, s := range bad {
//			log.Printf("slide content not rendered: %s", s.Label())
//		}
//	}
func (p *Presentation) UnsupportedShapes() []*UnsupportedShape {
	if p == nil {
		return nil
	}
	var out []*UnsupportedShape
	for _, slide := range p.GetAllSlides() {
		if slide == nil {
			continue
		}
		collectUnsupportedShapes(slide.GetShapes(), &out)
	}
	return out
}

// collectUnsupportedShapes appends every UnsupportedShape reachable from shapes,
// following groups, which can nest arbitrarily.
func collectUnsupportedShapes(shapes []Shape, out *[]*UnsupportedShape) {
	for _, s := range shapes {
		switch v := s.(type) {
		case *UnsupportedShape:
			*out = append(*out, v)
		case *GroupShape:
			collectUnsupportedShapes(v.shapes, out)
		}
	}
}

// --- rendering --------------------------------------------------------------

// Color palette for the placeholder. Amber is used deliberately: it is not part
// of the default chart or theme palettes, so it cannot be mistaken for real
// document content.
var (
	unsupportedFill   = color.RGBA{R: 255, G: 246, B: 224, A: 255}
	unsupportedBorder = color.RGBA{R: 214, G: 141, B: 17, A: 255}
	unsupportedText   = color.RGBA{R: 146, G: 96, B: 10, A: 255}
)

// renderUnsupported draws a conspicuous placeholder for a construct the reader
// could not represent. It is intentionally visible: drawing nothing would make
// the preview blank and hide the fact that part of the deck did not render,
// which is exactly the failure mode this type exists to prevent.
func (r *renderer) renderUnsupported(s *UnsupportedShape) {
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	if w <= 0 || h <= 0 {
		return
	}
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	label := s.Label()

	// paint draws at an explicit origin so the same code serves the rotated
	// path, where the temp buffer's origin is (0,0) rather than the slide's.
	paint := func(tr *renderer, ox, oy int) {
		rect := image.Rect(ox, oy, ox+w, oy+h)
		tr.fillRectBlend(rect, unsupportedFill)
		// A dashed frame reads as "not real content" and stays distinguishable
		// at low draft resolutions, where a lighter tint might wash out.
		tr.drawRectBorder(rect, unsupportedBorder, placeholderBorderWidth(w, h), BorderDash)
		tr.drawPlaceholderLabel(label, rect)
	}

	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()
	if rotation == 0 && !flipH && !flipV {
		paint(r, x, y)
		return
	}
	r.renderRotated(x, y, w, h, rotation, flipH, flipV, func(tmp *renderer) {
		paint(tmp, 0, 0)
	})
}

// placeholderBorderWidth keeps the dashed frame visible without letting it
// swallow the box: a fixed weight turns a 40x20 px placeholder into almost
// nothing but border.
func placeholderBorderWidth(w, h int) int {
	return maxInt(1, minInt(2, minInt(w, h)/12))
}

// drawPlaceholderLabel draws the label centred in rect, but only when a legible
// size fits inside the box.
//
// The label is generated by the library rather than taken from the document, so
// it must never run past the frame: a placeholder spilling over neighbouring
// content makes a preview look corrupted for reasons that are entirely the
// library's fault. Containment therefore rests on the fit check below, not on
// clipping the draw: image/draw does not clip a glyph blit from a uniform source
// to the destination's bounds, so drawing past the box would write into
// neighbouring pixels. A box too small for a legible label keeps only its dashed
// frame, which already says "nothing was rendered here", and which is also what
// the tests and Presentation.UnsupportedShapes report on.
func (r *renderer) drawPlaceholderLabel(label string, rect image.Rectangle) {
	if label == "" || rect.Dx() <= 0 || rect.Dy() <= 0 {
		return
	}
	// The padding absorbs glyph overhang, which can exceed the measured advance.
	face := r.placeholderLabelFace(label, float64(rect.Dx()-4), float64(rect.Dy()-4))
	if face == nil {
		return
	}
	r.drawStringCentered(label, face, unsupportedText, rect)
}

// minPlaceholderLabelPx is the smallest label size worth drawing. Below it the
// glyphs are illegible, so the dashed frame has to carry the meaning instead.
const minPlaceholderLabelPx = 7.0

// placeholderLabelFace returns a face in which label fits inside a maxW x maxH
// area, or nil when no size the library is willing to draw fits.
//
// Unlike getFace this never records a font fallback: the label is the library's
// own text, so a missing label font is not something the caller can act on, and
// reporting it would pollute the font diagnostics they use to judge the
// document itself.
func (r *renderer) placeholderLabelFace(label string, maxW, maxH float64) font.Face {
	if maxW < 4 || maxH < 4 {
		return nil
	}
	if basePx := r.fontSizePixels(NewFont()); basePx > 0 {
		for _, scale := range []float64{1, 0.8, 0.62, 0.48, 0.38} {
			px := basePx * scale
			if px < minPlaceholderLabelPx {
				break
			}
			face, _, _ := r.resolveFace(NewFont(), px, false)
			if face == nil {
				// No installed font matched at all, so only the bitmap face
				// below is left to try.
				break
			}
			if labelFits(face, label, maxW, maxH) {
				return face
			}
		}
	}
	if labelFits(basicfont.Face7x13, label, maxW, maxH) {
		return basicfont.Face7x13
	}
	return nil
}

// labelFits reports whether label fits inside a maxW x maxH area in the given
// face.
func labelFits(face font.Face, label string, maxW, maxH float64) bool {
	if face == nil {
		return false
	}
	m := face.Metrics()
	if float64((m.Ascent + m.Descent).Ceil()) > maxH {
		return false
	}
	return float64(font.MeasureString(face, label).Ceil()) <= maxW
}
