package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"strings"
)

// SmartArt (a "diagram") arrives as a <p:graphicFrame> whose <a:graphicData>
// names the diagram namespace and points at four auxiliary parts: the data
// model, the layout definition, the quick style and the colour transform.
// Those four describe *how* to lay a diagram out, and implementing that
// engine is out of scope.
//
// It does not have to be. Alongside them PowerPoint caches a fifth part,
// ppt/diagrams/drawingN.xml, holding the diagram's resolved geometry: every
// node already has an <a:xfrm>, a <a:prstGeom>, a fill, a line, a shadow and a
// text body. That is the same shape grammar a slide uses, so the existing
// slide scanner reads it as-is and the diagram renders exactly as PowerPoint
// last laid it out — no layout engine, no approximation.
//
// The relationship to that part is the one non-obvious part of the wiring: it
// is NOT reachable through the <dgm:relIds> element, which lists only
// r:dm/r:lo/r:qs/r:cs. The drawing is hung off the *slide's* relationship
// list instead, as a relationship whose type ends in "/diagramDrawing". A
// slide can carry several diagrams, and the frame's own xfrm is what says
// which one it is, so callers pass the frame geometry in.
const relTypeDiagramDrawing = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"

// isDiagramGraphicData reports whether a graphicFrame's <a:graphicData uri>
// names the diagram namespace, i.e. the frame is a SmartArt placeholder.
// The URI is matched on the trailing segment because Microsoft has shipped
// more than one spelling of the same namespace over the years.
func isDiagramGraphicData(uri string) bool {
	if uri == "" {
		return false
	}
	return strings.Contains(uri, "/diagram")
}

// readDiagramDrawing resolves the slide's diagramDrawing relationship and
// parses the cached drawing into a GroupShape sized to the graphic frame.
//
// It returns nil when there is no such relationship or the part cannot be
// read, so the caller falls back to the visible UnsupportedShape stand-in: a
// deck whose diagrams were never cached still reports the gap rather than
// drawing an empty box that reads like a correct render.
func (r *PPTXReader) readDiagramDrawing(zr *zip.Reader, rels []xmlRelForRead, slidePath string, pres *Presentation, frameCX, frameCY int64) *GroupShape {
	for _, rel := range rels {
		if rel.Type != relTypeDiagramDrawing && !strings.HasSuffix(rel.Type, "/diagramDrawing") {
			continue
		}
		target := rel.Target
		if !strings.HasPrefix(target, "ppt/") {
			dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
			target = resolveRelativePath(dir, target)
		}
		data, err := readFileFromZip(zr, target)
		if err != nil {
			return nil
		}
		return r.parseDiagramDrawing(data, rels, zr, target, pres, frameCX, frameCY)
	}
	return nil
}

// parseDiagramDrawing runs the ordinary slide scanner over a dsp:drawing part.
//
// The scanner matches elements on their local name, and dsp:sp / dsp:spPr /
// dsp:txBody / dsp:style have exactly the local names of the p:-prefixed
// shapes it already reads (sp, spPr, txBody, style), as does dsp:spTree. So
// the 3600-line scanner needs no changes at all: it walks to EOF, dsp:spTree
// opens the tree and every dsp:sp inside commits as a shape exactly as a
// <p:sp> would.
func (r *PPTXReader) parseDiagramDrawing(data []byte, rels []xmlRelForRead, zr *zip.Reader, path string, pres *Presentation, frameCX, frameCY int64) *GroupShape {
	holder := newSlide()
	decoder := xml.NewDecoder(bytes.NewReader(data))
	if err := r.parseSlideXML(decoder, holder, rels, zr, path, pres); err != nil {
		return nil
	}
	shapes := holder.GetShapes()
	if len(shapes) == 0 {
		return nil
	}

	g := NewGroupShape()
	g.name = "Diagram"
	for _, s := range shapes {
		g.AddShape(s)
	}

	// The drawing part's coordinates are already in the frame's EMU space —
	// measured against all ten diagrams of the customGeo corpus, the drawing's
	// bounding box equals the frame's <a:ext> to within 0.06%. So the child
	// coordinate space is the identity map: childOff=(0,0) and
	// childExt=frame extent make the renderer's group transform
	// (x' = off + (x - childOff) * w / childExt) reduce to x' = off + x.
	g.childOffX, g.childOffY = 0, 0
	g.childExtX, g.childExtY = frameCX, frameCY
	g.width, g.height = frameCX, frameCY
	return g
}
