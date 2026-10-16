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
// slide can carry several diagrams, and each frame is paired with its own
// drawing through the data part its r:dm names — see readDiagramDrawing.
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

// readDiagramDrawing resolves the diagramDrawing relationship that belongs to
// THIS graphic frame and parses the cached drawing into a GroupShape sized to
// the frame.
//
// dataRelID is the frame's <dgm:relIds r:dm> — the relationship id of the
// diagram's data part. A slide can carry several diagrams (customGeo's
// slide14 has two side by side), and each frame must render its own cached
// drawing: PowerPoint numbers the drawing to match the data part it resolves
// (data8.xml's diagram is cached as drawing8.xml), so the pairing walks
// frame r:dm → data part name → same-numbered drawing part. When the frame
// declares no data part or the numbers do not line up, the historical
// first-diagramDrawing behaviour stands (correct for the single-diagram
// slides it was written against).
//
// It returns nil when there is no such relationship or the part cannot be
// read, so the caller falls back to the visible UnsupportedShape stand-in: a
// deck whose diagrams were never cached still reports the gap rather than
// drawing an empty box that reads like a correct render.
func (r *PPTXReader) readDiagramDrawing(zr *zip.Reader, rels []xmlRelForRead, slidePath string, pres *Presentation, frameCX, frameCY int64, dataRelID string) *GroupShape {
	var candidates []xmlRelForRead
	for _, rel := range rels {
		if rel.Type == relTypeDiagramDrawing || strings.HasSuffix(rel.Type, "/diagramDrawing") {
			candidates = append(candidates, rel)
		}
	}
	if len(candidates) == 0 {
		return nil
	}
	if want := drawingPartForDataRel(rels, dataRelID); want != "" {
		for _, rel := range candidates {
			if strings.EqualFold(lastPathComponent(rel.Target), want) {
				return r.loadDiagramDrawing(zr, rels, slidePath, pres, frameCX, frameCY, rel)
			}
		}
	}
	return r.loadDiagramDrawing(zr, rels, slidePath, pres, frameCX, frameCY, candidates[0])
}

// drawingPartForDataRel maps a frame's data-part relationship id to the name
// of the drawing part PowerPoint caches for it: the drawing carries the same
// number as the data part it resolves (rId3 → data8.xml → drawing8.xml).
// The dsp root may also carry an explicit <dsp:dataModelRel r:dm> back-link,
// but customGeo's ten diagrams all omit it, and the numbering has held on
// every deck measured — so the number is the link that earns the code.
func drawingPartForDataRel(rels []xmlRelForRead, dataRelID string) string {
	if dataRelID == "" {
		return ""
	}
	for _, rel := range rels {
		if rel.ID != dataRelID {
			continue
		}
		base := lastPathComponent(rel.Target)
		if strings.HasPrefix(base, "data") && strings.HasSuffix(base, ".xml") {
			return "drawing" + base[len("data"):len(base)-len(".xml")] + ".xml"
		}
	}
	return ""
}

func (r *PPTXReader) loadDiagramDrawing(zr *zip.Reader, rels []xmlRelForRead, slidePath string, pres *Presentation, frameCX, frameCY int64, rel xmlRelForRead) *GroupShape {
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
