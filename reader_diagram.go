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

// StashedDiagram carries one SmartArt frame's own parts — the data model, the
// layout definition, the quick style and the colour transform, the cached
// drawing and the drawing's own relationship file with whatever media it
// names — byte for byte out of the source package. The reader renders the
// cached drawing; the writer puts these back and re-emits the graphicFrame
// whose relationship ids it filled in, so a read-write round trip neither
// flattens the diagram nor loses the parts editing it again would need.
type StashedDiagram struct {
	// parts maps the part path the source package carried to what the writer
	// puts back: bytes and the [Content_Types] entry each one needs.
	parts map[string]diagramPart
	// drawingPath is the cached drawing's part path; its blipFills resolve
	// through the relationship file that also lives in parts.
	drawingPath string
	// auxPaths names the part each auxiliary relationship points at, in
	// dm, lo, qs, cs order, so the writer's slide rels can emit the five
	// entries with r:dm → rIdN pointing at the same files the stash carries.
	auxPaths [4]string
}

// diagramPart is one stashed part.
type diagramPart struct {
	data        []byte
	contentType string
}

// stashDiagramParts copies a frame's parts out of the source package.
//
// dataRelID is the frame's <dgm:relIds r:dm> and drawingRelID the slide
// relationship the cached drawing was read through — the numbering rule that
// pairs data and drawing parts points at the same files here. Only readable
// parts are stashed; a part missing from the source cannot be written back.
func (r *PPTXReader) stashDiagramParts(zr *zip.Reader, slidePath string, rels []xmlRelForRead, dg *GroupShape, holder *UnsupportedShape, dataRelID, drawingRelID string, pres *Presentation, frameRelIDs map[string]bool) {
	if dg == nil && holder == nil {
		return
	}
	baseDir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
	resolve := func(target string) string {
		if strings.HasPrefix(target, "ppt/") {
			return target
		}
		return resolveRelativePath(baseDir, target)
	}
	st := &StashedDiagram{parts: make(map[string]diagramPart)}
	for _, rel := range rels {
		ct, slot, ok := auxDiagramContentType(rel.Type)
		if !ok {
			continue
		}
		// A frame's stash carries only the parts its own dgm:relIds named:
		// a slide with several diagrams numbered every part in one rels
		// table, and taking them all left every frame's auxPaths pointing at
		// the last data part (three frames all read the third diagram).
		if frameRelIDs != nil && !frameRelIDs[rel.ID] {
			continue
		}
		path := resolve(rel.Target)
		data, err := readFileFromZip(zr, path)
		if err != nil {
			continue
		}
		st.parts[path] = diagramPart{data: data, contentType: ct}
		st.auxPaths[slot] = path
		// The part's own relationship file travels with it when the source
		// package carried one.
		relsPath := strings.Replace(path, "/diagrams/", "/diagrams/_rels/", 1) + ".rels"
		if rd, rerr := readFileFromZip(zr, relsPath); rerr == nil {
			st.parts[relsPath] = diagramPart{data: rd,
				contentType: "application/vnd.openxmlformats-package.relationships+xml"}
		}
	}
	// The cached drawing, read through the slide relationship that produced
	// the rendered group.
	var drawingPath string
	for _, rel := range rels {
		if (drawingRelID != "" && rel.ID == drawingRelID) ||
			(drawingRelID == "" && strings.HasSuffix(rel.Type, "/diagramDrawing")) {
			drawingPath = resolve(rel.Target)
			break
		}
	}
	if drawingPath != "" {
		if data, err := readFileFromZip(zr, drawingPath); err == nil {
			st.parts[drawingPath] = diagramPart{data: data,
				contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml"}
			st.drawingPath = drawingPath
		}
	}
	// Whatever media the drawing's blipFills name comes along: the drawing's
	// own rels file resolves them, and both are stashed verbatim.
	if st.drawingPath != "" {
		drPath := strings.Replace(st.drawingPath, "/diagrams/", "/diagrams/_rels/", 1) + ".rels"
		if rd, err := readFileFromZip(zr, drPath); err == nil {
			st.parts[drPath] = diagramPart{data: rd,
				contentType: "application/vnd.openxmlformats-package.relationships+xml"}
			for _, dr := range parseRelsBytes(rd) {
				if !strings.Contains(strings.ToLower(dr.Type), "image") {
					continue
				}
				mpath := resolveRelativePath("ppt/diagrams", dr.Target)
				if md, merr := readFileFromZip(zr, mpath); merr == nil {
					st.parts[mpath] = diagramPart{data: md}
				}
			}
		}
	}
	if len(st.parts) == 0 {
		return
	}
	if dg != nil {
		dg.diagram = st
	} else if holder != nil {
		holder.diagram = st
	}
	// No cached drawing: the stand-in can still carry a real diagram —
	// parse the model out of the data part and the layout family out of
	// the layout part, and the renderer lays it out with the v1 engine
	// instead of showing the "Unsupported" box.
	if dg == nil && holder != nil {
		dataPath := st.auxPaths[0]
		if dataBytes, ok := st.parts[dataPath]; ok {
			if model, merr := parseSmartArtModel(dataBytes.data); merr == nil {
				if layoutPath := st.auxPaths[1]; layoutPath != "" {
					if layoutBytes, ok2 := st.parts[layoutPath]; ok2 {
						if alg := firstSmartArtAlg(layoutBytes.data); alg != "" {
							model.Algo = alg
						}
					}
				}
				holder.model = model
			}
		}
	}
}

// firstSmartArtAlg returns the layout family the layout part's root
// layoutNode drives — the first <dgm:alg type=".."> the document carries —
// mapped to one of the v1 engine's four families, or "" when nothing maps.
func firstSmartArtAlg(data []byte) string {
	var found string
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		if t, ok := token.(xml.StartElement); ok && t.Name.Local == "alg" {
			for _, attr := range t.Attr {
				if attr.Name.Local == "type" && found == "" {
					found = attr.Value
				}
			}
			if found != "" {
				break
			}
		}
	}
	switch found {
	case "lin", "snake", "cycle":
		return found
	case "hierChild", "hierRoot":
		return "hier"
	}
	return ""
}

// auxDiagramContentType names the [Content_Types] override for one of the
// four auxiliary diagram parts, or reports that the relationship names none.
func auxDiagramContentType(relType string) (string, int, bool) {
	switch {
	case strings.HasSuffix(relType, "/diagramData"):
		return "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml", 0, true
	case strings.HasSuffix(relType, "/diagramLayout"):
		return "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml", 1, true
	case strings.HasSuffix(relType, "/diagramQuickStyle"):
		return "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml", 2, true
	case strings.HasSuffix(relType, "/diagramColors"):
		return "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml", 3, true
	}
	return "", 0, false
}

// The full relationship types the four auxiliary diagram parts and the cached
// drawing go by. Readers match the trailing segment; the writer states these.
const (
	relTypeDiagramData       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
	relTypeDiagramLayout     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
	relTypeDiagramQuickStyle = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
	relTypeDiagramColors     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
)

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
	// The cached drawing's blip fills resolve against the DRAWING part's own
	// rels (drawingN.xml.rels), not the slide's: the two rels tables number
	// their ids independently, and smartart-simple's rId1 is a slideLayout
	// on one side and the ocean photo on the other.
	if drPath := strings.Replace(path, "/diagrams/", "/diagrams/_rels/", 1) + ".rels"; drPath != path {
		if rd, err := readFileFromZip(zr, drPath); err == nil {
			rels = parseRelsBytes(rd)
		}
	}
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

// parseRelsBytes parses a relationships file that is already in memory: the
// stash walks the cached drawing's own rels, and re-reading it from the zip
// would only repeat a lookup for a file the reader has in hand.
func parseRelsBytes(data []byte) []xmlRelForRead {
	var rels xmlRelsForRead
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil
	}
	return rels.Relationships
}

// diagramRels returns the relationships a stashed SmartArt frame needs, in
// the order every rel-id consumer walks them: the auxiliary parts that were
// actually read (dm, lo, qs, cs), then the cached drawing. A frame whose
// parts are gone from the source package yields an empty list and the writer
// skips the frame entirely, the way it skips any other unsupported construct.
func diagramRels(st *StashedDiagram) (types, targets []string) {
	if st == nil {
		return nil, nil
	}
	ordered := []string{relTypeDiagramData, relTypeDiagramLayout, relTypeDiagramQuickStyle, relTypeDiagramColors}
	for k, p := range st.auxPaths {
		if _, ok := st.parts[p]; !ok {
			continue
		}
		types = append(types, ordered[k])
		targets = append(targets, p)
	}
	if st.drawingPath != "" {
		if _, ok := st.parts[st.drawingPath]; ok {
			types = append(types, relTypeDiagramDrawing)
			targets = append(targets, st.drawingPath)
		}
	}
	return
}
