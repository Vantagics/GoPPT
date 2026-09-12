package gopresentation

import (
	"archive/zip"
	"fmt"
	"os"
	"sort"
	"strings"
)

// buildHyperlinkRelMap pre-computes the relationship IDs for all hyperlinks in a slide.
// This ensures the XML shape content and the .rels file use the same IDs.
//
// The walk has to stay identical to the one in writeSlideRels: both visit
// flattenShapes in order, add the image/chart relationships each shape consumes,
// then the hyperlinks of the paragraphs shapeParagraphs hands back. Only a link
// hyperlinkRel accepts is given an id, so the map, the r:id placeholder in the
// slide XML and the .rels file agree on which links exist.
func (w *PPTXWriter) buildHyperlinkRelMap(slide *Slide) map[*TextRun]string {
	m := make(map[*TextRun]string)
	relIdx := 2 // rId1 is slideLayout
	for _, shape := range flattenShapes(slide.shapes) {
		relIdx += countShapeRels(shape)
		for _, para := range shapeParagraphs(shape) {
			for _, elem := range para.elements {
				tr, ok := elem.(*TextRun)
				if !ok || tr.hyperlink == nil {
					continue
				}
				if _, _, ok := w.hyperlinkRel(tr.hyperlink); !ok {
					continue
				}
				m[tr] = fmt.Sprintf("rId%d", relIdx)
				relIdx++
			}
		}
	}
	return m
}

// hyperlinkRel maps a hyperlink to the relationship it needs: the relationship
// type, its target, and whether the link can be written at all.
//
// An internal link names a slide by number, so a number no slide backs is not
// writable — and it has to be rejected here, in one place, because three parts
// have to agree about which links count: the r:id placeholder in the slide XML,
// the id map, and the .rels file. Emitting a relationship for a link the other
// two skipped shifts every later id, which is how a package ends up with a
// reference nothing defines.
func (w *PPTXWriter) hyperlinkRel(h *Hyperlink) (relType, target string, ok bool) {
	if h == nil {
		return "", "", false
	}
	if h.IsInternal {
		n := h.SlideNumber
		if n < 1 || n > len(w.presentation.slides) {
			return "", "", false
		}
		return relTypeSlide, fmt.Sprintf("slide%d.xml", n), true
	}
	if h.URL == "" {
		return "", "", false
	}
	return relTypeHyperlink, h.URL, true
}

// countShapeRels returns the number of non-hyperlink relationship IDs consumed by a shape
// (images and charts each consume one relId).
func countShapeRels(shape Shape) int {
	switch s := shape.(type) {
	case *DrawingShape:
		if s.data != nil || s.path != "" {
			return 1
		}
	case *ChartShape:
		return 1
	}
	return 0
}

// shapeParagraphs returns the paragraphs for shapes that can contain hyperlinks.
//
// A table cell's runs carry hyperlinks through the same run emitter as a text
// box's, so they have to be enumerated here as well: this function is what keeps
// the id map, the slide XML and the .rels file in step, and a shape missing from
// it gets a placeholder no relationship ever resolves.
func shapeParagraphs(shape Shape) []*Paragraph {
	switch s := shape.(type) {
	case *RichTextShape:
		return s.paragraphs
	case *PlaceholderShape:
		return s.paragraphs
	case *TableShape:
		var out []*Paragraph
		for _, row := range s.rows {
			for _, cell := range row {
				out = append(out, cell.paragraphs...)
			}
		}
		return out
	}
	return nil
}

// countRelIdxBefore computes the relIdx for a target shape within a slide,
// counting all rels (images, charts, hyperlinks) for shapes before it in
// document order, which is the order flattenShapes defines.
//
// It is a fourth consumer of the hyperlink rule, after buildHyperlinkRelMap,
// writeSlideRels and the r:id placeholder itself, so it asks hyperlinkRel rather
// than deciding for itself: a predicate duplicated here drifts, and a picture
// numbered from the wrong count refers to a relationship that belongs to a
// hyperlink.
func (w *PPTXWriter) countRelIdxBefore(shapes []Shape, target Shape) int {
	relIdx := 2 // rId1 is slideLayout
	for _, shape := range flattenShapes(shapes) {
		if shape == target {
			break
		}
		relIdx += countShapeRels(shape)
		for _, para := range shapeParagraphs(shape) {
			for _, elem := range para.elements {
				if tr, ok := elem.(*TextRun); ok && tr.hyperlink != nil {
					if _, _, ok := w.hyperlinkRel(tr.hyperlink); ok {
						relIdx++
					}
				}
			}
		}
	}
	return relIdx
}

func (w *PPTXWriter) writeSlide(zw *zip.Writer, slide *Slide, slideNum int, hlinkRelMap map[*TextRun]string) error {

	var shapesXML strings.Builder
	shapeID := 2 // 1 is reserved for the group shape

	for _, shape := range slide.shapes {
		switch s := shape.(type) {
		case *PlaceholderShape:
			shapesXML.WriteString(w.writePlaceholderShapeXML(s, &shapeID))
		case *RichTextShape:
			shapesXML.WriteString(w.writeRichTextShapeXML(s, &shapeID))
		case *DrawingShape:
			shapesXML.WriteString(w.writeDrawingShapeXML(s, &shapeID, slideNum))
		case *TableShape:
			shapesXML.WriteString(w.writeTableShapeXML(s, &shapeID))
		case *AutoShape:
			shapesXML.WriteString(w.writeAutoShapeXML(s, &shapeID))
		case *LineShape:
			shapesXML.WriteString(w.writeLineShapeXML(s, &shapeID))
		case *ChartShape:
			shapesXML.WriteString(w.writeChartShapeXML(s, &shapeID, slideNum))
		case *GroupShape:
			shapesXML.WriteString(w.writeGroupShapeXML(s, &shapeID, slideNum))
		}
	}

	// Replace hyperlink placeholders with actual relationship IDs
	result := shapesXML.String()
	for tr, relID := range hlinkRelMap {
		placeholder := fmt.Sprintf("rId_hlink_%p", tr)
		result = strings.Replace(result, placeholder, relID, 1)
	}

	// Background XML
	bgXML := ""
	if slide.background != nil && slide.background.Type != FillNone {
		bgXML = "    <p:bg>\n      <p:bgPr>\n"
		bgXML += w.writeFillXML(slide.background)
		bgXML += "        <a:effectLst/>\n      </p:bgPr>\n    </p:bg>\n"
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:cSld>
%s    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
%s    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, bgXML, result)

	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/slides/slide%d.xml", slideNum), content)
}

func (w *PPTXWriter) writeSlideRels(zw *zip.Writer, slide *Slide, slideNum int, hlinkRelMap map[*TextRun]string) error {
	var rels strings.Builder
	fmt.Fprintf(&rels, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s" Target="../slideLayouts/slideLayout1.xml"/>`, nsRelationships, relTypeSlideLayout)

	relIdx := 2
	for _, shape := range flattenShapes(slide.shapes) {
		switch s := shape.(type) {
		case *DrawingShape:
			if s.data != nil || s.path != "" {
				imgIdx := w.getImageIndex(slide, s)
				ext := w.getImageExtension(s)
				fmt.Fprintf(&rels, `
  <Relationship Id="rId%d" Type="%s" Target="../media/image%d.%s"/>`,
					relIdx, relTypeImage, imgIdx, ext)
				relIdx++
			}
		case *ChartShape:
			chartIdx := w.getChartIndex(s)
			fmt.Fprintf(&rels, `
  <Relationship Id="rId%d" Type="%s" Target="../charts/chart%d.xml"/>`,
				relIdx, relTypeChart, chartIdx)
			relIdx++
		}
		// Hyperlinks in the shape's paragraphs — a table cell's included, which
		// is what shapeParagraphs exists to guarantee.
		for _, para := range shapeParagraphs(shape) {
			for _, elem := range para.elements {
				tr, ok := elem.(*TextRun)
				if !ok || tr.hyperlink == nil {
					continue
				}
				rid := hlinkRelMap[tr]
				if rid == "" {
					continue // hyperlinkRel refused it, so it was never given an id
				}
				relType, target, _ := w.hyperlinkRel(tr.hyperlink)
				if tr.hyperlink.IsInternal {
					// A jump to another slide is an internal relationship of
					// slide type; the action on the hlinkClick says to jump.
					fmt.Fprintf(&rels, `
  <Relationship Id="%s" Type="%s" Target="%s"/>`, rid, relType, target)
				} else {
					fmt.Fprintf(&rels, `
  <Relationship Id="%s" Type="%s" Target="%s" TargetMode="External"/>`,
						rid, relType, xmlEscape(target))
				}
				relIdx++
			}
		}
	}

	// Comments relationship
	if len(slide.comments) > 0 {
		fmt.Fprintf(&rels, `
  <Relationship Id="rId%d" Type="%s" Target="../comments/comment%d.xml"/>`,
			relIdx, relTypeComment, slideNum)
		relIdx++
	}

	// Notes slide relationship
	if slide.notes != "" {
		fmt.Fprintf(&rels, `
  <Relationship Id="rId%d" Type="%s" Target="../notesSlides/notesSlide%d.xml"/>`,
			relIdx, relTypeNotesSlide, slideNum)
		relIdx++
	}

	rels.WriteString(`
</Relationships>`)
	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum), rels.String())
}

func (w *PPTXWriter) getImageIndex(slide *Slide, target *DrawingShape) int {
	idx := 1
	for _, sl := range w.presentation.slides {
		for _, ds := range collectDrawingShapes(sl.shapes) {
			if ds == target {
				return idx
			}
			idx++
		}
	}
	return idx
}

// flattenShapes returns every shape of a slide in document order, with the
// children of a group following the group itself.
//
// Slide relationships — a picture, a chart, an external hyperlink — are numbered
// in this order, starting at rId2 because rId1 is the slide layout. Every part
// of the writer that has to agree on an rId therefore has to derive it from this
// one walk: the .rels file, the picture and chart emitters, and the hyperlink
// map. They used to walk only the top-level shapes, which is what left a picture
// or chart inside a group referring to an rId that no relationship defined and,
// in the chart's case, to a chart part that was never written.
func flattenShapes(shapes []Shape) []Shape {
	var out []Shape
	var visit func([]Shape)
	visit = func(list []Shape) {
		for _, s := range list {
			out = append(out, s)
			if g, ok := s.(*GroupShape); ok {
				visit(g.shapes)
			}
		}
	}
	visit(shapes)
	return out
}

// collectDrawingShapes returns all DrawingShapes from a shape list,
// including those nested inside GroupShapes (recursively).
func collectDrawingShapes(shapes []Shape) []*DrawingShape {
	var result []*DrawingShape
	for _, shape := range shapes {
		switch s := shape.(type) {
		case *DrawingShape:
			if s.data != nil || s.path != "" {
				result = append(result, s)
			}
		case *GroupShape:
			result = append(result, collectDrawingShapes(s.shapes)...)
		}
	}
	return result
}

// --- Rich Text Shape XML ---

// xfrmAttrs builds the attribute string for <a:xfrm> including rotation and flip.
func xfrmAttrs(b *BaseShape) string {
	var sb strings.Builder
	if b.rotation != 0 {
		fmt.Fprintf(&sb, ` rot="%d"`, b.rotation*60000)
	}
	if b.flipHorizontal {
		sb.WriteString(` flipH="1"`)
	}
	if b.flipVertical {
		sb.WriteString(` flipV="1"`)
	}
	return sb.String()
}

func (w *PPTXWriter) writeRichTextShapeXML(s *RichTextShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("TextBox %d", id)
	}

	xfAttrs := xfrmAttrs(&s.BaseShape)

	fillXML := w.writeFillXML(s.GetFill())
	borderXML := w.writeBorderXMLWithEnds(s.GetBorder(), s.headEnd, s.tailEnd)

	var paragraphsXML strings.Builder
	for _, para := range s.paragraphs {
		paragraphsXML.WriteString(w.writeParagraphXML(para))
	}

	descrAttr := ""
	if s.description != "" {
		descrAttr = fmt.Sprintf(` descr="%s"`, xmlEscape(s.description))
	}

	return fmt.Sprintf(`      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="%s"%s/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
%s
%s%s        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="%s" numCol="%d"%s>%s</a:bodyPr>
          <a:lstStyle/>
%s        </p:txBody>
      </p:sp>
`, id, xmlEscape(name), descrAttr, xfAttrs,
		s.offsetX, s.offsetY, s.width, s.height,
		shapeGeomXML("rect", nil, s.customPath, "          "),
		fillXML, borderXML,
		boolToWrap(s.wordWrap), s.columns, textAnchorAttr(s.textAnchor),
		normAutofitXML(s.fontScale),
		paragraphsXML.String())
}

func boolToWrap(wrap bool) string {
	if wrap {
		return "square"
	}
	return "none"
}

// textAnchorAttr returns the anchor attribute string for <a:bodyPr>.
func textAnchorAttr(anchor TextAnchorType) string {
	if anchor == "" || anchor == TextAnchorNone {
		return ""
	}
	return fmt.Sprintf(` anchor="%s"`, string(anchor))
}

// normAutofitXML returns the <a:normAutofit> child element for <a:bodyPr> if fontScale is set.
func normAutofitXML(fontScale int) string {
	if fontScale > 0 && fontScale != 100000 {
		return fmt.Sprintf(`<a:normAutofit fontScale="%d"/>`, fontScale)
	}
	return ""
}

func (w *PPTXWriter) writeParagraphXML(para *Paragraph) string {
	align := para.alignment
	algn := ""
	if align.Horizontal != "" {
		algn = fmt.Sprintf(` algn="%s"`, align.Horizontal)
	}

	// Indentation level
	if align.Level > 0 {
		algn += fmt.Sprintf(` lvl="%d"`, align.Level)
	}

	var elementsXML strings.Builder
	for _, elem := range para.elements {
		switch e := elem.(type) {
		case *TextRun:
			elementsXML.WriteString(w.writeTextRunXML(e))
		case *BreakElement:
			elementsXML.WriteString("          <a:br/>\n")
		}
	}

	spacing := ""
	if para.lineSpacing < 0 {
		// spcPct: stored as negative percentage * 1000
		spacing = fmt.Sprintf(`
            <a:lnSpc><a:spcPct val="%d"/></a:lnSpc>`, -para.lineSpacing)
	} else if para.lineSpacing > 0 {
		spacing = fmt.Sprintf(`
            <a:lnSpc><a:spcPts val="%d"/></a:lnSpc>`, para.lineSpacing)
	}
	if para.spaceBefore > 0 {
		spacing += fmt.Sprintf(`
            <a:spcBef><a:spcPts val="%d"/></a:spcBef>`, para.spaceBefore)
	}
	if para.spaceAfter > 0 {
		spacing += fmt.Sprintf(`
            <a:spcAft><a:spcPts val="%d"/></a:spcAft>`, para.spaceAfter)
	}

	// Bullet XML
	bulletXML := ""
	if para.bullet != nil {
		bulletXML = w.writeBulletXML(para.bullet)
	}

	return fmt.Sprintf(`          <a:p>
            <a:pPr%s>%s%s
            </a:pPr>
%s          </a:p>
`, algn, spacing, bulletXML, elementsXML.String())
}

// writeTextRunXML renders a text run inside a shape's <p:txBody>.
func (w *PPTXWriter) writeTextRunXML(tr *TextRun) string {
	return w.writeTextRunXMLAt(tr, "            ")
}

// writeTextRunXMLAt renders a text run with a caller-chosen indentation, so one
// emitter serves both a shape's <p:txBody> and a table cell's <a:txBody>.
//
// A table cell used to build its own <a:r> by hand, which dropped everything the
// run's font carried except the size: name, East Asian name, bold, italic,
// underline, strikethrough, colour, and the hyperlink. The reader reads all of
// them back, so every load-and-save quietly lost them.
func (w *PPTXWriter) writeTextRunXMLAt(tr *TextRun, indent string) string {
	inner := indent + "  "
	font := tr.font
	attrs := fmt.Sprintf(` lang="en-US" sz="%d" dirty="0"`, font.Size*100)

	if font.Bold {
		attrs += ` b="1"`
	}
	if font.Italic {
		attrs += ` i="1"`
	}
	if font.Underline != UnderlineNone && font.Underline != "" {
		attrs += fmt.Sprintf(` u="%s"`, font.Underline)
	}
	if font.Strikethrough {
		attrs += ` strike="sngStrike"`
	}

	solidFill := ""
	if font.Color.ARGB != "" {
		solidFill = fmt.Sprintf(`
%s<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, inner, colorRGB(font.Color))
	}

	latin := ""
	if font.Name != "" {
		latin = fmt.Sprintf(`
%s<a:latin typeface="%s"/>`, inner, xmlEscape(font.Name))
	}

	ea := ""
	if font.NameEA != "" {
		ea = fmt.Sprintf(`
%s<a:ea typeface="%s"/>`, inner, xmlEscape(font.NameEA))
	}

	// The relationship id is a placeholder the slide writer replaces once every
	// relationship has been numbered. Only a hyperlink hyperlinkRel accepts is
	// given a number, so a refused one is left out entirely rather than emitted
	// with an id nothing defines.
	hlink := ""
	if tr.hyperlink != nil {
		if _, _, ok := w.hyperlinkRel(tr.hyperlink); ok {
			action := ""
			if tr.hyperlink.IsInternal {
				// PowerPoint marks a jump to another slide with an action; the
				// r:id names the slide relationship it jumps to.
				action = ` action="` + actionSlideJump + `"`
			}
			hlink = fmt.Sprintf(`
%s<a:hlinkClick r:id="%s"%s/>`, inner, fmt.Sprintf("rId_hlink_%p", tr), action)
		}
	}

	return fmt.Sprintf(`%s<a:r>
%s<a:rPr%s>%s%s%s%s
%s</a:rPr>
%s<a:t>%s</a:t>
%s</a:r>
`, indent, inner, attrs, solidFill, latin, ea, hlink, inner, inner, xmlEscape(tr.text), indent)
}

// --- Geometry ---

// shapeGeomXML serialises the geometry child of a <p:spPr>: <a:custGeom> when
// the shape carries a custom path, otherwise <a:prstGeom> with its adjustment
// list.
//
// Custom geometry and preset geometry are alternatives in the schema, so a
// shape with a custom path must not also carry a prstGeom. The writer used to
// emit <a:prstGeom prst="rect"> unconditionally, which turned every freeform
// shape into a rectangle on save — the reader and the renderer both handled the
// path, only the byte that leaves the library did not.
func shapeGeomXML(prst string, adj map[string]int, cp *CustomGeomPath, indent string) string {
	if cp != nil && len(cp.Commands) > 0 {
		return custGeomXML(cp, indent)
	}
	return prstGeomXML(prst, adj, indent)
}

// prstGeomXML serialises <a:prstGeom> together with its adjustment list.
//
// Every caller used to emit a hardcoded <a:avLst/>, so a rounded rectangle's
// corner radius, an arrow's proportions, a callout's tail position and a bent
// connector's knee were read from the file, drawn by the renderer, and then
// dropped on save — PowerPoint reverted each of them to the preset default.
//
// The names are sorted because Go randomises map iteration: without it the same
// model produces different bytes on two consecutive saves.
func prstGeomXML(prst string, adj map[string]int, indent string) string {
	inner := indent + "  "
	var b strings.Builder
	fmt.Fprintf(&b, "%s<a:prstGeom prst=\"%s\">\n", indent, xmlEscape(prst))
	if len(adj) == 0 {
		fmt.Fprintf(&b, "%s<a:avLst/>\n", inner)
	} else {
		fmt.Fprintf(&b, "%s<a:avLst>\n", inner)
		for _, name := range sortedAdjustNames(adj) {
			fmt.Fprintf(&b, "%s  <a:gd name=\"%s\" fmla=\"val %d\"/>\n", inner, xmlEscape(name), adj[name])
		}
		fmt.Fprintf(&b, "%s</a:avLst>\n", inner)
	}
	fmt.Fprintf(&b, "%s</a:prstGeom>", indent)
	return b.String()
}

// sortedAdjustNames returns an adjustment map's keys in a stable order.
func sortedAdjustNames(adj map[string]int) []string {
	names := make([]string, 0, len(adj))
	for name := range adj {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}

// custGeomXML serialises a custom geometry path as <a:custGeom>.
//
// The child order is fixed by CT_CustomGeometry2D — avLst, gdLst, ahLst, cxnLst,
// rect, pathLst — and the four empty lists are emitted because PowerPoint
// writes them and the schema expects pathLst last. The commands mirror exactly
// what the reader parses, so a path survives a read/write round trip unchanged.
//
// A path whose coordinate space is unknown (w or h unset, which is what a path
// built through the API starts as) falls back to the shape's own extents, and
// the renderer draws nothing at all for a non-positive space. Negative
// coordinates are legal in a path, so they are written as they stand.
func custGeomXML(cp *CustomGeomPath, indent string) string {
	inner := indent + "  "
	w, h := cp.Width, cp.Height
	if w <= 0 || h <= 0 {
		// The caller's shape extents are not visible here, so fall back to the
		// path's own bounds: the smallest box that contains every point.
		w, h = pathBounds(cp)
	}
	var b strings.Builder
	fmt.Fprintf(&b, "%s<a:custGeom>\n", indent)
	fmt.Fprintf(&b, "%s<a:avLst/>\n", inner)
	fmt.Fprintf(&b, "%s<a:gdLst/>\n", inner)
	fmt.Fprintf(&b, "%s<a:ahLst/>\n", inner)
	fmt.Fprintf(&b, "%s<a:cxnLst/>\n", inner)
	fmt.Fprintf(&b, "%s<a:rect l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\n", inner)
	fmt.Fprintf(&b, "%s<a:pathLst>\n", inner)
	fmt.Fprintf(&b, "%s  <a:path w=\"%d\" h=\"%d\">\n", inner, w, h)
	for _, cmd := range cp.Commands {
		b.WriteString(pathCommandXML(cmd, inner+"    "))
	}
	fmt.Fprintf(&b, "%s  </a:path>\n", inner)
	fmt.Fprintf(&b, "%s</a:pathLst>\n", inner)
	fmt.Fprintf(&b, "%s</a:custGeom>", indent)
	return b.String()
}

// pathCommandXML serialises one path command. A command with fewer points than
// it requires is skipped rather than emitted half-formed: the reader appends
// whatever <a:pt> children it sees to the last command, so a truncated
// cubicBezTo would parse as a two-point curve and draw a wrong shape.
func pathCommandXML(cmd PathCommand, indent string) string {
	switch cmd.Type {
	case "moveTo", "lnTo":
		if len(cmd.Pts) < 1 {
			return ""
		}
		return fmt.Sprintf("%s<a:%s><a:pt x=\"%d\" y=\"%d\"/></a:%s>\n",
			indent, cmd.Type, cmd.Pts[0].X, cmd.Pts[0].Y, cmd.Type)
	case "cubicBezTo":
		if len(cmd.Pts) < 3 {
			return ""
		}
		return fmt.Sprintf("%s<a:cubicBezTo><a:pt x=\"%d\" y=\"%d\"/><a:pt x=\"%d\" y=\"%d\"/><a:pt x=\"%d\" y=\"%d\"/></a:cubicBezTo>\n",
			indent, cmd.Pts[0].X, cmd.Pts[0].Y, cmd.Pts[1].X, cmd.Pts[1].Y, cmd.Pts[2].X, cmd.Pts[2].Y)
	case "quadBezTo":
		if len(cmd.Pts) < 2 {
			return ""
		}
		return fmt.Sprintf("%s<a:quadBezTo><a:pt x=\"%d\" y=\"%d\"/><a:pt x=\"%d\" y=\"%d\"/></a:quadBezTo>\n",
			indent, cmd.Pts[0].X, cmd.Pts[0].Y, cmd.Pts[1].X, cmd.Pts[1].Y)
	case "arcTo":
		return fmt.Sprintf("%s<a:arcTo wR=\"%d\" hR=\"%d\" stAng=\"%d\" swAng=\"%d\"/>\n",
			indent, cmd.WR, cmd.HR, cmd.StAng, cmd.SwAng)
	case "close":
		return fmt.Sprintf("%s<a:close/>\n", indent)
	default:
		// An unknown command names a path element this library does not model.
		// Dropping it keeps the file valid; inventing an element would not.
		return ""
	}
}

// pathBounds returns the smallest box containing every point of a path, used
// when the path carries no coordinate space of its own. It never returns zero,
// so the result is always a usable <a:path w h>.
func pathBounds(cp *CustomGeomPath) (int64, int64) {
	maxX, maxY := int64(1), int64(1)
	for _, cmd := range cp.Commands {
		for _, p := range cmd.Pts {
			if p.X > maxX {
				maxX = p.X
			}
			if p.Y > maxY {
				maxY = p.Y
			}
		}
		if cmd.Type == "arcTo" {
			if cmd.WR > maxX {
				maxX = cmd.WR
			}
			if cmd.HR > maxY {
				maxY = cmd.HR
			}
		}
	}
	return maxX, maxY
}

// --- Drawing Shape XML ---

func (w *PPTXWriter) writeDrawingShapeXML(s *DrawingShape, shapeID *int, slideNum int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Picture %d", id)
	}

	// Find the relationship ID for this image within the current slide.
	// Must match the ordering in writeSlideRels exactly.
	currentSlide := w.presentation.slides[slideNum-1]
	relIdx := w.countRelIdxBefore(currentSlide.shapes, s)

	shadowXML := ""
	if s.shadow != nil && s.shadow.Visible {
		shadowXML = fmt.Sprintf(`
          <a:effectLst>
            <a:outerShdw blurRad="%d" dist="%d" dir="%d" algn="bl" rotWithShape="0">
              <a:srgbClr val="%s">
                <a:alpha val="%d"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>`,
			s.shadow.BlurRadius*12700,
			s.shadow.Distance*12700,
			s.shadow.Direction*60000,
			colorRGB(s.shadow.Color),
			s.shadow.Alpha*1000)
	}

	return fmt.Sprintf(`      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="%d" name="%s" descr="%s"/>
          <p:cNvPicPr>
            <a:picLocks noChangeAspect="1"/>
          </p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          %s
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>%s
        </p:spPr>
      </p:pic>
`, id, xmlEscape(name), xmlEscape(s.description),
		blipFillChildrenXML(relIdx, s.alpha, s.cropLeft, s.cropTop, s.cropRight, s.cropBottom),
		xfrmAttrs(&s.BaseShape),
		s.offsetX, s.offsetY, s.width, s.height,
		shadowXML)
}

// blipFillChildrenXML serialises the children of <p:blipFill> that describe the
// image itself: the reference, its crop and the picture's opacity.
//
// They are built together because CT_BlipFillProperties fixes their order —
// blip, then srcRect, then stretch — and the opacity is a child *of the blip*,
// so formatting the three independently would either reorder them or leave
// blank lines behind when one is absent. A crop or an opacity that is absent
// produces no element, which is what PowerPoint does.
//
// alpha is in 1/1000 of a percent, where 0 means fully opaque; a stored 100000
// is equally a no-op, so neither produces an <a:alphaModFix>. The srcRect
// values are signed percentages in the same 1/1000 unit, and negatives are
// meaningful (a picture cropped outwards), so they are written as they stand.
func blipFillChildrenXML(relIdx, alpha, left, top, right, bottom int) string {
	blip := fmt.Sprintf(`<a:blip r:embed="rId%d"/>`, relIdx)
	if alpha > 0 && alpha < 100000 {
		blip = fmt.Sprintf(`<a:blip r:embed="rId%d">
            <a:alphaModFix amt="%d"/>
          </a:blip>`, relIdx, alpha)
	}
	if left != 0 || top != 0 || right != 0 || bottom != 0 {
		blip += fmt.Sprintf(`
          <a:srcRect l="%d" t="%d" r="%d" b="%d"/>`, left, top, right, bottom)
	}
	return blip
}

// --- Auto Shape XML ---

func (w *PPTXWriter) writeAutoShapeXML(s *AutoShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Shape %d", id)
	}

	fillXML := w.writeFillXML(s.GetFill())
	borderXML := w.writeBorderXMLWithEnds(s.GetBorder(), s.headEnd, s.tailEnd)

	textXML := ""
	if s.text != "" {
		textXML = fmt.Sprintf(`
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>%s</a:t>
            </a:r>
          </a:p>
        </p:txBody>`, xmlEscape(s.text))
	}

	descrAttr := ""
	if s.description != "" {
		descrAttr = fmt.Sprintf(` descr="%s"`, xmlEscape(s.description))
	}

	return fmt.Sprintf(`      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="%s"%s/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
%s
%s%s        </p:spPr>%s
      </p:sp>
`, id, xmlEscape(name), descrAttr,
		xfrmAttrs(&s.BaseShape),
		s.offsetX, s.offsetY, s.width, s.height,
		shapeGeomXML(string(s.shapeType), s.adjustValues, nil, "          "),
		fillXML, borderXML, textXML)
}

// --- Line Shape XML ---

func (w *PPTXWriter) writeLineShapeXML(s *LineShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Line %d", id)
	}

	// Arrow ends, as they are written inside <a:ln> below.
	endsXML := lineEndXML(s.headEnd, s.tailEnd, "\n            ")

	prstGeom := "line"
	if s.connectorType != "" {
		prstGeom = s.connectorType
	}

	// Build dash style XML
	var dashXML string
	switch s.lineStyle {
	case BorderDash:
		dashXML = "\n            <a:prstDash val=\"dash\"/>"
	case BorderDot:
		dashXML = "\n            <a:prstDash val=\"dot\"/>"
	}

	return fmt.Sprintf(`      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvCxnSpPr/>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
%s
          <a:ln w="%d">
            <a:solidFill>
              <a:srgbClr val="%s"/>
            </a:solidFill>%s%s
          </a:ln>
        </p:spPr>
      </p:cxnSp>
`, id, xmlEscape(name),
		xfrmAttrs(&s.BaseShape),
		s.offsetX, s.offsetY, s.width, s.height,
		shapeGeomXML(prstGeom, s.adjustValues, s.customPath, "          "),
		int64(s.GetLineWidthEMU()),
		colorRGB(s.lineColor),
		dashXML, endsXML)
}

// lineEndXML serialises the arrow ends of an <a:ln>. Both arrow ends are drawn
// by the renderer and read back from the file, so a shape that carries them has
// to write them; prefix is placed before each element so the same helper serves
// an element written inline and one written on its own line.
//
// An absent end, or one whose type is ArrowNone, produces nothing: that is the
// schema default, and writing type="none" explicitly would be noise.
func lineEndXML(head, tail *LineEnd, prefix string) string {
	var b strings.Builder
	if head != nil && head.Type != ArrowNone && head.Type != "" {
		fmt.Fprintf(&b, "%s<a:headEnd type=\"%s\" w=\"%s\" len=\"%s\"/>", prefix, head.Type, head.Width, head.Length)
	}
	if tail != nil && tail.Type != ArrowNone && tail.Type != "" {
		fmt.Fprintf(&b, "%s<a:tailEnd type=\"%s\" w=\"%s\" len=\"%s\"/>", prefix, tail.Type, tail.Width, tail.Length)
	}
	return b.String()
}

// --- Table Shape XML ---

// tableCellAt returns the cell at (row, col), or nil when the row is shorter
// than the table's column count.
//
// A row can be short. The reader builds one row per <a:tr> and one cell per
// <a:tc> inside it without checking the count against <a:tblGrid>, and the
// renderer walks only the cells that are there. Indexing the row blindly made
// Save panic — recovered into an error, but the presentation could not be
// written at all — on a package Open had accepted.
func tableCellAt(t *TableShape, row, col int) *TableCell {
	if row < 0 || row >= len(t.rows) || col < 0 || col >= len(t.rows[row]) {
		return nil
	}
	return t.rows[row][col]
}

// tableColumnWidth returns the width of column i in EMU: the value read from
// the package when the table came from a file, an even share otherwise.
//
// The reader records the real <a:gridCol w> values, so recomputing an even
// split would resize every column of a presentation that was opened and saved.
func tableColumnWidth(t *TableShape, i int) int64 {
	if len(t.colWidths) == t.numCols && t.colWidths[i] > 0 {
		return t.colWidths[i]
	}
	if t.numCols > 0 {
		return t.width / int64(t.numCols)
	}
	return 0
}

// tableRowHeight is tableColumnWidth's counterpart for <a:tr h>.
func tableRowHeight(t *TableShape, i int) int64 {
	if len(t.rowHeights) == t.numRows && t.rowHeights[i] > 0 {
		return t.rowHeights[i]
	}
	if t.numRows > 0 {
		return t.height / int64(t.numRows)
	}
	return 0
}

// tableContinuationAttrs renders the attributes of a merge continuation cell.
// A position to the right of the anchor continues the column merge (hMerge),
// one below it continues the row merge (vMerge); a position that is both
// carries both.
func tableContinuationAttrs(right, below bool) string {
	attrs := ""
	if right {
		attrs += ` hMerge="1"`
	}
	if below {
		attrs += ` vMerge="1"`
	}
	return attrs
}

// tableMergeContinuationXML renders the empty cell PowerPoint writes for every
// position a merge covers other than its top-left anchor. The row still holds
// one <a:tc> per column, because that is how the grid is counted.
func tableMergeContinuationXML(attrs string) string {
	return fmt.Sprintf(`              <a:tc%s>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p/>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
`, attrs)
}

// tableCellBorderXML renders one side of a cell's border. tag is the qualified
// element name — a:lnL, a:lnR, a:lnT, a:lnB — because the element has to be in
// the drawingml namespace to mean anything. Our own reader matches on the local
// name and would accept an unqualified one, which is exactly why this is worth
// getting right and testing for.
//
// A side that is not drawn is left out rather than written as <a:noFill/>: the
// reader maps a missing side and an explicit noFill to the same BorderNone, and
// the writer emits no table style that a missing side could inherit one from.
func tableCellBorderXML(tag string, b *Border) string {
	if b == nil || b.Style == BorderNone {
		return ""
	}
	dash := `<a:prstDash val="solid"/>`
	switch b.Style {
	case BorderDash:
		dash = `<a:prstDash val="dash"/>`
	case BorderDot:
		dash = `<a:prstDash val="dot"/>`
	}
	width := ""
	if b.Width > 0 {
		// Border.Width is in points; the file wants EMU.
		width = fmt.Sprintf(` w="%d"`, b.Width*12700)
	}
	return fmt.Sprintf(`
                  <%s%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</%s>`,
		tag, width, colorRGB(b.Color), dash, tag)
}

// tableCellPrXML renders <a:tcPr>. Borders come before the fill, which is the
// order CT_TableCellProperties defines.
func tableCellPrXML(cell *TableCell) string {
	pr := ""
	if cell.border != nil {
		pr += tableCellBorderXML("a:lnL", cell.border.Left)
		pr += tableCellBorderXML("a:lnR", cell.border.Right)
		pr += tableCellBorderXML("a:lnT", cell.border.Top)
		pr += tableCellBorderXML("a:lnB", cell.border.Bottom)
	}
	if cell.fill != nil && cell.fill.Type == FillSolid {
		pr += fmt.Sprintf(`
                  <a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, colorRGB(cell.fill.Color))
	}
	return pr
}

func (w *PPTXWriter) writeTableShapeXML(s *TableShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Table %d", id)
	}

	var gridCols strings.Builder
	for i := 0; i < s.numCols; i++ {
		gridCols.WriteString(fmt.Sprintf(`            <a:gridCol w="%d"/>
`, tableColumnWidth(s, i)))
	}

	// A merge is written as the spanned cell plus an empty continuation cell
	// for every other position it covers, and the number of <a:tc> elements in
	// a row has to stay equal to the number of columns.
	//
	// The continuations are derived from the anchor's spans rather than only
	// from a cell's own hMerge/vMerge flags: SetColSpan and SetRowSpan are the
	// public API and API.md documents SetColSpan(2) alone, with nothing saying
	// the cell to its right has to be marked by hand. A table read back from a
	// file carries both, and deriving them again gives the same answer.
	claimed := make(map[[2]int]string, len(s.rows))

	var rowsXML strings.Builder
	for i := 0; i < s.numRows; i++ {
		rowsXML.WriteString(fmt.Sprintf(`            <a:tr h="%d">
`, tableRowHeight(s, i)))
		for j := 0; j < s.numCols; j++ {
			cell := tableCellAt(s, i, j)
			if cell == nil {
				// A short row is still a row: keep one cell per column.
				rowsXML.WriteString(tableMergeContinuationXML(""))
				continue
			}
			if attrs, ok := claimed[[2]int{i, j}]; ok {
				rowsXML.WriteString(tableMergeContinuationXML(attrs))
				continue
			}
			if cell.hMerge || cell.vMerge {
				// A continuation as it was read from the file, with nothing to
				// anchor it — whatever wrote it left the anchor out.
				rowsXML.WriteString(tableMergeContinuationXML(tableContinuationAttrs(cell.hMerge, cell.vMerge)))
				continue
			}

			colSpan, rowSpan := cell.colSpan, cell.rowSpan
			if colSpan < 1 {
				colSpan = 1
			}
			if rowSpan < 1 {
				rowSpan = 1
			}
			if j+colSpan > s.numCols {
				colSpan = s.numCols - j
			}
			if i+rowSpan > s.numRows {
				rowSpan = s.numRows - i
			}
			for r := i; r < i+rowSpan; r++ {
				for c := j; c < j+colSpan; c++ {
					if r == i && c == j {
						continue
					}
					claimed[[2]int{r, c}] = tableContinuationAttrs(c > j, r > i)
				}
			}

			spanAttrs := ""
			if colSpan > 1 {
				spanAttrs += fmt.Sprintf(` gridSpan="%d"`, colSpan)
			}
			if rowSpan > 1 {
				spanAttrs += fmt.Sprintf(` rowSpan="%d"`, rowSpan)
			}

			var cellText strings.Builder
			for _, para := range cell.paragraphs {
				cellText.WriteString("                <a:p>\n")
				for _, elem := range para.elements {
					if tr, ok := elem.(*TextRun); ok {
						// The same emitter a shape's text body uses, so a cell
						// run keeps its font, weight, colour and hyperlink.
						cellText.WriteString(w.writeTextRunXMLAt(tr, "                  "))
					}
				}
				cellText.WriteString("                </a:p>\n")
			}
			if cellText.Len() == 0 {
				// <a:txBody> requires at least one paragraph. A cell the reader
				// found no <a:p> in has none.
				cellText.WriteString("                <a:p/>\n")
			}

			rowsXML.WriteString(fmt.Sprintf(`              <a:tc%s>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
%s                </a:txBody>
                <a:tcPr>%s
                </a:tcPr>
              </a:tc>
`, spanAttrs, cellText.String(), tableCellPrXML(cell)))
		}
		rowsXML.WriteString("            </a:tr>\n")
	}

	return fmt.Sprintf(`      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvGraphicFramePr>
            <a:graphicFrameLocks noGrp="1"/>
          </p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="%d" y="%d"/>
          <a:ext cx="%d" cy="%d"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="1" bandRow="1"/>
              <a:tblGrid>
%s              </a:tblGrid>
%s            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
`, id, xmlEscape(name),
		s.offsetX, s.offsetY, s.width, s.height,
		gridCols.String(), rowsXML.String())
}

// --- Fill and Border helpers ---

func (w *PPTXWriter) writeFillXML(f *Fill) string {
	if f == nil {
		return ""
	}
	switch f.Type {
	case FillSolid:
		return fmt.Sprintf("          <a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>\n", colorRGB(f.Color))
	case FillGradientLinear:
		return fmt.Sprintf(`          <a:gradFill>
            <a:gsLst>
              <a:gs pos="0"><a:srgbClr val="%s"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="%s"/></a:gs>
            </a:gsLst>
            <a:lin ang="%d" scaled="1"/>
          </a:gradFill>
`, colorRGB(f.Color), colorRGB(f.EndColor), f.Rotation*60000)
	default:
		return ""
	}
}

// writeBorderXML serialises a Border as an <a:ln> element.
//
// Border.Width is in points (the reader converts from the EMU in the file with
// v/12700, and the renderer multiplies back by 12700), so it has to be
// converted to EMU here. Writing it raw made every border shrink to roughly
// 1/12700 of its intended width on a read/write round trip.
func (w *PPTXWriter) writeBorderXML(b *Border) string {
	return w.writeBorderXMLWithEnds(b, nil, nil)
}

// writeBorderXMLWithEnds is writeBorderXML for a shape that also carries arrow
// ends. A shape's <a:ln> owns its headEnd/tailEnd, so a shape whose line is
// drawn with arrowheads — an arc, or a freeform path — loses them if the ends
// are not emitted with the border they belong to.
//
// The child order is the one CT_LineProperties fixes: the fill, then the dash,
// then headEnd and tailEnd.
func (w *PPTXWriter) writeBorderXMLWithEnds(b *Border, head, tail *LineEnd) string {
	if b == nil || b.Style == BorderNone {
		return ""
	}
	widthEMU := maxInt(b.Width, 1) * 12700
	var dashXML string
	switch b.Style {
	case BorderDash:
		dashXML = "<a:prstDash val=\"dash\"/>"
	case BorderDot:
		dashXML = "<a:prstDash val=\"dot\"/>"
	}
	return fmt.Sprintf("          <a:ln w=\"%d\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>%s%s</a:ln>\n",
		widthEMU, colorRGB(b.Color), dashXML, lineEndXML(head, tail, ""))
}

// --- Media ---

func (w *PPTXWriter) writeMedia(zw *zip.Writer) error {
	imgIdx := 1
	for _, slide := range w.presentation.slides {
		for _, ds := range collectDrawingShapes(slide.shapes) {
			if ds.data != nil {
				ext := w.getImageExtension(ds)
				fw, err := zw.Create(fmt.Sprintf("ppt/media/image%d.%s", imgIdx, ext))
				if err != nil {
					return err
				}
				if _, err := fw.Write(ds.data); err != nil {
					return err
				}
				imgIdx++
			} else if ds.path != "" {
				info, err := os.Stat(ds.path)
				if err != nil {
					return fmt.Errorf("failed to stat image %s: %w", ds.path, err)
				}
				if info.Size() > maxImageFileSize {
					return fmt.Errorf("image file %s too large: %d bytes (max %d)", ds.path, info.Size(), maxImageFileSize)
				}
				data, err := os.ReadFile(ds.path)
				if err != nil {
					return fmt.Errorf("failed to read image %s: %w", ds.path, err)
				}
				ext := w.getImageExtension(ds)
				fw, err := zw.Create(fmt.Sprintf("ppt/media/image%d.%s", imgIdx, ext))
				if err != nil {
					return err
				}
				if _, err := fw.Write(data); err != nil {
					return err
				}
				imgIdx++
			}
		}
	}
	return nil
}

// getChartIndex returns the 1-based index of a chart part, numbering charts in
// the same document order that WriteTo emits them in. The two have to agree or a
// chart relationship points at another slide's chart.
func (w *PPTXWriter) getChartIndex(target *ChartShape) int {
	idx := 1
	for _, slide := range w.presentation.slides {
		for _, shape := range flattenShapes(slide.shapes) {
			if cs, ok := shape.(*ChartShape); ok {
				if cs == target {
					return idx
				}
				idx++
			}
		}
	}
	return idx
}

// --- Chart Shape XML ---

func (w *PPTXWriter) writeChartShapeXML(s *ChartShape, shapeID *int, slideNum int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Chart %d", id)
	}

	// Find chart rel ID — must match ordering in writeSlideRels exactly.
	relIdx := w.countRelIdxBefore(w.presentation.slides[slideNum-1].shapes, s)

	return fmt.Sprintf(`      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvGraphicFramePr>
            <a:graphicFrameLocks noGrp="1"/>
          </p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="%d" y="%d"/>
          <a:ext cx="%d" cy="%d"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId%d"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
`, id, xmlEscape(name),
		s.offsetX, s.offsetY, s.width, s.height,
		relIdx)
}

// --- Group Shape XML ---

func (w *PPTXWriter) writeGroupShapeXML(g *GroupShape, shapeID *int, slideNum int) string {
	id := *shapeID
	*shapeID++

	name := g.name
	if name == "" {
		name = fmt.Sprintf("Group %d", id)
	}

	var childXML strings.Builder
	for _, shape := range g.shapes {
		// Every shape kind the slide can hold has to be written here too.
		// A missing case does not fail loudly: the child simply disappears from
		// the saved file, so a group's contents are lost on a round trip.
		switch s := shape.(type) {
		case *PlaceholderShape:
			childXML.WriteString(w.writePlaceholderShapeXML(s, shapeID))
		case *RichTextShape:
			childXML.WriteString(w.writeRichTextShapeXML(s, shapeID))
		case *AutoShape:
			childXML.WriteString(w.writeAutoShapeXML(s, shapeID))
		case *LineShape:
			childXML.WriteString(w.writeLineShapeXML(s, shapeID))
		case *DrawingShape:
			childXML.WriteString(w.writeDrawingShapeXML(s, shapeID, slideNum))
		case *TableShape:
			childXML.WriteString(w.writeTableShapeXML(s, shapeID))
		case *ChartShape:
			childXML.WriteString(w.writeChartShapeXML(s, shapeID, slideNum))
		case *GroupShape:
			childXML.WriteString(w.writeGroupShapeXML(s, shapeID, slideNum))
		}
	}

	return fmt.Sprintf(`      <p:grpSp>
        <p:nvGrpSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
            <a:chOff x="%d" y="%d"/>
            <a:chExt cx="%d" cy="%d"/>
          </a:xfrm>
        </p:grpSpPr>
%s      </p:grpSp>
`, id, xmlEscape(name),
		xfrmAttrs(&g.BaseShape),
		g.offsetX, g.offsetY, g.width, g.height,
		g.offsetX, g.offsetY, g.width, g.height,
		childXML.String())
}

// --- Placeholder Shape XML ---

func (w *PPTXWriter) writePlaceholderShapeXML(s *PlaceholderShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Placeholder %d", id)
	}

	var paragraphsXML strings.Builder
	for _, para := range s.paragraphs {
		paragraphsXML.WriteString(w.writeParagraphXML(para))
	}

	return fmt.Sprintf(`      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph%s/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
%s        </p:txBody>
      </p:sp>
`, id, xmlEscape(name),
		placeholderAttrsXML(s.phType, s.phIdx),
		xfrmAttrs(&s.BaseShape),
		s.offsetX, s.offsetY, s.width, s.height,
		paragraphsXML.String())
}

// placeholderAttrsXML serialises the attributes of <p:ph>.
//
// The type is optional in the schema: a placeholder that inherits its type from
// the layout omits the attribute entirely, and such placeholders are common.
// ST_PlaceholderType is an enumeration with no empty member, so emitting
// type="" produces a file PowerPoint refuses to open cleanly — while this
// library's own reader accepts it, which is why no round-trip test can see it.
func placeholderAttrsXML(phType PlaceholderType, phIdx int) string {
	if phType == "" {
		return fmt.Sprintf(` idx="%d"`, phIdx)
	}
	return fmt.Sprintf(` type="%s" idx="%d"`, xmlEscape(string(phType)), phIdx)
}

// --- Notes Slide ---

func (w *PPTXWriter) writeNotesSlide(zw *zip.Writer, slide *Slide, slideNum int) error {
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Notes Placeholder"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
%s        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, notesBodyXML(slide.notes))

	if err := writeRawXMLToZip(zw, fmt.Sprintf("ppt/notesSlides/notesSlide%d.xml", slideNum), content); err != nil {
		return err
	}

	// Notes slide rels
	rels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s" Target="../slides/slide%d.xml"/>
</Relationships>`, nsRelationships, relTypeSlide, slideNum)
	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/notesSlides/_rels/notesSlide%d.xml.rels", slideNum), rels)
}

// notesBodyXML renders the notes text as the paragraphs of a text body.
//
// A note is stored as one paragraph per line, which is what PowerPoint writes
// and what the reader reads back: putting the whole note in a single run made
// every note a single paragraph, and left the line structure to survive only as
// escaped newlines a consumer has to notice. A blank line is a paragraph of its
// own; a line that ends in a space keeps it, which is what xml:space says.
func notesBodyXML(notes string) string {
	lines := strings.Split(strings.NewReplacer("\r\n", "\n", "\r", "\n").Replace(notes), "\n")

	var body strings.Builder
	for _, line := range lines {
		if line == "" {
			body.WriteString("          <a:p/>\n")
			continue
		}
		preserve := ""
		if strings.TrimSpace(line) != line {
			preserve = ` xml:space="preserve"`
		}
		fmt.Fprintf(&body, `          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t%s>%s</a:t>
            </a:r>
          </a:p>
`, preserve, xmlEscape(line))
	}
	return body.String()
}

// --- Bullet XML ---

func (w *PPTXWriter) writeBulletXML(b *Bullet) string {
	if b.Type == BulletTypeNone {
		return "\n              <a:buNone/>"
	}

	var sb strings.Builder

	// Bullet color
	if b.Color != nil {
		sb.WriteString(fmt.Sprintf("\n              <a:buClr><a:srgbClr val=\"%s\"/></a:buClr>", colorRGB(*b.Color)))
	}

	// Bullet size
	if b.Size > 0 && b.Size != 100 {
		sb.WriteString(fmt.Sprintf("\n              <a:buSzPct val=\"%d000\"/>", b.Size))
	}

	switch b.Type {
	case BulletTypeChar:
		fontAttr := ""
		if b.Font != "" {
			fontAttr = fmt.Sprintf("\n              <a:buFont typeface=\"%s\"/>", xmlEscape(b.Font))
		}
		sb.WriteString(fontAttr)
		// <a:buChar> without a character is not a bullet, and an empty
		// char="" is what makes PowerPoint offer to repair the file.
		char := b.Style
		if char == "" {
			char = defaultBulletChar
		}
		sb.WriteString(fmt.Sprintf("\n              <a:buChar char=\"%s\"/>", xmlEscape(char)))
	case BulletTypeNumeric, BulletTypeAutoNum:
		// The two constants describe the same thing to PowerPoint — there is
		// one element for a numbered bullet and no automatic/manual split — so
		// both write <a:buAutoNum>. AutoNum used to fall through the switch and
		// write nothing at all, while the renderer numbered the paragraph: the
		// preview showed "1." and the file had no bullet.
		format := b.NumFormat
		if format == "" {
			format = defaultNumFormat
		}
		startAt := b.StartAt
		if startAt < defaultBulletStart {
			startAt = defaultBulletStart
		}
		sb.WriteString(fmt.Sprintf("\n              <a:buAutoNum type=\"%s\" startAt=\"%d\"/>", xmlEscape(format), startAt))
	}

	return sb.String()
}
