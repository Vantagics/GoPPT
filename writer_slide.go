package gopresentation

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"
)

// buildHyperlinkRelMap pre-computes the relationship IDs for all hyperlinks in a slide.
// This ensures the XML shape content and the .rels file use the same IDs.
func (w *PPTXWriter) buildHyperlinkRelMap(slide *Slide) map[*TextRun]string {
	m := make(map[*TextRun]string)
	relIdx := 2 // rId1 is slideLayout
	for _, shape := range slide.shapes {
		switch s := shape.(type) {
		case *DrawingShape:
			if s.data != nil || s.path != "" {
				relIdx++
			}
		case *ChartShape:
			relIdx++
		}
		if bs, ok := shape.(*RichTextShape); ok {
			for _, para := range bs.paragraphs {
				for _, elem := range para.elements {
					if tr, ok := elem.(*TextRun); ok && tr.hyperlink != nil && !tr.hyperlink.IsInternal {
						m[tr] = fmt.Sprintf("rId%d", relIdx)
						relIdx++
					}
				}
			}
		}
	}
	return m
}

func (w *PPTXWriter) writeSlide(zw *zip.Writer, slide *Slide, slideNum int) error {
	// Pre-compute hyperlink relationship IDs so the XML references match the .rels file.
	hlinkRelMap := w.buildHyperlinkRelMap(slide)

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

func (w *PPTXWriter) writeSlideRels(zw *zip.Writer, slide *Slide, slideNum int) error {
	hlinkRelMap := w.buildHyperlinkRelMap(slide)

	rels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s" Target="../slideLayouts/slideLayout1.xml"/>`, nsRelationships, relTypeSlideLayout)

	relIdx := 2
	for _, shape := range slide.shapes {
		switch s := shape.(type) {
		case *DrawingShape:
			if s.data != nil || s.path != "" {
				imgIdx := w.getImageIndex(slide, s)
				ext := w.getImageExtension(s)
				rels += fmt.Sprintf(`
  <Relationship Id="rId%d" Type="%s" Target="../media/image%d.%s"/>`,
					relIdx, relTypeImage, imgIdx, ext)
				relIdx++
			}
		case *ChartShape:
			chartIdx := w.getChartIndex(s)
			rels += fmt.Sprintf(`
  <Relationship Id="rId%d" Type="%s" Target="../charts/chart%d.xml"/>`,
				relIdx, relTypeChart, chartIdx)
			relIdx++
		}
		// Handle hyperlinks in shapes
		if bs, ok := shape.(*RichTextShape); ok {
			for _, para := range bs.paragraphs {
				for _, elem := range para.elements {
					if tr, ok := elem.(*TextRun); ok && tr.hyperlink != nil {
						if !tr.hyperlink.IsInternal {
							rid := hlinkRelMap[tr]
							rels += fmt.Sprintf(`
  <Relationship Id="%s" Type="%s" Target="%s" TargetMode="External"/>`,
								rid, relTypeHyperlink, xmlEscape(tr.hyperlink.URL))
							relIdx++
						}
					}
				}
			}
		}
	}

	// Comments relationship
	if len(slide.comments) > 0 {
		rels += fmt.Sprintf(`
  <Relationship Id="rId%d" Type="%s" Target="../comments/comment%d.xml"/>`,
			relIdx, relTypeComment, slideNum)
		relIdx++
	}

	// Notes slide relationship
	if slide.notes != "" {
		rels += fmt.Sprintf(`
  <Relationship Id="rId%d" Type="%s" Target="../notesSlides/notesSlide%d.xml"/>`,
			relIdx, relTypeNotesSlide, slideNum)
		relIdx++
	}

	rels += `
</Relationships>`
	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum), rels)
}

func (w *PPTXWriter) getImageIndex(slide *Slide, target *DrawingShape) int {
	idx := 1
	for _, sl := range w.presentation.slides {
		for _, shape := range sl.shapes {
			if ds, ok := shape.(*DrawingShape); ok {
				if ds.data != nil || ds.path != "" {
					if ds == target {
						return idx
					}
					idx++
				}
			}
		}
	}
	return idx
}

// --- Rich Text Shape XML ---

func (w *PPTXWriter) writeRichTextShapeXML(s *RichTextShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("TextBox %d", id)
	}

	rotation := ""
	if s.rotation != 0 {
		rotation = fmt.Sprintf(` rot="%d"`, s.rotation*60000)
	}

	fillXML := w.writeFillXML(s.GetFill())
	borderXML := w.writeBorderXML(s.GetBorder())

	var paragraphsXML strings.Builder
	for _, para := range s.paragraphs {
		paragraphsXML.WriteString(w.writeParagraphXML(para))
	}

	return fmt.Sprintf(`      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
%s%s        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="%s" numCol="%d"/>
          <a:lstStyle/>
%s        </p:txBody>
      </p:sp>
`, id, xmlEscape(name), rotation,
		s.offsetX, s.offsetY, s.width, s.height,
		fillXML, borderXML,
		boolToWrap(s.wordWrap), s.columns,
		paragraphsXML.String())
}

func boolToWrap(wrap bool) string {
	if wrap {
		return "square"
	}
	return "none"
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
	if para.lineSpacing > 0 {
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

func (w *PPTXWriter) writeTextRunXML(tr *TextRun) string {
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
              <a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, font.Color.ARGB[2:])
	}

	latin := ""
	if font.Name != "" {
		latin = fmt.Sprintf(`
              <a:latin typeface="%s"/>`, xmlEscape(font.Name))
	}

	hlinkStart := ""
	hlinkEnd := ""
	if tr.hyperlink != nil && !tr.hyperlink.IsInternal {
		hlinkStart = fmt.Sprintf(`
              <a:hlinkClick r:id="rId_hlink_%p"/>`, tr)
	}

	return fmt.Sprintf(`            <a:r>
              <a:rPr%s>%s%s%s%s
              </a:rPr>
              <a:t>%s</a:t>
            </a:r>
`, attrs, solidFill, latin, hlinkStart, hlinkEnd, xmlEscape(tr.text))
}

// --- Drawing Shape XML ---

func (w *PPTXWriter) writeDrawingShapeXML(s *DrawingShape, shapeID *int, slideNum int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Picture %d", id)
	}

	// Find the relationship ID for this image within the current slide
	relIdx := 2 // rId1 is slideLayout
	currentSlide := w.presentation.slides[slideNum-1]
	for _, shape := range currentSlide.shapes {
		if ds, ok := shape.(*DrawingShape); ok && (ds.data != nil || ds.path != "") {
			if ds == s {
				break
			}
			relIdx++
		}
		if _, ok := shape.(*ChartShape); ok {
			relIdx++
		}
	}

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
			s.shadow.Color.ARGB[2:],
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
          <a:blip r:embed="rId%d"/>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>%s
        </p:spPr>
      </p:pic>
`, id, xmlEscape(name), xmlEscape(s.description),
		relIdx,
		s.offsetX, s.offsetY, s.width, s.height,
		shadowXML)
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
	borderXML := w.writeBorderXML(s.GetBorder())

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

	return fmt.Sprintf(`      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="%s">
            <a:avLst/>
          </a:prstGeom>
%s%s        </p:spPr>%s
      </p:sp>
`, id, xmlEscape(name),
		s.offsetX, s.offsetY, s.width, s.height,
		s.shapeType,
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

	return fmt.Sprintf(`      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvCxnSpPr/>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="line">
            <a:avLst/>
          </a:prstGeom>
          <a:ln w="%d">
            <a:solidFill>
              <a:srgbClr val="%s"/>
            </a:solidFill>
          </a:ln>
        </p:spPr>
      </p:cxnSp>
`, id, xmlEscape(name),
		s.offsetX, s.offsetY, s.width, s.height,
		int64(s.lineWidth)*12700,
		s.lineColor.ARGB[2:])
}

// --- Table Shape XML ---

func (w *PPTXWriter) writeTableShapeXML(s *TableShape, shapeID *int) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Table %d", id)
	}

	colWidth := int64(0)
	if s.numCols > 0 {
		colWidth = s.width / int64(s.numCols)
	}

	var gridCols strings.Builder
	for i := 0; i < s.numCols; i++ {
		gridCols.WriteString(fmt.Sprintf(`            <a:gridCol w="%d"/>
`, colWidth))
	}

	var rowsXML strings.Builder
	rowHeight := int64(0)
	if s.numRows > 0 {
		rowHeight = s.height / int64(s.numRows)
	}

	for i := 0; i < s.numRows; i++ {
		rowsXML.WriteString(fmt.Sprintf(`            <a:tr h="%d">
`, rowHeight))
		for j := 0; j < s.numCols; j++ {
			cell := s.rows[i][j]
			cellFill := ""
			if cell.fill != nil && cell.fill.Type == FillSolid {
				cellFill = fmt.Sprintf(`
                  <a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, cell.fill.Color.ARGB[2:])
			}

			var cellText strings.Builder
			for _, para := range cell.paragraphs {
				cellText.WriteString("                <a:p>\n")
				for _, elem := range para.elements {
					if tr, ok := elem.(*TextRun); ok {
						cellText.WriteString(fmt.Sprintf(`                  <a:r>
                    <a:rPr lang="en-US" sz="%d" dirty="0"/>
                    <a:t>%s</a:t>
                  </a:r>
`, tr.font.Size*100, xmlEscape(tr.text)))
					}
				}
				cellText.WriteString("                </a:p>\n")
			}

			rowsXML.WriteString(fmt.Sprintf(`              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
%s                </a:txBody>
                <a:tcPr>%s
                </a:tcPr>
              </a:tc>
`, cellText.String(), cellFill))
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
		return fmt.Sprintf("          <a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>\n", f.Color.ARGB[2:])
	case FillGradientLinear:
		return fmt.Sprintf(`          <a:gradFill>
            <a:gsLst>
              <a:gs pos="0"><a:srgbClr val="%s"/></a:gs>
              <a:gs pos="100000"><a:srgbClr val="%s"/></a:gs>
            </a:gsLst>
            <a:lin ang="%d" scaled="1"/>
          </a:gradFill>
`, f.Color.ARGB[2:], f.EndColor.ARGB[2:], f.Rotation*60000)
	default:
		return ""
	}
}

func (w *PPTXWriter) writeBorderXML(b *Border) string {
	if b == nil || b.Style == BorderNone {
		return ""
	}
	return fmt.Sprintf("          <a:ln w=\"%d\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill></a:ln>\n",
		b.Width, b.Color.ARGB[2:])
}

// --- Media ---

func (w *PPTXWriter) writeMedia(zw *zip.Writer) error {
	imgIdx := 1
	for _, slide := range w.presentation.slides {
		for _, shape := range slide.shapes {
			if ds, ok := shape.(*DrawingShape); ok {
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
	}
	return nil
}

func (w *PPTXWriter) getChartIndex(target *ChartShape) int {
	idx := 1
	for _, slide := range w.presentation.slides {
		for _, shape := range slide.shapes {
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

	// Find chart rel ID
	relIdx := 2 // rId1 is slideLayout
	for _, shape := range w.presentation.slides[slideNum-1].shapes {
		if ds, ok := shape.(*DrawingShape); ok && (ds.data != nil || ds.path != "") {
			relIdx++
		}
		if shape == s {
			break
		}
		if _, ok := shape.(*ChartShape); ok {
			relIdx++
		}
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
		switch s := shape.(type) {
		case *RichTextShape:
			childXML.WriteString(w.writeRichTextShapeXML(s, shapeID))
		case *AutoShape:
			childXML.WriteString(w.writeAutoShapeXML(s, shapeID))
		case *LineShape:
			childXML.WriteString(w.writeLineShapeXML(s, shapeID))
		case *DrawingShape:
			childXML.WriteString(w.writeDrawingShapeXML(s, shapeID, slideNum))
		}
	}

	return fmt.Sprintf(`      <p:grpSp>
        <p:nvGrpSpPr>
          <p:cNvPr id="%d" name="%s"/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
            <a:chOff x="%d" y="%d"/>
            <a:chExt cx="%d" cy="%d"/>
          </a:xfrm>
        </p:grpSpPr>
%s      </p:grpSp>
`, id, xmlEscape(name),
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
            <p:ph type="%s" idx="%d"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
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
		s.phType, s.phIdx,
		s.offsetX, s.offsetY, s.width, s.height,
		paragraphsXML.String())
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
          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>%s</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, xmlEscape(slide.notes))

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

// --- Bullet XML ---

func (w *PPTXWriter) writeBulletXML(b *Bullet) string {
	if b.Type == BulletTypeNone {
		return "\n              <a:buNone/>"
	}

	var sb strings.Builder

	// Bullet color
	if b.Color != nil {
		sb.WriteString(fmt.Sprintf("\n              <a:buClr><a:srgbClr val=\"%s\"/></a:buClr>", b.Color.ARGB[2:]))
	}

	// Bullet size
	if b.Size != 100 {
		sb.WriteString(fmt.Sprintf("\n              <a:buSzPct val=\"%d000\"/>", b.Size))
	}

	switch b.Type {
	case BulletTypeChar:
		fontAttr := ""
		if b.Font != "" {
			fontAttr = fmt.Sprintf("\n              <a:buFont typeface=\"%s\"/>", xmlEscape(b.Font))
		}
		sb.WriteString(fontAttr)
		sb.WriteString(fmt.Sprintf("\n              <a:buChar char=\"%s\"/>", xmlEscape(b.Style)))
	case BulletTypeNumeric:
		sb.WriteString(fmt.Sprintf("\n              <a:buAutoNum type=\"%s\" startAt=\"%d\"/>", b.NumFormat, b.StartAt))
	}

	return sb.String()
}
