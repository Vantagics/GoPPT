package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"strconv"
	"strings"
)

func (r *PPTXReader) readSlide(zr *zip.Reader, path string, pres *Presentation) (*Slide, error) {
	data, err := readFileFromZip(zr, path)
	if err != nil {
		return nil, err
	}

	slide := newSlide()
	decoder := xml.NewDecoder(bytes.NewReader(data))

	// Read slide relationships for images, charts, comments, notes
	relsPath := strings.Replace(path, "slides/", "slides/_rels/", 1) + ".rels"
	slideRels, _ := r.readRelationships(zr, relsPath)

	if err := r.parseSlideXML(decoder, slide, slideRels, zr, path); err != nil {
		return nil, err
	}

	// Read comments if relationship exists
	r.readSlideComments(zr, slide, slideRels, path)

	// Read notes if relationship exists
	r.readSlideNotes(zr, slide, slideRels, path)

	return slide, nil
}

func (r *PPTXReader) readSlideComments(zr *zip.Reader, slide *Slide, rels []xmlRelForRead, slidePath string) {
	for _, rel := range rels {
		if rel.Type == relTypeComment {
			target := rel.Target
			if !strings.HasPrefix(target, "ppt/") {
				dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
				target = resolveRelativePath(dir, target)
			}
			data, err := readFileFromZip(zr, target)
			if err != nil {
				continue
			}
			r.parseCommentsXML(data, slide)
		}
	}
}

func (r *PPTXReader) parseCommentsXML(data []byte, slide *Slide) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var currentComment *Comment
	var inText bool

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cm":
				currentComment = NewComment()
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "authorId":
						if v, err := strconv.Atoi(attr.Value); err == nil {
							currentComment.Author = &CommentAuthor{ID: v}
						}
					}
				}
			case "pos":
				if currentComment != nil {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "x":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentComment.PositionX = v
							}
						case "y":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentComment.PositionY = v
							}
						}
					}
				}
			case "text":
				if currentComment != nil {
					inText = true
				}
			}
		case xml.CharData:
			if inText && currentComment != nil {
				currentComment.Text = string(t)
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "cm":
				if currentComment != nil {
					slide.comments = append(slide.comments, currentComment)
					currentComment = nil
				}
			case "text":
				inText = false
			}
		}
	}
}

func (r *PPTXReader) readSlideNotes(zr *zip.Reader, slide *Slide, rels []xmlRelForRead, slidePath string) {
	for _, rel := range rels {
		if rel.Type == relTypeNotesSlide {
			target := rel.Target
			if !strings.HasPrefix(target, "ppt/") {
				dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
				target = resolveRelativePath(dir, target)
			}
			data, err := readFileFromZip(zr, target)
			if err != nil {
				continue
			}
			slide.notes = r.parseNotesXML(data)
		}
	}
}

func (r *PPTXReader) parseNotesXML(data []byte) string {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var inBody bool
	var inParagraph bool
	var inRun bool
	var inText bool
	var texts []string

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "txBody":
				inBody = true
			case "p":
				if inBody {
					inParagraph = true
				}
			case "r":
				if inParagraph {
					inRun = true
				}
			case "t":
				if inRun {
					inText = true
				}
			}
		case xml.CharData:
			if inText {
				texts = append(texts, string(t))
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "txBody":
				inBody = false
			case "p":
				inParagraph = false
			case "r":
				inRun = false
			case "t":
				inText = false
			}
		}
	}
	return strings.Join(texts, "")
}

func (r *PPTXReader) parseSlideXML(decoder *xml.Decoder, slide *Slide, rels []xmlRelForRead, zr *zip.Reader, slidePath string) error {
	type parseState struct {
		inSpTree       bool
		inSp           bool
		inPic          bool
		inCxnSp        bool
		inGraphicFrame bool
		inGrpSp        bool
		inTxBody       bool
		inParagraph    bool
		inRun          bool
		inRunProps     bool
		inText         bool
		inTbl          bool
		inTr           bool
		inTc           bool
		inTcTxBody     bool
		inTcParagraph  bool
		inTcRun        bool
		inTcText       bool
		inNvSpPr       bool
		inSolidFill    bool
		inSpPr         bool
		inLn           bool
		inPPr          bool
		inBg           bool
		inBgPr         bool
		inBgSolidFill  bool
		inBuClr        bool

		// Placeholder tracking
		isPlaceholder bool
		phType        string
		phIdx         int
	}

	state := &parseState{}
	var currentRichText *RichTextShape
	var currentDrawing *DrawingShape
	var currentLine *LineShape
	var currentTable *TableShape
	var currentGroup *GroupShape
	var currentPlaceholder *PlaceholderShape
	var currentParagraph *Paragraph
	var currentFont *Font
	var currentTableRow int
	var currentTableCol int

	var offX, offY, extCX, extCY int64
	var shapeName, shapeDescr string
	var flipH, flipV bool
	var shapeRotation int
	var prstGeom string
	var textAnchor TextAnchorType

	// Deferred shape-level fill (spPr solidFill comes before txBody)
	var pendingShapeFill *Fill

	// Group shape nesting depth
	grpDepth := 0

	// Saved group-level properties (child shapes overwrite the shared variables)
	type grpSaved struct {
		name     string
		descr    string
		offX     int64
		offY     int64
		extCX    int64
		extCY    int64
		flipH    bool
		flipV    bool
		rotation int
	}
	var savedGrp *grpSaved

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "bg":
				state.inBg = true
			case "bgPr":
				if state.inBg {
					state.inBgPr = true
				}
			case "spTree":
				state.inSpTree = true
			case "grpSp":
				if state.inSpTree {
					if state.inGrpSp {
						grpDepth++
					} else {
						state.inGrpSp = true
						grpDepth = 1
						currentGroup = NewGroupShape()
						offX, offY, extCX, extCY = 0, 0, 0, 0
						shapeName = ""
						prstGeom = ""
						shapeRotation = 0
					}
				}
			case "sp":
				if state.inSpTree || state.inGrpSp {
					state.inSp = true
					currentRichText = nil
					currentPlaceholder = nil
					state.isPlaceholder = false
					state.phType = ""
					state.phIdx = 0
					offX, offY, extCX, extCY = 0, 0, 0, 0
					shapeName = ""
					shapeDescr = ""
					prstGeom = ""
					shapeRotation = 0
					textAnchor = TextAnchorNone
					pendingShapeFill = nil
				}
			case "pic":
				if state.inSpTree || state.inGrpSp {
					state.inPic = true
					currentDrawing = NewDrawingShape()
					offX, offY, extCX, extCY = 0, 0, 0, 0
					shapeName = ""
					shapeDescr = ""
					prstGeom = ""
					shapeRotation = 0
				}
			case "cxnSp":
				if state.inSpTree || state.inGrpSp {
					state.inCxnSp = true
					currentLine = NewLineShape()
					offX, offY, extCX, extCY = 0, 0, 0, 0
					shapeName = ""
					prstGeom = ""
					shapeRotation = 0
				}
			case "graphicFrame":
				if state.inSpTree {
					state.inGraphicFrame = true
					offX, offY, extCX, extCY = 0, 0, 0, 0
					shapeName = ""
					prstGeom = ""
					shapeRotation = 0
				}
			case "tbl":
				if state.inGraphicFrame {
					state.inTbl = true
					currentTable = NewTableShape(0, 0)
					currentTable.rows = nil
					currentTableRow = -1
				}
			case "gridCol":
				if state.inTbl && currentTable != nil {
					currentTable.numCols++
				}
			case "tr":
				if state.inTbl && currentTable != nil {
					state.inTr = true
					currentTable.numRows++
					currentTable.rows = append(currentTable.rows, make([]*TableCell, 0))
					currentTableRow = len(currentTable.rows) - 1
					currentTableCol = -1
				}
			case "tc":
				if state.inTr && currentTable != nil {
					state.inTc = true
					currentTableCol++
					cell := NewTableCell()
					cell.paragraphs = nil
					if currentTableRow >= 0 && currentTableRow < len(currentTable.rows) {
						currentTable.rows[currentTableRow] = append(currentTable.rows[currentTableRow], cell)
					}
				}
			case "nvSpPr", "nvPicPr", "nvCxnSpPr", "nvGraphicFramePr", "nvGrpSpPr":
				state.inNvSpPr = true
			case "cNvPr":
				if state.inNvSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shapeName = attr.Value
						case "descr":
							shapeDescr = attr.Value
						}
					}
				}
			case "ph":
				if state.inNvSpPr && state.inSp {
					state.isPlaceholder = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							state.phType = attr.Value
						case "idx":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								state.phIdx = v
							}
						}
					}
				}
			case "txBody":
				if state.inSp {
					state.inTxBody = true
					if state.isPlaceholder {
						if currentPlaceholder == nil {
							currentPlaceholder = NewPlaceholderShape(PlaceholderType(state.phType))
							currentPlaceholder.phIdx = state.phIdx
							currentPlaceholder.paragraphs = nil
						}
					} else {
						if currentRichText == nil {
							currentRichText = NewRichTextShape()
							currentRichText.paragraphs = nil
						}
					}
				} else if state.inTc {
					state.inTcTxBody = true
				}
			case "bodyPr":
				if state.inTxBody && !state.isPlaceholder && currentRichText != nil {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "anchor":
							textAnchor = TextAnchorType(attr.Value)
						case "wrap":
							currentRichText.wordWrap = attr.Value == "square"
						case "numCol":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentRichText.columns = v
							}
						}
					}
				}
			case "p":
				if state.inTcTxBody {
					state.inTcParagraph = true
					currentParagraph = NewParagraph()
					if currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) &&
						currentTableCol < len(currentTable.rows[currentTableRow]) {
						cell := currentTable.rows[currentTableRow][currentTableCol]
						cell.paragraphs = append(cell.paragraphs, currentParagraph)
					}
				} else if state.inTxBody {
					state.inParagraph = true
					currentParagraph = NewParagraph()
					if state.isPlaceholder && currentPlaceholder != nil {
						currentPlaceholder.paragraphs = append(currentPlaceholder.paragraphs, currentParagraph)
					} else if currentRichText != nil {
						currentRichText.paragraphs = append(currentRichText.paragraphs, currentParagraph)
					}
				}
			case "pPr":
				if (state.inParagraph || state.inTcParagraph) && currentParagraph != nil {
					state.inPPr = true
					for _, attr := range t.Attr {
						if attr.Name.Local == "algn" {
							currentParagraph.alignment.Horizontal = HorizontalAlignment(attr.Value)
						}
					}
				}
			case "buNone":
				if state.inPPr && currentParagraph != nil {
					b := NewBullet()
					b.Type = BulletTypeNone
					currentParagraph.bullet = b
				}
			case "buChar":
				if state.inPPr && currentParagraph != nil {
					if currentParagraph.bullet == nil {
						currentParagraph.bullet = NewBullet()
					}
					currentParagraph.bullet.Type = BulletTypeChar
					for _, attr := range t.Attr {
						if attr.Name.Local == "char" {
							currentParagraph.bullet.Style = attr.Value
						}
					}
				}
			case "buAutoNum":
				if state.inPPr && currentParagraph != nil {
					if currentParagraph.bullet == nil {
						currentParagraph.bullet = NewBullet()
					}
					currentParagraph.bullet.Type = BulletTypeNumeric
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							currentParagraph.bullet.NumFormat = attr.Value
						case "startAt":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentParagraph.bullet.StartAt = v
							}
						}
					}
				}
			case "buFont":
				if state.inPPr && currentParagraph != nil {
					if currentParagraph.bullet == nil {
						currentParagraph.bullet = NewBullet()
					}
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" {
							currentParagraph.bullet.Font = attr.Value
						}
					}
				}
			case "buSzPct":
				if state.inPPr && currentParagraph != nil {
					if currentParagraph.bullet == nil {
						currentParagraph.bullet = NewBullet()
					}
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentParagraph.bullet.Size = v / 1000
							}
						}
					}
				}
			case "buClr":
				// Ensure bullet exists for color
				if state.inPPr && currentParagraph != nil {
					if currentParagraph.bullet == nil {
						currentParagraph.bullet = NewBullet()
					}
					state.inBuClr = true
				}
			case "spcBef":
				// Space before - handled by child spcPts
			case "spcAft":
				// Space after - handled by child spcPts
			case "r":
				if state.inTcParagraph {
					state.inTcRun = true
					currentFont = NewFont()
				} else if state.inParagraph {
					state.inRun = true
					currentFont = NewFont()
				}
			case "rPr":
				if state.inRun || state.inTcRun {
					state.inRunProps = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "sz":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentFont.Size = v / 100
							}
						case "b":
							currentFont.Bold = attr.Value == "1"
						case "i":
							currentFont.Italic = attr.Value == "1"
						case "u":
							currentFont.Underline = UnderlineType(attr.Value)
						case "strike":
							currentFont.Strikethrough = attr.Value == "sngStrike"
						}
					}
				}
			case "solidFill":
				if state.inRunProps {
					state.inSolidFill = true
				} else if state.inBgPr {
					state.inBgSolidFill = true
				} else if state.inSpPr && !state.inTxBody && !state.inLn {
					// Shape-level solid fill (not inside text body or line)
					state.inSolidFill = true
				} else if state.inLn {
					// Line solid fill
					state.inSolidFill = true
				}
			case "srgbClr":
				if state.inSolidFill && state.inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							currentFont.Color = NewColor("FF" + attr.Value)
						}
					}
				} else if state.inSolidFill && state.inLn && !state.inRunProps {
					// Line solid fill color (inside <a:ln>)
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							if state.inCxnSp && currentLine != nil {
								currentLine.lineColor = c
							}
						}
					}
				} else if state.inSolidFill && state.inSpPr && !state.inRunProps && !state.inTxBody && !state.inLn {
					// Shape-level solid fill color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							if state.inSp {
								if currentRichText != nil {
									currentRichText.GetFill().SetSolid(c)
								} else {
									// spPr comes before txBody, so defer the fill
									pendingShapeFill = NewFill()
									pendingShapeFill.SetSolid(c)
								}
							}
						}
					}
				} else if state.inBgSolidFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if slide.background == nil {
								slide.background = NewFill()
							}
							slide.background.SetSolid(NewColor("FF" + attr.Value))
						}
					}
				} else if state.inBuClr && currentParagraph != nil && currentParagraph.bullet != nil {
					// Bullet color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							currentParagraph.bullet.Color = &c
						}
					}
				}
			case "latin":
				if state.inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" {
							currentFont.Name = attr.Value
						}
					}
				}
			case "t":
				if state.inTcRun {
					state.inTcText = true
				} else if state.inRun {
					state.inText = true
				}
			case "br":
				if state.inParagraph && currentParagraph != nil {
					currentParagraph.CreateBreak()
				}
			case "xfrm":
				flipH = false
				flipV = false
				shapeRotation = 0
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "flipH":
						flipH = attr.Value == "1" || attr.Value == "true"
					case "flipV":
						flipV = attr.Value == "1" || attr.Value == "true"
					case "rot":
						// rotation in 60000ths of a degree
						if v, err := strconv.Atoi(attr.Value); err == nil {
							shapeRotation = v / 60000
						}
					}
				}
			case "off":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "x":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							offX = v
						}
					case "y":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							offY = v
						}
					}
				}
			case "ext":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							extCX = v
						}
					case "cy":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							extCY = v
						}
					}
				}
			case "blip":
				if state.inPic {
					for _, attr := range t.Attr {
						if attr.Name.Local == "embed" {
							for _, rel := range rels {
								if rel.ID == attr.Value {
									imgPath := rel.Target
									if !strings.HasPrefix(imgPath, "ppt/") {
										dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
										imgPath = resolveRelativePath(dir, imgPath)
									}
									imgData, err := readFileFromZip(zr, imgPath)
									if err == nil {
										currentDrawing.data = imgData
										currentDrawing.mimeType = guessMimeType(imgPath)
									}
									break
								}
							}
						}
					}
				}
			case "ln":
				if state.inSpPr {
					state.inLn = true
				}
				if state.inCxnSp && currentLine != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentLine.lineWidth = v / 12700
							}
						}
					}
				}
			case "spPr", "grpSpPr":
				if state.inSp || state.inPic || state.inCxnSp || state.inGrpSp {
					state.inSpPr = true
				}
			case "prstGeom":
				for _, attr := range t.Attr {
					if attr.Name.Local == "prst" {
						prstGeom = attr.Value
					}
				}
			}

		case xml.CharData:
			text := string(t)
			if state.inTcText && currentParagraph != nil {
				tr := currentParagraph.CreateTextRun(text)
				if currentFont != nil {
					tr.font = currentFont
				}
			} else if state.inText && currentParagraph != nil {
				tr := currentParagraph.CreateTextRun(text)
				if currentFont != nil {
					tr.font = currentFont
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "bg":
				state.inBg = false
			case "bgPr":
				state.inBgPr = false
				state.inBgSolidFill = false
			case "spTree":
				state.inSpTree = false
			case "grpSp":
				if state.inGrpSp {
					grpDepth--
					if grpDepth <= 0 {
						state.inGrpSp = false
						if currentGroup != nil && savedGrp != nil {
							currentGroup.name = savedGrp.name
							currentGroup.offsetX = savedGrp.offX
							currentGroup.offsetY = savedGrp.offY
							currentGroup.width = savedGrp.extCX
							currentGroup.height = savedGrp.extCY
							currentGroup.flipHorizontal = savedGrp.flipH
							currentGroup.flipVertical = savedGrp.flipV
							currentGroup.rotation = savedGrp.rotation
							slide.shapes = append(slide.shapes, currentGroup)
						}
						currentGroup = nil
						savedGrp = nil
					}
				}
			case "sp":
				if state.inSp {
					state.inSp = false
					if state.isPlaceholder && currentPlaceholder != nil {
						currentPlaceholder.name = shapeName
						currentPlaceholder.description = shapeDescr
						currentPlaceholder.offsetX = offX
						currentPlaceholder.offsetY = offY
						currentPlaceholder.width = extCX
						currentPlaceholder.height = extCY
						currentPlaceholder.flipHorizontal = flipH
						currentPlaceholder.flipVertical = flipV
						currentPlaceholder.rotation = shapeRotation
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentPlaceholder)
						} else {
							slide.shapes = append(slide.shapes, currentPlaceholder)
						}
						currentPlaceholder = nil
					} else if prstGeom != "" && prstGeom != "rect" {
						// Non-rect geometry → AutoShape
						autoShape := NewAutoShape()
						autoShape.name = shapeName
						autoShape.description = shapeDescr
						autoShape.offsetX = offX
						autoShape.offsetY = offY
						autoShape.width = extCX
						autoShape.height = extCY
						autoShape.flipHorizontal = flipH
						autoShape.flipVertical = flipV
						autoShape.rotation = shapeRotation
						autoShape.shapeType = AutoShapeType(prstGeom)
						// Copy text from richtext if any
						if currentRichText != nil && len(currentRichText.paragraphs) > 0 {
							var texts []string
							for _, para := range currentRichText.paragraphs {
								for _, elem := range para.elements {
									if tr, ok := elem.(*TextRun); ok {
										texts = append(texts, tr.text)
									}
								}
							}
							autoShape.text = joinNonEmpty(texts, "")
						}
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(autoShape)
						} else {
							slide.shapes = append(slide.shapes, autoShape)
						}
					} else if currentRichText != nil {
						currentRichText.name = shapeName
						currentRichText.description = shapeDescr
						currentRichText.offsetX = offX
						currentRichText.offsetY = offY
						currentRichText.width = extCX
						currentRichText.height = extCY
						currentRichText.flipHorizontal = flipH
						currentRichText.flipVertical = flipV
						currentRichText.rotation = shapeRotation
						currentRichText.textAnchor = textAnchor
						// Apply deferred shape-level fill (spPr comes before txBody)
						if pendingShapeFill != nil {
							currentRichText.fill = pendingShapeFill
							pendingShapeFill = nil
						}
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentRichText)
						} else {
							slide.shapes = append(slide.shapes, currentRichText)
						}
					}
					currentRichText = nil
					state.isPlaceholder = false
				}
			case "pic":
				if state.inPic {
					state.inPic = false
					if currentDrawing != nil {
						currentDrawing.name = shapeName
						currentDrawing.description = shapeDescr
						currentDrawing.offsetX = offX
						currentDrawing.offsetY = offY
						currentDrawing.width = extCX
						currentDrawing.height = extCY
						currentDrawing.flipHorizontal = flipH
						currentDrawing.flipVertical = flipV
						currentDrawing.rotation = shapeRotation
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentDrawing)
						} else {
							slide.shapes = append(slide.shapes, currentDrawing)
						}
					}
					currentDrawing = nil
				}
			case "cxnSp":
				if state.inCxnSp {
					state.inCxnSp = false
					if currentLine != nil {
						currentLine.name = shapeName
						currentLine.offsetX = offX
						currentLine.offsetY = offY
						currentLine.width = extCX
						currentLine.height = extCY
						currentLine.flipHorizontal = flipH
						currentLine.flipVertical = flipV
						currentLine.rotation = shapeRotation
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentLine)
						} else {
							slide.shapes = append(slide.shapes, currentLine)
						}
					}
					currentLine = nil
				}
			case "graphicFrame":
				if state.inGraphicFrame {
					state.inGraphicFrame = false
					if currentTable != nil {
						currentTable.name = shapeName
						currentTable.offsetX = offX
						currentTable.offsetY = offY
						currentTable.width = extCX
						currentTable.height = extCY
						slide.shapes = append(slide.shapes, currentTable)
					}
					currentTable = nil
				}
			case "tbl":
				state.inTbl = false
			case "tr":
				state.inTr = false
			case "tc":
				state.inTc = false
				state.inTcTxBody = false
			case "txBody":
				if state.inTc {
					state.inTcTxBody = false
				} else {
					state.inTxBody = false
				}
			case "p":
				if state.inTcParagraph {
					state.inTcParagraph = false
				} else {
					state.inParagraph = false
				}
				currentParagraph = nil
			case "pPr":
				state.inPPr = false
			case "r":
				if state.inTcRun {
					state.inTcRun = false
				} else {
					state.inRun = false
				}
				currentFont = nil
			case "rPr":
				state.inRunProps = false
				state.inSolidFill = false
			case "solidFill":
				state.inSolidFill = false
				state.inBgSolidFill = false
			case "spPr", "grpSpPr":
				state.inSpPr = false
				state.inLn = false
				// When the group's shape properties end, save position/size
				// before child shapes overwrite the shared variables.
				if t.Name.Local == "grpSpPr" && state.inGrpSp && savedGrp != nil {
					savedGrp.offX = offX
					savedGrp.offY = offY
					savedGrp.extCX = extCX
					savedGrp.extCY = extCY
					savedGrp.flipH = flipH
					savedGrp.flipV = flipV
					savedGrp.rotation = shapeRotation
				}
			case "ln":
				state.inLn = false
			case "buClr":
				state.inBuClr = false
			case "t":
				state.inText = false
				state.inTcText = false
			case "nvSpPr", "nvPicPr", "nvCxnSpPr", "nvGraphicFramePr", "nvGrpSpPr":
				state.inNvSpPr = false
				// When the group's non-visual properties end, save the group name
				// before child shapes overwrite the shared shapeName variable.
				if t.Name.Local == "nvGrpSpPr" && state.inGrpSp && currentGroup != nil && savedGrp == nil {
					savedGrp = &grpSaved{
						name:  shapeName,
						descr: shapeDescr,
					}
				}
			}
		}
	}

	return nil
}

func lastPathComponent(path string) string {
	parts := strings.Split(path, "/")
	return parts[len(parts)-1]
}

func resolveRelativePath(base, rel string) string {
	if strings.HasPrefix(rel, "/") {
		return strings.TrimPrefix(rel, "/")
	}

	baseParts := strings.Split(base, "/")
	relParts := strings.Split(rel, "/")

	result := make([]string, 0, len(baseParts)+len(relParts))
	result = append(result, baseParts...)

	for _, part := range relParts {
		if part == ".." {
			if len(result) > 0 {
				result = result[:len(result)-1]
			}
		} else if part != "." && part != "" {
			result = append(result, part)
		}
	}

	resolved := strings.Join(result, "/")

	// Security: ensure resolved path stays within the ppt/ directory to prevent
	// path traversal attacks via malicious relationship targets.
	if !strings.HasPrefix(resolved, "ppt/") && !strings.HasPrefix(resolved, "docProps/") && resolved != "[Content_Types].xml" && !strings.HasPrefix(resolved, "_rels/") {
		return "ppt/" + resolved
	}

	return resolved
}

func guessMimeType(path string) string {
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
