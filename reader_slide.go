package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
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

	if err := r.parseSlideXML(decoder, slide, slideRels, zr, path, pres); err != nil {
		return nil, err
	}

	// Apply slide layout inheritance for placeholders with missing position/size
	r.applyLayoutInheritance(zr, slide, slideRels, path, pres)

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

func (r *PPTXReader) parseSlideXML(decoder *xml.Decoder, slide *Slide, rels []xmlRelForRead, zr *zip.Reader, slidePath string, pres *Presentation) error {
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
		inTcRun         bool
		inTcText        bool
		inTcPr          bool
		inTcPrSolidFill bool
		inTcPrLn        bool
		tcPrLnSide      string // "L", "R", "T", "B" or "" for generic
		inNvSpPr       bool
		inSolidFill    bool
		inSpPr         bool
		inLn           bool
		inPPr          bool
		inBg           bool
		inBgPr         bool
		inBgSolidFill  bool
		inBuClr        bool

		// Spacing context tracking
		inSpcBef bool
		inSpcAft bool
		inLnSpc  bool

		// Color modifier tracking (for <a:alpha> inside <a:srgbClr> or <a:schemeClr>)
		inSrgbClr bool

		// defRPr tracking (default run properties inside pPr or lstStyle)
		inDefRPr       bool
		inLstStyle     bool
		inLstStyleLvl1 bool // inside lstStyle/lvl1pPr specifically

		// Placeholder tracking
		isPlaceholder bool
		phType        string
		phIdx         int

		// p:style / fontRef tracking
		inStyle   bool
		inFontRef bool

		// extLst tracking (to ignore hiddenFill etc.)
		inExtLst bool

		// blipFill inside spPr (shape image fill)
		inSpPrBlipFill bool

		// blipFill inside bgPr (slide background image)
		inBgBlipFill bool

		// gradFill tracking
		inGradFill    bool
		inGsLst       bool
		inGs          bool
		gradFillPos   int // current gs position (0-100000)
		inRunPropsGradFill bool // gradFill inside rPr (text color gradient)

		// avLst tracking (adjustment values for preset geometry)
		inAvLst bool

		// custGeom tracking
		inCustGeom  bool
		inPathLst   bool
		inCustPath  bool

		// effectLst / outerShdw tracking
		inEffectLst  bool
		inOuterShdw  bool
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

	// Pending custom geometry path
	var pendingCustomPath *CustomGeomPath
	var pendingPathCmds []PathCommand

	// Default font properties from defRPr (paragraph-level defaults)
	var defFont *Font

	// lstStyle-level default font (from <a:lstStyle>/<a:lvl1pPr>/<a:defRPr>)
	var lstStyleFont *Font

	// lastColor tracks the most recently parsed srgbClr/schemeClr so that child
	// elements like <a:alpha> can modify it.
	var lastColor *Color

	var offX, offY, extCX, extCY int64
	var chOffX, chOffY, chExtCX, chExtCY int64
	var shapeName, shapeDescr string
	var flipH, flipV bool
	var shapeRotation int
	var prstGeom string
	var textAnchor TextAnchorType
	var textDir string

	// Font color from <p:style>/<a:fontRef>/<a:schemeClr> (default text color for shape)
	var fontRefColor *Color

	// Deferred shape-level fill (spPr solidFill comes before txBody)
	var pendingShapeFill *Fill

	// Gradient fill stop colors
	var gradStopColors []Color
	var gradStopPositions []int
	var gradAngle int
	_ = gradStopColors
	_ = gradStopPositions
	_ = gradAngle

	// Deferred shape-level border (spPr ln comes before txBody)
	var pendingBorder *Border
	var pendingHeadEnd *LineEnd
	var pendingTailEnd *LineEnd

	// Deferred adjustment values from avLst
	var pendingAdjustValues map[string]int

	// Deferred shadow (spPr effectLst outerShdw)
	var pendingShadow *Shadow

	// Deferred blipFill image data (spPr blipFill for shapes)
	var pendingBlipFillData []byte
	var pendingBlipFillMime string

	// Background blipFill image data (bgPr blipFill)
	// TODO: use these to set slide.background as an image fill
	var bgBlipFillData []byte
	var bgBlipFillMime string
	_, _ = bgBlipFillData, bgBlipFillMime

	// Group shape nesting
	grpDepth := 0

	// Stack for nested groups
	type grpSaved struct {
		group    *GroupShape
		name     string
		descr    string
		offX     int64
		offY     int64
		extCX    int64
		extCY    int64
		chOffX   int64
		chOffY   int64
		chExtCX  int64
		chExtCY  int64
		flipH    bool
		flipV    bool
		rotation int
		grpFill  *Fill // solidFill from grpSpPr, inherited by child <a:grpFill/>
	}
	var grpStack []*grpSaved

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
					state.inGrpSp = true
					grpDepth++
					newGroup := NewGroupShape()
					grpStack = append(grpStack, &grpSaved{group: newGroup})
					currentGroup = newGroup
					offX, offY, extCX, extCY = 0, 0, 0, 0
					chOffX, chOffY, chExtCX, chExtCY = 0, 0, 0, 0
					shapeName = ""
					shapeDescr = ""
					prstGeom = ""
					shapeRotation = 0
					flipH, flipV = false, false
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
					textDir = ""
					pendingShapeFill = nil
					pendingBorder = nil
					pendingHeadEnd = nil
					pendingTailEnd = nil
					pendingAdjustValues = nil
					pendingShadow = nil
					pendingBlipFillData = nil
					pendingBlipFillMime = ""
					pendingCustomPath = nil
					fontRefColor = nil
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
					pendingCustomPath = nil
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
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								currentTable.colWidths = append(currentTable.colWidths, v)
							}
						}
					}
				}
			case "tr":
				if state.inTbl && currentTable != nil {
					state.inTr = true
					currentTable.numRows++
					currentTable.rows = append(currentTable.rows, make([]*TableCell, 0))
					currentTableRow = len(currentTable.rows) - 1
					currentTableCol = -1
					for _, attr := range t.Attr {
						if attr.Name.Local == "h" {
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								currentTable.rowHeights = append(currentTable.rowHeights, v)
							}
						}
					}
				}
			case "tc":
				if state.inTr && currentTable != nil {
					state.inTc = true
					currentTableCol++
					cell := NewTableCell()
					cell.paragraphs = nil
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "gridSpan":
							if v, err := strconv.Atoi(attr.Value); err == nil && v > 1 {
								cell.colSpan = v
							}
						case "rowSpan":
							if v, err := strconv.Atoi(attr.Value); err == nil && v > 1 {
								cell.rowSpan = v
							}
						case "hMerge":
							cell.hMerge = attr.Value == "1" || attr.Value == "true"
						case "vMerge":
							cell.vMerge = attr.Value == "1" || attr.Value == "true"
						}
					}
					if currentTableRow >= 0 && currentTableRow < len(currentTable.rows) {
						currentTable.rows[currentTableRow] = append(currentTable.rows[currentTableRow], cell)
					}
				}
			case "nvSpPr", "nvPicPr", "nvCxnSpPr", "nvGraphicFramePr", "nvGrpSpPr":
				state.inNvSpPr = true
			case "tcPr":
				if state.inTc && currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
					currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
					state.inTcPr = true
				}
			case "lnL":
				if state.inTcPr {
					state.inTcPrLn = true
					state.tcPrLnSide = "L"
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									currentTable.rows[currentTableRow][currentTableCol].border.Left.Width = v / 12700
									currentTable.rows[currentTableRow][currentTableCol].border.Left.Style = BorderSolid
								}
							}
						}
					}
				}
			case "lnR":
				if state.inTcPr {
					state.inTcPrLn = true
					state.tcPrLnSide = "R"
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									currentTable.rows[currentTableRow][currentTableCol].border.Right.Width = v / 12700
									currentTable.rows[currentTableRow][currentTableCol].border.Right.Style = BorderSolid
								}
							}
						}
					}
				}
			case "lnT":
				if state.inTcPr {
					state.inTcPrLn = true
					state.tcPrLnSide = "T"
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									currentTable.rows[currentTableRow][currentTableCol].border.Top.Width = v / 12700
									currentTable.rows[currentTableRow][currentTableCol].border.Top.Style = BorderSolid
								}
							}
						}
					}
				}
			case "lnB":
				if state.inTcPr {
					state.inTcPrLn = true
					state.tcPrLnSide = "B"
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									currentTable.rows[currentTableRow][currentTableCol].border.Bottom.Width = v / 12700
									currentTable.rows[currentTableRow][currentTableCol].border.Bottom.Style = BorderSolid
								}
							}
						}
					}
				}
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
					lstStyleFont = nil // reset for new text body
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
			case "lstStyle":
				if state.inTxBody {
					state.inLstStyle = true
				}
			case "lvl1pPr":
				if state.inLstStyle {
					state.inLstStyleLvl1 = true
				}
			case "bodyPr":
				if state.inTxBody {
					// Check if any inset attributes are present; if so, initialize to defaults first
					hasInsets := false
					for _, attr := range t.Attr {
						if attr.Name.Local == "lIns" || attr.Name.Local == "rIns" || attr.Name.Local == "tIns" || attr.Name.Local == "bIns" {
							hasInsets = true
							break
						}
					}
					if hasInsets {
						// Initialize to PowerPoint defaults before overriding
						if currentRichText != nil {
							currentRichText.insetLeft = 91440
							currentRichText.insetRight = 91440
							currentRichText.insetTop = 45720
							currentRichText.insetBottom = 45720
							currentRichText.insetsSet = true
						}
						if currentPlaceholder != nil {
							currentPlaceholder.insetLeft = 91440
							currentPlaceholder.insetRight = 91440
							currentPlaceholder.insetTop = 45720
							currentPlaceholder.insetBottom = 45720
							currentPlaceholder.insetsSet = true
						}
					}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "anchor":
							textAnchor = TextAnchorType(attr.Value)
						case "vert":
							textDir = attr.Value
							if currentRichText != nil {
								currentRichText.textDirection = attr.Value
							}
							if currentPlaceholder != nil {
								currentPlaceholder.textDirection = attr.Value
							}
						case "lIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								if currentRichText != nil {
									currentRichText.insetLeft = v
									currentRichText.insetsSet = true
								}
								if currentPlaceholder != nil {
									currentPlaceholder.insetLeft = v
									currentPlaceholder.insetsSet = true
								}
							}
						case "rIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								if currentRichText != nil {
									currentRichText.insetRight = v
									currentRichText.insetsSet = true
								}
								if currentPlaceholder != nil {
									currentPlaceholder.insetRight = v
									currentPlaceholder.insetsSet = true
								}
							}
						case "tIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								if currentRichText != nil {
									currentRichText.insetTop = v
									currentRichText.insetsSet = true
								}
								if currentPlaceholder != nil {
									currentPlaceholder.insetTop = v
									currentPlaceholder.insetsSet = true
								}
							}
						case "bIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								if currentRichText != nil {
									currentRichText.insetBottom = v
									currentRichText.insetsSet = true
								}
								if currentPlaceholder != nil {
									currentPlaceholder.insetBottom = v
									currentPlaceholder.insetsSet = true
								}
							}
						}
					}
					if state.isPlaceholder && currentPlaceholder != nil {
						for _, attr := range t.Attr {
							switch attr.Name.Local {
							case "wrap":
								currentPlaceholder.wordWrap = attr.Value == "square"
							case "numCol":
								if v, err := strconv.Atoi(attr.Value); err == nil {
									currentPlaceholder.columns = v
								}
							}
						}
					} else if currentRichText != nil {
						for _, attr := range t.Attr {
							switch attr.Name.Local {
							case "wrap":
								currentRichText.wordWrap = attr.Value == "square"
							case "numCol":
								if v, err := strconv.Atoi(attr.Value); err == nil {
									currentRichText.columns = v
								}
							}
						}
					}
				}
			case "normAutofit":
				// <a:normAutofit fontScale="62500"/> inside <a:bodyPr>
				if state.inTxBody {
					fontScaleVal := 100000 // default 100%
					for _, attr := range t.Attr {
						if attr.Name.Local == "fontScale" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								fontScaleVal = v
							}
						}
					}
					if state.isPlaceholder && currentPlaceholder != nil {
						currentPlaceholder.autoFit = AutoFitNormal
						currentPlaceholder.fontScale = fontScaleVal
					} else if currentRichText != nil {
						currentRichText.autoFit = AutoFitNormal
						currentRichText.fontScale = fontScaleVal
					}
				}
			case "spAutoFit":
				// <a:spAutoFit/> inside <a:bodyPr> — resize shape to fit text
				if state.inTxBody {
					if state.isPlaceholder && currentPlaceholder != nil {
						currentPlaceholder.autoFit = AutoFitShape
					} else if currentRichText != nil {
						currentRichText.autoFit = AutoFitShape
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
						switch attr.Name.Local {
						case "algn":
							currentParagraph.alignment.Horizontal = HorizontalAlignment(attr.Value)
						case "marL":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								currentParagraph.alignment.MarginLeft = v
							}
						case "marR":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								currentParagraph.alignment.MarginRight = v
							}
						case "indent":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								currentParagraph.alignment.Indent = v
							}
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
				// Space before paragraph
				if state.inPPr && currentParagraph != nil {
					state.inSpcBef = true
				}
			case "spcAft":
				// Space after paragraph
				if state.inPPr && currentParagraph != nil {
					state.inSpcAft = true
				}
			case "lnSpc":
				// Line spacing
				if state.inPPr && currentParagraph != nil {
					state.inLnSpc = true
				}
			case "spcPts":
				// Spacing in hundredths of a point (e.g. 1200 = 12pt)
				if currentParagraph != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if state.inSpcBef {
									currentParagraph.spaceBefore = v
								} else if state.inSpcAft {
									currentParagraph.spaceAfter = v
								} else if state.inLnSpc {
									currentParagraph.lineSpacing = v
								}
							}
						}
					}
				}
			case "spcPct":
				// Spacing as percentage (e.g. 150000 = 150%)
				if currentParagraph != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if state.inLnSpc {
									// Store as negative to distinguish from spcPts
									currentParagraph.lineSpacing = -v
								}
								// spcPct inside spcBef/spcAft with val=0 means no spacing;
								// non-zero percentage spacing before/after is rare and
								// would need line-height-relative calculation at render time.
							}
						}
					}
				}
			case "r":
				if state.inTcParagraph {
					state.inTcRun = true
					currentFont = NewFont()
					// PowerPoint default font size for table cell text is 18pt
					currentFont.Size = 18
				} else if state.inParagraph {
					state.inRun = true
					currentFont = NewFont()
					// PowerPoint default font size for text runs is 18pt (1800 hundredths)
					// when no size is specified in rPr, defRPr, or lstStyle.
					currentFont.Size = 18
					// Apply fontRef color from <p:style> as base default
					if fontRefColor != nil {
						currentFont.Color = *fontRefColor
					}
					// Apply lstStyle-level default font properties first
					if lstStyleFont != nil {
						if lstStyleFont.Size > 0 {
							currentFont.Size = lstStyleFont.Size
						}
						if lstStyleFont.Bold {
							currentFont.Bold = true
						}
						if lstStyleFont.Italic {
							currentFont.Italic = true
						}
						if lstStyleFont.Name != "Calibri" && lstStyleFont.Name != "" {
							currentFont.Name = lstStyleFont.Name
						}
						if lstStyleFont.NameEA != "" {
							currentFont.NameEA = lstStyleFont.NameEA
						}
						if lstStyleFont.Color.ARGB != "FF000000" && lstStyleFont.Color.ARGB != "" {
							currentFont.Color = lstStyleFont.Color
						}
					}
					// Apply paragraph-level default font properties (overrides lstStyle)
					if defFont != nil {
						if defFont.Size > 0 {
							currentFont.Size = defFont.Size
						}
						if defFont.Bold {
							currentFont.Bold = true
						}
						if defFont.Italic {
							currentFont.Italic = true
						}
						if defFont.Name != "Calibri" && defFont.Name != "" {
							currentFont.Name = defFont.Name
						}
						if defFont.NameEA != "" {
							currentFont.NameEA = defFont.NameEA
						}
						if defFont.Color.ARGB != "FF000000" && defFont.Color.ARGB != "" {
							currentFont.Color = defFont.Color
						}
					}
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
			case "defRPr":
				if state.inPPr || state.inLstStyleLvl1 {
					state.inDefRPr = true
					if state.inLstStyleLvl1 && !state.inPPr {
						// lstStyle/lvl1pPr-level defRPr
						lstStyleFont = NewFont()
						lstStyleFont.Size = 0
						for _, attr := range t.Attr {
							switch attr.Name.Local {
							case "sz":
								if v, err := strconv.Atoi(attr.Value); err == nil {
									lstStyleFont.Size = v / 100
								}
							case "b":
								lstStyleFont.Bold = attr.Value == "1"
							case "i":
								lstStyleFont.Italic = attr.Value == "1"
							}
						}
					} else {
						// pPr-level defRPr
						defFont = NewFont()
						defFont.Size = 0
						for _, attr := range t.Attr {
							switch attr.Name.Local {
							case "sz":
								if v, err := strconv.Atoi(attr.Value); err == nil {
									defFont.Size = v / 100
								}
							case "b":
								defFont.Bold = attr.Value == "1"
							case "i":
								defFont.Italic = attr.Value == "1"
							}
						}
					}
				}
			case "solidFill":
				if state.inExtLst {
					// Ignore solidFill inside extLst (e.g. hiddenFill)
				} else if state.inTcPr && !state.inTcPrLn {
					// Table cell solid fill
					state.inTcPrSolidFill = true
				} else if state.inTcPr && state.inTcPrLn {
					// Table cell border line solid fill
					state.inTcPrSolidFill = true
				} else if state.inDefRPr {
					state.inSolidFill = true
				} else if state.inRunProps {
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
			case "noFill":
				// <a:noFill/> inside spPr means the shape has no fill
				if state.inSpPr && !state.inTxBody && !state.inLn && !state.inExtLst {
					if state.inSp {
						pendingShapeFill = NewFill()
						pendingShapeFill.Type = FillNone
					}
				}
				// <a:noFill/> inside tcPr means the cell has no fill
				if state.inTcPr && !state.inTcPrLn {
					if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
						cell := currentTable.rows[currentTableRow][currentTableCol]
						cell.fill = NewFill()
						cell.fill.Type = FillNone
					}
				}
				// <a:noFill/> inside lnL/lnR/lnT/lnB means no border on that side
				if state.inTcPr && state.inTcPrLn {
					if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
						cell := currentTable.rows[currentTableRow][currentTableCol]
						if cell.border != nil {
							switch state.tcPrLnSide {
							case "L":
								cell.border.Left.Style = BorderNone
							case "R":
								cell.border.Right.Style = BorderNone
							case "T":
								cell.border.Top.Style = BorderNone
							case "B":
								cell.border.Bottom.Style = BorderNone
							}
						}
					}
				}
			case "grpFill":
				// <a:grpFill/> — inherit fill from parent group's grpSpPr
				if state.inSpPr && !state.inTxBody && !state.inLn && state.inSp && state.inGrpSp && len(grpStack) > 0 {
					gf := grpStack[len(grpStack)-1].grpFill
					if gf != nil {
						inherited := NewFill()
						*inherited = *gf
						pendingShapeFill = inherited
					}
				}
			case "gradFill":
				if state.inRunProps && currentFont != nil {
					// gradFill inside rPr — use first stop color as text color
					state.inRunPropsGradFill = true
					state.inGradFill = true
					gradStopColors = nil
					gradStopPositions = nil
					gradAngle = 0
				} else if state.inSpPr && !state.inTxBody && !state.inLn && !state.inExtLst {
					state.inGradFill = true
					gradStopColors = nil
					gradStopPositions = nil
					gradAngle = 0
				} else if state.inBgPr {
					state.inGradFill = true
					gradStopColors = nil
					gradStopPositions = nil
					gradAngle = 0
				}
			case "gsLst":
				if state.inGradFill {
					state.inGsLst = true
				}
			case "gs":
				if state.inGsLst {
					state.inGs = true
					state.gradFillPos = 0
					for _, attr := range t.Attr {
						if attr.Name.Local == "pos" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								state.gradFillPos = v
							}
						}
					}
				}
			case "lin":
				if state.inGradFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "ang" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								gradAngle = v / 60000
							}
						}
					}
				}
			case "blipFill":
				// <a:blipFill> inside spPr — shape has an image fill
				if state.inSpPr && state.inSp && !state.inTxBody && !state.inLn {
					state.inSpPrBlipFill = true
				} else if state.inBgPr {
					// <a:blipFill> inside bgPr — slide background image
					state.inBgBlipFill = true
				}
			case "extLst":
				if state.inSpPr {
					state.inExtLst = true
				}
			case "srgbClr":
				state.inSrgbClr = true
				lastColor = nil
				if state.inGs {
					// Gradient stop color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							gradStopColors = append(gradStopColors, c)
							gradStopPositions = append(gradStopPositions, state.gradFillPos)
							lastColor = &gradStopColors[len(gradStopColors)-1]
						}
					}
				} else if state.inOuterShdw && pendingShadow != nil {
					// Shadow color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							pendingShadow.Color = NewColor("FF" + attr.Value)
							lastColor = &pendingShadow.Color
						}
					}
				} else if state.inTcPrSolidFill {
					// Table cell fill or border color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							if state.inTcPrLn {
								// tcPr border line color — apply to the specific side
								lastColor = &c
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									cell := currentTable.rows[currentTableRow][currentTableCol]
									if cell.border != nil {
										switch state.tcPrLnSide {
										case "L":
											cell.border.Left.Color = c
											cell.border.Left.Style = BorderSolid
										case "R":
											cell.border.Right.Color = c
											cell.border.Right.Style = BorderSolid
										case "T":
											cell.border.Top.Color = c
											cell.border.Top.Style = BorderSolid
										case "B":
											cell.border.Bottom.Color = c
											cell.border.Bottom.Style = BorderSolid
										}
									}
								}
							} else {
								// tcPr cell fill color
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									cell := currentTable.rows[currentTableRow][currentTableCol]
									cell.fill = NewFill()
									cell.fill.SetSolid(c)
									lastColor = &cell.fill.Color
								}
							}
						}
					}
				} else if state.inFontRef {
					// <p:style>/<a:fontRef>/<a:srgbClr> — default text color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							fontRefColor = &c
							lastColor = fontRefColor
						}
					}
				} else if state.inSolidFill && state.inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							currentFont.Color = NewColor("FF" + attr.Value)
							lastColor = &currentFont.Color
						}
					}
				} else if state.inSolidFill && state.inLn && !state.inRunProps {
					// Line solid fill color (inside <a:ln>)
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							if state.inCxnSp && currentLine != nil {
								currentLine.lineColor = c
								lastColor = &currentLine.lineColor
							} else if state.inSp {
								if pendingBorder == nil {
									pendingBorder = &Border{Style: BorderSolid}
								}
								pendingBorder.Color = c
								lastColor = &pendingBorder.Color
							}
						}
					}
				} else if state.inSolidFill && state.inSpPr && !state.inRunProps && !state.inTxBody && !state.inLn {
					// Shape-level solid fill color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							if state.inGrpSp && !state.inSp && len(grpStack) > 0 {
								// solidFill inside grpSpPr — store as group fill
								f := NewFill()
								f.SetSolid(c)
								grpStack[len(grpStack)-1].grpFill = f
								lastColor = &grpStack[len(grpStack)-1].grpFill.Color
							} else if state.inSp {
								if currentRichText != nil {
									currentRichText.GetFill().SetSolid(c)
									lastColor = &currentRichText.GetFill().Color
								} else {
									// spPr comes before txBody, so defer the fill
									pendingShapeFill = NewFill()
									pendingShapeFill.SetSolid(c)
									lastColor = &pendingShapeFill.Color
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
							lastColor = &slide.background.Color
						}
					}
				} else if state.inBuClr && currentParagraph != nil && currentParagraph.bullet != nil {
					// Bullet color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							currentParagraph.bullet.Color = &c
							lastColor = currentParagraph.bullet.Color
						}
					}
				} else if state.inDefRPr && state.inSolidFill && state.inLstStyleLvl1 && lstStyleFont != nil {
					// lstStyle defRPr solidFill srgbClr
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							lstStyleFont.Color = NewColor("FF" + attr.Value)
							lastColor = &lstStyleFont.Color
						}
					}
				} else if state.inDefRPr && state.inSolidFill && defFont != nil {
					// defRPr solidFill srgbClr
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							defFont.Color = NewColor("FF" + attr.Value)
							lastColor = &defFont.Color
						}
					}
				}
			case "prstClr":
				state.inSrgbClr = true // reuse for alpha child handling
				lastColor = nil
				var prstName string
				for _, attr := range t.Attr {
					if attr.Name.Local == "val" {
						prstName = attr.Value
					}
				}
				c := presetColorToColor(prstName)
				if state.inGs {
					gradStopColors = append(gradStopColors, c)
					gradStopPositions = append(gradStopPositions, state.gradFillPos)
					lastColor = &gradStopColors[len(gradStopColors)-1]
				} else if state.inOuterShdw && pendingShadow != nil {
					pendingShadow.Color = c
					lastColor = &pendingShadow.Color
				} else if state.inFontRef {
					fontRefColor = &c
					lastColor = fontRefColor
				} else if state.inSolidFill && state.inRunProps && currentFont != nil && !state.inLn {
					currentFont.Color = c
					lastColor = &currentFont.Color
				} else if state.inSolidFill && state.inLn && !state.inRunProps {
					if state.inCxnSp && currentLine != nil {
						currentLine.lineColor = c
						lastColor = &currentLine.lineColor
					} else if state.inSp {
						if pendingBorder == nil {
							pendingBorder = &Border{Style: BorderSolid}
						}
						pendingBorder.Color = c
						lastColor = &pendingBorder.Color
					}
				} else if state.inSolidFill && state.inSpPr && !state.inRunProps && !state.inTxBody && !state.inLn {
					if state.inGrpSp && !state.inSp && len(grpStack) > 0 {
						f := NewFill()
						f.SetSolid(c)
						grpStack[len(grpStack)-1].grpFill = f
						lastColor = &grpStack[len(grpStack)-1].grpFill.Color
					} else if state.inSp {
						if currentRichText != nil {
							currentRichText.GetFill().SetSolid(c)
							lastColor = &currentRichText.GetFill().Color
						} else {
							pendingShapeFill = NewFill()
							pendingShapeFill.SetSolid(c)
							lastColor = &pendingShapeFill.Color
						}
					}
				} else if state.inBgSolidFill {
					if slide.background == nil {
						slide.background = NewFill()
					}
					slide.background.SetSolid(c)
					lastColor = &slide.background.Color
				} else if state.inBuClr && currentParagraph != nil && currentParagraph.bullet != nil {
					cc := c
					currentParagraph.bullet.Color = &cc
					lastColor = currentParagraph.bullet.Color
				} else if state.inDefRPr && state.inSolidFill && state.inLstStyleLvl1 && lstStyleFont != nil {
					lstStyleFont.Color = c
					lastColor = &lstStyleFont.Color
				} else if state.inDefRPr && state.inSolidFill && defFont != nil {
					defFont.Color = c
					lastColor = &defFont.Color
				}
			case "schemeClr":
				state.inSrgbClr = true // reuse for alpha child handling
				lastColor = nil
				if pres != nil && pres.themeColors != nil {
					var schemeName string
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							schemeName = attr.Value
						}
					}
					if argb, ok := pres.themeColors[schemeName]; ok && argb != "" {
						c := NewColor(argb)
						if state.inGs {
							gradStopColors = append(gradStopColors, c)
							gradStopPositions = append(gradStopPositions, state.gradFillPos)
							lastColor = &gradStopColors[len(gradStopColors)-1]
						} else if state.inOuterShdw && pendingShadow != nil {
							pendingShadow.Color = c
							lastColor = &pendingShadow.Color
						} else if state.inTcPrSolidFill {
							if state.inTcPrLn {
								lastColor = &c
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									cell := currentTable.rows[currentTableRow][currentTableCol]
									if cell.border != nil {
										switch state.tcPrLnSide {
										case "L":
											cell.border.Left.Color = c
											cell.border.Left.Style = BorderSolid
										case "R":
											cell.border.Right.Color = c
											cell.border.Right.Style = BorderSolid
										case "T":
											cell.border.Top.Color = c
											cell.border.Top.Style = BorderSolid
										case "B":
											cell.border.Bottom.Color = c
											cell.border.Bottom.Style = BorderSolid
										}
									}
								}
							} else {
								if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
									currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
									cell := currentTable.rows[currentTableRow][currentTableCol]
									cell.fill = NewFill()
									cell.fill.SetSolid(c)
									lastColor = &cell.fill.Color
								}
							}
						} else if state.inFontRef {
							// <p:style>/<a:fontRef>/<a:schemeClr> — default text color
							fontRefColor = &c
							lastColor = fontRefColor
						} else if state.inSolidFill && state.inRunProps && currentFont != nil {
							currentFont.Color = c
							lastColor = &currentFont.Color
						} else if state.inSolidFill && state.inLn && !state.inRunProps {
							if state.inCxnSp && currentLine != nil {
								currentLine.lineColor = c
								lastColor = &currentLine.lineColor
							} else if state.inSp {
								if pendingBorder == nil {
									pendingBorder = &Border{Style: BorderSolid}
								}
								pendingBorder.Color = c
								lastColor = &pendingBorder.Color
							}
						} else if state.inSolidFill && state.inSpPr && !state.inRunProps && !state.inTxBody && !state.inLn {
							if state.inGrpSp && !state.inSp && len(grpStack) > 0 {
								f := NewFill()
								f.SetSolid(c)
								grpStack[len(grpStack)-1].grpFill = f
								lastColor = &grpStack[len(grpStack)-1].grpFill.Color
							} else if state.inSp {
								if currentRichText != nil {
									currentRichText.GetFill().SetSolid(c)
									lastColor = &currentRichText.GetFill().Color
								} else {
									pendingShapeFill = NewFill()
									pendingShapeFill.SetSolid(c)
									lastColor = &pendingShapeFill.Color
								}
							}
						} else if state.inBgSolidFill {
							if slide.background == nil {
								slide.background = NewFill()
							}
							slide.background.SetSolid(c)
							lastColor = &slide.background.Color
						} else if state.inBuClr && currentParagraph != nil && currentParagraph.bullet != nil {
							currentParagraph.bullet.Color = &c
							lastColor = currentParagraph.bullet.Color
						} else if state.inDefRPr && state.inSolidFill && state.inLstStyleLvl1 && lstStyleFont != nil {
							lstStyleFont.Color = c
							lastColor = &lstStyleFont.Color
						} else if state.inDefRPr && state.inSolidFill && defFont != nil {
							defFont.Color = c
							lastColor = &defFont.Color
						}
					}
				}
			case "sysClr":
				// <a:sysClr val="window" lastClr="FFFFFF"/> — system color
				state.inSrgbClr = true // reuse for alpha/lumMod child handling
				lastColor = nil
				var sysLastClr string
				for _, attr := range t.Attr {
					if attr.Name.Local == "lastClr" {
						sysLastClr = attr.Value
					}
				}
				if sysLastClr != "" {
					c := NewColor("FF" + sysLastClr)
					if state.inOuterShdw && pendingShadow != nil {
						pendingShadow.Color = c
						lastColor = &pendingShadow.Color
					} else if state.inTcPrSolidFill {
						if state.inTcPrLn {
							lastColor = &c
							if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
								currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
								cell := currentTable.rows[currentTableRow][currentTableCol]
								if cell.border != nil {
									switch state.tcPrLnSide {
									case "L":
										cell.border.Left.Color = c
										cell.border.Left.Style = BorderSolid
									case "R":
										cell.border.Right.Color = c
										cell.border.Right.Style = BorderSolid
									case "T":
										cell.border.Top.Color = c
										cell.border.Top.Style = BorderSolid
									case "B":
										cell.border.Bottom.Color = c
										cell.border.Bottom.Style = BorderSolid
									}
								}
							}
						} else {
							if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
								currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
								cell := currentTable.rows[currentTableRow][currentTableCol]
								cell.fill = NewFill()
								cell.fill.SetSolid(c)
								lastColor = &cell.fill.Color
							}
						}
					} else if state.inFontRef {
						fontRefColor = &c
						lastColor = fontRefColor
					} else if state.inSolidFill && state.inRunProps && currentFont != nil {
						currentFont.Color = c
						lastColor = &currentFont.Color
					} else if state.inSolidFill && state.inLn && !state.inRunProps {
						if state.inCxnSp && currentLine != nil {
							currentLine.lineColor = c
							lastColor = &currentLine.lineColor
						} else if state.inSp {
							if pendingBorder == nil {
								pendingBorder = &Border{Style: BorderSolid}
							}
							pendingBorder.Color = c
							lastColor = &pendingBorder.Color
						}
					} else if state.inSolidFill && state.inSpPr && !state.inRunProps && !state.inTxBody && !state.inLn {
						if state.inGrpSp && !state.inSp && len(grpStack) > 0 {
							f := NewFill()
							f.SetSolid(c)
							grpStack[len(grpStack)-1].grpFill = f
							lastColor = &grpStack[len(grpStack)-1].grpFill.Color
						} else if state.inSp {
							if currentRichText != nil {
								currentRichText.GetFill().SetSolid(c)
								lastColor = &currentRichText.GetFill().Color
							} else {
								pendingShapeFill = NewFill()
								pendingShapeFill.SetSolid(c)
								lastColor = &pendingShapeFill.Color
							}
						}
					} else if state.inBgSolidFill {
						if slide.background == nil {
							slide.background = NewFill()
						}
						slide.background.SetSolid(c)
						lastColor = &slide.background.Color
					} else if state.inBuClr && currentParagraph != nil && currentParagraph.bullet != nil {
						currentParagraph.bullet.Color = &c
						lastColor = currentParagraph.bullet.Color
					} else if state.inDefRPr && state.inSolidFill && state.inLstStyleLvl1 && lstStyleFont != nil {
						lstStyleFont.Color = c
						lastColor = &lstStyleFont.Color
					} else if state.inDefRPr && state.inSolidFill && defFont != nil {
						defFont.Color = c
						lastColor = &defFont.Color
					}
				}
			case "alpha":
				// <a:alpha val="67000"/> means 67% opacity
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								// val is in 1/1000 of a percent, e.g. 67000 = 67%
								// For text run properties, skip val="0" to avoid
								// making text invisible (PowerPoint quirk).
								// For line/fill/gradient contexts, val="0" genuinely
								// means fully transparent.
								if v <= 0 && (state.inRunProps || state.inDefRPr) {
									continue
								}
								alpha := uint8(v * 255 / 100000)
								// Replace the alpha byte in the ARGB string
								alphaHex := fmt.Sprintf("%02X", alpha)
								lastColor.ARGB = alphaHex + lastColor.ARGB[2:]
								// Also update shadow Alpha when inside outerShdw
								if state.inOuterShdw && pendingShadow != nil {
									pendingShadow.Alpha = v / 1000 // convert to 0-100
								}
							}
						}
					}
				}
			case "lumMod":
				// Luminance modulation: multiply luminance by val/100000
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								applyLumMod(lastColor, float64(v)/100000.0)
							}
						}
					}
				}
			case "lumOff":
				// Luminance offset: add val/100000 to luminance
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								applyLumOff(lastColor, float64(v)/100000.0)
							}
						}
					}
				}
			case "tint":
				// Tint: blend toward white by val/100000
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								applyTint(lastColor, float64(v)/100000.0)
							}
						}
					}
				}
			case "shade":
				// Shade: blend toward black by val/100000
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								applyShade(lastColor, float64(v)/100000.0)
							}
						}
					}
				}
			case "latin":
				if state.inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							currentFont.Name = attr.Value
						}
					}
				} else if state.inDefRPr && state.inLstStyleLvl1 && lstStyleFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							lstStyleFont.Name = attr.Value
						}
					}
				} else if state.inDefRPr && defFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							defFont.Name = attr.Value
						}
					}
				}
			case "ea":
				// East Asian font
				if state.inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							currentFont.NameEA = attr.Value
						}
					}
				} else if state.inDefRPr && state.inLstStyleLvl1 && lstStyleFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							lstStyleFont.NameEA = attr.Value
						}
					}
				} else if state.inDefRPr && defFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							defFont.NameEA = attr.Value
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
			case "chOff":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "x":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							chOffX = v
						}
					case "y":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							chOffY = v
						}
					}
				}
			case "chExt":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							chExtCX = v
						}
					case "cy":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							chExtCY = v
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
				} else if state.inSpPrBlipFill {
					// <a:blip> inside <a:blipFill> inside <p:spPr> — shape image fill
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
										pendingBlipFillData = imgData
										pendingBlipFillMime = guessMimeType(imgPath)
									}
									break
								}
							}
						}
					}
				} else if state.inBgBlipFill {
					// <a:blip> inside <a:blipFill> inside <p:bgPr> — slide background image
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
										bgBlipFillData = imgData
										bgBlipFillMime = guessMimeType(imgPath)
									}
									break
								}
							}
						}
					}
				}
			case "alphaModFix":
				if state.inPic && currentDrawing != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "amt" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentDrawing.alpha = v
							}
						}
					}
				}
			case "srcRect":
				if state.inPic && currentDrawing != nil {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "l":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentDrawing.cropLeft = v
							}
						case "t":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentDrawing.cropTop = v
							}
						case "r":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentDrawing.cropRight = v
							}
						case "b":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentDrawing.cropBottom = v
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
								currentLine.lineWidthEMU = v
								currentLine.lineWidth = v / 12700
							}
						}
					}
				} else if state.inSp && state.inSpPr {
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if pendingBorder == nil {
									pendingBorder = &Border{Style: BorderSolid}
								}
								pendingBorder.Width = v / 12700
							}
						}
					}
				}
			case "headEnd":
				if state.inLn && state.inCxnSp && currentLine != nil {
					le := &LineEnd{Type: ArrowNone, Width: ArrowSizeMed, Length: ArrowSizeMed}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							le.Type = ArrowType(attr.Value)
						case "w":
							le.Width = ArrowSize(attr.Value)
						case "len":
							le.Length = ArrowSize(attr.Value)
						}
					}
					if le.Type != ArrowNone && le.Type != "" {
						currentLine.headEnd = le
					}
				} else if state.inLn && state.inSp {
					le := &LineEnd{Type: ArrowNone, Width: ArrowSizeMed, Length: ArrowSizeMed}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							le.Type = ArrowType(attr.Value)
						case "w":
							le.Width = ArrowSize(attr.Value)
						case "len":
							le.Length = ArrowSize(attr.Value)
						}
					}
					if le.Type != ArrowNone && le.Type != "" {
						pendingHeadEnd = le
					}
				}
			case "tailEnd":
				if state.inLn && state.inCxnSp && currentLine != nil {
					le := &LineEnd{Type: ArrowNone, Width: ArrowSizeMed, Length: ArrowSizeMed}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							le.Type = ArrowType(attr.Value)
						case "w":
							le.Width = ArrowSize(attr.Value)
						case "len":
							le.Length = ArrowSize(attr.Value)
						}
					}
					if le.Type != ArrowNone && le.Type != "" {
						currentLine.tailEnd = le
					}
				} else if state.inLn && state.inSp {
					le := &LineEnd{Type: ArrowNone, Width: ArrowSizeMed, Length: ArrowSizeMed}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							le.Type = ArrowType(attr.Value)
						case "w":
							le.Width = ArrowSize(attr.Value)
						case "len":
							le.Length = ArrowSize(attr.Value)
						}
					}
					if le.Type != ArrowNone && le.Type != "" {
						pendingTailEnd = le
					}
				}
			case "prstDash":
				if state.inLn && state.inCxnSp && currentLine != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							switch attr.Value {
							case "dash", "lgDash", "sysDash":
								currentLine.lineStyle = BorderDash
							case "dot", "sysDot":
								currentLine.lineStyle = BorderDot
							case "solid":
								currentLine.lineStyle = BorderSolid
							}
						}
					}
				} else if state.inLn && state.inSp {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							switch attr.Value {
							case "dash", "lgDash", "sysDash":
								if pendingBorder == nil {
									pendingBorder = &Border{Style: BorderDash}
								} else {
									pendingBorder.Style = BorderDash
								}
							case "dot", "sysDot":
								if pendingBorder == nil {
									pendingBorder = &Border{Style: BorderDot}
								} else {
									pendingBorder.Style = BorderDot
								}
							}
						}
					}
				}
			case "effectLst":
				if state.inSpPr && !state.inLn {
					state.inEffectLst = true
				}
			case "outerShdw":
				if state.inEffectLst {
					state.inOuterShdw = true
					pendingShadow = NewShadow()
					pendingShadow.Visible = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "blurRad":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								pendingShadow.BlurRadius = v / 12700
							}
						case "dist":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								pendingShadow.Distance = v / 12700
							}
						case "dir":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								pendingShadow.Direction = v / 60000
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
			case "custGeom":
				if state.inSpPr {
					state.inCustGeom = true
				}
			case "pathLst":
				if state.inCustGeom {
					state.inPathLst = true
				}
			case "path":
				if state.inPathLst {
					state.inCustPath = true
					pendingPathCmds = nil
					var pw, ph int64
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "w":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								pw = v
							}
						case "h":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								ph = v
							}
						}
					}
					pendingCustomPath = &CustomGeomPath{Width: pw, Height: ph}
				}
			case "moveTo":
				if state.inCustPath {
					pendingPathCmds = append(pendingPathCmds, PathCommand{Type: "moveTo"})
				}
			case "lnTo":
				if state.inCustPath {
					pendingPathCmds = append(pendingPathCmds, PathCommand{Type: "lnTo"})
				}
			case "cubicBezTo":
				if state.inCustPath {
					pendingPathCmds = append(pendingPathCmds, PathCommand{Type: "cubicBezTo"})
				}
			case "quadBezTo":
				if state.inCustPath {
					pendingPathCmds = append(pendingPathCmds, PathCommand{Type: "quadBezTo"})
				}
			case "arcTo":
				if state.inCustPath {
					var wR, hR, stAng, swAng int64
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "wR":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								wR = v
							}
						case "hR":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								hR = v
							}
						case "stAng":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								stAng = v
							}
						case "swAng":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								swAng = v
							}
						}
					}
					pendingPathCmds = append(pendingPathCmds, PathCommand{
						Type:  "arcTo",
						WR:    wR,
						HR:    hR,
						StAng: stAng,
						SwAng: swAng,
					})
				}
			case "close":
				if state.inCustPath {
					pendingPathCmds = append(pendingPathCmds, PathCommand{Type: "close"})
				}
			case "pt":
				if state.inCustPath && len(pendingPathCmds) > 0 {
					var px, py int64
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "x":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								px = v
							}
						case "y":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								py = v
							}
						}
					}
					last := &pendingPathCmds[len(pendingPathCmds)-1]
					last.Pts = append(last.Pts, PathPoint{X: px, Y: py})
				}
			case "avLst":
				if state.inSpPr {
					state.inAvLst = true
				}
			case "gd":
				if state.inAvLst {
					var gdName string
					var gdVal int
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							gdName = attr.Value
						case "fmla":
							// fmla is "val NNNNN"
							if strings.HasPrefix(attr.Value, "val ") {
								if v, err := strconv.Atoi(strings.TrimPrefix(attr.Value, "val ")); err == nil {
									gdVal = v
								}
							}
						}
					}
					if gdName != "" {
						if pendingAdjustValues == nil {
							pendingAdjustValues = make(map[string]int)
						}
						pendingAdjustValues[gdName] = gdVal
					}
				}
			case "style":
				// <p:style> element inside <p:sp> — provides default styling
				if state.inSp && !state.inSpPr && !state.inTxBody {
					state.inStyle = true
				}
			case "fontRef":
				// <a:fontRef> inside <p:style> — provides default text color
				if state.inStyle {
					state.inFontRef = true
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
				state.inBgBlipFill = false
			case "spTree":
				state.inSpTree = false
			case "grpSp":
				if state.inGrpSp {
					grpDepth--
					n := len(grpStack)
					if n > 0 {
						top := grpStack[n-1]
						grpStack = grpStack[:n-1]
						g := top.group
						if g != nil {
							g.name = top.name
							g.description = top.descr
							g.offsetX = top.offX
							g.offsetY = top.offY
							g.width = top.extCX
							g.height = top.extCY
							g.childOffX = top.chOffX
							g.childOffY = top.chOffY
							g.childExtX = top.chExtCX
							g.childExtY = top.chExtCY
							g.flipHorizontal = top.flipH
							g.flipVertical = top.flipV
							g.rotation = top.rotation
							g.groupFill = top.grpFill
							// Add to parent group or slide
							if len(grpStack) > 0 {
								parentGroup := grpStack[len(grpStack)-1].group
								parentGroup.AddShape(g)
							} else {
								slide.shapes = append(slide.shapes, g)
							}
						}
					}
					if grpDepth <= 0 {
						state.inGrpSp = false
						currentGroup = nil
					} else {
						// Restore currentGroup to parent
						currentGroup = grpStack[len(grpStack)-1].group
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
						// Apply deferred adjustment values
						if pendingAdjustValues != nil {
							autoShape.adjustValues = pendingAdjustValues
							pendingAdjustValues = nil
						}
						// Apply deferred shape-level fill
						if pendingShapeFill != nil {
							autoShape.fill = pendingShapeFill
							pendingShapeFill = nil
						}
						// Apply deferred shape-level border
						if pendingBorder != nil {
							autoShape.border = pendingBorder
							pendingBorder = nil
						}
						// Apply deferred shadow
						if pendingShadow != nil {
							autoShape.shadow = pendingShadow
							pendingShadow = nil
						}
																// Apply deferred arrow ends
										if pendingHeadEnd != nil {
											autoShape.headEnd = pendingHeadEnd
											pendingHeadEnd = nil
										}
										if pendingTailEnd != nil {
											autoShape.tailEnd = pendingTailEnd
											pendingTailEnd = nil
										}
// Copy paragraphs from richtext if any (preserves font info)
						if currentRichText != nil && len(currentRichText.paragraphs) > 0 {
							autoShape.paragraphs = currentRichText.paragraphs
							autoShape.textAnchor = textAnchor
							autoShape.textDirection = textDir
							autoShape.fontScale = currentRichText.fontScale
							// Copy text insets from richtext body properties
							if currentRichText.insetsSet {
								autoShape.insetLeft = currentRichText.insetLeft
								autoShape.insetRight = currentRichText.insetRight
								autoShape.insetTop = currentRichText.insetTop
								autoShape.insetBottom = currentRichText.insetBottom
								autoShape.insetsSet = true
							}
							// Default to middle anchor for AutoShapes (PowerPoint default)
							if autoShape.textAnchor == TextAnchorNone {
								autoShape.textAnchor = TextAnchorMiddle
							}
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
					} else if len(pendingBlipFillData) > 0 {
						// Shape has blipFill — convert to DrawingShape
						ds := NewDrawingShape()
						ds.name = shapeName
						ds.description = shapeDescr
						ds.offsetX = offX
						ds.offsetY = offY
						ds.width = extCX
						ds.height = extCY
						ds.flipHorizontal = flipH
						ds.flipVertical = flipV
						ds.rotation = shapeRotation
						ds.data = pendingBlipFillData
						ds.mimeType = pendingBlipFillMime
						pendingBlipFillData = nil
						pendingBlipFillMime = ""
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(ds)
						} else {
							slide.shapes = append(slide.shapes, ds)
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
						// Apply deferred shape-level border
						if pendingBorder != nil {
							currentRichText.border = pendingBorder
							pendingBorder = nil
						}
						// Apply deferred shadow
						if pendingShadow != nil {
							currentRichText.shadow = pendingShadow
							pendingShadow = nil
						}
						// Apply deferred arrow ends
						if pendingHeadEnd != nil {
							currentRichText.headEnd = pendingHeadEnd
							pendingHeadEnd = nil
						}
						if pendingTailEnd != nil {
							currentRichText.tailEnd = pendingTailEnd
							pendingTailEnd = nil
						}
						// Apply custom geometry path
						if pendingCustomPath != nil {
							currentRichText.customPath = pendingCustomPath
							pendingCustomPath = nil
						}
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentRichText)
						} else {
							slide.shapes = append(slide.shapes, currentRichText)
						}
					} else if pendingCustomPath != nil {
						// Shape has custom geometry but no text body — create a
						// RichTextShape to carry the custom path, fill, and border.
						rt := NewRichTextShape()
						rt.name = shapeName
						rt.description = shapeDescr
						rt.offsetX = offX
						rt.offsetY = offY
						rt.width = extCX
						rt.height = extCY
						rt.flipHorizontal = flipH
						rt.flipVertical = flipV
						rt.rotation = shapeRotation
						rt.customPath = pendingCustomPath
						pendingCustomPath = nil
						if pendingShapeFill != nil {
							rt.fill = pendingShapeFill
							pendingShapeFill = nil
						}
						if pendingBorder != nil {
							rt.border = pendingBorder
							pendingBorder = nil
						}
						if pendingShadow != nil {
							rt.shadow = pendingShadow
							pendingShadow = nil
						}
						if pendingHeadEnd != nil {
							rt.headEnd = pendingHeadEnd
							pendingHeadEnd = nil
						}
						if pendingTailEnd != nil {
							rt.tailEnd = pendingTailEnd
							pendingTailEnd = nil
						}
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(rt)
						} else {
							slide.shapes = append(slide.shapes, rt)
						}
					} else if prstGeom != "" && (pendingShapeFill != nil || pendingBorder != nil || pendingShadow != nil) {
						// Shape with geometry (including rect) that has fill or border
						// but no text body — create an AutoShape so it gets rendered.
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
						if pendingAdjustValues != nil {
							autoShape.adjustValues = pendingAdjustValues
							pendingAdjustValues = nil
						}
						if pendingShapeFill != nil {
							autoShape.fill = pendingShapeFill
							pendingShapeFill = nil
						}
						if pendingBorder != nil {
							autoShape.border = pendingBorder
							pendingBorder = nil
						}
						if pendingShadow != nil {
							autoShape.shadow = pendingShadow
							pendingShadow = nil
						}
																// Apply deferred arrow ends
										if pendingHeadEnd != nil {
											autoShape.headEnd = pendingHeadEnd
											pendingHeadEnd = nil
										}
										if pendingTailEnd != nil {
											autoShape.tailEnd = pendingTailEnd
											pendingTailEnd = nil
										}
if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(autoShape)
						} else {
							slide.shapes = append(slide.shapes, autoShape)
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
						currentLine.connectorType = prstGeom
						if pendingAdjustValues != nil {
							currentLine.adjustValues = pendingAdjustValues
							pendingAdjustValues = nil
						}
						if pendingCustomPath != nil {
							currentLine.customPath = pendingCustomPath
							pendingCustomPath = nil
						}
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
				state.inTcPr = false
				state.inTcPrSolidFill = false
				state.inTcPrLn = false
				state.tcPrLnSide = ""
			case "tcPr":
				state.inTcPr = false
				state.inTcPrSolidFill = false
				state.inTcPrLn = false
				state.tcPrLnSide = ""
			case "lnL", "lnR", "lnT", "lnB":
				if state.inTcPr {
					state.inTcPrLn = false
					state.inTcPrSolidFill = false
					state.tcPrLnSide = ""
				}
			case "txBody":
				if state.inTc {
					state.inTcTxBody = false
				} else {
					state.inTxBody = false
					state.inLstStyle = false
					state.inLstStyleLvl1 = false
					lstStyleFont = nil
				}
			case "p":
				if state.inTcParagraph {
					state.inTcParagraph = false
				} else {
					state.inParagraph = false
				}
				currentParagraph = nil
				defFont = nil
			case "pPr":
				state.inPPr = false
				state.inSpcBef = false
				state.inSpcAft = false
				state.inLnSpc = false
				state.inDefRPr = false
			case "spcBef":
				state.inSpcBef = false
			case "spcAft":
				state.inSpcAft = false
			case "lnSpc":
				state.inLnSpc = false
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
				state.inRunPropsGradFill = false
			case "defRPr":
				state.inDefRPr = false
				state.inSolidFill = false
			case "lstStyle":
				state.inLstStyle = false
				state.inLstStyleLvl1 = false
			case "lvl1pPr":
				state.inLstStyleLvl1 = false
			case "solidFill":
				state.inSolidFill = false
				state.inBgSolidFill = false
				state.inTcPrSolidFill = false
				state.inSrgbClr = false
				lastColor = nil
			case "gs":
				state.inGs = false
			case "gsLst":
				state.inGsLst = false
			case "gradFill":
				if state.inRunPropsGradFill && currentFont != nil && len(gradStopColors) >= 1 {
					// Use first gradient stop color as text color
					currentFont.Color = gradStopColors[0]
					state.inRunPropsGradFill = false
				} else if state.inGradFill && len(gradStopColors) >= 2 {
					startColor := gradStopColors[0]
					endColor := gradStopColors[len(gradStopColors)-1]
					if state.inBgPr {
						if slide.background == nil {
							slide.background = NewFill()
						}
						slide.background.SetGradientLinear(startColor, endColor, gradAngle)
					} else if state.inSpPr && state.inSp {
						pendingShapeFill = NewFill()
						pendingShapeFill.SetGradientLinear(startColor, endColor, gradAngle)
					}
				}
				state.inGradFill = false
			case "blipFill":
				state.inSpPrBlipFill = false
				state.inBgBlipFill = false
			case "srgbClr":
				state.inSrgbClr = false
			case "schemeClr":
				state.inSrgbClr = false
			case "outerShdw":
				state.inOuterShdw = false
			case "effectLst":
				state.inEffectLst = false
			case "spPr", "grpSpPr":
				state.inSpPr = false
				state.inLn = false
				state.inExtLst = false
				state.inEffectLst = false
				state.inOuterShdw = false
				state.inSpPrBlipFill = false
				// When the group's shape properties end, save position/size
				// before child shapes overwrite the shared variables.
				if t.Name.Local == "grpSpPr" && state.inGrpSp && len(grpStack) > 0 {
					top := grpStack[len(grpStack)-1]
					top.offX = offX
					top.offY = offY
					top.extCX = extCX
					top.extCY = extCY
					top.chOffX = chOffX
					top.chOffY = chOffY
					top.chExtCX = chExtCX
					top.chExtCY = chExtCY
					top.flipH = flipH
					top.flipV = flipV
					top.rotation = shapeRotation
				}
			case "ln":
				state.inLn = false
			case "extLst":
				state.inExtLst = false
			case "avLst":
				state.inAvLst = false
			case "custGeom":
				state.inCustGeom = false
			case "pathLst":
				state.inPathLst = false
			case "path":
				if state.inCustPath && pendingCustomPath != nil {
					pendingCustomPath.Commands = pendingPathCmds
					pendingPathCmds = nil
					state.inCustPath = false
				}
			case "buClr":
				state.inBuClr = false
			case "style":
				state.inStyle = false
				state.inFontRef = false
			case "fontRef":
				state.inFontRef = false
			case "t":
				state.inText = false
				state.inTcText = false
			case "nvSpPr", "nvPicPr", "nvCxnSpPr", "nvGraphicFramePr", "nvGrpSpPr":
				state.inNvSpPr = false
				// When the group's non-visual properties end, save the group name
				// before child shapes overwrite the shared shapeName variable.
				if t.Name.Local == "nvGrpSpPr" && state.inGrpSp && len(grpStack) > 0 {
					top := grpStack[len(grpStack)-1]
					if top.name == "" {
						top.name = shapeName
						top.descr = shapeDescr
					}
				}
			}
		}
	}

	// If slide has a blipFill background image, prepend as full-slide drawing
	if len(bgBlipFillData) > 0 && pres != nil {
		ds := NewDrawingShape()
		ds.data = bgBlipFillData
		ds.mimeType = bgBlipFillMime
		ds.offsetX = 0
		ds.offsetY = 0
		ds.width = pres.layout.CX
		ds.height = pres.layout.CY
		slide.shapes = append([]Shape{ds}, slide.shapes...)
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
	case strings.HasSuffix(lower, ".wmf"):
		return "image/x-wmf"
	case strings.HasSuffix(lower, ".emf"):
		return "image/x-emf"
	case strings.HasSuffix(lower, ".tiff"), strings.HasSuffix(lower, ".tif"):
		return "image/tiff"
	case strings.HasSuffix(lower, ".wdp"):
		return "image/vnd.ms-photo"
	default:
		return "image/png"
	}
}

// layoutPlaceholder holds position/size/font info extracted from a slide layout.
type layoutPlaceholder struct {
	phType string
	phIdx  int
	offX   int64
	offY   int64
	extCX  int64
	extCY  int64
	// Default font properties from defRPr
	fontName string
	fontEA   string
	fontSize int
	fontBold bool
	fontColor Color
	// Text insets from bodyPr
	insetLeft   int64
	insetRight  int64
	insetTop    int64
	insetBottom int64
	insetsSet   bool
}

// applyLayoutInheritance reads the slide layout and applies inherited properties
// to placeholders that have zero size (meaning they inherit from the layout).
func (r *PPTXReader) applyLayoutInheritance(zr *zip.Reader, slide *Slide, rels []xmlRelForRead, slidePath string, pres *Presentation) {
	// Find the slide layout relationship
	layoutPath := ""
	for _, rel := range rels {
		if rel.Type == relTypeSlideLayout {
			target := rel.Target
			if !strings.HasPrefix(target, "ppt/") {
				dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
				target = resolveRelativePath(dir, target)
			}
			layoutPath = target
			break
		}
	}
	if layoutPath == "" {
		return
	}

	data, err := readFileFromZip(zr, layoutPath)
	if err != nil {
		return
	}

	// Read layout relationships for images
	layoutRelsPath := strings.Replace(layoutPath, "slideLayouts/", "slideLayouts/_rels/", 1) + ".rels"
	layoutRels, _ := r.readRelationships(zr, layoutRelsPath)

	// Parse layout images and non-placeholder text shapes, prepend to slide (behind slide content)
	layoutImages := r.parseLayoutImages(data, layoutRels, zr, layoutPath, pres)
	if len(layoutImages) > 0 {
		slide.shapes = append(layoutImages, slide.shapes...)
	}

	// Parse layout to extract placeholder definitions
	layoutPHs := r.parseLayoutPlaceholders(data, pres)

	// Also parse layout background
	layoutBg, bgImage := r.parseLayoutBackground(data, layoutRels, zr, layoutPath, pres)

	// Apply layout background if slide has no background
	if slide.background == nil && layoutBg != nil {
		slide.background = layoutBg
	}
	// If layout has a blipFill background image, prepend as full-slide drawing.
	// Skip if slide already has a background image (first shape is a full-slide DrawingShape).
	if bgImage != nil && slide.background == nil {
		hasSlideBgImage := false
		if len(slide.shapes) > 0 {
			if ds, ok := slide.shapes[0].(*DrawingShape); ok && ds.offsetX == 0 && ds.offsetY == 0 && ds.width == pres.layout.CX && ds.height == pres.layout.CY {
				hasSlideBgImage = true
			}
		}
		if !hasSlideBgImage {
			bgImage.offsetX = 0
			bgImage.offsetY = 0
			bgImage.width = pres.layout.CX
			bgImage.height = pres.layout.CY
			slide.shapes = append([]Shape{bgImage}, slide.shapes...)
		}
	}

	if len(layoutPHs) == 0 {
		return
	}

	// Apply layout properties to slide placeholders
	for _, shape := range slide.shapes {
		ph, ok := shape.(*PlaceholderShape)
		if !ok {
			continue
		}

		// Find matching layout placeholder by type and idx
		var match *layoutPlaceholder
		for i := range layoutPHs {
			lp := &layoutPHs[i]
			if lp.phType == string(ph.phType) && lp.phIdx == ph.phIdx {
				match = lp
				break
			}
			// Also match by type alone if idx is 0 (default)
			if lp.phType == string(ph.phType) && ph.phIdx == 0 && lp.phIdx == 0 {
				match = lp
				break
			}
		}
		if match == nil {
			// Try matching by type only (ignoring idx)
			for i := range layoutPHs {
				lp := &layoutPHs[i]
				if lp.phType == string(ph.phType) {
					match = lp
					break
				}
			}
		}
		if match == nil {
			continue
		}

		// Apply position/size if placeholder has zero size
		if ph.width == 0 && ph.height == 0 {
			ph.offsetX = match.offX
			ph.offsetY = match.offY
			ph.width = match.extCX
			ph.height = match.extCY
		}

		// Apply text insets from layout if not set on the placeholder
		if !ph.insetsSet && match.insetsSet {
			ph.insetLeft = match.insetLeft
			ph.insetRight = match.insetRight
			ph.insetTop = match.insetTop
			ph.insetBottom = match.insetBottom
			ph.insetsSet = true
		}

		// Apply font properties to text runs that have default fonts
		if match.fontName != "" || match.fontEA != "" || match.fontSize > 0 {
			for _, para := range ph.paragraphs {
				for _, elem := range para.elements {
					tr, ok := elem.(*TextRun)
					if !ok || tr.font == nil {
						continue
					}
					// Apply layout font if the run has default Calibri/size=10
					if tr.font.Name == "Calibri" && match.fontName != "" {
						tr.font.Name = match.fontName
					}
					if tr.font.NameEA == "" && match.fontEA != "" {
						tr.font.NameEA = match.fontEA
					}
					if (tr.font.Size == 18 || tr.font.Size <= 10) && match.fontSize > 0 {
						tr.font.Size = match.fontSize
					}
					if match.fontBold {
						tr.font.Bold = true
					}
					if match.fontColor.ARGB != "" && match.fontColor.ARGB != "FF000000" {
						// Only apply if run has default black color
						if tr.font.Color.ARGB == "FF000000" {
							tr.font.Color = match.fontColor
						}
					}
				}
			}
		}
	}
}

// parseLayoutPlaceholders extracts placeholder definitions from a slide layout XML.
func (r *PPTXReader) parseLayoutPlaceholders(data []byte, pres *Presentation) []layoutPlaceholder {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var phs []layoutPlaceholder

	inSp := false
	inNvSpPr := false
	inSpPr := false
	inTxBody := false
	inLstStyle := false
	inDefRPr := false
	inDefSolidFill := false

	isPH := false
	var phType string
	var phIdx int
	var offX, offY, extCX, extCY int64
	var fontName, fontEA string
	var fontSize int
	var fontBold bool
	var fontColor Color
	var insetLeft, insetRight, insetTop, insetBottom int64
	var insetsSet bool

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sp":
				inSp = true
				isPH = false
				phType = ""
				phIdx = 0
				offX, offY, extCX, extCY = 0, 0, 0, 0
				fontName = ""
				fontEA = ""
				fontSize = 0
				fontBold = false
				fontColor = Color{}
				// Initialize to PowerPoint defaults
				insetLeft, insetRight = 91440, 91440
				insetTop, insetBottom = 45720, 45720
				insetsSet = false
			case "nvSpPr":
				if inSp {
					inNvSpPr = true
				}
			case "ph":
				if inNvSpPr {
					isPH = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							phType = attr.Value
						case "idx":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								phIdx = v
							}
						}
					}
				}
			case "spPr":
				if inSp && !inNvSpPr {
					inSpPr = true
				}
			case "off":
				if inSpPr {
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
				}
			case "ext":
				if inSpPr {
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
				}
			case "txBody":
				if inSp {
					inTxBody = true
				}
			case "bodyPr":
				if inTxBody {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "lIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								insetLeft = v
								insetsSet = true
							}
						case "rIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								insetRight = v
								insetsSet = true
							}
						case "tIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								insetTop = v
								insetsSet = true
							}
						case "bIns":
							if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
								insetBottom = v
								insetsSet = true
							}
						}
					}
				}
			case "lstStyle":
				if inTxBody {
					inLstStyle = true
				}
			case "defRPr":
				if inLstStyle {
					inDefRPr = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "sz":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								fontSize = v / 100
							}
						case "b":
							fontBold = attr.Value == "1"
						}
					}
				}
			case "solidFill":
				if inDefRPr {
					inDefSolidFill = true
				}
			case "srgbClr":
				if inDefSolidFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							fontColor = NewColor("FF" + attr.Value)
						}
					}
				}
			case "sysClr":
				if inDefSolidFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "lastClr" {
							fontColor = NewColor("FF" + attr.Value)
						}
					}
				}
			case "schemeClr":
				// Handle scheme colors in layout placeholder defRPr
				if inDefSolidFill {
					var schemeName string
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							schemeName = attr.Value
						}
					}
					if pres != nil && pres.themeColors != nil {
						if argb, ok := pres.themeColors[schemeName]; ok && argb != "" {
							fontColor = NewColor(argb)
						}
					}
				}
			case "latin":
				if inDefRPr {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" {
							fontName = attr.Value
						}
					}
				}
			case "ea":
				if inDefRPr {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" {
							fontEA = attr.Value
						}
					}
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "sp":
				if inSp && isPH {
					phs = append(phs, layoutPlaceholder{
						phType:      phType,
						phIdx:       phIdx,
						offX:        offX,
						offY:        offY,
						extCX:       extCX,
						extCY:       extCY,
						fontName:    fontName,
						fontEA:      fontEA,
						fontSize:    fontSize,
						fontBold:    fontBold,
						fontColor:   fontColor,
						insetLeft:   insetLeft,
						insetRight:  insetRight,
						insetTop:    insetTop,
						insetBottom: insetBottom,
						insetsSet:   insetsSet,
					})
				}
				inSp = false
				inSpPr = false
				inTxBody = false
				inLstStyle = false
				inDefRPr = false
			case "nvSpPr":
				inNvSpPr = false
			case "spPr":
				inSpPr = false
			case "txBody":
				inTxBody = false
				inLstStyle = false
			case "lstStyle":
				inLstStyle = false
			case "defRPr":
				inDefRPr = false
				inDefSolidFill = false
			case "solidFill":
				inDefSolidFill = false
			}
		}
	}

	return phs
}

// parseLayoutBackground extracts the background fill from a slide layout XML.
// It also handles blipFill backgrounds by returning an image shape via bgImage.
func (r *PPTXReader) parseLayoutBackground(data []byte, rels []xmlRelForRead, zr *zip.Reader, layoutPath string, pres *Presentation) (*Fill, *DrawingShape) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	inBg := false
	inBgPr := false
	inSolidFill := false
	inBlipFill := false

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "bg":
				inBg = true
			case "bgPr":
				if inBg {
					inBgPr = true
				}
			case "solidFill":
				if inBgPr {
					inSolidFill = true
				}
			case "blipFill":
				if inBgPr {
					inBlipFill = true
				}
			case "blip":
				if inBlipFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "embed" {
							for _, rel := range rels {
								if rel.ID == attr.Value {
									imgPath := rel.Target
									if !strings.HasPrefix(imgPath, "ppt/") {
										dir := strings.TrimSuffix(layoutPath, "/"+lastPathComponent(layoutPath))
										imgPath = resolveRelativePath(dir, imgPath)
									}
									imgData, err := readFileFromZip(zr, imgPath)
									if err == nil {
										ds := NewDrawingShape()
										ds.data = imgData
										ds.mimeType = guessMimeType(imgPath)
										return nil, ds
									}
									break
								}
							}
						}
					}
				}
			case "srgbClr":
				if inSolidFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							fill := NewFill()
							fill.SetSolid(NewColor("FF" + attr.Value))
							return fill, nil
						}
					}
				}
			case "schemeClr":
				if inSolidFill {
					var schemeName string
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							schemeName = attr.Value
						}
					}
					if pres != nil && pres.themeColors != nil {
						if argb, ok := pres.themeColors[schemeName]; ok && argb != "" {
							fill := NewFill()
							fill.SetSolid(NewColor(argb))
							return fill, nil
						}
					}
					// Fallback: treat bg1 as white
					if schemeName == "bg1" {
						fill := NewFill()
						fill.SetSolid(ColorWhite)
						return fill, nil
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "bg":
				return nil, nil // bg found but no recognized fill
			case "bgPr":
				inBgPr = false
			case "solidFill":
				inSolidFill = false
			case "blipFill":
				inBlipFill = false
			}
		}
	}
	return nil, nil
}

// parseLayoutImages extracts image shapes and non-placeholder text shapes from a slide layout XML.
func (r *PPTXReader) parseLayoutImages(data []byte, rels []xmlRelForRead, zr *zip.Reader, layoutPath string, pres *Presentation) []Shape {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var shapes []Shape

	inPic := false
	inSp := false
	inCxnSp := false
	inSpPr := false
	inNvSpPr := false
	inNvPr := false
	inTxBody := false
	inLstStyle := false
	inLstStyleLvl1 := false
	inDefRPr := false
	inDefSolidFill := false
	inParagraph := false
	inRun := false
	inRunProps := false
	inRunSolidFill := false
	inText := false
	inPPr := false
	inPPrDefRPr := false
	inPPrDefSolidFill := false
	inLn := false
	inLnSolidFill := false
	isPH := false
	var offX, offY, extCX, extCY int64
	var embedID string
	var flipH, flipV bool
	var picAlpha int // alphaModFix amount for pic blip
	var cropL, cropT, cropR, cropB int // srcRect crop percentages

	// For cxnSp (line connector) shapes
	var currentLine *LineShape

	// For non-placeholder sp shapes
	var currentRichText *RichTextShape
	var currentParagraph *Paragraph
	var currentFont *Font
	var lstStyleFont *Font
	var defFont *Font
	var textAnchor TextAnchorType

	// Color tracking for schemeClr inside rPr
	inSrgbClr := false
	var lastColor *Color

	// Font color from <p:style>/<a:fontRef>/<a:schemeClr>
	inStyle := false
	inFontRef := false
	var fontRefColor *Color

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pic":
				inPic = true
				offX, offY, extCX, extCY = 0, 0, 0, 0
				embedID = ""
				picAlpha = 0
				cropL, cropT, cropR, cropB = 0, 0, 0, 0
			case "sp":
				inSp = true
				isPH = false
				offX, offY, extCX, extCY = 0, 0, 0, 0
				flipH, flipV = false, false
				currentRichText = nil
				lstStyleFont = nil
				defFont = nil
				textAnchor = TextAnchorNone
				fontRefColor = nil
			case "cxnSp":
				inCxnSp = true
				offX, offY, extCX, extCY = 0, 0, 0, 0
				flipH, flipV = false, false
				currentLine = NewLineShape()
			case "nvSpPr":
				if inSp {
					inNvSpPr = true
				}
			case "nvCxnSpPr":
				if inCxnSp {
					inNvSpPr = true
				}
			case "nvPr":
				if inNvSpPr {
					inNvPr = true
				}
			case "ph":
				if inNvPr {
					isPH = true
				}
			case "spPr":
				if inPic || (inSp && !inNvSpPr) || (inCxnSp && !inNvSpPr) {
					inSpPr = true
				}
			case "off":
				if inSpPr {
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
				}
			case "ext":
				if inSpPr {
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
				}
			case "blip":
				if inPic {
					for _, attr := range t.Attr {
						if attr.Name.Local == "embed" {
							embedID = attr.Value
						}
					}
				}
			case "alphaModFix":
				if inPic {
					for _, attr := range t.Attr {
						if attr.Name.Local == "amt" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								picAlpha = v
							}
						}
					}
				}
			case "srcRect":
				if inPic {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "l":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cropL = v
							}
						case "t":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cropT = v
							}
						case "r":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cropR = v
							}
						case "b":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cropB = v
							}
						}
					}
				}
			case "xfrm":
				if inSpPr {
					flipH = false
					flipV = false
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "flipH":
							flipH = attr.Value == "1" || attr.Value == "true"
						case "flipV":
							flipV = attr.Value == "1" || attr.Value == "true"
						}
					}
				}
			case "ln":
				if inSpPr && inCxnSp {
					inLn = true
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if currentLine != nil {
									currentLine.lineWidthEMU = v
									currentLine.lineWidth = v / 12700
								}
							}
						}
					}
				}
			case "headEnd":
				if inLn && inCxnSp && currentLine != nil {
					le := &LineEnd{Type: ArrowNone, Width: ArrowSizeMed, Length: ArrowSizeMed}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							le.Type = ArrowType(attr.Value)
						case "w":
							le.Width = ArrowSize(attr.Value)
						case "len":
							le.Length = ArrowSize(attr.Value)
						}
					}
					if le.Type != ArrowNone && le.Type != "" {
						currentLine.headEnd = le
					}
				}
			case "tailEnd":
				if inLn && inCxnSp && currentLine != nil {
					le := &LineEnd{Type: ArrowNone, Width: ArrowSizeMed, Length: ArrowSizeMed}
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							le.Type = ArrowType(attr.Value)
						case "w":
							le.Width = ArrowSize(attr.Value)
						case "len":
							le.Length = ArrowSize(attr.Value)
						}
					}
					if le.Type != ArrowNone && le.Type != "" {
						currentLine.tailEnd = le
					}
				}
			case "prstDash":
				if inLn && inCxnSp && currentLine != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							switch attr.Value {
							case "dash", "lgDash", "sysDash":
								currentLine.lineStyle = BorderDash
							case "dot", "sysDot":
								currentLine.lineStyle = BorderDot
							case "solid":
								currentLine.lineStyle = BorderSolid
							}
						}
					}
				}
			case "txBody":
				if inSp && !isPH {
					inTxBody = true
					currentRichText = NewRichTextShape()
					currentRichText.paragraphs = nil
				}
			case "bodyPr":
				if inTxBody && currentRichText != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "anchor" {
							textAnchor = TextAnchorType(attr.Value)
						}
					}
				}
			case "normAutofit":
				if inTxBody && currentRichText != nil {
					fontScaleVal := 100000
					for _, attr := range t.Attr {
						if attr.Name.Local == "fontScale" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								fontScaleVal = v
							}
						}
					}
					currentRichText.autoFit = AutoFitNormal
					currentRichText.fontScale = fontScaleVal
				}
			case "lstStyle":
				if inTxBody {
					inLstStyle = true
				}
			case "lvl1pPr":
				if inLstStyle {
					inLstStyleLvl1 = true
				}
			case "defRPr":
				if inLstStyleLvl1 {
					inDefRPr = true
					lstStyleFont = NewFont()
					lstStyleFont.Size = 0
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "sz":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								lstStyleFont.Size = v / 100
							}
						case "b":
							lstStyleFont.Bold = attr.Value == "1"
						case "i":
							lstStyleFont.Italic = attr.Value == "1"
						}
					}
				} else if inPPr {
					inPPrDefRPr = true
					defFont = NewFont()
					defFont.Size = 0
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "sz":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								defFont.Size = v / 100
							}
						case "b":
							defFont.Bold = attr.Value == "1"
						case "i":
							defFont.Italic = attr.Value == "1"
						}
					}
				}
			case "solidFill":
				if inDefRPr {
					inDefSolidFill = true
				} else if inPPrDefRPr {
					inPPrDefSolidFill = true
				} else if inRunProps {
					inRunSolidFill = true
				} else if inLn {
					inLnSolidFill = true
				}
			case "srgbClr":
				inSrgbClr = true
				lastColor = nil
				if inFontRef {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							fontRefColor = &c
							lastColor = fontRefColor
						}
					}
				} else if inLnSolidFill && currentLine != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							currentLine.lineColor = NewColor("FF" + attr.Value)
							lastColor = &currentLine.lineColor
						}
					}
				} else if inDefSolidFill && lstStyleFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							lstStyleFont.Color = NewColor("FF" + attr.Value)
							lastColor = &lstStyleFont.Color
						}
					}
				} else if inPPrDefSolidFill && defFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							defFont.Color = NewColor("FF" + attr.Value)
							lastColor = &defFont.Color
						}
					}
				} else if inRunSolidFill && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							currentFont.Color = NewColor("FF" + attr.Value)
							lastColor = &currentFont.Color
						}
					}
				}
			case "prstClr":
				inSrgbClr = true
				lastColor = nil
				var prstName string
				for _, attr := range t.Attr {
					if attr.Name.Local == "val" {
						prstName = attr.Value
					}
				}
				c := presetColorToColor(prstName)
				if inFontRef {
					fontRefColor = &c
					lastColor = fontRefColor
				} else if inLnSolidFill && currentLine != nil {
					currentLine.lineColor = c
					lastColor = &currentLine.lineColor
				} else if inDefSolidFill && lstStyleFont != nil {
					lstStyleFont.Color = c
					lastColor = &lstStyleFont.Color
				} else if inPPrDefSolidFill && defFont != nil {
					defFont.Color = c
					lastColor = &defFont.Color
				} else if inRunSolidFill && currentFont != nil {
					currentFont.Color = c
					lastColor = &currentFont.Color
				}
			case "schemeClr":
				inSrgbClr = true
				lastColor = nil
				if pres != nil && pres.themeColors != nil {
					var schemeName string
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							schemeName = attr.Value
						}
					}
					if argb, ok := pres.themeColors[schemeName]; ok && argb != "" {
						c := NewColor(argb)
						if inFontRef {
							fontRefColor = &c
							lastColor = fontRefColor
						} else if inLnSolidFill && currentLine != nil {
							currentLine.lineColor = c
							lastColor = &currentLine.lineColor
						} else if inDefSolidFill && lstStyleFont != nil {
							lstStyleFont.Color = c
							lastColor = &lstStyleFont.Color
						} else if inPPrDefSolidFill && defFont != nil {
							defFont.Color = c
							lastColor = &defFont.Color
						} else if inRunSolidFill && currentFont != nil {
							currentFont.Color = c
							lastColor = &currentFont.Color
						}
					}
				}
			case "sysClr":
				// <a:sysClr val="window" lastClr="FFFFFF"/> — system color
				inSrgbClr = true
				lastColor = nil
				var sysLastClr string
				for _, attr := range t.Attr {
					if attr.Name.Local == "lastClr" {
						sysLastClr = attr.Value
					}
				}
				if sysLastClr != "" {
					c := NewColor("FF" + sysLastClr)
					if inFontRef {
						fontRefColor = &c
						lastColor = fontRefColor
					} else if inLnSolidFill && currentLine != nil {
						currentLine.lineColor = c
						lastColor = &currentLine.lineColor
					} else if inDefSolidFill && lstStyleFont != nil {
						lstStyleFont.Color = c
						lastColor = &lstStyleFont.Color
					} else if inPPrDefSolidFill && defFont != nil {
						defFont.Color = c
						lastColor = &defFont.Color
					} else if inRunSolidFill && currentFont != nil {
						currentFont.Color = c
						lastColor = &currentFont.Color
					}
				}
			case "alpha":
				if inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								// For text run properties, skip val="0" to avoid
								// making text invisible. For line/fill contexts,
								// val="0" genuinely means fully transparent.
								if v <= 0 && (inRunProps || inDefRPr) {
									continue
								}
								alpha := uint8(v * 255 / 100000)
								alphaHex := fmt.Sprintf("%02X", alpha)
								lastColor.ARGB = alphaHex + lastColor.ARGB[2:]
							}
						}
					}
				}
			case "latin":
				if inDefRPr && lstStyleFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							lstStyleFont.Name = attr.Value
						}
					}
				} else if inPPrDefRPr && defFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							defFont.Name = attr.Value
						}
					}
				} else if inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							currentFont.Name = attr.Value
						}
					}
				}
			case "ea":
				if inDefRPr && lstStyleFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							lstStyleFont.NameEA = attr.Value
						}
					}
				} else if inPPrDefRPr && defFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							defFont.NameEA = attr.Value
						}
					}
				} else if inRunProps && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && !strings.HasPrefix(attr.Value, "+") {
							currentFont.NameEA = attr.Value
						}
					}
				}
			case "p":
				if inTxBody && currentRichText != nil {
					inParagraph = true
					currentParagraph = NewParagraph()
					currentRichText.paragraphs = append(currentRichText.paragraphs, currentParagraph)
				}
			case "pPr":
				if inParagraph {
					inPPr = true
					for _, attr := range t.Attr {
						if attr.Name.Local == "algn" && currentParagraph != nil {
							currentParagraph.alignment.Horizontal = HorizontalAlignment(attr.Value)
						}
					}
				}
			case "r":
				if inParagraph {
					inRun = true
					currentFont = NewFont()
					currentFont.Size = 18
					// Apply fontRef color from <p:style> as base default
					if fontRefColor != nil {
						currentFont.Color = *fontRefColor
					}
					// Apply lstStyle defaults
					if lstStyleFont != nil {
						if lstStyleFont.Size > 0 {
							currentFont.Size = lstStyleFont.Size
						}
						if lstStyleFont.Bold {
							currentFont.Bold = true
						}
						if lstStyleFont.Italic {
							currentFont.Italic = true
						}
						if lstStyleFont.Name != "Calibri" && lstStyleFont.Name != "" {
							currentFont.Name = lstStyleFont.Name
						}
						if lstStyleFont.NameEA != "" {
							currentFont.NameEA = lstStyleFont.NameEA
						}
						if lstStyleFont.Color.ARGB != "FF000000" && lstStyleFont.Color.ARGB != "" {
							currentFont.Color = lstStyleFont.Color
						}
					}
					// Apply pPr defRPr defaults
					if defFont != nil {
						if defFont.Size > 0 {
							currentFont.Size = defFont.Size
						}
						if defFont.Bold {
							currentFont.Bold = true
						}
						if defFont.Italic {
							currentFont.Italic = true
						}
						if defFont.Name != "Calibri" && defFont.Name != "" {
							currentFont.Name = defFont.Name
						}
						if defFont.NameEA != "" {
							currentFont.NameEA = defFont.NameEA
						}
						if defFont.Color.ARGB != "FF000000" && defFont.Color.ARGB != "" {
							currentFont.Color = defFont.Color
						}
					}
				}
			case "rPr":
				if inRun {
					inRunProps = true
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
						}
					}
				}
			case "t":
				if inRun {
					inText = true
				}
			case "style":
				// <p:style> inside <p:sp>
				if inSp && !inSpPr && !inTxBody {
					inStyle = true
				}
			case "fontRef":
				if inStyle {
					inFontRef = true
				}
			}

		case xml.CharData:
			if inText && currentParagraph != nil && currentFont != nil {
				text := string(t)
				tr := currentParagraph.CreateTextRun(text)
				tr.font = currentFont
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "pic":
				if inPic && embedID != "" {
					for _, rel := range rels {
						if rel.ID == embedID {
							imgPath := rel.Target
							if !strings.HasPrefix(imgPath, "ppt/") {
								dir := strings.TrimSuffix(layoutPath, "/"+lastPathComponent(layoutPath))
								imgPath = resolveRelativePath(dir, imgPath)
							}
							imgData, err := readFileFromZip(zr, imgPath)
							if err == nil {
								ds := NewDrawingShape()
								ds.offsetX = offX
								ds.offsetY = offY
								ds.width = extCX
								ds.height = extCY
								ds.data = imgData
								ds.mimeType = guessMimeType(imgPath)
								ds.alpha = picAlpha
								ds.cropLeft = cropL
								ds.cropTop = cropT
								ds.cropRight = cropR
								ds.cropBottom = cropB
								shapes = append(shapes, ds)
							}
							break
						}
					}
				}
				inPic = false
				inSpPr = false
			case "sp":
				if inSp && !isPH && currentRichText != nil && len(currentRichText.paragraphs) > 0 {
					// Check if there's actual text content
					hasText := false
					for _, p := range currentRichText.paragraphs {
						for _, elem := range p.elements {
							if tr, ok := elem.(*TextRun); ok && tr.text != "" {
								hasText = true
								break
							}
						}
						if hasText {
							break
						}
					}
					if hasText {
						currentRichText.offsetX = offX
						currentRichText.offsetY = offY
						currentRichText.width = extCX
						currentRichText.height = extCY
						currentRichText.textAnchor = textAnchor
						shapes = append(shapes, currentRichText)
					}
				}
				inSp = false
				inSpPr = false
				inTxBody = false
				inLstStyle = false
				inLstStyleLvl1 = false
				inDefRPr = false
				inDefSolidFill = false
				currentRichText = nil
				lstStyleFont = nil
				defFont = nil
			case "spPr":
				inSpPr = false
				inLn = false
			case "nvSpPr", "nvCxnSpPr":
				inNvSpPr = false
			case "nvPr":
				inNvPr = false
			case "txBody":
				inTxBody = false
				inLstStyle = false
				inLstStyleLvl1 = false
			case "lstStyle":
				inLstStyle = false
				inLstStyleLvl1 = false
			case "lvl1pPr":
				inLstStyleLvl1 = false
			case "defRPr":
				inDefRPr = false
				inDefSolidFill = false
				inPPrDefRPr = false
				inPPrDefSolidFill = false
			case "solidFill":
				inDefSolidFill = false
				inPPrDefSolidFill = false
				inRunSolidFill = false
				inLnSolidFill = false
				inSrgbClr = false
				lastColor = nil
			case "srgbClr", "schemeClr":
				inSrgbClr = false
			case "style":
				inStyle = false
				inFontRef = false
			case "fontRef":
				inFontRef = false
			case "cxnSp":
				if inCxnSp && currentLine != nil {
					currentLine.offsetX = offX
					currentLine.offsetY = offY
					currentLine.width = extCX
					currentLine.height = extCY
					currentLine.flipHorizontal = flipH
					currentLine.flipVertical = flipV
					if currentLine.lineWidth == 0 {
						currentLine.lineWidth = 1
					}
					shapes = append(shapes, currentLine)
				}
				inCxnSp = false
				inSpPr = false
				inLn = false
				currentLine = nil
			case "ln":
				inLn = false
				inLnSolidFill = false
			case "p":
				inParagraph = false
				currentParagraph = nil
				defFont = nil
			case "pPr":
				inPPr = false
				inPPrDefRPr = false
				inPPrDefSolidFill = false
			case "r":
				inRun = false
				currentFont = nil
			case "rPr":
				inRunProps = false
				inRunSolidFill = false
			case "t":
				inText = false
			}
		}
	}

	return shapes
}
