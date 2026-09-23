package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
	"time"
)

func (r *PPTXReader) readSlide(zr *zip.Reader, path string, pres *Presentation, commentAuthors map[int]*CommentAuthor) (*Slide, error) {
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
	r.readSlideComments(zr, slide, slideRels, path, commentAuthors)

	// Read notes if relationship exists
	r.readSlideNotes(zr, slide, slideRels, path)

	// Read ink annotations last so they render on top, as PowerPoint draws them.
	r.readSlideInk(zr, slide, slideRels, path)

	return slide, nil
}

func (r *PPTXReader) readSlideComments(zr *zip.Reader, slide *Slide, rels []xmlRelForRead, slidePath string, authors map[int]*CommentAuthor) {
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
			r.parseCommentsXML(data, slide, authors)
		}
	}
}

// commentAuthorsPath is where the package-level comment author table lives.
//
// ECMA-376 fixes the part name, so it is read directly rather than through a
// relationship: PowerPoint writes it here and does not reference it from
// presentation.xml.rels, and the writer in this package emits the same path.
const commentAuthorsPath = "ppt/commentAuthors.xml"

// readCommentAuthors reads the author table keyed by author id.
//
// It is needed because <p:cm> carries only an authorId — the name, initials and
// colour index live in this separate part. Without it a comment can only record
// a numeric id, and since the writer groups comments into authors by name, every
// author would collapse into one blank-named entry on the next save.
//
// A missing or unreadable part yields a nil map; comments then keep a bare id,
// which is what the reader did before this part was read at all.
func (r *PPTXReader) readCommentAuthors(zr *zip.Reader) map[int]*CommentAuthor {
	data, err := readFileFromZip(zr, commentAuthorsPath)
	if err != nil {
		return nil
	}

	authors := make(map[int]*CommentAuthor)
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "cmAuthor" {
			continue
		}

		author := &CommentAuthor{}
		hasID := false
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "id":
				if v, err := strconv.Atoi(attr.Value); err == nil {
					author.ID = v
					hasID = true
				}
			case "name":
				author.Name = attr.Value
			case "initials":
				author.Initials = attr.Value
			case "clrIdx":
				if v, err := strconv.Atoi(attr.Value); err == nil {
					author.ColorIdx = v
				}
			}
		}
		if hasID {
			authors[author.ID] = author
		}
	}
	return authors
}

// commentDateLayouts lists the timestamp shapes a comment may carry, most
// specific first. This writer emits milliseconds with no zone and treats the
// value as UTC; other producers may add a zone or omit the fraction.
var commentDateLayouts = []string{
	"2006-01-02T15:04:05.000",
	"2006-01-02T15:04:05",
	time.RFC3339Nano,
	time.RFC3339,
}

func parseCommentDate(value string) (time.Time, bool) {
	for _, layout := range commentDateLayouts {
		if t, err := time.Parse(layout, value); err == nil {
			return t.UTC(), true
		}
	}
	return time.Time{}, false
}

// parseCommentsXML reads one slide's comment list.
//
// <p:cm> carries the author id and timestamp, while the author's name lives in a
// separate package part, so authors is the lookup filled in by
// readCommentAuthors.
//
// Comment text comes from the <a:t> runs inside the <p:text> body — the shape
// the schema requires (p:text is a CT_TextBody) and the one PowerPoint writes.
// A body with no runs, such as the bare <p:text>text</p:text> this library used
// to emit, is still read through the fallback below. Reading runs rather than
// any character data inside <p:text> also matters because these parts are
// normally indented: the whitespace that precedes </p:text> would otherwise be
// taken as the comment text.
func (r *PPTXReader) parseCommentsXML(data []byte, slide *Slide, authors map[int]*CommentAuthor) {
	decoder := xml.NewDecoder(bytes.NewReader(data))

	var (
		current  *Comment
		inText   bool
		inPara   bool
		inRun    bool
		runs     []string // runs of the paragraph being read
		paras    []string // completed paragraphs
		fallback strings.Builder
	)

	resetBody := func() {
		inText = false
		inPara = false
		inRun = false
		runs = runs[:0]
		paras = paras[:0]
		fallback.Reset()
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cm":
				current = NewComment()
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "authorId":
						if v, err := strconv.Atoi(attr.Value); err == nil {
							if a, ok := authors[v]; ok {
								current.Author = a
							} else {
								current.Author = &CommentAuthor{ID: v}
							}
						}
					case "dt":
						if d, ok := parseCommentDate(attr.Value); ok {
							current.Date = d
						}
					}
				}
			case "pos":
				if current != nil {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "x":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								current.PositionX = v
							}
						case "y":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								current.PositionY = v
							}
						}
					}
				}
			case "text":
				if current != nil {
					resetBody()
					inText = true
				}
			case "p":
				if inText {
					runs = runs[:0]
					inPara = true
				}
			case "t":
				if inText && inPara {
					inRun = true
				}
			}
		case xml.CharData:
			switch {
			case inRun:
				runs = append(runs, string(t))
			case inText:
				// Character data directly inside <p:text>.
				fallback.WriteString(string(t))
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "cm":
				if current != nil {
					slide.comments = append(slide.comments, current)
					current = nil
				}
			case "text":
				if current != nil {
					if len(paras) > 0 {
						current.Text = strings.Join(paras, "\n")
					} else {
						current.Text = fallback.String()
					}
				}
				resetBody()
			case "p":
				if inText && inPara {
					paras = append(paras, strings.Join(runs, ""))
					runs = runs[:0]
					inPara = false
				}
			case "t":
				inRun = false
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

// setCellBorderStyle applies a line style to one side of a cell's border.
//
// A side of a cell border is not addressed by anything the caller reads out of
// the file: it is named by the <a:lnL>/<a:lnR>/<a:lnT>/<a:lnB> element the
// value sits inside, which the scanner records as a side marker. Nothing
// outside the scanner can reach a single side, so the mapping lives here.
// dedupeGradStops collapses the raw gs list into ordered stops. Repeated
// positions keep the LAST colour written there (PowerPoint's reading); the
// result is sorted by position.
func dedupeGradStops(colors []Color, positions []int) []GradStop {
	var out []GradStop
	index := make(map[int]int)
	for i, c := range colors {
		if idx, ok := index[positions[i]]; ok {
			out[idx].Color = c
		} else {
			index[positions[i]] = len(out)
			out = append(out, GradStop{Pos: positions[i], Color: c})
		}
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Pos < out[j].Pos })
	return out
}

func setCellBorderStyle(cell *TableCell, side string, style BorderStyle) {
	if cell == nil || cell.border == nil {
		return
	}
	var b *Border
	switch side {
	case "L":
		b = cell.border.Left
	case "R":
		b = cell.border.Right
	case "T":
		b = cell.border.Top
	case "B":
		b = cell.border.Bottom
	}
	if b != nil {
		b.Style = style
	}
}

// applyPPrAttrs folds the attributes of an <a:pPr> into a paragraph.
//
// Two slide readers walk this markup — parseSlideXML for a slide and
// parseLayoutImages for the text shapes a layout contributes — and both have to
// understand the same attribute set. The layout reader knew only algn, so a text
// shape on a layout kept its alignment and quietly lost its hanging indent, its
// margins and its outline level. The renderer indents by marL/marR/indent, so
// the preview drew such a paragraph flush against the shape's left edge.
//
// lvl is read here for the first time: the writer has always emitted it, and
// nothing consumed it, so a paragraph's outline level could not survive a load.
func applyPPrAttrs(p *Paragraph, attrs []xml.Attr) {
	if p == nil {
		return
	}
	if p.alignment == nil {
		// The attribute block below needs somewhere to put the values, and the
		// renderer treats a missing Alignment the same way it treats a default
		// one, so building it here keeps a hand-built paragraph readable.
		//
		// It is deliberately left zero rather than NewAlignment(), whose
		// Horizontal is HorizontalLeft. Stating "left" for a paragraph whose
		// markup stated nothing is a value the file does not contain: it makes
		// the writer emit algn="l" on save, and it hid the master's algn="ctr"
		// from the inheritance ladder, because a title that had already been
		// told "left" looks like one that was told on purpose.
		p.alignment = &Alignment{}
	}
	for _, attr := range attrs {
		switch attr.Name.Local {
		case "algn":
			p.alignment.Horizontal = HorizontalAlignment(attr.Value)
		case "marL":
			if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
				p.alignment.MarginLeft = v
				p.alignment.marLSet = true
			}
		case "marR":
			if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
				p.alignment.MarginRight = v
			}
		case "indent":
			if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
				p.alignment.Indent = v
				p.alignment.indentSet = true
			}
		case "lvl":
			if v, err := strconv.Atoi(attr.Value); err == nil {
				p.alignment.Level = v
			}
		}
	}
}

// parseNotesXML reads the notes text out of a notes slide part.
//
// The text of a note is a text body like any other: one <a:p> per paragraph,
// each holding runs. Joining the runs and dropping the paragraph boundaries
// loses the line breaks, which is most of what a note is — PowerPoint writes a
// note the user typed across three lines as three paragraphs, and reading it
// back as one run of text cannot be undone on the next save. The comments
// parser alongside this one has always joined its paragraphs with "\n".
//
// <a:br/> is a line break inside a paragraph. The model holds notes as a single
// string, so a newline is the only way to carry one.
func (r *PPTXReader) parseNotesXML(data []byte) string {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var inBody, inParagraph, inRun, inText bool
	var runs []string  // runs of the paragraph being read
	var paras []string // completed paragraphs
	// bodyStart and bodyHasRun scope the paragraphs to one text body. A notes
	// slide can carry a placeholder for the slide number, the date or the
	// footer, each with a body of its own whose "text" is a field, and none of
	// that is what the speaker typed.
	var bodyStart int
	var bodyHasRun bool

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
				bodyStart = len(paras)
				bodyHasRun = false
			case "p":
				if inBody {
					runs = runs[:0]
					inParagraph = true
				}
			case "r":
				// Gated on the paragraph so a slide number or date field, which
				// holds its <a:t> directly without a run, is not read as notes.
				inRun = inParagraph
			case "br":
				if inParagraph {
					runs = append(runs, "\n")
				}
			case "t":
				if inRun {
					inText = true
				}
			}
		case xml.CharData:
			if inText {
				runs = append(runs, string(t))
				bodyHasRun = true
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "txBody":
				inBody = false
				if !bodyHasRun {
					paras = paras[:bodyStart]
				}
			case "p":
				if inBody && inParagraph {
					paras = append(paras, strings.Join(runs, ""))
					runs = runs[:0]
					inParagraph = false
				}
			case "r":
				inRun = false
			case "t":
				inText = false
			}
		}
	}
	return strings.Join(paras, "\n")
}

// maxGroupDepth caps how deeply the reader will nest groups inside a slide.
//
// The cap exists because nesting depth is chosen entirely by the file, while
// the consumers of the resulting tree recurse over it: the renderer descends
// through renderGroup, and the writer through writeGroupShapeXML. A few
// megabytes of <p:grpSp> elements can therefore nest deeply enough to exhaust
// the goroutine stack, and a stack overflow is a fatal error that recover
// cannot turn into a *PanicError — the very failure mode the fail-closed
// boundary exists to prevent.
//
// Past the cap the subtree is still parsed, because the XML token stream has to
// stay balanced, but into a group that is never attached to the slide, so it is
// unreachable and gets collected. Real decks nest a handful of levels deep; 64
// is far beyond anything PowerPoint produces, so the cap only ever trips on
// input that is malformed or hostile.
const maxGroupDepth = 64

// readTransition consumes <p:transition> from its start element through its
// matching end tag and returns what it describes, or nil when it names no effect
// this package models — an empty element, or one of the PowerPoint 2010 effects
// that live in the p14 namespace rather than in CT_SlideTransition.
//
// It reads its own subtree instead of going through the slide parser's switch
// because the effect is a child element and there are twenty-one possible
// children. Attributes are kept as the document states them; only the ones the
// schema gives a default are left nil, so a caller can tell "unstated" from
// "stated as the default".
func (r *PPTXReader) readTransition(decoder *xml.Decoder, start xml.StartElement) *Transition {
	tr := &Transition{}
	for _, a := range start.Attr {
		switch {
		case a.Name.Space == "" && a.Name.Local == "spd":
			tr.Speed = TransitionSpeed(a.Value)
		case a.Name.Space == "" && a.Name.Local == "advClick":
			advClick := a.Value == "1" || a.Value == "true"
			tr.AdvanceOnClick = &advClick
		case a.Name.Space == "" && a.Name.Local == "advTm":
			if n, err := strconv.Atoi(a.Value); err == nil {
				tr.AdvanceAfterTime = n
			}
		case a.Name.Space == nsPowerPoint2010 && a.Name.Local == "dur":
			if n, err := strconv.Atoi(a.Value); err == nil {
				tr.Duration = n
			}
		}
	}

	depth := 0
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch tok := token.(type) {
		case xml.StartElement:
			if depth == 0 {
				if typ, ok := transitionTypeForElement(tok.Name.Local); ok {
					tr.Type = typ
					for _, a := range tok.Attr {
						switch a.Name.Local {
						case "dir":
							// One field holds both meanings: horz/vert for
							// blinds, checker, comb and randomBar, a side or
							// corner for the rest, in/out for split and zoom.
							tr.Direction = TransitionDirection(a.Value)
						case "orient":
							tr.Orientation = TransitionOrientation(a.Value)
						case "thruBlk":
							tr.ThroughBlack = a.Value == "1" || a.Value == "true"
						case "spokes":
							if n, err := strconv.Atoi(a.Value); err == nil {
								tr.Spokes = n
							}
						}
					}
				}
			}
			depth++
		case xml.EndElement:
			if depth == 0 && tok.Name.Local == start.Name.Local {
				if tr.Type == TransitionNone {
					return nil
				}
				return tr
			}
			depth--
		}
	}
	if tr.Type == TransitionNone {
		return nil
	}
	return tr
}

func (r *PPTXReader) parseSlideXML(decoder *xml.Decoder, slide *Slide, rels []xmlRelForRead, zr *zip.Reader, slidePath string, pres *Presentation) error {
	type parseState struct {
		inSpTree        bool
		inSp            bool
		inPic           bool
		inCxnSp         bool
		inGraphicFrame  bool
		inGrpSp         bool
		inTxBody        bool
		inParagraph     bool
		inRun           bool
		inBr            bool          // inside <a:br> — its rPr sizes the blank line
		brSize          int           // the <a:br>'s own rPr sz, hundredths of a point
		curBreak        *BreakElement // the break currently open
		inFld           bool
		inRunProps      bool
		inText          bool
		inTbl           bool
		inTr            bool
		inTc            bool
		inTcTxBody      bool
		inTcParagraph   bool
		inTcRun         bool
		inTcText        bool
		inTcPr          bool
		inTcPrSolidFill bool
		inTcPrLn        bool
		tcPrLnSide      string // "L", "R", "T", "B" or "" for generic
		inNvSpPr        bool
		inSolidFill     bool
		inSpPr          bool
		inLn            bool
		inDuotone       bool // inside <a:duotone> under a pic's <a:blip>
		inPPr           bool
		inBg            bool
		inBgPr          bool
		inBgSolidFill   bool
		inBuClr         bool

		// Markup compatibility. A slide transition with a duration is wrapped
		// in mc:AlternateContent, whose mc:Choice carries the p14:dur attribute
		// and whose mc:Fallback repeats the transition without it. The Choice
		// copy is the one to keep, so the reader has to know which branch it is
		// inside.
		inAltContent  bool
		inAltChoice   bool
		inAltFallback bool

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
		inGradFill         bool
		inGsLst            bool
		inGs               bool
		gradFillPos        int  // current gs position (0-100000)
		inRunPropsGradFill bool // gradFill inside rPr (text color gradient)

		// avLst tracking (adjustment values for preset geometry)
		inAvLst bool

		// custGeom tracking
		inCustGeom bool
		inPathLst  bool
		inCustPath bool

		// effectLst / outerShdw tracking
		inEffectLst bool
		inOuterShdw bool

		// <a:sp3d> inside a connector's spPr: the 3D frame that gives the
		// line its bevel highlight (coolSlant on the flow diagrams).
		inSp3d bool

		// <a:clrChange> inside a picture's <a:blip>: a per-pixel colour
		// replacement the composited photo depends on.
		inClrChange bool
		inClrFrom   bool
		inClrTo     bool
		clrFromHex  string // 6 hex digits, the colour being replaced
		clrToHex    string // 6 hex digits, the replacement colour
		clrToAlpha  int    // 1/1000 of a percent; -1 when none declared

		// <p:style> fill/line references: the theme styling a shape falls
		// back to when its own <p:spPr> declares neither.
		inFillRef       bool
		inLnRef         bool
		styleFillScheme string // scheme colour name, e.g. "accent1"
		styleLnScheme   string
		styleFillIdx    int             // the fillRef's idx (0 = no theme fill)
		styleLnIdx      int             // the lnRef's idx (picks the line weight)
		styleEffectIdx  int             // the effectRef's idx (0 = no theme effect)
		styleFillOps    []themeColorOp  // transforms on the fillRef's schemeClr
		styleLnOps      []themeColorOp  // transforms on the lnRef's schemeClr
		styleOpsTarget  *[]themeColorOp // which of the two the current colour feeds
		inStyleScheme   bool            // collecting those transforms

		// <a:tableStyleId> character data
		inTableStyleID bool
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
	// currentHyperlink is the <a:hlinkClick> of the run being read. It belongs
	// to the run whose <a:rPr> carries it, and is cleared when that run starts.
	var currentHyperlink *Hyperlink
	// currentFldType is the type of the <a:fld> being read ("" outside one).
	// It is stamped onto the TextRun the field's <a:t> produces.
	var currentFldType string
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
	var shapeHidden bool
	var flipH, flipV bool
	var shapeRotation int
	var prstGeom string

	// Chart graphicFrame tracking. A chart is a graphicFrame whose
	// graphicData references a chart part through a relationship id.
	var chartRelID string
	var graphicDataIsChart bool
	// graphicDataURI is the raw a:graphicData uri of the graphicFrame being
	// read. It classifies the frame when it turns out to be neither a table nor
	// a chart, so the stand-in can name what it replaced.
	var graphicDataURI string
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
	// Gradient path geometry: <a:path path=...>, <a:fillToRect> and
	// <a:tileRect> insets (l, t, r, b) in 0..100000 gradient-box units.
	var gradPathKind string
	var gradFillTo [4]int
	var gradTileTo [4]int

	// Deferred shape-level border (spPr ln comes before txBody)
	var pendingBorder *Border
	var pendingHeadEnd *LineEnd
	var pendingTailEnd *LineEnd

	// lineColorExplicit records that the connector's <a:ln> declared its own
	// colour; only then does the <p:style> lnRef stay out of the way.
	var lineColorExplicit bool

	// Deferred adjustment values from avLst
	var pendingAdjustValues map[string]int

	// Deferred shadow (spPr effectLst outerShdw)
	var pendingShadow *Shadow

	// duotoneIdx is the slot (0 or 1) the next <a:duotone> colour lands in.
	// The element is a bare sequence of colour choices; the parser only knows
	// which one it is reading by counting.
	var duotoneIdx int

	// Deferred blipFill image data (spPr blipFill for shapes)
	var pendingBlipFillData []byte
	var pendingBlipFillMime string
	var pendingBlipFillLumBright, pendingBlipFillLumContrast int

	// Background blipFill image data (bgPr blipFill). The picture becomes a
	// full-slide DrawingShape prepended to the shape list; its crop and opacity
	// have to be collected here because they are siblings of the <a:blip>
	// rather than children of a <p:pic>, which is what the picture cases below
	// are keyed on.
	var bgBlipFillData []byte
	var bgBlipFillMime string
	var bgCropLeft, bgCropTop, bgCropRight, bgCropBottom int
	var bgAlpha int
	var bgLumBright, bgLumContrast int

	// Group shape nesting
	grpDepth := 0

	// Stack for nested groups
	type grpSaved struct {
		group    *GroupShape
		name     string
		descr    string
		hidden   bool
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
		// detached marks a group nested deeper than maxGroupDepth. Its subtree
		// is still parsed but is never attached to the slide; see maxGroupDepth.
		detached bool
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
			case "AlternateContent":
				if !state.inSpTree && !state.inGrpSp {
					state.inAltContent = true
					state.inAltChoice = false
					state.inAltFallback = false
				}
			case "Choice":
				if state.inAltContent {
					state.inAltChoice = true
				}
			case "Fallback":
				if state.inAltContent {
					state.inAltFallback = true
				}
			case "transition":
				// <p:transition> is a direct child of <p:sld>. It is read as a
				// whole subtree rather than through this switch, because the
				// effect it names is one of its children and there are twenty-one
				// of them.
				if !state.inSpTree && !state.inGrpSp {
					if tr := r.readTransition(decoder, t); tr != nil && !state.inAltFallback {
						slide.transition = tr
					} else if tr != nil && slide.transition == nil {
						// An mc:Fallback with no mc:Choice before it is all the
						// document has; take it.
						slide.transition = tr
					}
				}
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
					grpStack = append(grpStack, &grpSaved{group: newGroup, detached: grpDepth > maxGroupDepth})
					currentGroup = newGroup
					offX, offY, extCX, extCY = 0, 0, 0, 0
					chOffX, chOffY, chExtCX, chExtCY = 0, 0, 0, 0
					shapeName = ""
					shapeHidden = false
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
					shapeHidden = false
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
					shapeHidden = false
					shapeDescr = ""
					prstGeom = ""
					shapeRotation = 0
					state.inDuotone = false
					duotoneIdx = 0
					pendingBorder = nil
					pendingShadow = nil
				}
			case "duotone":
				// <a:duotone> maps every pixel onto the line between two
				// colours at its luma. Only the picture context is kept: a
				// duotone on a shape's blipFill or a background is rarer and
				// has no model to land in yet.
				if state.inPic && currentDrawing != nil {
					state.inDuotone = true
					duotoneIdx = 0
				}
			case "cxnSp":
				if state.inSpTree || state.inGrpSp {
					state.inCxnSp = true
					currentLine = NewLineShape()
					offX, offY, extCX, extCY = 0, 0, 0, 0
					shapeName = ""
					shapeHidden = false
					prstGeom = ""
					shapeRotation = 0
					pendingCustomPath = nil
					lineColorExplicit = false
					state.inFillRef = false
					state.inLnRef = false
					state.styleFillScheme = ""
					state.styleLnScheme = ""
				}
			case "graphicFrame":
				if state.inSpTree {
					state.inGraphicFrame = true
					offX, offY, extCX, extCY = 0, 0, 0, 0
					shapeName = ""
					shapeHidden = false
					prstGeom = ""
					shapeRotation = 0
					chartRelID = ""
					graphicDataIsChart = false
					graphicDataURI = ""
				}
			case "graphicData":
				if state.inGraphicFrame {
					for _, attr := range t.Attr {
						if attr.Name.Local == "uri" {
							graphicDataURI = attr.Value
							graphicDataIsChart = strings.Contains(attr.Value, "/chart")
						}
					}
				}
			case "chart":
				// <c:chart r:id="rIdN"/> inside a chart graphicData.
				if state.inGraphicFrame && graphicDataIsChart {
					for _, attr := range t.Attr {
						if attr.Name.Local == "id" {
							chartRelID = attr.Value
						}
					}
				}
			case "tbl":
				if state.inGraphicFrame {
					state.inTbl = true
					currentTable = NewTableShape(0, 0)
					// NewTableShape defaults firstRow/bandRow on for
					// API-created tables; a file that omits a flag means
					// false (xsd:boolean has no default). slide21's six
					// bandRow-only tables must not pick up the first-row
					// dark header.
					currentTable.firstRow = false
					currentTable.lastRow = false
					currentTable.firstCol = false
					currentTable.lastCol = false
					currentTable.bandRow = false
					currentTable.bandCol = false
					currentTable.rows = nil
					currentTableRow = -1
				}
			case "tblPr":
				if state.inTbl && currentTable != nil {
					for _, attr := range t.Attr {
						v := attr.Value == "1" || attr.Value == "true"
						switch attr.Name.Local {
						case "firstRow":
							currentTable.firstRow = v
						case "lastRow":
							currentTable.lastRow = v
						case "firstCol":
							currentTable.firstCol = v
						case "lastCol":
							currentTable.lastCol = v
						case "bandRow":
							currentTable.bandRow = v
						case "bandCol":
							currentTable.bandCol = v
						}
					}
				}
			case "tableStyleId":
				if state.inTbl && currentTable != nil {
					state.inTableStyleID = true
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
					// The cell's insets and vertical anchor live on the tag
					// itself, so they have to be read here and not from a
					// child: they decide the width a cell's text may wrap
					// into, and a missing default re-wraps the whole table.
					cellRef := currentTable.rows[currentTableRow][currentTableCol]
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "marL":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cellRef.marginL = v
							}
						case "marR":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cellRef.marginR = v
							}
						case "marT":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cellRef.marginT = v
							}
						case "marB":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cellRef.marginB = v
							}
						case "anchor":
							switch attr.Value {
							case "t", "ctr", "b":
								cellRef.anchor = attr.Value
							}
						}
					}
				}
			case "lnL":
				if state.inTcPr {
					state.inTcPrLn = true
					state.tcPrLnSide = "L"
					if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
						currentTable.rows[currentTableRow][currentTableCol].border.leftDeclared = true
					}
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
					if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
						currentTable.rows[currentTableRow][currentTableCol].border.rightDeclared = true
					}
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
					if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
						currentTable.rows[currentTableRow][currentTableCol].border.topDeclared = true
					}
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
					if currentTable != nil && currentTableRow >= 0 && currentTableCol >= 0 &&
						currentTableRow < len(currentTable.rows) && currentTableCol < len(currentTable.rows[currentTableRow]) {
						currentTable.rows[currentTableRow][currentTableCol].border.bottomDeclared = true
					}
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
						case "hidden":
							// <p:cNvPr hidden="1"> — PowerPoint keeps the shape
							// in the file but never draws it. true/1 both mean
							// hidden per the XML boolean rules.
							shapeHidden = attr.Value == "1" || attr.Value == "true"
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
							// The anchor belongs to the shape, not only to the
							// local that the RichTextShape branch copies out.
							// A placeholder never read it back, so its text was
							// re-anchored to the top of the shape on save.
							textAnchor = TextAnchorType(attr.Value)
							if currentRichText != nil {
								currentRichText.textAnchor = textAnchor
							}
							if currentPlaceholder != nil {
								currentPlaceholder.textAnchor = textAnchor
							}
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
				// <a:normAutofit fontScale="62500" lnSpcReduction="10000"/> inside <a:bodyPr>
				if state.inTxBody {
					fontScaleVal := 100000 // default 100%
					lnSpcReductionVal := 0 // raw file value in thousandths of a percent, 0 = no reduction
					for _, attr := range t.Attr {
						if attr.Name.Local == "fontScale" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								fontScaleVal = v
							}
						}
						if attr.Name.Local == "lnSpcReduction" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								lnSpcReductionVal = v
							}
						}
					}
					if state.isPlaceholder && currentPlaceholder != nil {
						currentPlaceholder.autoFit = AutoFitNormal
						currentPlaceholder.fontScale = fontScaleVal
						currentPlaceholder.lnSpcReduction = lnSpcReductionVal
					} else if currentRichText != nil {
						currentRichText.autoFit = AutoFitNormal
						currentRichText.fontScale = fontScaleVal
						currentRichText.lnSpcReduction = lnSpcReductionVal
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
					applyPPrAttrs(currentParagraph, t.Attr)
				}
			case "endParaRPr":
				// A runless paragraph still ends with run properties, and its
				// sz is the only size stated anywhere in it — PowerPoint draws
				// the empty line box at that height and resolves its inherited
				// space percentage against it. Self-closing in practice, but
				// the attribute is read on Start either way.
				if (state.inParagraph || state.inTcParagraph) && currentParagraph != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "sz" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentParagraph.endParaRPrSize = v
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
					// A bullet font is a typeface like any other, theme
					// reference included: PowerPoint writes "+mj-lt" here as
					// readily as it does on a run.
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentParagraph.bullet.Font = n
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
								} else if state.inSpcBef {
									// A declared percentage rides the paragraph
									// raw; the renderer resolves it against the
									// paragraph's own font size, because that
									// size is not known until the runs close.
									currentParagraph.spaceBeforePct = v
								}
								// spcPct inside spcAft with val=0 means no spacing;
								// non-zero percentage spacing after is rare and
								// would need line-height-relative calculation at render time.
							}
						}
					}
				}
			case "r", "fld":
				// A hyperlink belongs to a single run. Clearing it here keeps the
				// previous run's link from following this one, and leaves a run
				// whose <a:rPr> has no hlinkClick linkless rather than inheriting.
				currentHyperlink = nil
				if t.Name.Local == "fld" {
					// <a:fld> wraps pPr/rPr/t exactly like a run (CT_TextField),
					// so it rides the run machinery; only its type is kept, so
					// the renderer can re-evaluate the field per slide and the
					// writer can emit a field instead of a literal run. Without
					// this branch the slide-number placeholder's <a:t> was never
					// read — inRun stays false — and every page bottom-right
					// rendered empty.
					state.inFld = true
					currentFldType = ""
					for _, attr := range t.Attr {
						if attr.Name.Local == "type" {
							currentFldType = attr.Value
						}
					}
				}
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
			case "hlinkClick":
				// The hyperlink of the run whose properties these are. Guarded on
				// inRunProps because <p:cNvPr> carries an hlinkClick too — a
				// shape-level link, which is a different element the model does
				// not place here.
				if state.inRunProps {
					currentHyperlink = parseHyperlinkClick(t, rels)
					// hlinkClick sits after solidFill in CT_TextCharacterProperties,
					// so by the time it arrives the run colour is whatever the fill
					// said — and PowerPoint still paints the run in the theme's
					// hlink colour (slide07: rPr solidFill tx1 renders 0000FF).
					// A linked run is a themed link, not a coloured run.
					if currentHyperlink != nil && pres != nil && pres.themeColors != nil {
						if hl, ok := pres.themeColors["hlink"]; ok && hl != "" && currentFont != nil {
							currentFont.Color = NewColor(hl)
						}
					}
				}
			case "rPr":
				if state.inBr {
					// The <a:rPr> inside a <a:br> sizes the blank line the
					// break produces — it is not a text run's font.
					for _, attr := range t.Attr {
						if attr.Name.Local == "sz" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								state.brSize = v
							}
						}
					}
				} else if state.inRun || state.inTcRun {
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
						case "baseline":
							// baseline is a shift in thousandths of a percent:
							// +30000 raises the run (superscript), -25000 drops it
							// (subscript) — the two values PowerPoint writes. The
							// model keeps booleans; the renderer applies the shift.
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if v > 0 {
									currentFont.Superscript = true
								} else if v < 0 {
									currentFont.Subscript = true
								}
							}
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
				// <a:noFill/> inside <a:ln> turns the outline off outright —
				// record a BorderNone pending border so the <p:style> lnRef
				// fallback at sp-end stays out of the way (a themed band with
				// ln noFill must not grow the lnRef's 2pt accent line).
				if state.inLn && !state.inTcPr {
					pendingBorder = &Border{Style: BorderNone}
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
								cell.border.leftDeclared = true
							case "R":
								cell.border.Right.Style = BorderNone
								cell.border.rightDeclared = true
							case "T":
								cell.border.Top.Style = BorderNone
								cell.border.topDeclared = true
							case "B":
								cell.border.Bottom.Style = BorderNone
								cell.border.bottomDeclared = true
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
					gradPathKind = ""
					gradFillTo = [4]int{}
					gradTileTo = [4]int{}
				} else if state.inSpPr && !state.inTxBody && !state.inLn && !state.inExtLst {
					state.inGradFill = true
					gradStopColors = nil
					gradStopPositions = nil
					gradAngle = 0
					gradPathKind = ""
					gradFillTo = [4]int{}
					gradTileTo = [4]int{}
				} else if state.inBgPr {
					state.inGradFill = true
					gradStopColors = nil
					gradStopPositions = nil
					gradAngle = 0
					gradPathKind = ""
					gradFillTo = [4]int{}
					gradTileTo = [4]int{}
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
			case "fillToRect":
				// Path-gradient focus rectangle: insets (l, t, r, b) from
				// the gradient box, in 0..100000 box units. The focus is the
				// centre of the rectangle they carve out.
				if state.inGradFill {
					for _, attr := range t.Attr {
						if v, err := strconv.Atoi(attr.Value); err == nil {
							switch attr.Name.Local {
							case "l":
								gradFillTo[0] = v
							case "t":
								gradFillTo[1] = v
							case "r":
								gradFillTo[2] = v
							case "b":
								gradFillTo[3] = v
							}
						}
					}
				}
			case "tileRect":
				// Tile rectangle insets — negative values grow the tile
				// beyond the gradient box, which moves the gradient's t=1
				// boundary (and with it the whole colour ramp) outward.
				if state.inGradFill {
					for _, attr := range t.Attr {
						if v, err := strconv.Atoi(attr.Value); err == nil {
							switch attr.Name.Local {
							case "l":
								gradTileTo[0] = v
							case "t":
								gradTileTo[1] = v
							case "r":
								gradTileTo[2] = v
							case "b":
								gradTileTo[3] = v
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
				if state.inClrFrom || state.inClrTo {
					// The colour a <a:clrChange> rule replaces or paints
					// with. Captured as hex, not as a Color — no alpha is
					// implied until the rule's own <a:alpha> says so.
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if state.inClrFrom {
								state.clrFromHex = attr.Value
							} else {
								state.clrToHex = attr.Value
							}
						}
					}
				} else if state.inGs {
					// Gradient stop color
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							c := NewColor("FF" + attr.Value)
							gradStopColors = append(gradStopColors, c)
							gradStopPositions = append(gradStopPositions, state.gradFillPos)
							lastColor = &gradStopColors[len(gradStopColors)-1]
						}
					}
				} else if state.inDuotone && currentDrawing != nil {
					// Duotone endpoint. lastColor points at the slot so the
					// transform cases (tint/satMod/...) fold into it the same
					// way they do for gradient stops.
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if duotoneIdx == 0 {
								currentDrawing.duotoneA = NewColor("FF" + attr.Value)
								currentDrawing.hasDuotone = true
								lastColor = &currentDrawing.duotoneA
							} else {
								currentDrawing.duotoneB = NewColor("FF" + attr.Value)
								lastColor = &currentDrawing.duotoneB
							}
							duotoneIdx++
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
								lineColorExplicit = true
								lastColor = &currentLine.lineColor
							} else if state.inSp || state.inPic {
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
				if state.inDuotone && currentDrawing != nil {
					if duotoneIdx == 0 {
						currentDrawing.duotoneA = c
						currentDrawing.hasDuotone = true
						lastColor = &currentDrawing.duotoneA
					} else {
						currentDrawing.duotoneB = c
						lastColor = &currentDrawing.duotoneB
					}
					duotoneIdx++
				} else if state.inGs {
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
						lineColorExplicit = true
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
				if state.inFillRef || state.inLnRef {
					// The colour the style reference points at. Captured by
					// name and resolved when <p:style> closes, because a
					// reference is a fallback, not a property: an explicit
					// fill in <p:spPr> still outranks it. The transforms on
					// this schemeClr ride along — a lnRef naming
					// "dk1 shade 50%" is not plain dk1. Each reference keeps
					// its own op list: fillRef starts AFTER lnRef inside
					// <p:style>, and a shared list would let the fill wipe
					// the line's shade before the line is resolved.
					state.inStyleScheme = true
					if state.inFillRef {
						state.styleFillOps = nil
						state.styleOpsTarget = &state.styleFillOps
					} else {
						state.styleLnOps = nil
						state.styleOpsTarget = &state.styleLnOps
					}
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if state.inFillRef {
								state.styleFillScheme = attr.Value
							} else {
								state.styleLnScheme = attr.Value
							}
						}
					}
					break
				}
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
							lineColorExplicit = true
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
						lineColorExplicit = true
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
				if state.inClrTo {
					// The replacement's own opacity — <a:alpha val="0"/> is
					// PowerPoint's "knock the colour out entirely".
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								state.clrToAlpha = v
							}
						}
					}
					break
				}
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
				if state.inStyleScheme {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								*state.styleOpsTarget = append(*state.styleOpsTarget, themeColorOp{op: "lumMod", val: float64(v) / 100000.0})
							}
						}
					}
				}
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
				if state.inStyleScheme {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								*state.styleOpsTarget = append(*state.styleOpsTarget, themeColorOp{op: "lumOff", val: float64(v) / 100000.0})
							}
						}
					}
				}
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
				if state.inStyleScheme {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								*state.styleOpsTarget = append(*state.styleOpsTarget, themeColorOp{op: "tint", val: float64(v) / 100000.0})
							}
						}
					}
				}
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
				if state.inStyleScheme {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								*state.styleOpsTarget = append(*state.styleOpsTarget, themeColorOp{op: "shade", val: float64(v) / 100000.0})
							}
						}
					}
				}
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								applyShade(lastColor, float64(v)/100000.0)
							}
						}
					}
				}
			case "satMod":
				// Saturation modulation — the Office fill/line styles pair it
				// with tint/shade on every stop. On the achromatic phClr most
				// references carry it is a no-op; on accents it is not.
				if state.inStyleScheme {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								*state.styleOpsTarget = append(*state.styleOpsTarget, themeColorOp{op: "satMod", val: float64(v) / 100000.0})
							}
						}
					}
				}
				// Slide-level colour (gradient stops, solid fills): apply in
				// document order like the other transforms, with the
				// no-S-clamp semantics PowerPoint uses (applySatMod).
				if state.inSrgbClr && lastColor != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								applySatMod(lastColor, float64(v)/100000.0)
							}
						}
					}
				}
			case "latin":
				if state.inRunProps && currentFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentFont.Name = n
					}
				} else if state.inDefRPr && state.inLstStyleLvl1 && lstStyleFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						lstStyleFont.Name = n
					}
				} else if state.inDefRPr && defFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						defFont.Name = n
					}
				}
			case "ea":
				// East Asian font
				if state.inRunProps && currentFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentFont.NameEA = n
					}
				} else if state.inDefRPr && state.inLstStyleLvl1 && lstStyleFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						lstStyleFont.NameEA = n
					}
				} else if state.inDefRPr && defFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						defFont.NameEA = n
					}
				}
			case "sym":
				// Symbol font: the declaration that carries the run's PUA
				// characters (U+F000-F0FF). PowerPoint writes it for runs the
				// user typed with a symbol font selected; without it the
				// arrows and bullets those runs hold draw as tofu.
				if state.inRunProps && currentFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentFont.NameSym = n
					}
				} else if state.inDefRPr && state.inLstStyleLvl1 && lstStyleFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						lstStyleFont.NameSym = n
					}
				} else if state.inDefRPr && defFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						defFont.NameSym = n
					}
				}
			case "t":
				if state.inTcRun {
					state.inTcText = true
				} else if state.inRun {
					state.inText = true
				}
			case "br":
				// A break belongs to a paragraph wherever that paragraph lives.
				// Gating it on "inside a shape" alone meant a cell's <a:br/> was
				// never read at all: the cell branch sets inTcParagraph, and
				// inParagraph is only set for a shape's own text body.
				if (state.inParagraph || state.inTcParagraph) && currentParagraph != nil {
					state.curBreak = currentParagraph.CreateBreak()
					state.inBr = true
					state.brSize = 0
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
				// <a:blip r:embed="rIdN"> — the image reference proper.
				//
				// The one element serves three unrelated contexts, told apart
				// by what encloses it rather than by anything on the tag: a
				// picture, a shape's picture fill, and a slide background.
				if id := embedRelID(t); id != "" {
					if state.inPic && currentDrawing != nil {
						currentDrawing.data, currentDrawing.mimeType = imageRelData(rels, slidePath, zr, id)
					} else if state.inSpPrBlipFill {
						pendingBlipFillData, pendingBlipFillMime = imageRelData(rels, slidePath, zr, id)
					} else if state.inBgBlipFill {
						bgBlipFillData, bgBlipFillMime = imageRelData(rels, slidePath, zr, id)
					}
				}
			case "svgBlip":
				// <asvg:svgBlip r:embed="rIdN"/>, a child of <a:blip> in the
				// Microsoft SVG extension (ext uri
				// 96DAC541-7B7A-43D3-8B79-37D633B846F1).
				//
				// PowerPoint writes this next to a raster reference on the
				// parent <a:blip>, so the raster is normally already in hand
				// and the nil guards below leave it alone. But SVG artwork can
				// also be written with the parent blip left bare and the sole
				// reference in here — decks that inline SVG art tend to do
				// exactly that. Without this case such a picture reads as zero
				// bytes and nothing is drawn, with no sign that anything was
				// skipped.
				if id := embedRelID(t); id != "" {
					if state.inPic && currentDrawing != nil {
						if currentDrawing.data == nil {
							currentDrawing.data, currentDrawing.mimeType = imageRelData(rels, slidePath, zr, id)
						}
					} else if state.inSpPrBlipFill {
						if pendingBlipFillData == nil {
							pendingBlipFillData, pendingBlipFillMime = imageRelData(rels, slidePath, zr, id)
						}
					} else if state.inBgBlipFill {
						if bgBlipFillData == nil {
							bgBlipFillData, bgBlipFillMime = imageRelData(rels, slidePath, zr, id)
						}
					}
				}
			case "alphaModFix":
				// <a:alphaModFix amt="50000"/> is a child of <a:blip>, so it
				// occurs both under a <p:pic> and under a background's
				// blipFill. Only the first was collected, which left a
				// translucent background image fully opaque.
				if state.inPic && currentDrawing != nil {
					if v, ok := intAttrValue(t, "amt"); ok {
						currentDrawing.alpha = v
					}
				} else if state.inBgBlipFill {
					if v, ok := intAttrValue(t, "amt"); ok {
						bgAlpha = v
					}
				}
			case "lum":
				// <a:lum bright=".." contrast=".."/> is a child of <a:blip>
				// and adjusts the picture's own pixels: a deck that darkens a
				// photo by 20% is unrecognisable without it, since the whole
				// image then renders a fifth brighter than PowerPoint draws
				// it. Collected in all three blip contexts, like alphaModFix.
				br, okBr := intAttrValue(t, "bright")
				co, okCo := intAttrValue(t, "contrast")
				if state.inPic && currentDrawing != nil {
					if okBr {
						currentDrawing.lumBright = br
					}
					if okCo {
						currentDrawing.lumContrast = co
					}
				} else if state.inSpPrBlipFill {
					if okBr {
						pendingBlipFillLumBright = br
					}
					if okCo {
						pendingBlipFillLumContrast = co
					}
				} else if state.inBgBlipFill {
					if okBr {
						bgLumBright = br
					}
					if okCo {
						bgLumContrast = co
					}
				}
			case "clrChange":
				// <a:clrChange><a:clrFrom>…</a:clrFrom><a:clrTo>…</a:clrTo>
				// </a:clrChange> is a child of <a:blip> and recolours the
				// picture's pixels: everything in clrFrom becomes clrTo. The
				// target usually carries <a:alpha val="0"/>, which is how
				// PowerPoint writes "make this colour transparent" — stacked
				// photos depend on that to composite. Parsed only for a
				// picture; a blip fill on a shape or a background is still
				// left alone (recorded gap).
				if state.inPic && currentDrawing != nil {
					state.inClrChange = true
					state.inClrFrom = false
					state.inClrTo = false
					state.clrFromHex = ""
					state.clrToHex = ""
					state.clrToAlpha = -1
				}
			case "clrFrom":
				if state.inClrChange {
					state.inClrFrom = true
					state.inClrTo = false
				}
			case "clrTo":
				if state.inClrChange {
					state.inClrTo = true
					state.inClrFrom = false
				}
			case "srcRect":
				if state.inPic && currentDrawing != nil {
					l, top, r, b := cropFromAttrs(t)
					currentDrawing.cropLeft, currentDrawing.cropTop = l, top
					currentDrawing.cropRight, currentDrawing.cropBottom = r, b
				} else if state.inBgBlipFill {
					bgCropLeft, bgCropTop, bgCropRight, bgCropBottom = cropFromAttrs(t)
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
				} else if (state.inSp || state.inPic) && state.inSpPr {
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
			case "sp3d":
				if state.inSpPr {
					state.inSp3d = true
				}
			case "bevelT":
				if state.inSp3d && state.inCxnSp && currentLine != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "prst" && attr.Value != "" {
							currentLine.SetBevelTop(attr.Value)
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
				} else if state.inTcPr && state.inTcPrLn {
					// The dash style of one side of a cell border. The side is
					// named by the <a:lnX> element this sits inside, so it is
					// read from the scanner's side marker rather than here.
					for _, attr := range t.Attr {
						if attr.Name.Local != "val" {
							continue
						}
						style := BorderSolid
						switch attr.Value {
						case "dash", "lgDash", "sysDash":
							style = BorderDash
						case "dot", "sysDot":
							style = BorderDot
						}
						setCellBorderStyle(tableCellAt(currentTable, currentTableRow, currentTableCol), state.tcPrLnSide, style)
					}
				}
			case "effectLst":
				// Two homes: inside <p:spPr> it is the shape's own shadow
				// (attached when the shape closes), inside <a:rPr> it is the
				// run's text shadow (attached to currentFont when the element
				// closes). A run is never inside spPr, so the two cannot be
				// confused.
				if (state.inSpPr && !state.inLn) || state.inRunProps {
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
				} else if state.inGradFill {
					// <a:path path="circle"> — the gradient-path kind. The
					// custGeom branch above owns this element name inside
					// pathLst; gradFill's <a:path> carries only the kind.
					for _, attr := range t.Attr {
						if attr.Name.Local == "path" {
							gradPathKind = attr.Value
						}
					}
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
				// <p:style> element inside <p:sp> or <p:cxnSp> — provides
				// default styling
				if (state.inSp || state.inCxnSp) && !state.inSpPr && !state.inTxBody {
					state.inStyle = true
					state.inFillRef = false
					state.inLnRef = false
					state.styleFillScheme = ""
					state.styleLnScheme = ""
					state.styleEffectIdx = 0
				}
			case "fillRef":
				// <a:fillRef idx="N"><a:schemeClr val="accent1"/></a:fillRef>
				// — the theme fill style the shape inherits. The idx picks a
				// style from the theme's fill style list (1 subtle … 3
				// intense); only its colour is taken here, applied as a solid
				// fill, because the real style bodies are gradients the theme
				// defines (recorded gap).
				//
				// idx="0" means *no* theme fill: the scheme colour inside the
				// element is a placeholder PowerPoint ignores. Treating it as
				// a real fill painted slide34's un-filled rightBrace solid
				// accent1 — PowerPoint draws only its outline.
				if state.inStyle {
					if fillRefIdx(t) == 0 {
						state.inFillRef = false
						state.styleFillIdx = 0
					} else {
						state.inFillRef = true
						state.styleFillScheme = ""
						state.styleFillIdx = fillRefIdx(t)
						state.styleFillOps = nil
					}
				}
			case "lnRef":
				// <a:lnRef idx="N"><a:schemeClr val="…"/></a:lnRef> — same
				// idea for the outline.
				if state.inStyle {
					state.inLnRef = true
					state.styleLnScheme = ""
					state.styleLnIdx = fillRefIdx(t)
					state.styleLnOps = nil
				}
			case "fontRef":
				// <a:fontRef> inside <p:style> — provides default text color
				if state.inStyle {
					state.inFontRef = true
				}
			case "effectRef":
				// <a:effectRef idx="N"><a:schemeClr …/></a:effectRef> — the
				// theme effect style the shape inherits. Only the index is
				// kept: the shadow it names is resolved when <p:style>
				// closes, and only where the shape's own <p:spPr> declared
				// no <a:effectLst>.
				if state.inStyle {
					state.styleEffectIdx = fillRefIdx(t)
				}
			}

		case xml.CharData:
			text := string(t)
			if state.inTableStyleID && currentTable != nil {
				currentTable.styleGUID += text
			}
			if (state.inTcText || state.inText) && currentParagraph != nil {
				tr := currentParagraph.CreateTextRun(text)
				if currentFont != nil {
					tr.font = currentFont
				}
				if currentHyperlink != nil {
					tr.SetHyperlink(currentHyperlink)
				}
				if state.inFld {
					tr.fieldType = currentFldType
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "AlternateContent":
				state.inAltContent = false
				state.inAltChoice = false
				state.inAltFallback = false
			case "Choice":
				state.inAltChoice = false
			case "Fallback":
				state.inAltFallback = false
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
							g.hidden = top.hidden
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
							// Add to parent group or slide — unless this group was
							// nested past maxGroupDepth, in which case it is
							// dropped rather than attached anywhere.
							switch {
							case top.detached:
								// Intentionally discarded; see maxGroupDepth.
							case len(grpStack) > 0:
								parentGroup := grpStack[len(grpStack)-1].group
								parentGroup.AddShape(g)
							default:
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
						currentPlaceholder.hidden = shapeHidden
						currentPlaceholder.description = shapeDescr
						currentPlaceholder.offsetX = offX
						currentPlaceholder.offsetY = offY
						currentPlaceholder.width = extCX
						currentPlaceholder.height = extCY
						currentPlaceholder.flipHorizontal = flipH
						currentPlaceholder.flipVertical = flipV
						currentPlaceholder.rotation = shapeRotation
						// A placeholder carries the same <p:style> fallback a
						// regular shape does: slide18's accent6 banner names no
						// fill in its <p:spPr> and lives entirely off its
						// fillRef/lnRef/effectRef. The non-placeholder commits
						// below hand the pending values to their shape; this
						// branch dropped them, so a themed placeholder rendered
						// with no fill and no border at all (its fontRef-white
						// text invisible on the white background).
						if pendingShapeFill != nil {
							currentPlaceholder.fill = pendingShapeFill
							pendingShapeFill = nil
						}
						if pendingBorder != nil {
							currentPlaceholder.border = pendingBorder
							pendingBorder = nil
						}
						if pendingShadow != nil {
							currentPlaceholder.shadow = pendingShadow
							pendingShadow = nil
						}
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
						autoShape.hidden = shapeHidden
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
							autoShape.lnSpcReduction = currentRichText.lnSpcReduction
							// The body properties are read into the temporary
							// text body whatever the shape turns out to be, so
							// they have to travel with it: a shape whose wrap
							// or auto-fit was dropped here wrapped its text in
							// the renderer however the deck said not to.
							autoShape.wordWrap = currentRichText.wordWrap
							autoShape.columns = currentRichText.columns
							autoShape.autoFit = currentRichText.autoFit
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
						ds.hidden = shapeHidden
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
						ds.lumBright = pendingBlipFillLumBright
						ds.lumContrast = pendingBlipFillLumContrast
						pendingBlipFillData = nil
						pendingBlipFillMime = ""
						pendingBlipFillLumBright, pendingBlipFillLumContrast = 0, 0
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(ds)
						} else {
							slide.shapes = append(slide.shapes, ds)
						}
					} else if currentRichText != nil {
						currentRichText.name = shapeName
						currentRichText.hidden = shapeHidden
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
						rt.hidden = shapeHidden
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
						autoShape.hidden = shapeHidden
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
			case "clrFrom":
				state.inClrFrom = false
			case "clrTo":
				state.inClrTo = false
			case "clrChange":
				if state.inClrChange {
					state.inClrChange = false
					// A rule with no source colour replaces nothing; keep
					// the order the file wrote them in, because stacked
					// photos can chain several knock-outs.
					if currentDrawing != nil && state.clrFromHex != "" {
						currentDrawing.recolors = append(currentDrawing.recolors, colorReplace{
							From:    state.clrFromHex,
							To:      state.clrToHex,
							ToAlpha: state.clrToAlpha,
						})
					}
					state.clrFromHex = ""
					state.clrToHex = ""
					state.clrToAlpha = -1
				}
			case "pic":
				if state.inPic {
					state.inPic = false
					if currentDrawing != nil {
						currentDrawing.name = shapeName
						currentDrawing.hidden = shapeHidden
						currentDrawing.description = shapeDescr
						currentDrawing.offsetX = offX
						currentDrawing.offsetY = offY
						currentDrawing.width = extCX
						currentDrawing.height = extCY
						currentDrawing.flipHorizontal = flipH
						currentDrawing.flipVertical = flipV
						currentDrawing.rotation = shapeRotation
						// The frame geometry, line and shadow ride in the
						// pic's spPr and were collected by the same pending*
						// variables a shape uses.
						currentDrawing.presetGeom = prstGeom
						if pendingBorder != nil {
							currentDrawing.border = pendingBorder
							pendingBorder = nil
						}
						if pendingShadow != nil {
							currentDrawing.shadow = pendingShadow
							pendingShadow = nil
						}
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentDrawing)
						} else {
							slide.shapes = append(slide.shapes, currentDrawing)
						}
					}
					currentDrawing = nil
				}
			case "duotone":
				state.inDuotone = false
			case "cxnSp":
				if state.inCxnSp {
					state.inCxnSp = false
					if currentLine != nil {
						currentLine.name = shapeName
						currentLine.hidden = shapeHidden
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
					isTable := currentTable != nil
					if isTable {
						currentTable.name = shapeName
						currentTable.hidden = shapeHidden
						currentTable.offsetX = offX
						currentTable.offsetY = offY
						currentTable.width = extCX
						currentTable.height = extCY
						// Tables are the same kind of object as every other
						// shape here, so they belong inside the enclosing group
						// like the rest. Appending them to the slide instead
						// left their child-space coordinates untransformed,
						// which placed a grouped table in the wrong spot.
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(currentTable)
						} else {
							slide.shapes = append(slide.shapes, currentTable)
						}
					}
					currentTable = nil

					// A chart graphicFrame: resolve the chart part and turn it
					// into a ChartShape so the rasterizer can draw it.
					hadChartRef := chartRelID != ""
					isChart := false
					if hadChartRef {
						var themeColors map[string]string
						if pres != nil {
							themeColors = pres.themeColors
						}
						if cs := r.readChartShape(zr, rels, slidePath, chartRelID, themeColors); cs != nil {
							cs.name = shapeName
							cs.hidden = shapeHidden
							cs.offsetX = offX
							cs.offsetY = offY
							cs.width = extCX
							cs.height = extCY
							cs.rotation = ((shapeRotation % 360) + 360) % 360
							cs.flipHorizontal = flipH
							cs.flipVertical = flipV
							if state.inGrpSp && currentGroup != nil {
								currentGroup.AddShape(cs)
							} else {
								slide.shapes = append(slide.shapes, cs)
							}
							isChart = true
						}
					}
					chartRelID = ""
					graphicDataIsChart = false

					// Anything else — SmartArt, an OLE object, a chart whose part
					// could not be read — used to be dropped here, which left a
					// blank region in the preview that nobody could tell apart
					// from a correctly rendered empty shape. Keep a visible
					// stand-in so the gap is obvious and reportable.
					if !isTable && !isChart {
						reason := classifyGraphicData(graphicDataURI)
						if hadChartRef {
							reason = "chart (its part could not be read)"
						}
						ph := NewUnsupportedShape(reason)
						ph.name = shapeName
						ph.hidden = shapeHidden
						ph.offsetX = offX
						ph.offsetY = offY
						ph.width = extCX
						ph.height = extCY
						ph.rotation = ((shapeRotation % 360) + 360) % 360
						ph.flipHorizontal = flipH
						ph.flipVertical = flipV
						ph.SetContentType(graphicDataURI)
						if state.inGrpSp && currentGroup != nil {
							currentGroup.AddShape(ph)
						} else {
							slide.shapes = append(slide.shapes, ph)
						}
					}
					graphicDataURI = ""
				}
			case "tableStyleId":
				state.inTableStyleID = false
				if currentTable != nil {
					currentTable.styleGUID = strings.TrimSpace(currentTable.styleGUID)
				}
			case "tbl":
				state.inTbl = false
				if currentTable != nil {
					applyTableStyleFill(currentTable, pres)
					applyTableStyleBorders(currentTable, pres)
				}
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
			case "br":
				if state.curBreak != nil {
					state.curBreak.rprSize = state.brSize
				}
				state.inBr = false
				state.curBreak = nil
				state.brSize = 0
			case "fld":
				state.inFld = false
				state.inRun = false
				currentFldType = ""
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
					// Fold the full stop list into the fill. Duplicate
					// positions collapse to the LAST stop written there —
					// PowerPoint's own reading (measured with a variant deck:
					// a pos=0 red followed by three pos=0 purples renders
					// purple, the red never shows). Linear gradients keep
					// their legacy angle; path gradients carry the kind and
					// the fillToRect/tileRect insets.
					stops := dedupeGradStops(gradStopColors, gradStopPositions)
					if len(stops) >= 2 {
						if state.inBgPr {
							if slide.background == nil {
								slide.background = NewFill()
							}
							slide.background.SetGradientStops(stops, gradAngle, gradPathKind, gradFillTo, gradTileTo)
						} else if state.inSpPr && state.inSp {
							pendingShapeFill = NewFill()
							pendingShapeFill.SetGradientStops(stops, gradAngle, gradPathKind, gradFillTo, gradTileTo)
						}
					}
				}
				state.inGradFill = false
				gradPathKind = ""
				gradFillTo = [4]int{}
				gradTileTo = [4]int{}
			case "blipFill":
				state.inSpPrBlipFill = false
				state.inBgBlipFill = false
			case "srgbClr":
				state.inSrgbClr = false
			case "schemeClr":
				state.inSrgbClr = false
				// The transforms a style-reference colour carries are now
				// complete; they stay pending in styleFillOps/styleLnOps
				// until <p:style> closes and the fill/line is resolved.
				state.inStyleScheme = false
			case "outerShdw":
				state.inOuterShdw = false
			case "effectLst":
				state.inEffectLst = false
				// The run's text shadow completes here: hand the pending
				// outerShdw to the font being built. The shape path attaches
				// the same variable when the shape closes, so only consume it
				// in a run context.
				if state.inRunProps && pendingShadow != nil && currentFont != nil {
					currentFont.Shadow = pendingShadow
					pendingShadow = nil
				}
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
			case "sp3d":
				state.inSp3d = false
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
			case "fillRef", "lnRef":
				// The ref elements reset their own flag when they close.
				// Without this, inFillRef stayed true for the whole
				// <p:style>, and the schemeClr inside effectRef and fontRef
				// was captured as the fill reference's colour — a fontRef
				// naming lt1 turned the fillRef's accent into white.
				state.inFillRef = false
				state.inLnRef = false
				state.inStyleScheme = false
			case "style":
				state.inStyle = false
				state.inFontRef = false
				state.inFillRef = false
				state.inLnRef = false
				state.inStyleScheme = false
				// The effect reference is the same kind of fallback: where
				// the shape's own <p:spPr> carried no <a:effectLst>, the
				// theme's effect style supplies the outer shadow. A copy, not
				// the theme's own pointer — the pending value is handed to
				// exactly one shape.
				if state.inSp && pres != nil && pendingShadow == nil && state.styleEffectIdx >= 1 && state.styleEffectIdx <= len(pres.themeEffectStyles) {
					if ts := pres.themeEffectStyles[state.styleEffectIdx-1].shadow; ts != nil {
						sh := *ts
						pendingShadow = &sh
					}
				}
				state.styleEffectIdx = 0
				// The style reference is a fallback: only where the shape's
				// own <p:spPr> declared nothing does the theme colour become
				// the fill. Applied to the pending values so the shape that
				// <p:sp> closes next picks them up through its usual path.
				//
				// The idx is an index into the theme's fillStyleLst, not a
				// solid-fill instruction: idx 2 is a three-stop gradient
				// whose stops are the reference colour tinted, and resolving
				// it as a solid painted every themed gradient text box flat
				// black (slide07 drew #000 where PowerPoint drew a
				// 236→188 grey ramp). resolveThemeFillStyle falls back to
				// the plain solid when no theme styles were parsed.
				if state.inSp && pres != nil && pendingShapeFill == nil && state.styleFillScheme != "" {
					pendingShapeFill = resolveThemeFillStyle(pres, state.styleFillIdx, state.styleFillScheme, state.styleFillOps)
				}
				if state.inSp && pres != nil && pendingBorder == nil && state.styleLnScheme != "" {
					if argb, ok := pres.themeColors[state.styleLnScheme]; ok && argb != "" {
						lineColor := applyColorOps(NewColor(argb), state.styleLnOps)
						width := 1 // points
						if state.styleLnIdx >= 1 && state.styleLnIdx <= len(pres.themeLnStyles) {
							lnStyle := pres.themeLnStyles[state.styleLnIdx-1]
							lineColor = applyColorOps(lineColor, lnStyle.ops)
							if lnStyle.widthEMU > 0 {
								width = (lnStyle.widthEMU + 6350) / 12700
							}
						}
						pendingBorder = NewBorder()
						pendingBorder.Style = BorderSolid
						pendingBorder.Width = width
						pendingBorder.Color = lineColor
					}
				}
				// A connector's colour lives on its <a:ln> too, but when the
				// ln declares only a width — the way PowerPoint writes themed
				// arrows — the colour the lnRef names is all there is.
				if state.inCxnSp && currentLine != nil && pres != nil && !lineColorExplicit && state.styleLnScheme != "" {
					if argb, ok := pres.themeColors[state.styleLnScheme]; ok && argb != "" {
						currentLine.lineColor = NewColor(argb)
					}
				}
				state.styleFillScheme = ""
				state.styleLnScheme = ""
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
						top.hidden = shapeHidden
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
		// A background picture carries its crop and opacity on the same
		// siblings a <p:pic> uses, so carry them over rather than drawing the
		// whole image stretched.
		ds.cropLeft, ds.cropTop = bgCropLeft, bgCropTop
		ds.cropRight, ds.cropBottom = bgCropRight, bgCropBottom
		ds.alpha = bgAlpha
		ds.lumBright = bgLumBright
		ds.lumContrast = bgLumContrast
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

// embedRelID returns the r:embed relationship id carried by an element, or ""
// when it carries none.
//
// Only the local name is compared: the reverse namespace may be bound to any
// prefix, and xml.Decoder hands back the local name either way.
func embedRelID(t xml.StartElement) string {
	for _, attr := range t.Attr {
		if attr.Name.Local == "embed" {
			return attr.Value
		}
	}
	return ""
}

// intAttrValue returns the named attribute as an int. It reports false when the
// attribute is absent or not an integer, so callers can leave the previous
// value (or the zero value) alone instead of writing a silent 0.
func intAttrValue(t xml.StartElement, local string) (int, bool) {
	for _, attr := range t.Attr {
		if attr.Name.Local == local {
			if v, err := strconv.Atoi(attr.Value); err == nil {
				return v, true
			}
			return 0, false
		}
	}
	return 0, false
}

// fillRefIdx reads the idx of an <a:fillRef>; zero when absent. idx="0" is
// the theme's "no fill" entry, which is why the missing attribute also reads
// as zero.
func fillRefIdx(t xml.StartElement) int {
	v, _ := intAttrValue(t, "idx")
	return v
}

// cropFromAttrs reads the l/t/r/b of an <a:srcRect>. The values are signed
// percentages in 1/1000 of a percent, and PowerPoint uses negatives for a
// picture cropped outwards (the "fill" mode), so they must not be clamped to
// zero.
func cropFromAttrs(t xml.StartElement) (left, top, right, bottom int) {
	left, _ = intAttrValue(t, "l")
	top, _ = intAttrValue(t, "t")
	right, _ = intAttrValue(t, "r")
	bottom, _ = intAttrValue(t, "b")
	return left, top, right, bottom
}

// imageRelData resolves an image relationship to the bytes of the part it
// names, along with the MIME type implied by that part's extension.
//
// parseHyperlinkClick resolves an <a:hlinkClick> to a hyperlink.
//
// The element names its target through a relationship, so the slide's own
// relationship list is what gives the link a meaning, and the action attribute
// says which kind of link it is. PowerPoint writes a jump to another slide as
// action="ppaction://hlinksldjump" plus a slide-type relationship whose target
// is slideN.xml, and an external link as a hyperlink-type relationship with no
// action.
//
// Anything else — the show-navigation actions, or an id this slide's
// relationships do not define — yields no hyperlink at all. That is deliberate:
// the alternative is a link the model cannot act on, which the writer would then
// re-emit as an external relationship pointing somewhere that does not exist.
//
// An external target is kept exactly as it was written. Scheme validation
// belongs to NewHyperlink, which guards URLs a caller supplies; a reader that
// dropped the links whose scheme it disliked would lose content from the very
// document it was asked to preserve.
func parseHyperlinkClick(se xml.StartElement, rels []xmlRelForRead) *Hyperlink {
	var relID, action string
	for _, attr := range se.Attr {
		switch {
		case attr.Name.Space == nsOfficeDocRels && attr.Name.Local == "id":
			relID = attr.Value
		case attr.Name.Space == "" && attr.Name.Local == "action":
			action = attr.Value
		}
	}
	if relID == "" {
		return nil
	}
	for _, rel := range rels {
		if rel.ID != relID {
			continue
		}
		switch {
		case action == actionSlideJump:
			if n := slideNumberFromRelTarget(rel.Target); n > 0 {
				return NewInternalHyperlink(n)
			}
		case action == "" && strings.HasSuffix(rel.Type, "/hyperlink") && rel.Target != "":
			return &Hyperlink{URL: rel.Target}
		}
		return nil
	}
	return nil
}

// slideNumberFromRelTarget extracts N from a relationship target such as
// "slide3.xml" or "../slides/slide3.xml". It returns 0 when the target does not
// name a slide, which is what makes it usable as a validity check.
func slideNumberFromRelTarget(target string) int {
	base := lastPathComponent(target)
	if !strings.HasPrefix(base, "slide") || !strings.HasSuffix(base, ".xml") {
		return 0
	}
	n, err := strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(base, "slide"), ".xml"))
	if err != nil || n < 1 {
		return 0
	}
	return n
}

// A nil result means the picture cannot be drawn: the id is unknown, the part
// is missing, or the package cannot be read. Callers record that as "no image",
// which the renderer turns into a labelled placeholder — a blank rectangle
// looks exactly like a picture that was meant to be blank, so the failure has
// to be visible.
//
// Targets are normally relative to the part owning the relationship
// ("media/image6.svg" from "ppt/slides/slide8.xml"), but package-absolute
// targets ("ppt/media/image6.png") are also written; both are handled.
func imageRelData(rels []xmlRelForRead, slidePath string, zr *zip.Reader, relID string) ([]byte, string) {
	if relID == "" {
		return nil, ""
	}
	for _, rel := range rels {
		if rel.ID != relID {
			continue
		}
		imgPath := rel.Target
		if !strings.HasPrefix(imgPath, "ppt/") {
			dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
			imgPath = resolveRelativePath(dir, imgPath)
		}
		data, err := readFileFromZip(zr, imgPath)
		if err != nil {
			return nil, ""
		}
		return data, guessMimeType(imgPath)
	}
	return nil, ""
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
	fontName  string
	fontEA    string
	fontSize  int
	fontBold  bool
	fontColor Color
	// Horizontal alignment from the placeholder's own <a:lstStyle> lvl1pPr
	// algn attribute. The master's slide-number placeholder is where
	// algn="r" lives in a deck PowerPoint wrote — its page number is
	// right-aligned inside the placeholder box — and no <p:txStyles> table
	// carries it.
	alignH   string
	alignSet bool
	// Text insets from bodyPr
	insetLeft   int64
	insetRight  int64
	insetTop    int64
	insetBottom int64
	insetsSet   bool
	// Vertical anchoring from bodyPr. A placeholder inherits its anchor from
	// the layout and then the master the same way it inherits its insets: the
	// master's title placeholder is where "anchor=ctr" lives in a deck
	// PowerPoint wrote, because the layout and the slide both leave bodyPr
	// empty.
	anchorSet bool
	anchor    TextAnchorType
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

	// The master sits one rung above the layout and is the same for every slide
	// that shares it, so it is read once and cached on the presentation.
	if !pres.masterRead {
		r.readMaster(zr, layoutRels, pres)
	}

	// Parse layout images and non-placeholder text shapes, prepend to slide (behind slide content)
	layoutImages := r.parseLayoutImages(data, layoutRels, zr, layoutPath, pres)
	if len(layoutImages) > 0 {
		slide.shapes = append(layoutImages, slide.shapes...)
	}

	// Parse layout to extract placeholder definitions
	layoutPHs := r.parsePlaceholderDefs(data, pres)

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

	// Apply inherited properties to slide placeholders. The ladder, farthest
	// ancestor first, is master <p:txStyles> → master placeholder → layout
	// placeholder → the slide itself. Every step below fills in only what the
	// rungs nearer the slide left at the library's default, so applying them in
	// this order lets each nearer rung win.
	for _, shape := range slide.shapes {
		ph, ok := shape.(*PlaceholderShape)
		if !ok {
			continue
		}

		match := matchPlaceholderDef(layoutPHs, ph)
		masterMatch := matchPlaceholderDef(pres.masterPlaceholders, ph)

		// Geometry takes the nearest definition that has any, rather than
		// applying the master first and locking the layout out. Both are empty
		// for the title in a deck PowerPoint wrote, so the master is often the
		// only rung that answers at all: without it a title placeholder is drawn
		// at 0,0 with zero size, which shrinks it, moves it left and re-wraps it.
		if ph.width == 0 && ph.height == 0 {
			for _, def := range []*layoutPlaceholder{match, masterMatch} {
				if def != nil && (def.extCX != 0 || def.extCY != 0) {
					ph.offsetX = def.offX
					ph.offsetY = def.offY
					ph.width = def.extCX
					ph.height = def.extCY
					break
				}
			}
		}

		// Text insets: layout preferred, then master. Same reasoning as above,
		// but here the guard on insetsSet makes the order self-enforcing.
		applyPlaceholderInsets(ph, match)
		applyPlaceholderInsets(ph, masterMatch)

		// Horizontal alignment, layout preferred, then master — applied before
		// the master's txStyles so the placeholder's own lstStyle wins the
		// ladder, the way PowerPoint resolves it. The master's sldNum
		// placeholder is the case that needs this: its algn="r" is the only
		// right-aligned paragraph in the deck.
		applyPlaceholderAlign(ph, masterMatch)
		applyPlaceholderAlign(ph, match)

		// Vertical anchoring, same ladder and same reasoning as the insets:
		// the slide's own anchor wins, then the layout's, then the master's.
		applyPlaceholderAnchor(ph, match)
		applyPlaceholderAnchor(ph, masterMatch)

		// Fonts, farthest first so that each nearer rung can still override.
		applyMasterTextStyles(ph, pres.masterTextStyles)
		applyPlaceholderFont(ph, masterMatch)
		applyPlaceholderFont(ph, match)
	}
}

// runSizeIsDefault reports whether a run is still carrying a size the reader
// substituted for markup that declared none.
//
// There are two such sizes and both have to be recognised: 18, which the reader
// assigns a placeholder run with no sz, and 10, which NewFont starts at. Testing
// only for 10 is what left a 44pt title at 18pt even once the master was read.
func runSizeIsDefault(size int) bool { return size == 18 || size <= 10 }

// matchPlaceholderDef finds the definition a slide placeholder inherits from.
//
// A slide placeholder that declares an index matches by index first: OOXML
// inheritance keys on <p:ph idx>, and PowerPoint-written slides routinely drop
// the type attribute on content placeholders (<p:ph idx="1"/> with no type),
// which must still find the master's type="body" idx="1" definition. Requiring
// type AND idx to both equal left those slides matching nothing, so the body
// fell back to a default-size box and re-wrapped every paragraph (deck
// 00022693 slide17: the full-width body rendered as a one-word-per-line
// column). The type walk — including the alias list — only runs for
// index-less placeholders or when no definition carries the index.
//
// PowerPoint treats ctrTitle as the title's centred variant and subTitle as
// the body's, and the master only ever defines title/body. Without the
// aliases a ctrTitle placeholder matched nothing on the master rung and lost
// the master's anchor="ctr" (slide1's title sat 32px too low because it
// rendered top-anchored).
func matchPlaceholderDef(defs []layoutPlaceholder, ph *PlaceholderShape) *layoutPlaceholder {
	if ph.phIdx > 0 {
		for i := range defs {
			if defs[i].phIdx == ph.phIdx {
				return &defs[i]
			}
		}
	}
	for i := range defs {
		d := &defs[i]
		if d.phType == string(ph.phType) && d.phIdx == ph.phIdx {
			return d
		}
	}
	aliases := placeholderTypeAliases(string(ph.phType))
	for _, alias := range aliases {
		for i := range defs {
			if defs[i].phType == alias {
				return &defs[i]
			}
		}
	}
	return nil
}

// placeholderTypeAliases returns the definition type names that can satisfy a
// placeholder of the given type, most specific first. Only the two title/body
// variants have aliases; every other type matches itself alone.
func placeholderTypeAliases(t string) []string {
	switch t {
	case "ctrTitle":
		return []string{"ctrTitle", "title"}
	case "subTitle":
		return []string{"subTitle", "body"}
	default:
		return []string{t}
	}
}

// applyPlaceholderInsets copies a definition's text insets onto the placeholder
// if the placeholder has not been given any of its own.
func applyPlaceholderInsets(ph *PlaceholderShape, def *layoutPlaceholder) {
	if ph == nil || def == nil || ph.insetsSet || !def.insetsSet {
		return
	}
	ph.insetLeft = def.insetLeft
	ph.insetRight = def.insetRight
	ph.insetTop = def.insetTop
	ph.insetBottom = def.insetBottom
	ph.insetsSet = true
}

// applyPlaceholderAlign copies a definition's lstStyle alignment onto the
// placeholder's paragraphs that declared none of their own. Runs before
// applyMasterTextStyles so the placeholder's own lstStyle outranks the
// txStyles tables in the ladder.
func applyPlaceholderAlign(ph *PlaceholderShape, def *layoutPlaceholder) {
	if ph == nil || def == nil || !def.alignSet || def.alignH == "" {
		return
	}
	for _, para := range ph.paragraphs {
		if para.alignment == nil || para.alignment.Horizontal == "" {
			if para.alignment == nil {
				para.alignment = NewAlignment()
			}
			para.alignment.Horizontal = HorizontalAlignment(def.alignH)
		}
	}
}

// applyPlaceholderAnchor copies a definition's vertical anchor onto the
// placeholder if neither the slide nor a nearer rung of the ladder gave it one.//
// TextAnchorNone ("") is what the reader stores for "no anchor attribute", so
// it doubles as the unset marker; an explicit anchor="t" reads back as "t" and
// is therefore never overwritten — which matters, because an explicit top is
// the slide overriding the master's centre.
func applyPlaceholderAnchor(ph *PlaceholderShape, def *layoutPlaceholder) {
	if ph == nil || def == nil || !def.anchorSet || ph.textAnchor != TextAnchorNone {
		return
	}
	ph.textAnchor = def.anchor
}

// applyPlaceholderFont copies a definition's font onto the runs of ph that have
// not been given one of their own.
//
// "Have not been given one" is read off the library's own defaults, because the
// reader cannot tell a run that declared Calibri 10 from one that declared
// nothing: it fills in both the same way. 18 is the size the reader assigns a
// placeholder run with no sz, 10 is what NewFont starts at.
func applyPlaceholderFont(ph *PlaceholderShape, def *layoutPlaceholder) {
	if ph == nil || def == nil {
		return
	}
	if def.fontName == "" && def.fontEA == "" && def.fontSize == 0 && !def.fontBold {
		return
	}
	for _, para := range ph.paragraphs {
		for _, elem := range para.elements {
			tr, ok := elem.(*TextRun)
			if !ok || tr.font == nil {
				continue
			}
			if tr.font.Name == "Calibri" && def.fontName != "" {
				tr.font.Name = def.fontName
			}
			if tr.font.NameEA == "" && def.fontEA != "" {
				tr.font.NameEA = def.fontEA
			}
			if runSizeIsDefault(tr.font.Size) && def.fontSize > 0 {
				tr.font.Size = def.fontSize
			}
			if def.fontBold {
				tr.font.Bold = true
			}
			if def.fontColor.ARGB != "" && def.fontColor.ARGB != "FF000000" && tr.font.Color.ARGB == "FF000000" {
				tr.font.Color = def.fontColor
			}
		}
	}
}

// masterLevelStyle is one <a:lvlNpPr> of a slide master's <p:titleStyle>,
// <p:bodyStyle> or <p:otherStyle>.
//
// This is where a real deck's placeholder text is actually styled. A slide says
// <p:ph type="title"/> and one run of text; the layout's placeholder
// <a:lstStyle> is usually empty; the 44pt centred title comes from the master
// and nowhere else. Reading none of it leaves the title at the library default,
// which both shrinks it and moves it left — and because the default measures
// differently, it wraps onto a second line that PowerPoint never draws.
type masterLevelStyle struct {
	fontName string
	fontEA   string
	size     int // points; 0 means the master says nothing
	bold     bool
	italic   bool
	align    string
	marL     int64
	indent   int64
	color    Color
	// Space before each paragraph, in hundredths of a point — the units the
	// paragraph model stores spcPts in. A master usually declares it as a
	// percentage of the line (<a:spcPct val="20000"/>), which is resolved at
	// layout time against the paragraph's own run size — the COM variants on
	// slide34 proved the base is the run's size and not the level's defRPr
	// (enlarging defRPr 28→48pt moved nothing), so a percentage cannot be
	// turned into points here where only the level size is known.
	spcBef int
	// spcBefPct keeps the raw percentage when the level declared spcPct
	// rather than spcPts; applyMasterTextStyles defers it to the paragraph.
	spcBefPct int
	// The level's bullet, exactly as the master declares it: buChar carries
	// the character and buFont its typeface, buAutoNum an ordered list (with
	// its startAt), buNone an explicit "no bullet". An empty buChar with no
	// auto-num and no buNone means the level declares nothing, and the
	// paragraph keeps whatever nearer rung of the ladder gave it.
	buChar    string
	buFont    string
	buAutoNum string
	buStartAt int
	buNone    bool
}

// masterTextStyles is a slide master's <p:txStyles>, indexed by outline level
// 1..9: title is <p:titleStyle>, body <p:bodyStyle>, other <p:otherStyle>.
type masterTextStyles struct {
	title [10]*masterLevelStyle
	body  [10]*masterLevelStyle
	other [10]*masterLevelStyle
}

// styleFor walks the ladder PowerPoint walks: the style a placeholder of this
// type inherits at this outline level, falling back to level 1 and then to the
// "other" list, both of which PowerPoint does too.
func (m *masterTextStyles) styleFor(phType string, level int) *masterLevelStyle {
	if m == nil {
		return nil
	}
	// The chrome placeholders are not outline placeholders: PowerPoint styles
	// them only from their own placeholder <a:lstStyle> (the master's sldNum
	// placeholder says algn="r" sz="1200" there and nowhere else) and never
	// from <p:txStyles>. Falling through to body/other styled the page number
	// with otherStyle's 24pt bullet list — a bullet glyph and a left-aligned,
	// oversized number exactly where PowerPoint draws none of that.
	switch phType {
	case "sldNum", "dt", "ftr":
		return nil
	}
	table := m.body
	if phType == "title" || phType == "ctrTitle" {
		table = m.title
	}
	if level < 1 || level > 9 {
		level = 1
	}
	if s := table[level]; s != nil {
		return s
	}
	if s := table[1]; s != nil {
		return s
	}
	if s := m.other[level]; s != nil {
		return s
	}
	return m.other[1]
}

// readMaster reads the slide master the layout points at: the placeholder
// geometry and the <p:txStyles> a placeholder inherits when neither the slide
// nor the layout gives it either.
//
// Both are read together and once, because they come from the same part and are
// the same for every slide that shares the master.
func (r *PPTXReader) readMaster(zr *zip.Reader, layoutRels []xmlRelForRead, pres *Presentation) {
	pres.masterRead = true
	path := ""
	for _, rel := range layoutRels {
		if rel.Type == relTypeSlideMaster {
			path = rel.Target
			break
		}
	}
	if path == "" {
		path = "ppt/slideMasters/slideMaster1.xml"
	} else if !strings.HasPrefix(path, "ppt/") {
		path = resolveRelativePath("ppt/slideLayouts", path)
	}
	data, err := readFileFromZip(zr, path)
	if err != nil {
		pres.masterTextStyles = &masterTextStyles{}
		return
	}
	pres.masterPlaceholders = r.parsePlaceholderDefs(data, pres)
	pres.masterTextStyles = parseMasterTextStyles(data, pres)
}

// parseMasterTextStyles reads <p:txStyles> out of a slide master.
func parseMasterTextStyles(data []byte, pres *Presentation) *masterTextStyles {
	m := &masterTextStyles{}
	decoder := xml.NewDecoder(bytes.NewReader(data))

	var table *[10]*masterLevelStyle
	var cur *masterLevelStyle
	inDefRPr := false
	inSolidFill := false
	inSpcBef := false
	spcBefPct := 0

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "titleStyle":
				table = &m.title
			case "bodyStyle":
				table = &m.body
			case "otherStyle":
				table = &m.other
			case "lvl1pPr", "lvl2pPr", "lvl3pPr", "lvl4pPr", "lvl5pPr",
				"lvl6pPr", "lvl7pPr", "lvl8pPr", "lvl9pPr":
				if table == nil {
					continue
				}
				lvl := int(t.Name.Local[3] - '0')
				cur = &masterLevelStyle{}
				table[lvl] = cur
				// The alignment, the margins and the indent are attributes of
				// the level element itself: a list style has no <a:pPr> child
				// to carry them, unlike a paragraph in a text body. Reading
				// them off a child left the title's algn="ctr" unread, so the
				// title kept the left alignment it had been defaulted to.
				//
				// The attributes go through the same reader the slide and
				// layout scanners use; a third hand-rolled one would drift
				// from them.
				p := NewParagraph()
				applyPPrAttrs(p, t.Attr)
				cur.align = string(p.alignment.Horizontal)
				cur.marL = p.alignment.MarginLeft
				cur.indent = p.alignment.Indent
			case "defRPr":
				inDefRPr = true
				if cur == nil {
					continue
				}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "sz":
						if v, err := strconv.Atoi(attr.Value); err == nil {
							cur.size = v / 100
						}
					case "b":
						cur.bold = attr.Value == "1" || attr.Value == "true"
					case "i":
						cur.italic = attr.Value == "1" || attr.Value == "true"
					}
				}
			case "latin":
				if inDefRPr && cur != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						cur.fontName = n
					}
				}
			case "ea":
				if inDefRPr && cur != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						cur.fontEA = n
					}
				}
			case "solidFill":
				inSolidFill = true
			case "spcBef":
				// Space before each paragraph. PowerPoint's own masters write
				// it as a percentage of the line (<a:spcPct val="20000"/>),
				// less often as absolute points; both are captured here and
				// the percentage is converted when the level closes, because
				// defRPr — which carries the size the percentage needs — ends
				// after spcBef begins.
				if cur != nil {
					inSpcBef = true
					spcBefPct = 0
				}
			case "spcPts":
				if inSpcBef && cur != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cur.spcBef = v
							}
						}
					}
				}
			case "spcPct":
				if inSpcBef && cur != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								spcBefPct = v
							}
						}
					}
				}
			case "buChar":
				// The bullet elements are children of the level element, next
				// to spcBef and defRPr — the same grammar a paragraph's pPr
				// uses, one rung up the ladder.
				if cur != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "char" {
							cur.buChar = attr.Value
						}
					}
				}
			case "buFont":
				if cur != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						cur.buFont = n
					}
				}
			case "buNone":
				if cur != nil {
					cur.buNone = true
				}
			case "buAutoNum":
				if cur != nil {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "type":
							cur.buAutoNum = attr.Value
						case "startAt":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								cur.buStartAt = v
							}
						}
					}
				}
			case "srgbClr":
				if inSolidFill && cur != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							cur.color = NewColor("FF" + attr.Value)
						}
					}
				}
			case "schemeClr":
				if inSolidFill && cur != nil && pres != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local != "val" {
							continue
						}
						if argb, ok := pres.themeColors[attr.Value]; ok && argb != "" {
							cur.color = NewColor(argb)
						}
					}
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "defRPr":
				inDefRPr = false
			case "solidFill":
				inSolidFill = false
			case "spcBef":
				inSpcBef = false
			case "lvl1pPr", "lvl2pPr", "lvl3pPr", "lvl4pPr", "lvl5pPr",
				"lvl6pPr", "lvl7pPr", "lvl8pPr", "lvl9pPr":
				// Keep the percentage raw: the base it rides is the
				// paragraph's own run size — slide34's COM variants moved
				// every gap by 20% × 1.2 × 32pt when the master's 20% was
				// tripled, and by nothing when the level's defRPr grew — so
				// resolving it here against the level size (the old bake)
				// fixed the number for the wrong reason and starved every
				// paragraph whose runs differ from the level default.
				if cur != nil && cur.spcBef == 0 && spcBefPct > 0 {
					cur.spcBefPct = spcBefPct
				}
				cur = nil
			case "titleStyle", "bodyStyle", "otherStyle":
				table = nil
			case "txStyles":
				return m
			}
		}
	}
	return m
}

// applyMasterTextStyles is the master rung of the inheritance ladder: it fills
// in what neither the slide nor the layout said. Everything here is a fallback
// — a paragraph or run that already carries a value keeps it, because the
// layout and the slide are nearer ancestors than the master.
func applyMasterTextStyles(ph *PlaceholderShape, m *masterTextStyles) {
	if ph == nil || m == nil {
		return
	}
	for pi, para := range ph.paragraphs {
		// lvl is 0-based on the paragraph, lvlNpPr is 1-based on the master:
		// a level-0 paragraph is styled by lvl1pPr, a level-1 one by lvl2pPr.
		level := 1
		if para.alignment != nil && para.alignment.Level > 0 {
			level = para.alignment.Level + 1
		}
		s := m.styleFor(string(ph.phType), level)
		if s == nil {
			continue
		}

		if para.alignment == nil {
			para.alignment = NewAlignment()
		}
		if para.alignment.Horizontal == "" && s.align != "" {
			para.alignment.Horizontal = HorizontalAlignment(s.align)
		}
		// marL/indent: a paragraph that stated the attribute explicitly keeps
		// it even when the stated value is 0 — slide27's body paragraphs say
		// marL="0" indent="0" to break free of the master's hanging indent,
		// and baking the master's marL in shifted the whole text block right.
		if !para.alignment.marLSet && s.marL != 0 {
			para.alignment.MarginLeft = s.marL
		}
		if !para.alignment.indentSet && s.indent != 0 {
			para.alignment.Indent = s.indent
		}
		// Space before paragraphs. The inherited value reaches paragraphs two
		// and onward only — a COM experiment on the comparison deck (explicit
		// spcPts 0 on the first paragraph moved nothing, on the second moved
		// one line) shows PowerPoint keeps the first paragraph flush with the
		// text inset unless the paragraph itself declares spacing. A declared
		// spcPts already sits in spaceBefore, a declared spcPct in
		// spaceBeforePct; both apply to the first paragraph. The master's own
		// percentage stays raw — its base is the paragraph's run size at
		// layout time, not the level's defRPr (slide34's vB variant).
		if pi > 0 && para.spaceBefore == 0 && para.spaceBeforePct == 0 {
			if s.spcBef > 0 {
				para.spaceBefore = s.spcBef
			} else if s.spcBefPct > 0 {
				para.inheritedSpaceBeforePct = s.spcBefPct
			}
		}

		// Bullets inherit on the same ladder. A paragraph that declared
		// nothing (bullet == nil) takes the level's bullet; one that declared
		// <a:buChar>, <a:buAutoNum> or <a:buNone> keeps it — the reader marks
		// buNone as a bullet of type None, so a non-nil bullet always means
		// "the slide spoke". A paragraph with no runs stays bulletless too:
		// PowerPoint draws no glyph for an empty paragraph, only the line.
		if para.bullet == nil {
			hasText := false
			for _, elem := range para.elements {
				if _, ok := elem.(*TextRun); ok {
					hasText = true
					break
				}
			}
			if hasText {
				switch {
				case s.buChar != "":
					b := NewBullet()
					b.Type = BulletTypeChar
					b.Style = s.buChar
					if s.buFont != "" {
						b.Font = s.buFont
					}
					para.bullet = b
				case s.buAutoNum != "":
					b := NewBullet()
					b.Type = BulletTypeNumeric
					b.NumFormat = s.buAutoNum
					if s.buStartAt > 0 {
						b.StartAt = s.buStartAt
					}
					if s.buFont != "" {
						b.Font = s.buFont
					}
					para.bullet = b
				}
			}
		}

		for _, elem := range para.elements {
			tr, ok := elem.(*TextRun)
			if !ok || tr.font == nil {
				continue
			}
			if tr.font.Name == "Calibri" && s.fontName != "" {
				tr.font.Name = s.fontName
			}
			if tr.font.NameEA == "" && s.fontEA != "" {
				tr.font.NameEA = s.fontEA
			}
			// 18 is the size the reader gives a placeholder run whose markup
			// declared none, 10 is what NewFont starts at.
			if runSizeIsDefault(tr.font.Size) && s.size > 0 {
				tr.font.Size = s.size
			}
			if s.bold {
				tr.font.Bold = true
			}
			if s.italic {
				tr.font.Italic = true
			}
			if s.color.ARGB != "" && s.color.ARGB != "FF000000" && tr.font.Color.ARGB == "FF000000" {
				tr.font.Color = s.color
			}
		}

		// A runless paragraph's line box rides its <a:endParaRPr> (round 26),
		// but PowerPoint-written empty paragraphs usually declare no sz there
		// — the size then comes from the same ladder the sibling runs resolve
		// through. slide36's blank line between two groups advances a full
		// 1.2 × 32pt line in the export where the 14px fallback left it 71px
		// short, dragging everything below it up. Fill it from the level
		// style when the paragraph itself said nothing; a paragraph with
		// runs keeps endParaRPrSize 0 (the writer must not grow the XML).
		if para.endParaRPrSize == 0 && s.size > 0 {
			hasText := false
			for _, elem := range para.elements {
				if _, ok := elem.(*TextRun); ok {
					hasText = true
					break
				}
			}
			if !hasText {
				para.endParaRPrSize = s.size * 100
			}
		}
	}
}

// parsePlaceholderDefs extracts placeholder definitions from a slide layout or
// a slide master XML.
//
// One function serves both because they are the same markup: a master's
// <p:sp> with a <p:ph> carries the geometry a placeholder inherits, and so does
// a layout's. The master is the farther rung of the ladder, and it is the only
// one that answers for the title in a deck whose layout leaves the title's
// <p:spPr/> empty — which is not rare, it is what PowerPoint writes.
func (r *PPTXReader) parsePlaceholderDefs(data []byte, pres *Presentation) []layoutPlaceholder {
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
	var phAnchor TextAnchorType
	var phAnchorSet bool
	var phAlign string
	var phAlignSet bool

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
				phAnchor = TextAnchorNone
				phAnchorSet = false
				phAlign = ""
				phAlignSet = false
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
						case "anchor":
							// Only the three anchors the model carries; the
							// schema's remaining values ("just", "dist",
							// "justLow") have no model counterpart and inherit
							// as nothing rather than being guessed.
							phAnchorSet = true
							switch attr.Value {
							case "t":
								phAnchor = TextAnchorTop
							case "ctr":
								phAnchor = TextAnchorMiddle
							case "b":
								phAnchor = TextAnchorBottom
							default:
								phAnchor = TextAnchorNone
							}
						}
					}
				}
			case "lstStyle":
				if inTxBody {
					inLstStyle = true
				}
			case "lvl1pPr":
				// The placeholder's own level-1 paragraph properties. The
				// alignment here is the one a slide's placeholder inherits —
				// the master's sldNum placeholder carries algn="r" exactly
				// here — and none of the <p:txStyles> tables say it.
				if inLstStyle {
					for _, attr := range t.Attr {
						if attr.Name.Local == "algn" {
							phAlign = attr.Value
							phAlignSet = true
						}
					}
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
			case "tint", "shade":
				// A transform nested in the defRPr's scheme colour — the
				// master's sldNum placeholder shades tx1 this way
				// (<a:tint val="75000"/>) and the page number inherits the
				// grey that results. Mixing in linear light, per the COM
				// variants that fixed applyTint.
				if inDefSolidFill && fontColor.ARGB != "" {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil && v > 0 {
								amount := float64(v) / 100000
								if t.Name.Local == "tint" {
									applyTint(&fontColor, amount)
								} else {
									applyShade(&fontColor, amount)
								}
							}
						}
					}
				}
			case "latin":
				// The layout's placeholder font is the one place that used to
				// store "+mj-lt" verbatim, so a layout inherited a font name
				// that no font file answers to and the run fell back — with
				// the fallback's metrics, which move where lines wrap.
				if inDefRPr {
					if n := typefaceOf(pres, t.Attr); n != "" {
						fontName = n
					}
				}
			case "ea":
				if inDefRPr {
					if n := typefaceOf(pres, t.Attr); n != "" {
						fontEA = n
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
						anchorSet:   phAnchorSet,
						anchor:      phAnchor,
						alignH:      phAlign,
						alignSet:    phAlignSet,
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

	// The background picture is not returned the moment its <a:blip> is seen:
	// <a:srcRect> and <a:alphaModFix> are siblings that follow it, so bailing
	// out there is what made a cropped layout background render stretched.
	var bgDS *DrawingShape

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
			case "blip", "svgBlip":
				// Both tags name the background part through r:embed; which one
				// carries it depends on whether the artwork is a raster or an
				// SVG written through the Microsoft extension. A raster on
				// <a:blip> wins over the SVG fallback, so the first hit is kept.
				if inBlipFill && bgDS == nil {
					if data, mime := imageRelData(rels, layoutPath, zr, embedRelID(t)); data != nil {
						bgDS = NewDrawingShape()
						bgDS.data = data
						bgDS.mimeType = mime
					}
				}
			case "srcRect":
				if inBlipFill && bgDS != nil {
					bgDS.cropLeft, bgDS.cropTop, bgDS.cropRight, bgDS.cropBottom = cropFromAttrs(t)
				}
			case "alphaModFix":
				if inBlipFill && bgDS != nil {
					if v, ok := intAttrValue(t, "amt"); ok {
						bgDS.alpha = v
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
				if bgDS != nil {
					return nil, bgDS
				}
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
	inFld := false
	fldType := ""
	inRunProps := false
	inRunSolidFill := false
	inText := false
	inPPr := false
	inPPrDefRPr := false
	inPPrDefSolidFill := false
	inLn := false
	inLnSolidFill := false
	isPH := false
	shapeHidden := false
	var offX, offY, extCX, extCY int64
	var embedID string
	var flipH, flipV bool
	var picAlpha int                   // alphaModFix amount for pic blip
	var cropL, cropT, cropR, cropB int // srcRect crop percentages
	var picLumBright, picLumContrast int

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
				picLumBright, picLumContrast = 0, 0
				cropL, cropT, cropR, cropB = 0, 0, 0, 0
				shapeHidden = false
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
				shapeHidden = false
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
			case "cNvPr":
				if inNvSpPr {
					for _, attr := range t.Attr {
						if attr.Name.Local == "hidden" {
							shapeHidden = attr.Value == "1" || attr.Value == "true"
						}
					}
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
					if id := embedRelID(t); id != "" {
						embedID = id
					}
				}
			case "svgBlip":
				// SVG-only artwork skips the raster reference on <a:blip> and
				// names the part here instead (see parseSlideXML). Recorded only
				// when the parent blip named nothing, so a raster fallback still
				// takes precedence.
				if inPic && embedID == "" {
					embedID = embedRelID(t)
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
			case "lum":
				// <a:lum bright contrast> under a picture's blip — the same
				// adjustment parseSlideXML collects, read here so a picture
				// that lives on a slide layout is not the one place in the
				// deck that ignores it.
				if inPic {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "bright":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								picLumBright = v
							}
						case "contrast":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								picLumContrast = v
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
					lnSpcReductionVal := 0
					for _, attr := range t.Attr {
						if attr.Name.Local == "fontScale" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								fontScaleVal = v
							}
						}
						if attr.Name.Local == "lnSpcReduction" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								lnSpcReductionVal = v
							}
						}
					}
					currentRichText.autoFit = AutoFitNormal
					currentRichText.fontScale = fontScaleVal
					currentRichText.lnSpcReduction = lnSpcReductionVal
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
					if n := typefaceOf(pres, t.Attr); n != "" {
						lstStyleFont.Name = n
					}
				} else if inPPrDefRPr && defFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						defFont.Name = n
					}
				} else if inRunProps && currentFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentFont.Name = n
					}
				}
			case "ea":
				if inDefRPr && lstStyleFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						lstStyleFont.NameEA = n
					}
				} else if inPPrDefRPr && defFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						defFont.NameEA = n
					}
				} else if inRunProps && currentFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentFont.NameEA = n
					}
				}
			case "sym":
				// Same declaration the main scanner reads: a layout-sourced
				// symbol run without it draws its PUA characters as tofu.
				if inDefRPr && lstStyleFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						lstStyleFont.NameSym = n
					}
				} else if inPPrDefRPr && defFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						defFont.NameSym = n
					}
				} else if inRunProps && currentFont != nil {
					if n := typefaceOf(pres, t.Attr); n != "" {
						currentFont.NameSym = n
					}
				}
			case "br":
				// A manual line break is a paragraph element and the renderer
				// draws one, but this reader never looked for it, so a text
				// shape contributed by a layout was drawn as a single long line.
				if inParagraph && currentParagraph != nil {
					currentParagraph.CreateBreak()
				}
			case "p":
				if inTxBody && currentRichText != nil {
					inParagraph = true
					currentParagraph = NewParagraph()
					currentRichText.paragraphs = append(currentRichText.paragraphs, currentParagraph)
				}
			case "pPr":
				if inParagraph && currentParagraph != nil {
					inPPr = true
					applyPPrAttrs(currentParagraph, t.Attr)
				}
			case "endParaRPr":
				// Same read as the slide scanner: an empty paragraph's only
				// size statement lives here, and the layout scanner must not
				// drop it or its empty lines collapse to the fallback height.
				if inParagraph && currentParagraph != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "sz" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								currentParagraph.endParaRPrSize = v
							}
						}
					}
				}
			case "r", "fld":
				if t.Name.Local == "fld" {
					// Same field handling as the slide scanner: an <a:fld> in a
					// layout text body wraps run content, so ride the run
					// machinery and keep the type for the renderer.
					inFld = true
					fldType = ""
					for _, attr := range t.Attr {
						if attr.Name.Local == "type" {
							fldType = attr.Value
						}
					}
				}
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
						case "baseline":
							// Same shift the main scanner reads: without it a
							// superscript picked up from a layout draws flat.
							if v, err := strconv.Atoi(attr.Value); err == nil {
								if v > 0 {
									currentFont.Superscript = true
								} else if v < 0 {
									currentFont.Subscript = true
								}
							}
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
				if inFld {
					tr.fieldType = fldType
				}
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
								ds.lumBright = picLumBright
								ds.lumContrast = picLumContrast
								ds.cropLeft = cropL
								ds.cropTop = cropT
								ds.cropRight = cropR
								ds.cropBottom = cropB
								ds.hidden = shapeHidden
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
						currentRichText.hidden = shapeHidden
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
					currentLine.hidden = shapeHidden
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
			case "fld":
				inFld = false
				inRun = false
				fldType = ""
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
