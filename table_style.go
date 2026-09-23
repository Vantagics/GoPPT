package gopresentation

import (
	"bytes"
	"encoding/xml"
	"io"
	"strconv"
	"strings"

	"archive/zip"
)

// tableStyleLine is one <a:ln> declaration inside a style part's <a:tcBdr>,
// one of left/right/top/bottom/insideH/insideV. Declared distinguishes "the
// element is absent" (the wholeTbl part's line shows through) from an explicit
// <a:noFill/> (the line is suppressed even when another cell's declaration is
// heavier).
type tableStyleLine struct {
	declared bool
	noFill   bool
	width    int // points, from the ln w attribute
	scheme   string
	tint     float64 // -1: none; else 0..1
	shade    float64 // -1: none; else 0..1
	lumMod   float64 // -1: none; else 0..1
	lumOff   float64 // -1: none; else 0..1
}

// tableStyleBorders holds the six line declarations of one part's tcBdr.
type tableStyleBorders struct {
	left, right, top, bottom, insideH, insideV tableStyleLine
}

// tableStyleBand is one named part of a table style — wholeTbl, a band, or an
// edge row/column. A part may declare a fill (a scheme colour plus the colour
// transforms written inside it), a border treatment (six tcBdr lines), and a
// text treatment (bold, and the tcTxStyle colour).
type tableStyleBand struct {
	hasFill    bool
	scheme     string  // scheme colour name, e.g. "accent1"
	tint       float64 // -1: none; else 0..1
	shade      float64 // -1: none; else 0..1
	lumMod     float64 // -1: none; else 0..1
	lumOff     float64 // -1: none; else 0..1
	bold       bool
	hasTextMod bool
	// tcTxStyle's own colour reference (a schemeClr sibling of the fontRef,
	// not a fill colour). Empty when the part declares none.
	textScheme string
	textTint   float64 // -1: none; else 0..1
	textShade  float64 // -1: none; else 0..1
	textLumMod float64 // -1: none; else 0..1
	textLumOff float64 // -1: none; else 0..1
	borders    tableStyleBorders
}

// tableStyle is one <a:tblStyle> from ppt/tableStyles.xml, keyed by the GUID a
// table's <a:tableStyleId> names.
type tableStyle struct {
	id       string
	wholeTbl tableStyleBand
	band1H   tableStyleBand
	band2H   tableStyleBand
	band1V   tableStyleBand
	band2V   tableStyleBand
	firstRow tableStyleBand
	lastRow  tableStyleBand
	firstCol tableStyleBand
	lastCol  tableStyleBand
}

// readTableStyles reads ppt/tableStyles.xml into the presentation. The styles
// resolve a table's fills when its cells declare none of their own.
func (r *PPTXReader) readTableStyles(zr *zip.Reader, pres *Presentation) {
	data, err := readFileFromZip(zr, "ppt/tableStyles.xml")
	if err != nil {
		return
	}
	styles, def := parseTableStyles(data)
	pres.tableStyles = styles
	pres.defaultTableStyle = def
}

// parseTableStyles parses a <a:tblStyleLst> document. The list's def attribute
// is returned separately: it is the style a table with no <a:tableStyleId> of
// its own falls back to.
func parseTableStyles(data []byte) (styles map[string]*tableStyle, def string) {
	styles = make(map[string]*tableStyle)
	dec := xml.NewDecoder(bytes.NewReader(data))
	dec.Strict = false
	var cur *tableStyle
	var band *tableStyleBand
	// colour-transform accumulation for the band's scheme colour
	var inScheme bool
	var inFill bool
	// tcBdr line accumulation: the current side element and its <a:ln>
	var inTcBdr bool
	var inTcBdrSide string
	var bdrLine *tableStyleLine
	var bdrW int
	// tcTxStyle colour accumulation
	var inTcTx bool
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return styles, def
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblStyleLst":
				def = attrVal(t, "def")
			case "tblStyle":
				cur = &tableStyle{}
				band = nil
				for _, a := range t.Attr {
					if a.Name.Local == "styleId" {
						cur.id = a.Value
					}
				}
			case "wholeTbl":
				if cur != nil {
					band = &cur.wholeTbl
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "band1H":
				if cur != nil {
					band = &cur.band1H
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "band2H":
				if cur != nil {
					band = &cur.band2H
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "band1V":
				if cur != nil {
					band = &cur.band1V
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "band2V":
				if cur != nil {
					band = &cur.band2V
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "firstRow":
				if cur != nil {
					band = &cur.firstRow
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "lastRow":
				if cur != nil {
					band = &cur.lastRow
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "firstCol":
				if cur != nil {
					band = &cur.firstCol
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "lastCol":
				if cur != nil {
					band = &cur.lastCol
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1,
						textTint: -1, textShade: -1, textLumMod: -1, textLumOff: -1}
				}
			case "fill":
				// The band's own fill; the colours inside it are the ones
				// that count (a tcBdr's line colours are also schemeClr).
				inFill = true
			case "tcStyle":
				// container; nothing to do
			case "tcBdr":
				inTcBdr = true
				inTcBdrSide = ""
				bdrLine = nil
			case "left", "right", "top", "bottom", "insideH", "insideV":
				if inTcBdr && band != nil {
					inTcBdrSide = t.Name.Local
					switch inTcBdrSide {
					case "left":
						bdrLine = &band.borders.left
					case "right":
						bdrLine = &band.borders.right
					case "top":
						bdrLine = &band.borders.top
					case "bottom":
						bdrLine = &band.borders.bottom
					case "insideH":
						bdrLine = &band.borders.insideH
					case "insideV":
						bdrLine = &band.borders.insideV
					}
					*bdrLine = tableStyleLine{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "ln":
				if inTcBdrSide != "" && bdrLine != nil {
					bdrLine.declared = true
					bdrW = 0
					for _, a := range t.Attr {
						if a.Name.Local == "w" {
							if v, err := strconv.Atoi(a.Value); err == nil {
								bdrW = v
							}
						}
					}
				}
			case "tcTxStyle":
				if band != nil {
					band.hasTextMod = true
					inTcTx = true
					for _, a := range t.Attr {
						if a.Name.Local == "b" {
							band.bold = a.Value == "1" || a.Value == "on" || a.Value == "true"
						}
					}
				}
			case "schemeClr":
				if bdrLine != nil && inTcBdrSide != "" {
					// a tcBdr line's colour reference
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							bdrLine.scheme = a.Value
						}
					}
				} else if inTcTx && band != nil {
					// the tcTxStyle's own colour reference
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							band.textScheme = a.Value
						}
					}
				} else if band != nil && inFill {
					band.hasFill = true
					inScheme = true
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							band.scheme = a.Value
						}
					}
				}
			case "srgbClr":
				// A literal colour instead of a scheme reference; resolve
				// it through a synthetic name no theme carries. It can sit
				// directly in the band's fill (no schemeClr wrapper), so the
				// guard is the fill, not the scheme — tcBdr line colours and
				// tcTxStyle text colours both live outside <a:fill> and stay
				// excluded.
				if band != nil && inFill {
					band.hasFill = true
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							band.scheme = "srgb:" + a.Value
						}
					}
				} else if bdrLine != nil && inTcBdrSide != "" {
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							bdrLine.scheme = "srgb:" + a.Value
						}
					}
				} else if inTcTx && band != nil {
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							band.textScheme = "srgb:" + a.Value
						}
					}
				}
			case "noFill":
				if bdrLine != nil && inTcBdrSide != "" {
					bdrLine.noFill = true
				}
			case "tint":
				if bdrLine != nil && inTcBdrSide != "" {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						bdrLine.tint = v / 100000.0
					}
				} else if inTcTx && band != nil {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.textTint = v / 100000.0
					}
				} else if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.tint = v / 100000.0
					}
				}
			case "shade":
				if bdrLine != nil && inTcBdrSide != "" {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						bdrLine.shade = v / 100000.0
					}
				} else if inTcTx && band != nil {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.textShade = v / 100000.0
					}
				} else if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.shade = v / 100000.0
					}
				}
			case "lumMod":
				if bdrLine != nil && inTcBdrSide != "" {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						bdrLine.lumMod = v / 100000.0
					}
				} else if inTcTx && band != nil {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.textLumMod = v / 100000.0
					}
				} else if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.lumMod = v / 100000.0
					}
				}
			case "lumOff":
				if bdrLine != nil && inTcBdrSide != "" {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						bdrLine.lumOff = v / 100000.0
					}
				} else if inTcTx && band != nil {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.textLumOff = v / 100000.0
					}
				} else if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.lumOff = v / 100000.0
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "schemeClr", "srgbClr":
				inScheme = false
			case "fill":
				inFill = false
			case "ln":
				if bdrLine != nil && inTcBdrSide != "" {
					bdrLine.width = bdrW / 12700
				}
			case "left", "right", "top", "bottom", "insideH", "insideV":
				if inTcBdr {
					inTcBdrSide = ""
					bdrLine = nil
				}
			case "tcBdr":
				inTcBdr = false
				inTcBdrSide = ""
				bdrLine = nil
			case "tcTxStyle":
				inTcTx = false
			case "tblStyle":
				if cur != nil && cur.id != "" {
					styles[cur.id] = cur
				}
				cur = nil
				band = nil
			case "wholeTbl", "band1H", "band2H", "band1V", "band2V",
				"firstRow", "lastRow", "firstCol", "lastCol":
				band = nil
				bdrLine = nil
				inTcBdrSide = ""
				inTcTx = false
			}
		}
	}
	return styles, def
}

func attrVal(t xml.StartElement, name string) string {
	for _, a := range t.Attr {
		if a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}

// applyColorTransforms folds a scheme colour's transforms into an ARGB value.
//
// The four transforms do not share a colour space. tint and shade mix with
// white / black in linear light with the value counting the colour that
// survives (see applyTint), while lumMod and lumOff scale and shift the HLS
// luminance — the r29 COM probe confirms both halves on the same slide: with
// accent1 #4F81BD, lumMod 50000 renders #254061 and lumMod 60000 + lumOff
// 40000 renders #95B3D7, both of which are the HLS answers (a linear reading
// would say #385D8A and #7BA1CF), whereas the tint values are the linear ones.
//
// When a colour carries both kinds the tint is resolved first, so the HLS
// luminance that lumMod scales is the tinted colour's.
func applyColorTransforms(c Color, tint, shade, lumMod, lumOff float64) Color {
	if tint >= 0 {
		applyTint(&c, tint)
	}
	if shade >= 0 {
		applyShade(&c, shade)
	}
	if lumMod >= 0 || lumOff >= 0 {
		h, l, s := rgbToHLS(float64(c.GetRed())/255, float64(c.GetGreen())/255, float64(c.GetBlue())/255)
		if lumMod >= 0 {
			l = l * lumMod
		}
		if lumOff >= 0 {
			l = l + lumOff
		}
		if l > 1 {
			l = 1
		}
		if l < 0 {
			l = 0
		}
		r, g, b := hlsToRGB(h, l, s)
		c = Color{ARGB: fmtARGB(uint8(r*255+0.5), uint8(g*255+0.5), uint8(b*255+0.5))}
	}
	return c
}

func fmtARGB(r, g, b uint8) string {
	const hex = "0123456789ABCDEF"
	out := [8]byte{'F', 'F', 0, 0, 0, 0, 0, 0}
	out[2] = hex[r>>4]
	out[3] = hex[r&0xF]
	out[4] = hex[g>>4]
	out[5] = hex[g&0xF]
	out[6] = hex[b>>4]
	out[7] = hex[b&0xF]
	return string(out[:])
}

// rgbToHLS / hlsToRGB are the classic double-hexcone conversions, on 0..1.
func rgbToHLS(r, g, b float64) (h, l, s float64) {
	mx := maxf(r, g, b)
	mn := minf(r, g, b)
	l = (mx + mn) / 2
	if mx == mn {
		return 0, l, 0
	}
	d := mx - mn
	if l > 0.5 {
		s = d / (2 - mx - mn)
	} else {
		s = d / (mx + mn)
	}
	switch mx {
	case r:
		h = (g - b) / d
		if g < b {
			h += 6
		}
	case g:
		h = (b-r)/d + 2
	default:
		h = (r-g)/d + 4
	}
	h /= 6
	return h, l, s
}

func hlsToRGB(h, l, s float64) (r, g, b float64) {
	if s == 0 {
		return l, l, l
	}
	var q float64
	if l < 0.5 {
		q = l * (1 + s)
	} else {
		q = l + s - l*s
	}
	p := 2*l - q
	return hueToRGB(p, q, h+1.0/3.0), hueToRGB(p, q, h), hueToRGB(p, q, h-1.0/3.0)
}

func hueToRGB(p, q, t float64) float64 {
	if t < 0 {
		t += 1
	}
	if t > 1 {
		t -= 1
	}
	switch {
	case t < 1.0/6.0:
		return p + (q-p)*6*t
	case t < 0.5:
		return q
	case t < 2.0/3.0:
		return p + (q-p)*(2.0/3.0-t)*6
	}
	return p
}

func maxf(a, b, c float64) float64 {
	if b > a {
		a = b
	}
	if c > a {
		a = c
	}
	return a
}

func minf(a, b, c float64) float64 {
	if b < a {
		a = b
	}
	if c < a {
		a = c
	}
	return a
}

// applyTableStyleFill gives every cell that declares no fill of its own the
// fill its position earns under the table's style: an edge row/column part
// when the matching tblPr flag is on, a band otherwise, wholeTbl as the base.
func applyTableStyleFill(s *TableShape, pres *Presentation) {
	if s == nil || pres == nil || len(pres.tableStyles) == 0 {
		return
	}
	st := pres.tableStyles[s.styleGUID]
	if st == nil {
		st = pres.tableStyles[pres.defaultTableStyle]
	}
	if st == nil {
		return
	}
	for ri, row := range s.rows {
		// Data-row index for banding: the rows that survive the edge flags.
		dataRow := ri
		if s.firstRow && ri == 0 {
			dataRow = -1
		} else if s.lastRow && ri == s.numRows-1 {
			dataRow = -2
		} else if s.firstRow && ri > 0 {
			dataRow = ri - 1
		}
		for ci, cell := range row {
			// The tcTxStyle colour applies to every cell of the band
			// regardless of fills: the firstRow part of the r33 deck says
			// lt1, and its header text is white while we painted it black.
			// Only runs without a colour of their own (the reader's default
			// black) take it.
			if cell != nil {
				band, _ := tableStyleBandFor(s, st, ri, ci, dataRow)
				applyBandTextColor(cell, pres, band)
			}
			// FillNone covers both "never declared" and an explicit
			// <a:noFill/>; the two are not distinguished here, and the style
			// wins over either.
			if cell == nil || cell.fill != nil && cell.fill.Type != FillNone {
				continue
			}
			band, ok := tableStyleBandFor(s, st, ri, ci, dataRow)
			if !ok {
				continue
			}
			if !band.hasFill || band.scheme == "" {
				// wholeTbl is the fill *under* the whole table, so a band
				// that declares no fill of its own shows it rather than
				// nothing. The r29 deck is the proof: its band2H is empty
				// and its wholeTbl is accent1 tint 20000, which is exactly
				// the #E9EDF4 PowerPoint paints on every even row — with
				// the fallback missing those rows came out plain white.
				band = st.wholeTbl
				if !band.hasFill || band.scheme == "" {
					continue
				}
			}
			argb := resolveStyleColor(pres, band.scheme)
			if argb == "" {
				continue
			}
			c := applyColorTransforms(NewColor(argb), band.tint, band.shade, band.lumMod, band.lumOff)
			cell.fill = NewFill()
			cell.fill.SetSolid(c)
			if band.hasTextMod && band.bold {
				for _, para := range cell.paragraphs {
					for _, elem := range para.elements {
						if tr, ok := elem.(*TextRun); ok && tr.font != nil && !tr.font.Bold {
							tr.font.Bold = true
						}
					}
				}
			}
		}
	}
}

// tableStyleBandFor picks the band a cell's position earns. The bool reports
// whether any part applied.
func tableStyleBandFor(s *TableShape, st *tableStyle, ri, ci, dataRow int) (tableStyleBand, bool) {
	// Edge parts first: a flagged edge claims its row or column outright.
	if ri == 0 && s.firstRow {
		return st.firstRow, true
	}
	if ri == s.numRows-1 && s.lastRow {
		return st.lastRow, true
	}
	if ci == 0 && s.firstCol {
		return st.firstCol, true
	}
	if ci == s.numCols-1 && s.lastCol {
		return st.lastCol, true
	}
	// Bands. Horizontal bands alternate over the data rows (or, with the row
	// flags off, over every row).
	if s.bandRow {
		idx := dataRow
		if idx < 0 {
			idx = ri
		}
		if idx%2 == 0 {
			return st.band1H, true
		}
		return st.band2H, true
	}
	if s.bandCol {
		if ci%2 == 0 {
			return st.band1V, true
		}
		return st.band2V, true
	}
	return st.wholeTbl, true
}

// resolveStyleColor maps a style's colour reference to an ARGB string. Scheme
// names go through the theme; the "srgb:RRGGBB" pseudo-name is what the
// parser stores for a literal colour.
func resolveStyleColor(pres *Presentation, scheme string) string {
	if argb, ok := pres.themeColors[scheme]; ok && argb != "" {
		return argb
	}
	if v, ok := strings.CutPrefix(scheme, "srgb:"); ok && len(v) == 6 {
		return "FF" + v
	}
	return ""
}

// applyBandTextColor gives a cell's runs the tcTxStyle colour of the band
// their position earns, but only where a run carries no colour of its own —
// the reader leaves an unstyled run at the default black, and "FF000000" is
// the same convention the lstStyle defaults already use.
func applyBandTextColor(cell *TableCell, pres *Presentation, band tableStyleBand) {
	if !band.hasTextMod || band.textScheme == "" {
		return
	}
	argb := resolveStyleColor(pres, band.textScheme)
	if argb == "" {
		return
	}
	c := applyColorTransforms(NewColor(argb), band.textTint, band.textShade, band.textLumMod, band.textLumOff)
	for _, para := range cell.paragraphs {
		for _, elem := range para.elements {
			tr, ok := elem.(*TextRun)
			if !ok || tr.font == nil {
				continue
			}
			if tr.font.Color.ARGB == "" || tr.font.Color.ARGB == "FF000000" {
				tr.font.Color = c
			}
		}
	}
}

// applyTableStyleBorders gives every cell the lines its position earns under
// the table's style, for the sides the cell's own <a:tcPr> left undeclared.
//
// The mapping is positional, per the style part the cell lands in: a part's
// left/right/top/bottom line applies to that side of its cells (which for
// wholeTbl means the table's perimeter, since only perimeter cells have that
// side "outward"), and insideH/insideV to the interior edges. A part that
// declares no line for a side shows the wholeTbl part's line for it — the
// same fallback the fills already follow. Where two cells declare conflicting
// lines for one shared edge, the heavier line wins (firstRow's 3pt bottom
// separator beats the 1pt insideH of the row below it).
func applyTableStyleBorders(s *TableShape, pres *Presentation) {
	if s == nil || pres == nil || len(pres.tableStyles) == 0 {
		return
	}
	st := pres.tableStyles[s.styleGUID]
	if st == nil {
		st = pres.tableStyles[pres.defaultTableStyle]
	}
	if st == nil {
		return
	}
	lineFor := func(part, whole tableStyleLine) tableStyleLine {
		if part.declared {
			return part
		}
		return whole
	}
	for ri, row := range s.rows {
		dataRow := ri
		if s.firstRow && ri == 0 {
			dataRow = -1
		} else if s.lastRow && ri == s.numRows-1 {
			dataRow = -2
		} else if s.firstRow && ri > 0 {
			dataRow = ri - 1
		}
		for ci, cell := range row {
			if cell == nil || cell.border == nil {
				continue
			}
			band, _ := tableStyleBandFor(s, st, ri, ci, dataRow)
			lastRow := ri == s.numRows-1
			lastCol := ci == s.numCols-1
			edges := [4]tableStyleLine{
				lineFor(band.borders.top, pickWholeTblEdge(&st.wholeTbl.borders, ri == 0, true)),
				lineFor(band.borders.bottom, pickWholeTblEdge(&st.wholeTbl.borders, lastRow, true)),
				lineFor(band.borders.left, pickWholeTblEdge(&st.wholeTbl.borders, ci == 0, false)),
				lineFor(band.borders.right, pickWholeTblEdge(&st.wholeTbl.borders, lastCol, false)),
			}
			// Interior edges compete with the neighbour's declaration: the
			// heavier line wins (the firstRow part's 3pt bottom beats the
			// 1pt insideH the row below resolves for its top).
			if ri > 0 {
				edges[0] = heavierLine(edges[0], neighbourLine(s, st, ri-1, ci, dataRow, "bottom"))
			}
			if !lastRow {
				edges[1] = heavierLine(edges[1], neighbourLine(s, st, ri+1, ci, dataRow, "top"))
			}
			if ci > 0 {
				edges[2] = heavierLine(edges[2], neighbourLine(s, st, ri, ci-1, dataRow, "right"))
			}
			if !lastCol {
				edges[3] = heavierLine(edges[3], neighbourLine(s, st, ri, ci+1, dataRow, "left"))
			}
			if !cell.border.topDeclared {
				setCellSide(pres, cell.border.Top, edges[0])
			}
			if !cell.border.bottomDeclared {
				setCellSide(pres, cell.border.Bottom, edges[1])
			}
			if !cell.border.leftDeclared {
				setCellSide(pres, cell.border.Left, edges[2])
			}
			if !cell.border.rightDeclared {
				setCellSide(pres, cell.border.Right, edges[3])
			}
		}
	}
}

// pickWholeTblEdge picks the wholeTbl line an interior edge falls back to:
// insideH for horizontal edges, insideV for vertical ones.
func pickWholeTblEdge(b *tableStyleBorders, outer bool, horizontal bool) tableStyleLine {
	if outer {
		if horizontal {
			return b.top
		}
		return b.left
	}
	if horizontal {
		return b.insideH
	}
	return b.insideV
}

// neighbourLine resolves the line a neighbouring cell declares toward this
// cell across their shared edge, so a heavy separator (firstRow's bottom)
// wins over the lighter insideH the row below resolves for its top.
func neighbourLine(s *TableShape, st *tableStyle, ri, ci, dataRow int, toward string) tableStyleLine {
	if ri < 0 || ri >= len(s.rows) || ci < 0 || ci >= len(s.rows[ri]) {
		return tableStyleLine{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
	}
	band, _ := tableStyleBandFor(s, st, ri, ci, dataRow)
	whole := &st.wholeTbl.borders
	switch toward {
	case "bottom":
		return lineFallback(band.borders.bottom, whole.bottom, whole.insideH, ri == s.numRows-1)
	case "top":
		return lineFallback(band.borders.top, whole.top, whole.insideH, ri == 0)
	case "right":
		return lineFallback(band.borders.right, whole.right, whole.insideV, ci == s.numCols-1)
	case "left":
		return lineFallback(band.borders.left, whole.left, whole.insideV, ci == 0)
	}
	return tableStyleLine{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
}

// lineFallback resolves one side of a band: the band's own line, else the
// wholeTbl line for that side when the edge is on the table's perimeter,
// else the wholeTbl interior line.
func lineFallback(part, wholeOuter, wholeInner tableStyleLine, outer bool) tableStyleLine {
	if part.declared {
		return part
	}
	if outer {
		return wholeOuter
	}
	return wholeInner
}

// heavierLine picks the line that wins a shared edge: the heavier width; an
// explicit noFill weighs nothing but still beats an undeclared side.
func heavierLine(a, b tableStyleLine) tableStyleLine {
	if !b.declared {
		return a
	}
	if !a.declared {
		return b
	}
	if b.width > a.width {
		return b
	}
	return a
}

// setCellSide writes one resolved style line into a cell's side, leaving the
// side untouched when the style declares nothing for it.
func setCellSide(pres *Presentation, b *Border, line tableStyleLine) {
	if !line.declared {
		return
	}
	if line.noFill || line.scheme == "" {
		b.Style = BorderNone
		b.Width = 0
		return
	}
	argb := resolveStyleColor(pres, line.scheme)
	if argb == "" {
		return
	}
	b.Style = BorderSolid
	b.Width = line.width
	b.Color = applyColorTransforms(NewColor(argb), line.tint, line.shade, line.lumMod, line.lumOff)
}
