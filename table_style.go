package gopresentation

import (
	"bytes"
	"encoding/xml"
	"io"
	"strconv"
	"strings"

	"archive/zip"
)

// tableStyleBand is one named part of a table style — wholeTbl, a band, or an
// edge row/column. A part may declare a fill (a scheme colour plus the colour
// transforms written inside it) and a text treatment (bold).
type tableStyleBand struct {
	hasFill    bool
	scheme     string  // scheme colour name, e.g. "accent1"
	tint       float64 // -1: none; else 0..1
	shade      float64 // -1: none; else 0..1
	lumMod     float64 // -1: none; else 0..1
	lumOff     float64 // -1: none; else 0..1
	bold       bool
	hasTextMod bool
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
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "band1H":
				if cur != nil {
					band = &cur.band1H
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "band2H":
				if cur != nil {
					band = &cur.band2H
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "band1V":
				if cur != nil {
					band = &cur.band1V
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "band2V":
				if cur != nil {
					band = &cur.band2V
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "firstRow":
				if cur != nil {
					band = &cur.firstRow
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "lastRow":
				if cur != nil {
					band = &cur.lastRow
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "firstCol":
				if cur != nil {
					band = &cur.firstCol
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "lastCol":
				if cur != nil {
					band = &cur.lastCol
					*band = tableStyleBand{tint: -1, shade: -1, lumMod: -1, lumOff: -1}
				}
			case "fill":
				// The band's own fill; the colours inside it are the ones
				// that count (a tcBdr's line colours are also schemeClr).
				inFill = true
			case "tcStyle", "tcBdr":
				// container; nothing to do
			case "tcTxStyle":
				if band != nil {
					band.hasTextMod = true
					for _, a := range t.Attr {
						if a.Name.Local == "b" {
							band.bold = a.Value == "1" || a.Value == "on" || a.Value == "true"
						}
					}
				}
			case "schemeClr":
				if band != nil && inFill {
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
				}
			case "tint":
				if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.tint = v / 100000.0
					}
				}
			case "shade":
				if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.shade = v / 100000.0
					}
				}
			case "lumMod":
				if band != nil && inScheme {
					if v, err := strconv.ParseFloat(attrVal(t, "val"), 64); err == nil {
						band.lumMod = v / 100000.0
					}
				}
			case "lumOff":
				if band != nil && inScheme {
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
			case "tblStyle":
				if cur != nil && cur.id != "" {
					styles[cur.id] = cur
				}
				cur = nil
				band = nil
			case "wholeTbl", "band1H", "band2H", "band1V", "band2V",
				"firstRow", "lastRow", "firstCol", "lastCol":
				band = nil
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
// tint and shade move the HLS luminance toward white / black, lumMod and
// lumOff scale and shift it — the same four the table styles use.
func applyColorTransforms(c Color, tint, shade, lumMod, lumOff float64) Color {
	h, l, s := rgbToHLS(float64(c.GetRed())/255, float64(c.GetGreen())/255, float64(c.GetBlue())/255)
	if tint >= 0 {
		l = l*(1-tint) + tint
	}
	if shade >= 0 {
		l = l * (1 - shade)
	}
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
	return Color{ARGB: fmtARGB(uint8(r*255+0.5), uint8(g*255+0.5), uint8(b*255+0.5))}
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
				continue
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
