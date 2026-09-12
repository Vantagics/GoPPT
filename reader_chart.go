package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"strconv"
	"strings"
)

// This file implements reading of native chart parts.
//
// In OOXML a chart embedded in a slide is a <p:graphicFrame> whose
// <a:graphicData> points at a separate chart part (ppt/charts/chartN.xml) via a
// relationship. The chart part uses the chart namespace (c:) and holds the
// chart type, the series data and the axes.
//
// GoPPT's renderer has always been able to rasterize a ChartShape, but nothing
// ever produced one when reading a file, so charts silently disappeared on a
// write→read round-trip. The code below resolves the relationship, parses the
// chart part and rebuilds an equivalent ChartShape so the existing rasterizer
// can draw it.

// readChartShape resolves a chart relationship from the slide's relationships
// and parses the referenced chart part. It returns nil when the relationship or
// the part cannot be resolved so callers can simply skip the graphicFrame.
func (r *PPTXReader) readChartShape(zr *zip.Reader, rels []xmlRelForRead, slidePath, relID string, themeColors map[string]string) *ChartShape {
	if relID == "" {
		return nil
	}
	for _, rel := range rels {
		if rel.ID != relID {
			continue
		}
		// Accept the canonical relationship type as well as vendor variants
		// that still end in "/chart".
		if rel.Type != "" && rel.Type != relTypeChart && !strings.HasSuffix(rel.Type, "/chart") {
			return nil
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
		return parseChartXML(data, themeColors)
	}
	return nil
}

// seriesPoints collects <c:pt> entries keyed by their point index. Indexes are
// not guaranteed to be contiguous (blank cells are skipped), so gaps are
// preserved as empty entries to keep categories and values aligned.
type seriesPoints struct {
	pts map[int]string
}

func (sp *seriesPoints) set(idx int, v string) {
	if sp.pts == nil {
		sp.pts = make(map[int]string)
	}
	sp.pts[idx] = v
}

func (sp *seriesPoints) empty() bool { return len(sp.pts) == 0 }

// ordered returns the collected values ordered by point index, preserving gaps.
func (sp *seriesPoints) ordered() []string {
	if len(sp.pts) == 0 {
		return nil
	}
	maxIdx := 0
	for idx := range sp.pts {
		if idx > maxIdx {
			maxIdx = idx
		}
	}
	out := make([]string, maxIdx+1)
	for idx, v := range sp.pts {
		out[idx] = v
	}
	return out
}

// orderedFloats is ordered() converted to float64; non-numeric or missing
// entries become 0.
func (sp *seriesPoints) orderedFloats() []float64 {
	raw := sp.ordered()
	if raw == nil {
		return nil
	}
	out := make([]float64, len(raw))
	for i, s := range raw {
		if f, err := strconv.ParseFloat(strings.TrimSpace(s), 64); err == nil {
			out[i] = f
		}
	}
	return out
}

// chartSeriesLabels mirrors the per-series <c:dLbls> block.
type chartSeriesLabels struct {
	showValue   bool
	showCatName bool
	showPercent bool
	showSerName bool
	separator   string
	position    string
}

// parseChartXML parses a chart part (chartSpace) into a ChartShape. It returns
// nil when no supported plot type with usable series is present.
func parseChartXML(data []byte, themeColors map[string]string) *ChartShape {
	chart := NewChartShape()
	// A chart part only shows a title/legend when it declares one, so start
	// from the "absent" state: a missing element means "not shown".
	chart.title.Visible = false
	chart.legend.Visible = false

	dec := xml.NewDecoder(bytes.NewReader(data))
	dec.Strict = false

	var (
		inPlotArea   bool
		inSeries     bool
		inSeriesSpPr bool
		inSeriesLn   bool
		inChartSpPr  bool
		inChartLn    bool
		chartNoFill  bool
		inMarker     bool
		inDLbls      bool
		inGridlines  bool
		inAxis       bool
		axisIsValue  bool
		inColor      bool
		inV          bool
		inT          bool
		inSeparator  bool
		titleTarget  string // "", "chart" or "axis"
		valueCtx     string // "", "serTitle", "cat", "val", "xval", "yval"

		plotType    string
		barDir      string
		barGrouping string
		gapWidth    *int
		overlap     *int
		holeSize    *int
		lineSmooth  bool

		cats      *seriesPoints
		vals      *seriesPoints
		xVals     *seriesPoints
		yVals     *seriesPoints
		serTitle  strings.Builder
		serFill   *Color
		serLine   *Color
		serLw     int
		serMark   *SeriesMarker
		serSmooth bool
		serDlbls  *chartSeriesLabels

		curAxis      *ChartAxis
		curGridlines *Gridlines
		gridColor    *Color
		gridWidth    int

		chartFillColor *Color
		chartLineColor *Color
		chartLineWidth int

		ptIdx  int
		series []*ChartSeries

		vBuf       strings.Builder
		tBuf       strings.Builder
		sepBuf     strings.Builder
		chartTitle strings.Builder
		axisTitle  strings.Builder

		pendColor  *Color
		pendTarget string
	)

	var (
		chartTitleText string
	)

	resetSeries := func() {
		cats = &seriesPoints{}
		vals = &seriesPoints{}
		xVals = &seriesPoints{}
		yVals = &seriesPoints{}
		serTitle.Reset()
		serFill = nil
		serLine = nil
		serLw = 0
		serMark = nil
		serSmooth = false
		serDlbls = nil
	}

	finishSeries := func() {
		if cats == nil {
			return
		}
		catList := cats.ordered()
		valList := vals.ordered()
		// Scatter/bubble charts store X/Y instead of cat/val.
		if len(catList) == 0 && !xVals.empty() {
			catList = xVals.ordered()
		}
		if len(valList) == 0 && !yVals.empty() {
			valList = yVals.ordered()
		}

		if len(catList) == 0 {
			// No explicit categories: synthesise 1..n so the series stays
			// renderable and the ordering is deterministic.
			n := len(valList)
			catList = make([]string, n)
			for i := 0; i < n; i++ {
				catList[i] = strconv.Itoa(i + 1)
			}
		}
		if len(catList) == 0 {
			return
		}

		floats := make([]float64, len(catList))
		for i := range catList {
			if i < len(valList) {
				if f, err := strconv.ParseFloat(strings.TrimSpace(valList[i]), 64); err == nil {
					floats[i] = f
				}
			}
		}

		s := NewChartSeriesOrdered(strings.TrimSpace(serTitle.String()), catList, floats)
		if serFill != nil {
			s.FillColor = *serFill
		} else if serLine != nil {
			// Line/scatter series encode their colour on the outline.
			s.FillColor = *serLine
		}
		if serLine != nil || serLw > 0 {
			outline := &SeriesOutline{Width: serLw}
			if serLine != nil {
				outline.Color = *serLine
			}
			s.Outline = outline
		}
		if serMark != nil {
			s.Marker = serMark
		}
		if serSmooth {
			lineSmooth = true
		}
		if serDlbls != nil {
			s.ShowValue = serDlbls.showValue
			s.ShowCategoryName = serDlbls.showCatName
			s.ShowPercentage = serDlbls.showPercent
			s.ShowSeriesName = serDlbls.showSerName
			s.LabelPosition = serDlbls.position
			if serDlbls.separator != "" {
				s.Separator = serDlbls.separator
			}
		}
		series = append(series, s)
	}

	newAxis := func(isValue bool) {
		curAxis = NewChartAxis()
		axisIsValue = isValue
		axisTitle.Reset()
		inAxis = true
	}

	// commitColor stores the colour parsed for the pending <c:spPr>/<a:ln>
	// context. Every target the colour scanner can set has to appear here: a
	// missing case drops the colour silently, which is how the chart-area fill
	// and outline went unread for so long — the scanner routed them to
	// "chartFill"/"chartLine" but nothing consumed those, so chart.fill and
	// chart.border stayed nil for every document opened from disk.
	commitColor := func(c Color) {
		switch pendTarget {
		case "serFill":
			cc := c
			serFill = &cc
		case "serLine":
			cc := c
			serLine = &cc
		case "gridline":
			cc := c
			gridColor = &cc
		case "chartFill":
			cc := c
			chartFillColor = &cc
		case "chartLine":
			cc := c
			chartLineColor = &cc
		}
	}

	for {
		token, err := dec.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			name := t.Name.Local
			val := attrValue(t, "val")

			switch name {
			case "plotArea":
				inPlotArea = true

			case "barChart":
				if plotType == "" {
					plotType = "bar"
				}
			case "bar3DChart":
				if plotType == "" {
					plotType = "bar3D"
				}
			case "lineChart":
				if plotType == "" {
					plotType = "line"
				}
			case "areaChart":
				if plotType == "" {
					plotType = "area"
				}
			case "pieChart", "ofPieChart":
				if plotType == "" {
					plotType = "pie"
				}
			case "pie3DChart":
				if plotType == "" {
					plotType = "pie3D"
				}
			case "doughnutChart":
				if plotType == "" {
					plotType = "doughnut"
				}
			case "scatterChart", "bubbleChart":
				if plotType == "" {
					plotType = "scatter"
				}
			case "radarChart":
				if plotType == "" {
					plotType = "radar"
				}

			case "ser":
				inSeries = true
				resetSeries()

			case "tx":
				if inSeries && !inDLbls {
					valueCtx = "serTitle"
				}
			case "cat":
				if inSeries {
					valueCtx = "cat"
				}
			case "val":
				if inSeries {
					valueCtx = "val"
				}
			case "xVal":
				if inSeries {
					valueCtx = "xval"
				}
			case "yVal":
				if inSeries {
					valueCtx = "yval"
				}

			case "pt":
				ptIdx = 0
				if v := attrValue(t, "idx"); v != "" {
					if n, err := strconv.Atoi(v); err == nil {
						ptIdx = n
					}
				}

			case "v":
				inV = true
				vBuf.Reset()

			case "spPr":
				switch {
				case inSeries && !inMarker:
					inSeriesSpPr = true
				case !inSeries && !inAxis && !inPlotArea:
					// <c:chartSpace><c:spPr> holds the chart area fill/line.
					inChartSpPr = true
				}
			case "ln":
				switch {
				case inSeriesSpPr:
					inSeriesLn = true
				case inChartSpPr:
					inChartLn = true
				}
				if inSeriesSpPr || inChartSpPr {
					if w := attrValue(t, "w"); w != "" {
						if n, err := strconv.Atoi(w); err == nil {
							// Line width is stored in EMU; Gridlines/series
							// widths are expressed in points elsewhere.
							if pt := (n + 6350) / 12700; pt > 0 {
								switch {
								case inGridlines:
									gridWidth = pt
								case inChartSpPr:
									chartLineWidth = pt
								default:
									serLw = pt
								}
							}
						}
					}
				}
			case "noFill":
				if inChartSpPr && !inChartLn {
					chartNoFill = true
				}

			case "srgbClr", "schemeClr", "sysClr", "prstClr":
				if !inSeriesSpPr && !inGridlines && !inChartSpPr {
					break
				}
				if c, ok := chartColorFromElement(t, themeColors); ok {
					inColor = true
					pendColor = &c
					switch {
					case inSeriesSpPr && inSeriesLn:
						pendTarget = "serLine"
					case inSeriesSpPr:
						pendTarget = "serFill"
					case inGridlines:
						pendTarget = "gridline"
					case inChartSpPr && inChartLn:
						pendTarget = "chartLine"
					case inChartSpPr:
						pendTarget = "chartFill"
					default:
						pendTarget = ""
					}
				}

			case "alpha":
				if inColor && pendColor != nil {
					if pct, err := strconv.Atoi(val); err == nil {
						pendColor.ARGB = applyAlphaPercent(pendColor.ARGB, pct)
					}
				}

			case "marker":
				if inSeries {
					inMarker = true
					serMark = &SeriesMarker{}
				}
			case "symbol":
				if inMarker && serMark != nil {
					serMark.Symbol = val
				}
			case "size":
				if inMarker && serMark != nil {
					if n, err := strconv.Atoi(val); err == nil {
						serMark.Size = n
					}
				}

			case "smooth":
				if inSeries {
					serSmooth = val == "1" || val == "true"
				}

			case "dLbls":
				if inSeries {
					inDLbls = true
					serDlbls = &chartSeriesLabels{}
				}
			case "showVal":
				if serDlbls != nil {
					serDlbls.showValue = isXMLTrue(val)
				}
			case "showCatName":
				if serDlbls != nil {
					serDlbls.showCatName = isXMLTrue(val)
				}
			case "showPercent":
				if serDlbls != nil {
					serDlbls.showPercent = isXMLTrue(val)
				}
			case "showSerName":
				if serDlbls != nil {
					serDlbls.showSerName = isXMLTrue(val)
				}
			case "dLblPos":
				if serDlbls != nil {
					serDlbls.position = val
				}
			case "separator":
				inSeparator = true
				sepBuf.Reset()

			case "barDir":
				if inPlotArea {
					barDir = val
				}
			case "grouping":
				if inPlotArea && !inSeries {
					barGrouping = val
				}
			case "gapWidth":
				if inPlotArea {
					if n, err := strconv.Atoi(val); err == nil {
						gapWidth = &n
					}
				}
			case "overlap":
				if inPlotArea {
					if n, err := strconv.Atoi(val); err == nil {
						overlap = &n
					}
				}
			case "holeSize":
				if inPlotArea {
					if n, err := strconv.Atoi(val); err == nil {
						holeSize = &n
					}
				}

			case "catAx":
				newAxis(false)
			case "valAx":
				newAxis(true)
			case "axId", "axPos":
				// Only needed for layout, which is recomputed from the frame.
			case "orientation":
				if inAxis && curAxis != nil {
					curAxis.ReversedOrder = val == "maxMin"
				}
			case "delete":
				if inAxis && curAxis != nil {
					curAxis.Visible = !isXMLTrue(val)
				}
			case "min":
				if inAxis && curAxis != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						curAxis.SetMinBounds(f)
					}
				}
			case "max":
				if inAxis && curAxis != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						curAxis.SetMaxBounds(f)
					}
				}
			case "majorUnit":
				if inAxis && curAxis != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						curAxis.SetMajorUnit(f)
					}
				}
			case "minorUnit":
				if inAxis && curAxis != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						curAxis.SetMinorUnit(f)
					}
				}
			case "crosses":
				if inAxis && curAxis != nil && val != "" {
					curAxis.CrossesAt = val
				}
			case "tickLblPos":
				if inAxis && curAxis != nil && val != "" {
					curAxis.TickLabelPos = val
				}
			case "majorTickMark":
				if inAxis && curAxis != nil && val != "" {
					curAxis.MajorTickMark = val
				}
			case "minorTickMark":
				if inAxis && curAxis != nil && val != "" {
					curAxis.MinorTickMark = val
				}

			case "majorGridlines":
				if inAxis {
					inGridlines = true
					curGridlines = NewGridlines()
					gridColor = nil
					gridWidth = 0
				}
			case "minorGridlines":
				if inAxis {
					inGridlines = true
					curGridlines = NewGridlines()
					gridColor = nil
					gridWidth = 0
				}

			case "title":
				if inAxis {
					titleTarget = "axis"
					axisTitle.Reset()
				} else {
					titleTarget = "chart"
					chartTitle.Reset()
				}

			case "legend":
				chart.legend.Visible = true
			case "legendPos":
				if val != "" {
					chart.legend.Position = LegendPosition(val)
				}

			case "dispBlanksAs":
				if val != "" {
					chart.displayBlankAs = val
				}

			case "t":
				if titleTarget != "" || valueCtx == "serTitle" {
					inT = true
					tBuf.Reset()
				}
			}

		case xml.CharData:
			if inV {
				vBuf.WriteString(string(t))
			}
			if inT {
				tBuf.WriteString(string(t))
			}
			if inSeparator {
				sepBuf.WriteString(string(t))
			}

		case xml.EndElement:
			name := t.Name.Local
			switch name {
			case "plotArea":
				inPlotArea = false

			case "ser":
				finishSeries()
				inSeries = false
				valueCtx = ""

			case "tx", "cat", "val", "xVal", "yVal":
				valueCtx = ""

			case "spPr":
				inSeriesSpPr = false
				inSeriesLn = false
				// inChartSpPr is not per-series state, so it needs clearing
				// here as well: leaving it set made every later colour in the
				// part look like the chart area's.
				inChartSpPr = false
				inChartLn = false
			case "ln":
				inSeriesLn = false
				inChartLn = false

			case "srgbClr", "schemeClr", "sysClr", "prstClr":
				if inColor && pendColor != nil {
					commitColor(*pendColor)
				}
				inColor = false
				pendColor = nil
				pendTarget = ""

			case "v":
				inV = false
				value := strings.TrimSpace(vBuf.String())
				switch valueCtx {
				case "serTitle":
					serTitle.WriteString(value)
				case "cat":
					if cats != nil {
						cats.set(ptIdx, value)
					}
				case "val":
					if vals != nil {
						vals.set(ptIdx, value)
					}
				case "xval":
					if xVals != nil {
						xVals.set(ptIdx, value)
					}
				case "yval":
					if yVals != nil {
						yVals.set(ptIdx, value)
					}
				}

			case "marker":
				inMarker = false
			case "dLbls":
				inDLbls = false
			case "separator":
				inSeparator = false
				if serDlbls != nil {
					if s := strings.TrimSpace(sepBuf.String()); s != "" {
						serDlbls.separator = s
					}
				}

			case "majorGridlines":
				if inAxis && curAxis != nil && curGridlines != nil {
					if gridColor != nil {
						curGridlines.Color = *gridColor
					}
					if gridWidth > 0 {
						curGridlines.Width = gridWidth
					}
					curAxis.MajorGridlines = curGridlines
				}
				inGridlines = false
				curGridlines = nil

			case "minorGridlines":
				if inAxis && curAxis != nil && curGridlines != nil {
					if gridColor != nil {
						curGridlines.Color = *gridColor
					}
					if gridWidth > 0 {
						curGridlines.Width = gridWidth
					}
					curAxis.MinorGridlines = curGridlines
				}
				inGridlines = false
				curGridlines = nil

			case "catAx", "valAx":
				if curAxis != nil {
					if title := strings.TrimSpace(axisTitle.String()); title != "" {
						curAxis.Title = title
					}
					if axisIsValue {
						chart.plotArea.axisY = curAxis
					} else {
						chart.plotArea.axisX = curAxis
					}
				}
				curAxis = nil
				inAxis = false

			case "title":
				if titleTarget == "chart" {
					chartTitleText = strings.TrimSpace(chartTitle.String())
				}
				titleTarget = ""

			case "t":
				if inT {
					inT = false
					text := tBuf.String()
					switch {
					case titleTarget == "chart":
						chartTitle.WriteString(text)
					case titleTarget == "axis":
						axisTitle.WriteString(text)
					case valueCtx == "serTitle":
						serTitle.WriteString(text)
					}
				}
			}
		}
	}

	// --- Assemble the chart type -------------------------------------------
	if plotType == "" || len(series) == 0 {
		return nil
	}

	switch plotType {
	case "bar", "bar3D":
		bc := NewBarChart()
		if barDir != "" {
			bc.BarDirection = barDir
		}
		if barGrouping != "" {
			bc.SetBarGrouping(barGrouping)
		}
		if gapWidth != nil {
			bc.SetGapWidthPercent(*gapWidth)
		}
		if overlap != nil {
			bc.SetOverlapPercent(*overlap)
		}
		for _, s := range series {
			bc.AddSeries(s)
		}
		if plotType == "bar3D" {
			// The rasterizer draws 3D bars with the 2D bar pipeline, but the
			// chart type is preserved so a round-trip keeps the flavour.
			chart.plotArea.SetType(&Bar3DChart{BarChart: *bc})
		} else {
			chart.plotArea.SetType(bc)
		}
	case "line":
		lc := NewLineChart()
		lc.IsSmooth = lineSmooth
		for _, s := range series {
			lc.AddSeries(s)
		}
		chart.plotArea.SetType(lc)
	case "area":
		ac := NewAreaChart()
		for _, s := range series {
			ac.AddSeries(s)
		}
		chart.plotArea.SetType(ac)
	case "pie":
		pc := NewPieChart()
		for _, s := range series {
			pc.AddSeries(s)
		}
		chart.plotArea.SetType(pc)
	case "pie3D":
		p3 := NewPie3DChart()
		for _, s := range series {
			p3.AddSeries(s)
		}
		chart.plotArea.SetType(p3)
	case "doughnut":
		dc := NewDoughnutChart()
		if holeSize != nil && *holeSize >= 10 && *holeSize <= 90 {
			dc.HoleSize = *holeSize
		}
		for _, s := range series {
			dc.AddSeries(s)
		}
		chart.plotArea.SetType(dc)
	case "scatter":
		sc := NewScatterChart()
		sc.IsSmooth = lineSmooth
		for _, s := range series {
			sc.AddSeries(s)
		}
		chart.plotArea.SetType(sc)
	case "radar":
		rc := NewRadarChart()
		for _, s := range series {
			rc.AddSeries(s)
		}
		chart.plotArea.SetType(rc)
	default:
		return nil
	}

	if chartTitleText != "" {
		chart.title.Text = chartTitleText
		chart.title.Visible = true
	}

	// Chart area fill / outline, if the author set one. The zero value (nil)
	// means "no fill / no border", which is PowerPoint's default.
	if !chartNoFill && chartFillColor != nil {
		chart.fill = (&Fill{}).SetSolid(*chartFillColor)
	}
	if chartLineColor != nil {
		width := chartLineWidth
		if width <= 0 {
			width = 1
		}
		chart.border = (&Border{}).SetSolidFill(*chartLineColor).SetWidth(width)
	}

	return chart
}

// attrValue returns the value of the named attribute, or "" when absent.
func attrValue(se xml.StartElement, local string) string {
	for _, attr := range se.Attr {
		if attr.Name.Local == local {
			return attr.Value
		}
	}
	return ""
}

// isXMLTrue reports whether an OOXML boolean attribute value is true.
func isXMLTrue(v string) bool {
	return v == "1" || v == "true" || v == "on"
}

// applyAlphaPercent rewrites the alpha channel of an 8-char ARGB string given
// an OOXML alpha percentage (1000ths of a percent, e.g. 50000 = 50%).
func applyAlphaPercent(argb string, pct int) string {
	if len(argb) != 8 {
		return argb
	}
	if pct < 0 {
		pct = 0
	}
	if pct > 100000 {
		pct = 100000
	}
	a := (pct*255 + 50000) / 100000
	return hexByte(uint8(a)) + argb[2:]
}

// hexByte formats a byte as two uppercase hex digits.
func hexByte(b uint8) string {
	const digits = "0123456789ABCDEF"
	return string([]byte{digits[b>>4], digits[b&0x0F]})
}

// themeAccentDefaults are used when a chart references theme colours but the
// presentation has no readable theme (or none was supplied). They match the
// default Office theme so charts still render in colour.
var themeSchemeDefaults = map[string]string{
	"dk1":      "FF000000",
	"lt1":      "FFFFFFFF",
	"dk2":      "FF44546A",
	"lt2":      "FFE7E6E6",
	"tx1":      "FF000000",
	"bg1":      "FFFFFFFF",
	"tx2":      "FF44546A",
	"bg2":      "FFE7E6E6",
	"accent1":  "FF4472C4",
	"accent2":  "FFED7D31",
	"accent3":  "FFA5A5A5",
	"accent4":  "FFFFC000",
	"accent5":  "FF5B9BD5",
	"accent6":  "FF70AD47",
	"hlink":    "FF0563C1",
	"folHlink": "FF954F72",
	"phClr":    "FF000000",
	"dk1-link": "FF000000",
	"lt1-link": "FFFFFFFF",
}

// chartColorFromElement resolves an <a:srgbClr>/<a:schemeClr>/<a:sysClr>/
// <a:prstClr> element to a Color.
func chartColorFromElement(se xml.StartElement, themeColors map[string]string) (Color, bool) {
	switch se.Name.Local {
	case "srgbClr":
		v := attrValue(se, "val")
		if v == "" {
			return Color{}, false
		}
		return NewColor(v), true

	case "schemeClr":
		name := attrValue(se, "val")
		if name == "" {
			return Color{}, false
		}
		if themeColors != nil {
			if argb, ok := themeColors[name]; ok && argb != "" {
				return NewColor(argb), true
			}
		}
		if argb, ok := themeSchemeDefaults[name]; ok {
			return NewColor(argb), true
		}
		return Color{}, false

	case "sysClr":
		if lastClr := attrValue(se, "lastClr"); lastClr != "" {
			return NewColor(lastClr), true
		}
		if v := attrValue(se, "val"); v != "" {
			// windowText/window/... cannot be resolved without a theme.
			if argb, ok := themeSchemeDefaults[v]; ok {
				return NewColor(argb), true
			}
		}
		return Color{}, false

	case "prstClr":
		v := attrValue(se, "val")
		if v == "" {
			return Color{}, false
		}
		return presetColorToColor(v), true
	}
	return Color{}, false
}
