package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"math"
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
		inPlotArea     bool
		inSeries       bool
		inSeriesSpPr   bool
		inSeriesLn     bool
		inChartSpPr    bool
		inChartLn      bool
		chartNoFill    bool
		inMarker       bool
		inDLbls        bool
		inGridlines    bool
		inAxis         bool
		axisIsValue    bool
		inAxisSpPr     bool
		inAxisLn       bool
		inColor        bool
		inV            bool
		inT            bool
		inSeparator    bool
		inLegend       bool
		inManualLayout bool
		inChartElem    bool
		titleTarget    string // "", "chart" or "axis"
		txPrTarget     string // "", "axis", "legend", "series" or "chartspace"
		valueCtx       string // "", "serTitle", "cat", "val", "xval", "yval"

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
		serDash   string
		serMark   *SeriesMarker
		serSmooth bool
		serDlbls  *chartSeriesLabels
		inDPt     bool
		dPtIdx    int
		dPtFill   map[int]Color

		curAxis      *ChartAxis
		curGridlines *Gridlines
		gridColor    *Color
		gridWidth    int
		gridNoFill   bool

		// The axis line's own <c:spPr><a:ln>: width in points, an explicit
		// colour, or <a:noFill> (chart5's left value axis declares noFill and
		// PowerPoint draws no line there at all).
		axisLnW      int
		axisLnNoFill bool
		axisLnColor  *Color

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

		// Colour-transform accumulation for the colour currently being read.
		ctTint, ctShade, ctLumMod, ctLumOff float64
	)

	var (
		chartTitleText string

		// Typefaces declared by the title run currently being scanned. Chart
		// text declares them on <a:rPr><a:latin>/<a:ea>, exactly as slide text
		// does, so they are collected while the title is open and applied when
		// it closes. Without this the part's font was dropped on read and every
		// chart string came back as the model default.
		runLatin string
		runEA    string

		// Run properties stated by that same title run. They are applied with
		// the typefaces, on the matching close tag.
		runSize            int
		runBold, runItalic bool
		runBoldSet         bool
		runItalicSet       bool

		// The <c:txPr> currently open: what its paragraph defaults state, and
		// the series font it accumulates (a series' data labels are read into
		// a font of their own, applied when the series closes).
		txPrLatin, txPrEA    string
		txPrSize             int
		txPrBold, txPrItalic bool
		txPrBoldSet          bool
		txPrItalicSet        bool
		serFont              *Font

		// The chartSpace-level <c:txPr> is the default text size for every
		// chart element that does not declare its own (chart3 of deck
		// 00022823 states 20pt there and nothing on its value axis, and
		// PowerPoint renders 20pt tick labels). sizedFonts records the fonts
		// that did declare a size so the default skips them.
		chartDefSize int
		sizedFonts   map[*Font]bool
	)
	sizedFonts = make(map[*Font]bool)

	resetSeries := func() {
		cats = &seriesPoints{}
		vals = &seriesPoints{}
		xVals = &seriesPoints{}
		yVals = &seriesPoints{}
		serTitle.Reset()
		serFill = nil
		serLine = nil
		serLw = 0
		serDash = ""
		serMark = nil
		serSmooth = false
		serDlbls = nil
		serFont = nil
		inDPt = false
		dPtFill = nil
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
			// Blank/missing points stay NaN so the renderer breaks the line
			// there instead of plotting a fake zero (chart1's "Cost per
			// Genome" stops at idx 20; COM shows the line ending, not
			// hugging the axis floor).
			floats[i] = math.NaN()
			if i < len(valList) {
				if f, err := strconv.ParseFloat(strings.TrimSpace(valList[i]), 64); err == nil {
					floats[i] = f
				}
			}
		}

		s := NewChartSeriesOrdered(strings.TrimSpace(serTitle.String()), catList, floats)
		if serFont != nil {
			// The series carried a <c:txPr>, which is where its data labels'
			// font lives. NewChartSeriesOrdered seeds a default font, so the
			// parsed one replaces it whole.
			s.Font = serFont
		}
		if serFill != nil {
			s.FillColor = *serFill
		} else if serLine != nil {
			// Line/scatter series encode their colour on the outline.
			s.FillColor = *serLine
		}
		if dPtFill != nil {
			s.PointColors = dPtFill
		}
		if serLine != nil || serLw > 0 {
			outline := &SeriesOutline{Width: serLw}
			if serLine != nil {
				outline.Color = *serLine
			}
			s.Outline = outline
		}
		if serDash != "" {
			s.LineDash = serDash
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
		// PowerPoint's default tick-mark style for chart axes is "out":
		// deck 00022823 chart1/chart2 omit <c:majorTickMark> and the COM
		// export shows ticks, while chart3/4/5 declare val="none" and show
		// none. An explicit element below still overwrites this.
		curAxis.MajorTickMark = TickMarkOutside
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
		case "dPtFill":
			if dPtFill == nil {
				dPtFill = make(map[int]Color)
			}
			dPtFill[dPtIdx] = c
		case "serLine":
			cc := c
			serLine = &cc
		case "gridline":
			cc := c
			gridColor = &cc
		case "axisLn":
			cc := c
			axisLnColor = &cc
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

			case "chart":
				inChartElem = true

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
				case inAxis && !inGridlines:
					// The axis line's own spPr. Gridlines also live inside an
					// axis but carry their own capture path.
					inAxisSpPr = true
				case !inSeries && !inAxis && !inPlotArea:
					// <c:chartSpace><c:spPr> holds the chart area fill/line.
					inChartSpPr = true
				}
			case "ln":
				switch {
				case inSeriesSpPr:
					inSeriesLn = true
				case inAxisSpPr:
					inAxisLn = true
					if w := attrValue(t, "w"); w != "" {
						if n, err := strconv.Atoi(w); err == nil {
							if pt := (n + 6350) / 12700; pt > 0 {
								axisLnW = pt
							}
						}
					}
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
				switch {
				case inAxisSpPr && inAxisLn:
					axisLnNoFill = true
				case inChartSpPr && !inChartLn:
					chartNoFill = true
				case inGridlines:
					gridNoFill = true
				}

			case "srgbClr", "schemeClr", "sysClr", "prstClr":
				if !inSeriesSpPr && !inGridlines && !inChartSpPr && !inAxisSpPr {
					break
				}
				if c, ok := chartColorFromElement(t, themeColors); ok {
					inColor = true
					cc := c
					pendColor = &cc
					ctTint, ctShade, ctLumMod, ctLumOff = -1, -1, -1, -1
					switch {
					case inDPt && inSeriesSpPr && !inSeriesLn:
						pendTarget = "dPtFill"
					case inSeriesSpPr && inSeriesLn:
						pendTarget = "serLine"
					case inSeriesSpPr:
						pendTarget = "serFill"
					case inGridlines:
						pendTarget = "gridline"
					case inAxisSpPr && inAxisLn:
						pendTarget = "axisLn"
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

			case "tint", "shade", "lumMod", "lumOff":
				// Colour transforms inside a chart colour reference —
				// chart1's "Moore's Law" stroke is bg1 with lumMod 50%,
				// i.e. mid grey, not the plain white the raw lookup gives.
				if inColor && pendColor != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						switch name {
						case "tint":
							ctTint = f / 100000.0
						case "shade":
							ctShade = f / 100000.0
						case "lumMod":
							ctLumMod = f / 100000.0
						case "lumOff":
							ctLumOff = f / 100000.0
						}
					}
				}

			case "dPt":
				// Per-data-point override (a highlighted bar in a series).
				// Its <c:spPr> must not be mistaken for the series fill: it
				// comes after the series <c:spPr> and used to overwrite it,
				// painting every bar the highlight colour.
				if inSeries {
					inDPt = true
					dPtIdx = 0
					if v := attrValue(t, "idx"); v != "" {
						if n, err := strconv.Atoi(v); err == nil {
							dPtIdx = n
						}
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

			case "x", "y", "w", "h":
				// manualLayout fractions (exponent notation included).
				if inManualLayout && chart.plotArea.layout != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil {
						switch name {
						case "x":
							chart.plotArea.layout.x = f
						case "y":
							chart.plotArea.layout.y = f
						case "w":
							chart.plotArea.layout.w = f
						case "h":
							chart.plotArea.layout.h = f
						}
					}
				}

			case "catAx", "dateAx":
				// A date axis is a category axis whose labels are date
				// serial numbers; the model treats them the same way.
				newAxis(false)
			case "valAx":
				newAxis(true)
			case "axPos":
				// The axis position decides where the axis is parked: a
				// value axis along the bottom/top of a scatter chart is
				// still the X axis.
				if inAxis && curAxis != nil {
					curAxis.Position = val
				}
			case "logBase":
				if inAxis && curAxis != nil {
					if f, err := strconv.ParseFloat(val, 64); err == nil && f > 1 {
						curAxis.LogBase = f
					}
				}
			case "date1904":
				chart.date1904 = isXMLTrue(val)
			case "prstDash":
				// The dash preset of the series stroke; gridlines and axes
				// carry their own, which we don't model.
				if inSeries && inSeriesSpPr && inSeriesLn {
					serDash = val
				}
			case "axId":
				// Only needed for layout, which is recomputed from the frame.
			case "orientation":
				if inAxis && curAxis != nil {
					curAxis.ReversedOrder = val == "maxMin"
				}
			case "numFmt":
				// Axis tick label format ("#,##0", "0.0%", "General", ...).
				// Data-point caches carry their own numFmt; only the axis
				// one drives the tick labels the renderer draws. "General"
				// stays empty so the model keeps a single "unformatted" state.
				if inAxis && curAxis != nil {
					if fc := attrValue(t, "formatCode"); fc != "" && fc != "General" {
						curAxis.NumberFormat = fc
					}
				}
			case "manualLayout":
				// <c:plotArea><c:layout><c:manualLayout> pins the inner plot
				// rect as fractions of the chart frame. Legend layouts are
				// not modelled, and an axis title's own manualLayout (chart5
				// pins "Seconds" at x=0.0077) must not overwrite the plot's,
				// so only read it directly inside the plot area.
				if inPlotArea && !inLegend && !inAxis && titleTarget == "" {
					inManualLayout = true
					if chart.plotArea.layout == nil {
						chart.plotArea.layout = &chartManualLayout{}
					}
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
					gridNoFill = false
				}
			case "minorGridlines":
				if inAxis {
					inGridlines = true
					curGridlines = NewGridlines()
					gridColor = nil
					gridWidth = 0
					gridNoFill = false
				}

			case "title":
				runLatin, runEA = "", ""
				runSize, runBoldSet, runItalicSet = 0, false, false
				if inAxis {
					titleTarget = "axis"
					axisTitle.Reset()
				} else {
					titleTarget = "chart"
					chartTitle.Reset()
				}

			case "txPr":
				// A <c:txPr> holds the default run properties for the text of
				// the element it sits in: an axis' tick labels, the legend
				// entries, or a series' data labels. PowerPoint reads a chart
				// label's font from here, so the writer emits it and the reader
				// has to read it back — the two used to disagree, with the
				// rasteriser resolving the model's font while PowerPoint fell
				// back to the theme.
				txPrLatin, txPrEA = "", ""
				txPrSize, txPrBoldSet, txPrItalicSet = 0, false, false
				switch {
				case inAxis:
					txPrTarget = "axis"
				case inSeries && inDLbls:
					txPrTarget = "series"
					if serFont == nil {
						serFont = NewFont()
					}
				case inLegend:
					txPrTarget = "legend"
				case !inChartElem:
					// The chartSpace-level <c:txPr> after </c:chart>: the
					// chart-wide text default.
					txPrTarget = "chartspace"
				default:
					// Somewhere we do not model, e.g. a chart-group level
					// <c:dLbls>: read nothing rather than guess a target.
					txPrTarget = ""
				}

			case "defRPr":
				if txPrTarget == "" {
					break
				}
				if v := attrValue(t, "sz"); v != "" {
					if n, err := strconv.Atoi(v); err == nil && n > 0 {
						txPrSize = n
					}
				}
				if v := attrValue(t, "b"); v != "" {
					txPrBold, txPrBoldSet = isXMLTrue(v), true
				}
				if v := attrValue(t, "i"); v != "" {
					txPrItalic, txPrItalicSet = isXMLTrue(v), true
				}

			case "rPr":
				// A title states its font on the run itself rather than in a
				// <c:txPr>, so the same properties are collected here and
				// applied when the title closes.
				if titleTarget == "" {
					break
				}
				runSize, runBoldSet, runItalicSet = 0, false, false
				if v := attrValue(t, "sz"); v != "" {
					if n, err := strconv.Atoi(v); err == nil && n > 0 {
						runSize = n
					}
				}
				if v := attrValue(t, "b"); v != "" {
					runBold, runBoldSet = isXMLTrue(v), true
				}
				if v := attrValue(t, "i"); v != "" {
					runItalic, runItalicSet = isXMLTrue(v), true
				}

			case "latin", "ea":
				// <a:rPr> names the faces a run uses and <a:defRPr> names the
				// defaults for a whole text element, so both the title runs and
				// every <c:txPr> have to be read: the writer emits the font on
				// both, and a part that states it only in a <c:txPr> would
				// otherwise come back as the model default.
				typeface := attrValue(t, "typeface")
				// "+mn-lt" and friends are theme references, not font names.
				if typeface == "" || strings.HasPrefix(typeface, "+") {
					break
				}
				if txPrTarget != "" {
					if name == "latin" {
						if txPrLatin == "" {
							txPrLatin = typeface
						}
					} else if txPrEA == "" {
						txPrEA = typeface
					}
					break
				}
				if titleTarget == "" {
					break
				}
				if name == "latin" {
					if runLatin == "" {
						runLatin = typeface
					}
				} else if runEA == "" {
					runEA = typeface
				}

			case "legend":
				chart.legend.Visible = true
				// The legend's <c:txPr> carries the entry font; the flag pairs
				// with the close handler below so the target is only ever
				// chosen for text that really is inside a legend.
				inLegend = true
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
				inAxisSpPr = false
				inAxisLn = false
			case "ln":
				inSeriesLn = false
				inChartLn = false
				inAxisLn = false

			case "srgbClr", "schemeClr", "sysClr", "prstClr":
				if inColor && pendColor != nil {
					*pendColor = applyColorTransforms(*pendColor, ctTint, ctShade, ctLumMod, ctLumOff)
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
			case "dPt":
				inDPt = false
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
						curGridlines.ColorSet = true
					}
					curGridlines.NoFill = gridNoFill
					if gridWidth > 0 {
						curGridlines.Width = gridWidth
					}
					curAxis.MajorGridlines = curGridlines
				}
				inGridlines = false
				curGridlines = nil
				gridNoFill = false

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

			case "catAx", "valAx", "dateAx":
				if curAxis != nil {
					if title := strings.TrimSpace(axisTitle.String()); title != "" {
						curAxis.Title = title
					}
					// The axis line's explicit stroke: noFill kills the line
					// entirely (chart5's left value axis), otherwise the
					// declared colour and width replace the 134-grey default.
					if axisLnNoFill {
						curAxis.OutlineNoFill = true
					} else {
						if axisLnColor != nil {
							curAxis.OutlineColor = *axisLnColor
						}
						if axisLnW > 0 {
							curAxis.OutlineWidth = axisLnW
						}
					}
					switch {
					case !axisIsValue:
						chart.plotArea.axisX = curAxis
					case plotType == "scatter" &&
						(curAxis.Position == "b" || curAxis.Position == "t"):
						// Two value axes meet only in scatter/bubble
						// charts; there the one along the horizontal edge
						// is the X axis. A horizontal bar chart's value
						// axis also sits at "b" but must stay axisY — the
						// renderer flips the bar geometry itself.
						chart.plotArea.axisX = curAxis
					default:
						chart.plotArea.axisY = curAxis
					}
				}
				curAxis = nil
				inAxis = false
				axisLnW = 0
				axisLnNoFill = false
				axisLnColor = nil

			case "title":
				switch {
				case titleTarget == "chart":
					chartTitleText = strings.TrimSpace(chartTitle.String())
					applyChartTextFont(chart.title.Font, runLatin, runEA)
					applyChartFontAttributes(chart.title.Font, runSize,
						runBold, runBoldSet, runItalic, runItalicSet)
					if runSize > 0 {
						sizedFonts[chart.title.Font] = true
					}
				case titleTarget == "axis" && curAxis != nil:
					applyChartTextFont(curAxis.Font, runLatin, runEA)
					applyChartFontAttributes(curAxis.Font, runSize,
						runBold, runBoldSet, runItalic, runItalicSet)
					if runSize > 0 {
						sizedFonts[curAxis.Font] = true
					}
				}
				titleTarget = ""

			case "txPr":
				// Commit the paragraph defaults to whichever element owned the
				// <c:txPr>. Every target the start handler can choose has to
				// appear here: a missing case drops the font silently, which is
				// exactly how the label fonts went unread.
				switch {
				case txPrTarget == "axis" && curAxis != nil:
					applyChartTextFont(curAxis.Font, txPrLatin, txPrEA)
					applyChartFontAttributes(curAxis.Font, txPrSize,
						txPrBold, txPrBoldSet, txPrItalic, txPrItalicSet)
					if txPrSize > 0 {
						sizedFonts[curAxis.Font] = true
					}
				case txPrTarget == "legend":
					applyChartTextFont(chart.legend.Font, txPrLatin, txPrEA)
					applyChartFontAttributes(chart.legend.Font, txPrSize,
						txPrBold, txPrBoldSet, txPrItalic, txPrItalicSet)
					if txPrSize > 0 {
						sizedFonts[chart.legend.Font] = true
					}
				case txPrTarget == "series" && serFont != nil:
					applyChartTextFont(serFont, txPrLatin, txPrEA)
					applyChartFontAttributes(serFont, txPrSize,
						txPrBold, txPrBoldSet, txPrItalic, txPrItalicSet)
					if txPrSize > 0 {
						sizedFonts[serFont] = true
					}
				case txPrTarget == "chartspace":
					chartDefSize = txPrSize
				}
				txPrTarget = ""

			case "chart":
				inChartElem = false

			case "legend":
				inLegend = false

			case "manualLayout":
				inManualLayout = false

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

	// The chartSpace-level <c:txPr> size is every chart font's default: any
	// axis/legend/series/title font that declared no size of its own inherits
	// it (chart3 of deck 00022823: value-axis tick labels render at the 20pt
	// chart default, not the 10pt model fallback).
	if chartDefSize > 0 {
		ds := chartDefSize / 100
		fonts := []*Font{chart.plotArea.axisX.Font, chart.plotArea.axisY.Font,
			chart.legend.Font, chart.title.Font}
		for _, s := range series {
			fonts = append(fonts, s.Font)
		}
		for _, f := range fonts {
			if f != nil && !sizedFonts[f] {
				f.Size = ds
			}
		}
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

// applyChartFontAttributes copies the run properties a chart text element
// stated into the model font, leaving anything it did not state alone. Size is
// in hundredths of a point in the file, as everywhere else in the format.
func applyChartFontAttributes(f *Font, size int, bold, boldSet, italic, italicSet bool) {
	if f == nil {
		return
	}
	if size > 0 {
		f.Size = size / 100
	}
	if boldSet {
		f.Bold = bold
	}
	if italicSet {
		f.Italic = italic
	}
}

// applyChartTextFont copies the typefaces a chart text run declared into the
// model. Empty names leave the existing face alone, so a run that declares only
// <a:latin> does not wipe an East Asian face the caller set.
func applyChartTextFont(f *Font, latin, ea string) {
	if f == nil {
		return
	}
	if latin != "" {
		f.Name = latin
	}
	if ea != "" {
		f.NameEA = ea
	}
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
