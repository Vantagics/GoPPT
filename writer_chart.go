package gopresentation

import (
	"archive/zip"
	"fmt"
	"strings"
)

// getChartSeries extracts the series list from any chart type.
func getChartSeries(ct ChartType) []*ChartSeries {
	switch c := ct.(type) {
	case *BarChart:
		return c.Series
	case *Bar3DChart:
		return c.Series
	case *LineChart:
		return c.Series
	case *AreaChart:
		return c.Series
	case *PieChart:
		return c.Series
	case *Pie3DChart:
		return c.Series
	case *DoughnutChart:
		return c.Series
	case *ScatterChart:
		return c.Series
	case *RadarChart:
		return c.Series
	default:
		return nil
	}
}

// getCategories returns the union of all categories across series.
func getCategories(series []*ChartSeries) []string {
	if len(series) == 0 {
		return nil
	}
	return series[0].Categories
}

func (w *PPTXWriter) writeChartPart(zw *zip.Writer, chart *ChartShape, chartIdx int) error {
	ct := chart.plotArea.chartType
	if ct == nil {
		return nil
	}

	series := getChartSeries(ct)
	categories := getCategories(series)

	var chartTypeXML strings.Builder
	chartTypeName := ct.GetChartTypeName()

	switch c := ct.(type) {
	case *BarChart:
		chartTypeXML.WriteString(w.writeBarChartXML(c, categories))
	case *Bar3DChart:
		chartTypeXML.WriteString(w.writeBar3DChartXML(c, categories))
	case *LineChart:
		chartTypeXML.WriteString(w.writeLineChartXML(c, categories))
	case *AreaChart:
		chartTypeXML.WriteString(w.writeAreaChartXML(c, categories))
	case *PieChart:
		chartTypeXML.WriteString(w.writePieChartXML(c, categories))
	case *Pie3DChart:
		chartTypeXML.WriteString(w.writePie3DChartXML(c, categories))
	case *DoughnutChart:
		chartTypeXML.WriteString(w.writeDoughnutChartXML(c, categories))
	case *ScatterChart:
		chartTypeXML.WriteString(w.writeScatterChartXML(c, categories))
	case *RadarChart:
		chartTypeXML.WriteString(w.writeRadarChartXML(c, categories))
	}
	_ = chartTypeName

	// Title XML
	titleXML := ""
	if chart.title.Visible && chart.title.Text != "" {
		f := chart.title.Font
		if f == nil {
			f = NewFont()
		}
		runProps := chartRunPropsXML(f, "a:rPr", chartTextRunAttrs(f))
		titleXML = fmt.Sprintf(`  <c:title>
    <c:tx>
      <c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r>
            %s
            <a:t>%s</a:t>
          </a:r>
        </a:p>
      </c:rich>
    </c:tx>
    <c:overlay val="0"/>
  </c:title>
`, runProps, xmlEscape(chart.title.Text))
	} else if !chart.title.Visible {
		titleXML = `  <c:autoTitleDeleted val="1"/>
`
	}

	// Legend XML. <c:txPr> follows <c:overlay> in CT_Legend, and it is where
	// PowerPoint reads the legend entry font from.
	legendXML := ""
	if chart.legend.Visible {
		legendXML = fmt.Sprintf(`  <c:legend>
    <c:legendPos val="%s"/>
    <c:overlay val="0"/>
%s  </c:legend>
`, chart.legend.Position, chartTxPrXML("    ", chart.legend.Font))
	}

	// Axis XML
	axisXML := ""
	if !isPieType(ct) {
		axisXML = w.writeAxesXML(chart)
	}

	// Chart-area fill and outline. The reader fills chart.fill/chart.border
	// from <c:chartSpace><c:spPr> and the rasteriser draws them, but the writer
	// had no counterpart, so an author's chart-area shading was silently gone
	// the next time the deck was opened.
	chartSpPrXML := w.writeChartAreaSpPrXML(chart)

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="%s" xmlns:r="%s">
  <c:chart>
%s    <c:plotArea>
      <c:layout/>
%s%s    </c:plotArea>
%s    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="%s"/>
  </c:chart>
%s</c:chartSpace>`,
		nsDrawingML, nsOfficeDocRels,
		titleXML,
		chartTypeXML.String(), axisXML,
		legendXML,
		chart.displayBlankAs,
		chartSpPrXML)

	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/charts/chart%d.xml", chartIdx), content)
}

// writeChartAreaSpPrXML renders the chart area's <c:spPr>.
//
// Element order matters: in CT_ChartSpace, <c:spPr> is a sibling that follows
// <c:chart>, not a child of it, so it cannot share the slot the title uses.
func (w *PPTXWriter) writeChartAreaSpPrXML(chart *ChartShape) string {
	fillXML := w.writeFillXML(chart.fill)
	borderXML := w.writeBorderXML(chart.border)
	if fillXML == "" && borderXML == "" {
		return ""
	}
	return "  <c:spPr>\n" + fillXML + borderXML + "  </c:spPr>\n"
}

func boolToXML(v bool) string {
	if v {
		return "1"
	}
	return "0"
}

// chartRunPropsXML renders the run properties of a chart text element.
//
// elem is the element to emit: <a:rPr> for a run inside <c:title>, and
// <a:defRPr> for the paragraph default a <c:txPr> carries. One function emits
// both so the two ways a chart states a font cannot drift apart.
//
// Chart text carries the same font model as slide text: the Latin face on
// <a:latin> and the East Asian face on <a:ea>. The chart writer emitted the
// size and bold flag only, so an author's font choice was silently dropped on
// save. PowerPoint then fell back to its theme font, and the preview
// rasteriser — which has to read the font back out of the part — had no East
// Asian face to use, so every Chinese chart label drew as a .notdef box.
func chartRunPropsXML(f *Font, elem, attrs string) string {
	if attrs != "" {
		attrs = " " + attrs
	}
	latin, ea := "", ""
	if f != nil && f.Name != "" {
		latin = fmt.Sprintf(`<a:latin typeface="%s"/>`, xmlEscape(f.Name))
	}
	if f != nil && f.NameEA != "" {
		ea = fmt.Sprintf(`<a:ea typeface="%s"/>`, xmlEscape(f.NameEA))
	}
	if latin == "" && ea == "" {
		return fmt.Sprintf(`<%s%s/>`, elem, attrs)
	}
	return fmt.Sprintf(`<%s%s>%s%s</%s>`, elem, attrs, latin, ea, elem)
}

// chartTextRunAttrs is the attribute set every chart text element carries.
//
// Shared by the title runs and the <c:txPr> paragraph defaults so both paths
// agree on what a chart font is. The size, bold and italic flags are exactly
// the ones the rasteriser reads back (fontFaceFor takes bold and italic), and
// an unstated size is omitted rather than written as sz="0".
func chartTextRunAttrs(f *Font) string {
	attrs := `lang="en-US"`
	if f == nil {
		return attrs
	}
	if f.Size > 0 {
		attrs += fmt.Sprintf(` sz="%d"`, f.Size*100)
	}
	return attrs + fmt.Sprintf(` b="%s" i="%s"`, boolToXML(f.Bold), boolToXML(f.Italic))
}

// chartTxPrXML renders the <c:txPr> a chart text element uses for its label
// font: axis tick labels, legend entries and data labels.
//
// PowerPoint reads a chart label's font from here and nowhere else, so
// omitting it made every axis and legend label fall back to the theme font
// even though the preview looked right — the rasteriser resolves the face
// itself and had the model's font to resolve. The element is only written when
// the font states something (a nil font, or one with no face and no size,
// would be an empty <c:txPr> that says less than its absence).
func chartTxPrXML(indent string, f *Font) string {
	if f == nil || (f.Name == "" && f.NameEA == "" && f.Size <= 0) {
		return ""
	}
	return fmt.Sprintf("%s<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>%s</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>\n",
		indent, chartRunPropsXML(f, "a:defRPr", chartTextRunAttrs(f)))
}

// chartTickMarksXML renders the two tick-mark settings an axis may carry, in
// schema order.
//
// The model has always held MajorTickMark/MinorTickMark and the reader has
// always parsed them, but the writer emitted neither, so a document's tick
// marks reverted to the model default every time it was saved.
func chartTickMarksXML(indent, major, minor string) string {
	out := ""
	if major != "" {
		out += fmt.Sprintf("%s<c:majorTickMark val=\"%s\"/>\n", indent, major)
	}
	if minor != "" {
		out += fmt.Sprintf("%s<c:minorTickMark val=\"%s\"/>\n", indent, minor)
	}
	return out
}

// chartTitleXML renders an axis <c:title> whose single run carries the axis
// font. It is shared by both axes so their run properties cannot drift apart.
func chartTitleXML(text string, f *Font) string {
	return fmt.Sprintf(`        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>%s<a:t>%s</a:t></a:r></a:p></c:rich></c:tx></c:title>
`, chartRunPropsXML(f, "a:rPr", chartTextRunAttrs(f)), xmlEscape(text))
}

func isPieType(ct ChartType) bool {
	switch ct.(type) {
	case *PieChart, *Pie3DChart, *DoughnutChart:
		return true
	}
	return false
}

func (w *PPTXWriter) writeAxesXML(chart *ChartShape) string {
	axX := chart.plotArea.axisX
	axY := chart.plotArea.axisY

	// The element order here is fixed by CT_CatAx and CT_ValAx: gridlines and
	// the title precede the tick settings, <c:txPr> sits immediately before the
	// crossing axis, and a value axis' units follow the crossing pair. The
	// reader matches on the local name and does not validate order, so a part
	// PowerPoint would reject still round-trips here — which is how the order
	// drifted in the first place.
	catAxisXML := fmt.Sprintf(`      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="%s"/></c:scaling>
        <c:delete val="%s"/>
        <c:axPos val="b"/>
`, w.axisOrientation(axX), boolToXML(!axX.Visible))

	if axX.MajorGridlines != nil {
		catAxisXML += w.writeGridlinesXML("c:majorGridlines", axX.MajorGridlines)
	}
	if axX.MinorGridlines != nil {
		catAxisXML += w.writeGridlinesXML("c:minorGridlines", axX.MinorGridlines)
	}
	if axX.Title != "" {
		catAxisXML += chartTitleXML(axX.Title, axX.Font)
	}
	if axX.NumberFormat != "" && axX.NumberFormat != "General" {
		catAxisXML += fmt.Sprintf("        <c:numFmt formatCode=%q sourceLinked=\"0\"/>\n", axX.NumberFormat)
	}
	catAxisXML += chartTickMarksXML("        ", axX.MajorTickMark, axX.MinorTickMark)
	catAxisXML += fmt.Sprintf("        <c:tickLblPos val=\"%s\"/>\n", axX.TickLabelPos)
	catAxisXML += chartTxPrXML("        ", axX.Font)
	catAxisXML += fmt.Sprintf(`        <c:crossAx val="2"/>
        <c:crosses val="%s"/>
      </c:catAx>
`, axX.CrossesAt)

	valAxisXML := fmt.Sprintf(`      <c:valAx>
        <c:axId val="2"/>
        <c:scaling>
          <c:orientation val="%s"/>`, w.axisOrientation(axY))

	// CT_Scaling puts max before min.
	if axY.MaxBounds != nil {
		valAxisXML += fmt.Sprintf(`
          <c:max val="%g"/>`, *axY.MaxBounds)
	}
	if axY.MinBounds != nil {
		valAxisXML += fmt.Sprintf(`
          <c:min val="%g"/>`, *axY.MinBounds)
	}
	valAxisXML += `
        </c:scaling>
`
	valAxisXML += fmt.Sprintf(`        <c:delete val="%s"/>
        <c:axPos val="l"/>
`, boolToXML(!axY.Visible))

	if axY.MajorGridlines != nil {
		valAxisXML += w.writeGridlinesXML("c:majorGridlines", axY.MajorGridlines)
	}
	if axY.MinorGridlines != nil {
		valAxisXML += w.writeGridlinesXML("c:minorGridlines", axY.MinorGridlines)
	}
	if axY.Title != "" {
		valAxisXML += chartTitleXML(axY.Title, axY.Font)
	}
	if axY.NumberFormat != "" && axY.NumberFormat != "General" {
		valAxisXML += fmt.Sprintf("        <c:numFmt formatCode=%q sourceLinked=\"0\"/>\n", axY.NumberFormat)
	}
	valAxisXML += chartTickMarksXML("        ", axY.MajorTickMark, axY.MinorTickMark)
	valAxisXML += fmt.Sprintf("        <c:tickLblPos val=\"%s\"/>\n", axY.TickLabelPos)
	valAxisXML += chartTxPrXML("        ", axY.Font)
	valAxisXML += fmt.Sprintf(`        <c:crossAx val="1"/>
        <c:crosses val="%s"/>
`, axY.CrossesAt)

	if axY.MajorUnit != nil {
		valAxisXML += fmt.Sprintf("        <c:majorUnit val=\"%g\"/>\n", *axY.MajorUnit)
	}
	if axY.MinorUnit != nil {
		valAxisXML += fmt.Sprintf("        <c:minorUnit val=\"%g\"/>\n", *axY.MinorUnit)
	}
	valAxisXML += "      </c:valAx>\n"

	return catAxisXML + valAxisXML
}

func (w *PPTXWriter) axisOrientation(ax *ChartAxis) string {
	if ax.ReversedOrder {
		return "maxMin"
	}
	return "minMax"
}

func (w *PPTXWriter) writeGridlinesXML(tag string, gl *Gridlines) string {
	return fmt.Sprintf(`        <%s>
          <c:spPr>
            <a:ln w="%d">
              <a:solidFill><a:srgbClr val="%s"/></a:solidFill>
            </a:ln>
          </c:spPr>
        </%s>
`, tag, gl.Width*12700, colorRGB(gl.Color), tag)
}

// writeSeriesSpPrXML renders a series' <c:spPr>, carrying its fill and/or its
// outline.
//
// Where a series' colour lives depends on the chart type. A line or scatter
// series is stroked, so PowerPoint reads its colour from <c:spPr><a:ln> and
// ignores a bare <a:solidFill> there; the reader agrees, mapping that element
// to Outline and the fill element to FillColor. Writing a line series' colour
// as a fill therefore only looked correct because the reader falls back to
// FillColor when no outline colour is present, while PowerPoint silently
// dropped it. Other chart types fill their series markers and may additionally
// carry an outline, so both elements can appear.
func (w *PPTXWriter) writeSeriesSpPrXML(s *ChartSeries, lineSeries bool) string {
	if s == nil {
		return ""
	}
	// A line series' colour is authored on the outline, so an outline colour
	// wins over any fill colour the model also happens to hold.
	lineColor := s.FillColor
	if s.Outline != nil && s.Outline.Color.ARGB != "" {
		lineColor = s.Outline.Color
	}

	var b strings.Builder
	if lineSeries {
		if lineColor.ARGB != "" {
			b.WriteString(lineElementXML(lineColor, outlineWidthPt(s)))
		}
	} else {
		if s.FillColor.ARGB != "" {
			b.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, colorRGB(s.FillColor)))
		}
		if s.Outline != nil && lineColor.ARGB != "" {
			b.WriteString(lineElementXML(lineColor, outlineWidthPt(s)))
		}
	}
	if b.Len() == 0 {
		return ""
	}
	return "          <c:spPr>" + b.String() + "</c:spPr>\n"
}

// outlineWidthPt returns a series outline width in points, or 0 when the
// document never stated one.
func outlineWidthPt(s *ChartSeries) int {
	if s == nil || s.Outline == nil || s.Outline.Width <= 0 {
		return 0
	}
	return s.Outline.Width
}

// lineElementXML renders an <a:ln> stroke.
//
// widthPt is in points, matching Border.Width and the reader's v/12700
// conversion, so it becomes EMU with the same 12700 factor. An unstated width
// omits the attribute entirely rather than inventing one, leaving PowerPoint to
// apply its own default.
func lineElementXML(c Color, widthPt int) string {
	wAttr := ""
	if widthPt > 0 {
		wAttr = fmt.Sprintf(` w="%d"`, widthPt*12700)
	}
	return fmt.Sprintf(`<a:ln%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>`, wAttr, colorRGB(c))
}

// writeSeriesXML renders the <c:ser> elements shared by every chart type.
//
// withMarker emits the series marker; lineSeries selects the stroked-series
// shape, putting the series colour on the outline instead of a fill. They are
// separate because a radar chart has markers but colours its polygons with a
// fill, while a line chart stroked from the outline is the only type whose
// colour belongs in <a:ln>.
func (w *PPTXWriter) writeSeriesXML(series []*ChartSeries, categories []string, withMarker, lineSeries bool) string {
	var sb strings.Builder
	for idx, s := range series {
		spPrXML := w.writeSeriesSpPrXML(s, lineSeries)

		sb.WriteString(fmt.Sprintf(`        <c:ser>
          <c:idx val="%d"/>
          <c:order val="%d"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>
%s`, idx, idx, xmlEscape(s.Title), spPrXML))

		// Per-point overrides. CT_BarSer/CT_PieSer put the <c:dPt> run after
		// <c:spPr> and before <c:dLbls>; each carries the point's fill.
		for i := 0; i < len(categories); i++ {
			c, ok := s.PointColors[i]
			if !ok || c.ARGB == "" || c.ARGB == "00000000" {
				continue
			}
			sb.WriteString(fmt.Sprintf(`          <c:dPt>
            <c:idx val="%d"/>
            <c:bubble3D val="0"/>
            <c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>
          </c:dPt>
`, i, colorRGB(c)))
		}

		// Data labels. CT_DLbls orders <c:txPr> before <c:dLblPos> and the show
		// flags, and <c:separator> last. The label font goes in the txPr, which
		// is also where PowerPoint looks for it.
		if s.ShowValue || s.ShowCategoryName || s.ShowPercentage || s.ShowSeriesName {
			sb.WriteString("          <c:dLbls>\n")
			sb.WriteString(chartTxPrXML("            ", s.Font))
			if s.LabelPosition != "" {
				sb.WriteString(fmt.Sprintf("            <c:dLblPos val=\"%s\"/>\n", s.LabelPosition))
			}
			if s.ShowValue {
				sb.WriteString("            <c:showVal val=\"1\"/>\n")
			}
			if s.ShowCategoryName {
				sb.WriteString("            <c:showCatName val=\"1\"/>\n")
			}
			if s.ShowSeriesName {
				sb.WriteString("            <c:showSerName val=\"1\"/>\n")
			}
			if s.ShowPercentage {
				sb.WriteString("            <c:showPercent val=\"1\"/>\n")
			}
			if s.Separator != "" && s.Separator != "," {
				sb.WriteString(fmt.Sprintf("            <c:separator>%s</c:separator>\n", xmlEscape(s.Separator)))
			}
			sb.WriteString("          </c:dLbls>\n")
		}

		// Categories
		if len(categories) > 0 {
			sb.WriteString("          <c:cat>\n            <c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache>\n")
			sb.WriteString(fmt.Sprintf("              <c:ptCount val=\"%d\"/>\n", len(categories)))
			for i, cat := range categories {
				sb.WriteString(fmt.Sprintf("              <c:pt idx=\"%d\"><c:v>%s</c:v></c:pt>\n", i, xmlEscape(cat)))
			}
			sb.WriteString("            </c:strCache></c:strRef>\n          </c:cat>\n")
		}

		// Values
		sb.WriteString("          <c:val>\n            <c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache>\n")
		sb.WriteString(fmt.Sprintf("              <c:formatCode>General</c:formatCode>\n              <c:ptCount val=\"%d\"/>\n", len(categories)))
		for i, cat := range categories {
			val := s.Values[cat]
			sb.WriteString(fmt.Sprintf("              <c:pt idx=\"%d\"><c:v>%g</c:v></c:pt>\n", i, val))
		}
		sb.WriteString("            </c:numCache></c:numRef>\n          </c:val>\n")

		if withMarker && s.Marker != nil {
			sb.WriteString(fmt.Sprintf("          <c:marker><c:symbol val=\"%s\"/><c:size val=\"%d\"/></c:marker>\n",
				s.Marker.Symbol, s.Marker.Size))
		}

		sb.WriteString("        </c:ser>\n")
	}
	return sb.String()
}

func (w *PPTXWriter) writeBarChartXML(c *BarChart, cats []string) string {
	return fmt.Sprintf(`      <c:barChart>
        <c:barDir val="%s"/>
        <c:grouping val="%s"/>
        <c:varyColors val="0"/>
%s        <c:gapWidth val="%d"/>
        <c:overlap val="%d"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
`, c.BarDirection, c.BarGrouping, w.writeSeriesXML(c.Series, cats, false, false),
		c.GapWidthPercent, c.OverlapPercent)
}

func (w *PPTXWriter) writeBar3DChartXML(c *Bar3DChart, cats []string) string {
	return fmt.Sprintf(`      <c:bar3DChart>
        <c:barDir val="%s"/>
        <c:grouping val="%s"/>
        <c:varyColors val="0"/>
%s        <c:gapWidth val="%d"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:bar3DChart>
`, c.BarDirection, c.BarGrouping, w.writeSeriesXML(c.Series, cats, false, false),
		c.GapWidthPercent)
}

func (w *PPTXWriter) writeLineChartXML(c *LineChart, cats []string) string {
	smooth := "0"
	if c.IsSmooth {
		smooth = "1"
	}
	seriesXML := w.writeSeriesXML(c.Series, cats, true, true)
	// Add smooth to each series
	seriesXML = strings.ReplaceAll(seriesXML, "</c:ser>",
		fmt.Sprintf("          <c:smooth val=\"%s\"/>\n        </c:ser>", smooth))

	return fmt.Sprintf(`      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
%s        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
`, seriesXML)
}

func (w *PPTXWriter) writeAreaChartXML(c *AreaChart, cats []string) string {
	return fmt.Sprintf(`      <c:areaChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
%s        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:areaChart>
`, w.writeSeriesXML(c.Series, cats, false, false))
}

func (w *PPTXWriter) writePieChartXML(c *PieChart, cats []string) string {
	return fmt.Sprintf(`      <c:pieChart>
        <c:varyColors val="1"/>
%s      </c:pieChart>
`, w.writeSeriesXML(c.Series, cats, false, false))
}

func (w *PPTXWriter) writePie3DChartXML(c *Pie3DChart, cats []string) string {
	return fmt.Sprintf(`      <c:pie3DChart>
        <c:varyColors val="1"/>
%s      </c:pie3DChart>
`, w.writeSeriesXML(c.Series, cats, false, false))
}

func (w *PPTXWriter) writeDoughnutChartXML(c *DoughnutChart, cats []string) string {
	return fmt.Sprintf(`      <c:doughnutChart>
        <c:varyColors val="1"/>
%s        <c:holeSize val="%d"/>
      </c:doughnutChart>
`, w.writeSeriesXML(c.Series, cats, false, false), c.HoleSize)
}

func (w *PPTXWriter) writeScatterChartXML(c *ScatterChart, cats []string) string {
	smooth := "0"
	if c.IsSmooth {
		smooth = "1"
	}

	var sb strings.Builder
	for idx, s := range c.Series {
		// A scatter series is plotted as a stroked line with markers, so its
		// colour belongs on the outline, like a line chart's.
		spPrXML := w.writeSeriesSpPrXML(s, true)
		sb.WriteString(fmt.Sprintf(`        <c:ser>
          <c:idx val="%d"/>
          <c:order val="%d"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>
%s`, idx, idx, xmlEscape(s.Title), spPrXML))

		// X values
		sb.WriteString("          <c:xVal>\n            <c:numRef><c:f>Sheet1!$A$2</c:f><c:numCache>\n")
		sb.WriteString(fmt.Sprintf("              <c:formatCode>General</c:formatCode>\n              <c:ptCount val=\"%d\"/>\n", len(cats)))
		for i, cat := range cats {
			sb.WriteString(fmt.Sprintf("              <c:pt idx=\"%d\"><c:v>%s</c:v></c:pt>\n", i, xmlEscape(cat)))
		}
		sb.WriteString("            </c:numCache></c:numRef>\n          </c:xVal>\n")

		// Y values
		sb.WriteString("          <c:yVal>\n            <c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache>\n")
		sb.WriteString(fmt.Sprintf("              <c:formatCode>General</c:formatCode>\n              <c:ptCount val=\"%d\"/>\n", len(cats)))
		for i, cat := range cats {
			val := s.Values[cat]
			sb.WriteString(fmt.Sprintf("              <c:pt idx=\"%d\"><c:v>%g</c:v></c:pt>\n", i, val))
		}
		sb.WriteString("            </c:numCache></c:numRef>\n          </c:yVal>\n")

		sb.WriteString(fmt.Sprintf("          <c:smooth val=\"%s\"/>\n", smooth))
		sb.WriteString("        </c:ser>\n")
	}

	return fmt.Sprintf(`      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:varyColors val="0"/>
%s        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:scatterChart>
`, sb.String())
}

func (w *PPTXWriter) writeRadarChartXML(c *RadarChart, cats []string) string {
	return fmt.Sprintf(`      <c:radarChart>
        <c:radarStyle val="marker"/>
        <c:varyColors val="0"/>
%s        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:radarChart>
`, w.writeSeriesXML(c.Series, cats, true, false))
}
