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
		titleXML = fmt.Sprintf(`  <c:title>
    <c:tx>
      <c:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="%d" b="%s"/>
            <a:t>%s</a:t>
          </a:r>
        </a:p>
      </c:rich>
    </c:tx>
    <c:overlay val="0"/>
  </c:title>
`, chart.title.Font.Size*100, boolToXML(chart.title.Font.Bold), xmlEscape(chart.title.Text))
	} else if !chart.title.Visible {
		titleXML = `  <c:autoTitleDeleted val="1"/>
`
	}

	// Legend XML
	legendXML := ""
	if chart.legend.Visible {
		legendXML = fmt.Sprintf(`  <c:legend>
    <c:legendPos val="%s"/>
    <c:overlay val="0"/>
  </c:legend>
`, chart.legend.Position)
	}

	// Axis XML
	axisXML := ""
	if !isPieType(ct) {
		axisXML = w.writeAxesXML(chart)
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="%s" xmlns:r="%s">
  <c:chart>
%s%s    <c:plotArea>
      <c:layout/>
%s%s    </c:plotArea>
%s    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="%s"/>
  </c:chart>
</c:chartSpace>`,
		nsDrawingML, nsOfficeDocRels,
		titleXML, "",
		chartTypeXML.String(), axisXML,
		legendXML,
		chart.displayBlankAs)

	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/charts/chart%d.xml", chartIdx), content)
}

func boolToXML(v bool) string {
	if v {
		return "1"
	}
	return "0"
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

	catAxisXML := fmt.Sprintf(`      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="%s"/></c:scaling>
        <c:delete val="%s"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:crosses val="%s"/>
        <c:tickLblPos val="%s"/>
`, w.axisOrientation(axX), boolToXML(!axX.Visible), axX.CrossesAt, axX.TickLabelPos)

	if axX.Title != "" {
		catAxisXML += fmt.Sprintf(`        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx></c:title>
`, xmlEscape(axX.Title))
	}
	if axX.MajorGridlines != nil {
		catAxisXML += w.writeGridlinesXML("c:majorGridlines", axX.MajorGridlines)
	}
	catAxisXML += "      </c:catAx>\n"

	valAxisXML := fmt.Sprintf(`      <c:valAx>
        <c:axId val="2"/>
        <c:scaling>
          <c:orientation val="%s"/>`, w.axisOrientation(axY))

	if axY.MinBounds != nil {
		valAxisXML += fmt.Sprintf(`
          <c:min val="%g"/>`, *axY.MinBounds)
	}
	if axY.MaxBounds != nil {
		valAxisXML += fmt.Sprintf(`
          <c:max val="%g"/>`, *axY.MaxBounds)
	}
	valAxisXML += `
        </c:scaling>
`
	valAxisXML += fmt.Sprintf(`        <c:delete val="%s"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
        <c:crosses val="%s"/>
        <c:tickLblPos val="%s"/>
`, boolToXML(!axY.Visible), axY.CrossesAt, axY.TickLabelPos)

	if axY.MajorUnit != nil {
		valAxisXML += fmt.Sprintf(`        <c:majorUnit val="%g"/>
`, *axY.MajorUnit)
	}
	if axY.MinorUnit != nil {
		valAxisXML += fmt.Sprintf(`        <c:minorUnit val="%g"/>
`, *axY.MinorUnit)
	}
	if axY.Title != "" {
		valAxisXML += fmt.Sprintf(`        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx></c:title>
`, xmlEscape(axY.Title))
	}
	if axY.MajorGridlines != nil {
		valAxisXML += w.writeGridlinesXML("c:majorGridlines", axY.MajorGridlines)
	}
	if axY.MinorGridlines != nil {
		valAxisXML += w.writeGridlinesXML("c:minorGridlines", axY.MinorGridlines)
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

func (w *PPTXWriter) writeSeriesXML(series []*ChartSeries, categories []string, withMarker bool) string {
	var sb strings.Builder
	for idx, s := range series {
		fillXML := ""
		if s.FillColor.ARGB != "" {
			fillXML = fmt.Sprintf(`          <c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>
`, colorRGB(s.FillColor))
		}

		sb.WriteString(fmt.Sprintf(`        <c:ser>
          <c:idx val="%d"/>
          <c:order val="%d"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>
%s`, idx, idx, xmlEscape(s.Title), fillXML))

		// Data labels
		if s.ShowValue || s.ShowCategoryName || s.ShowPercentage || s.ShowSeriesName {
			sb.WriteString("          <c:dLbls>\n")
			if s.ShowValue {
				sb.WriteString("            <c:showVal val=\"1\"/>\n")
			}
			if s.ShowCategoryName {
				sb.WriteString("            <c:showCatName val=\"1\"/>\n")
			}
			if s.ShowPercentage {
				sb.WriteString("            <c:showPercent val=\"1\"/>\n")
			}
			if s.ShowSeriesName {
				sb.WriteString("            <c:showSerName val=\"1\"/>\n")
			}
			if s.Separator != "" && s.Separator != "," {
				sb.WriteString(fmt.Sprintf("            <c:separator>%s</c:separator>\n", xmlEscape(s.Separator)))
			}
			if s.LabelPosition != "" {
				sb.WriteString(fmt.Sprintf("            <c:dLblPos val=\"%s\"/>\n", s.LabelPosition))
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
`, c.BarDirection, c.BarGrouping, w.writeSeriesXML(c.Series, cats, false),
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
`, c.BarDirection, c.BarGrouping, w.writeSeriesXML(c.Series, cats, false),
		c.GapWidthPercent)
}

func (w *PPTXWriter) writeLineChartXML(c *LineChart, cats []string) string {
	smooth := "0"
	if c.IsSmooth {
		smooth = "1"
	}
	seriesXML := w.writeSeriesXML(c.Series, cats, true)
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
`, w.writeSeriesXML(c.Series, cats, false))
}

func (w *PPTXWriter) writePieChartXML(c *PieChart, cats []string) string {
	return fmt.Sprintf(`      <c:pieChart>
        <c:varyColors val="1"/>
%s      </c:pieChart>
`, w.writeSeriesXML(c.Series, cats, false))
}

func (w *PPTXWriter) writePie3DChartXML(c *Pie3DChart, cats []string) string {
	return fmt.Sprintf(`      <c:pie3DChart>
        <c:varyColors val="1"/>
%s      </c:pie3DChart>
`, w.writeSeriesXML(c.Series, cats, false))
}

func (w *PPTXWriter) writeDoughnutChartXML(c *DoughnutChart, cats []string) string {
	return fmt.Sprintf(`      <c:doughnutChart>
        <c:varyColors val="1"/>
%s        <c:holeSize val="%d"/>
      </c:doughnutChart>
`, w.writeSeriesXML(c.Series, cats, false), c.HoleSize)
}

func (w *PPTXWriter) writeScatterChartXML(c *ScatterChart, cats []string) string {
	smooth := "0"
	if c.IsSmooth {
		smooth = "1"
	}

	var sb strings.Builder
	for idx, s := range c.Series {
		fillXML := ""
		if s.FillColor.ARGB != "" {
			fillXML = fmt.Sprintf(`          <c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>
`, colorRGB(s.FillColor))
		}
		sb.WriteString(fmt.Sprintf(`        <c:ser>
          <c:idx val="%d"/>
          <c:order val="%d"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>
%s`, idx, idx, xmlEscape(s.Title), fillXML))

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
`, w.writeSeriesXML(c.Series, cats, true))
}
