package gopresentation

import (
	"image/color"
	"strings"
	"testing"
)

// Chart-area and series styling, end to end.
//
// The chart path had three separate places where a declared style was silently
// thrown away: the reader never committed the chart-area colour it had already
// parsed, the writer never emitted the chart-area <c:spPr> at all, and the
// series outline was dropped on write while the rasteriser computed its width
// with an expression that cancelled to zero. Each of those fails quietly — the
// chart still renders, just not the way the document asked — so each one gets
// an assertion here rather than relying on the golden images, which would
// happily compare a uniformly-thin line against a uniformly-thin line.

// TestChartAreaStyleIsRead covers the reader half. The colour scanner routes a
// <c:chartSpace><c:spPr> colour to the "chartFill"/"chartLine" targets, but for
// those to reach the model every target has to be handled where the colour is
// committed; two of the five were missing, so a chart-area fill parsed into
// nothing.
func TestChartAreaStyleIsRead(t *testing.T) {
	const xmlData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill>
    <a:ln w="19050"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>`

	chart := parseChartXML([]byte(xmlData), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil for a valid chart part")
	}

	if chart.fill == nil || chart.fill.Type != FillSolid {
		t.Fatalf("chart area fill = %+v, want a solid fill", chart.fill)
	}
	if got, want := argbToRGBA(chart.fill.Color), (color.RGBA{R: 0xEE, G: 0xEE, B: 0xEE, A: 0xFF}); got != want {
		t.Errorf("chart area fill colour = %v, want %v", got, want)
	}

	if chart.border == nil {
		t.Fatal("chart area outline was not read")
	}
	if got, want := argbToRGBA(chart.border.Color), (color.RGBA{R: 0x44, G: 0x55, B: 0x66, A: 0xFF}); got != want {
		t.Errorf("chart area outline colour = %v, want %v", got, want)
	}
	// 19050 EMU is 1.5pt, which the model rounds to whole points.
	if chart.border.Width != 2 {
		t.Errorf("chart area outline width = %d, want 2 (points)", chart.border.Width)
	}
}

// TestChartAreaStyleSurvivesRoundTrip covers the writer half: the reader and the
// rasteriser both understand chart.fill/chart.border, so a deck whose chart-area
// shading vanished on save was losing it in the writer and nowhere else.
func TestChartAreaStyleSurvivesRoundTrip(t *testing.T) {
	p := New()
	chart := p.GetActiveSlide().CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
	line := NewLineChart()
	line.AddSeries(NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20}))
	chart.GetPlotArea().SetType(line)
	chart.GetTitle().SetVisible(false)
	chart.GetLegend().Visible = false

	// The reader only ever produces a solid chart-area fill, so set one the
	// same way rather than testing a gradient the round trip cannot carry.
	chart.fill = (&Fill{}).SetSolid(NewColor("EEEEEE"))
	chart.border = (&Border{}).SetSolidFill(NewColor("445566")).SetWidth(2)

	got := firstChartShape(t, roundTrip(t, p))
	if got.fill == nil || got.fill.Type != FillSolid {
		t.Fatalf("chart area fill after round trip = %+v, want a solid fill", got.fill)
	}
	if got, want := argbToRGBA(got.fill.Color), (color.RGBA{R: 0xEE, G: 0xEE, B: 0xEE, A: 0xFF}); got != want {
		t.Errorf("chart area fill colour = %v, want %v", got, want)
	}
	if got.border == nil {
		t.Fatal("chart area outline was lost on save")
	}
	if got, want := argbToRGBA(got.border.Color), (color.RGBA{R: 0x44, G: 0x55, B: 0x66, A: 0xFF}); got != want {
		t.Errorf("chart area outline colour = %v, want %v", got, want)
	}
	if got.border.Width != 2 {
		t.Errorf("chart area outline width = %d, want 2 (points)", got.border.Width)
	}
}

// TestLineSeriesOutlineSurvivesRoundTrip pins both halves of the series
// outline. The width is the part that used to disappear completely, and it is
// written in EMU, so the assertion on the emitted attribute checks the unit as
// well as the presence: a width written as raw points would come back 12700
// times too thin.
func TestLineSeriesOutlineSurvivesRoundTrip(t *testing.T) {
	p := New()
	chart := p.GetActiveSlide().CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
	line := NewLineChart()
	ser := NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20})
	ser.Outline = &SeriesOutline{Width: 3, Color: NewColor("FF0000")}
	line.AddSeries(ser)
	chart.GetPlotArea().SetType(line)
	chart.GetTitle().SetVisible(false)
	chart.GetLegend().Visible = false

	pkg := writeToBytes(t, p)
	part, ok := zipParts(t, pkg)["ppt/charts/chart1.xml"]
	if !ok {
		t.Fatal("chart part was not written")
	}
	if !strings.Contains(string(part), `<a:ln w="38100">`) {
		t.Errorf("chart part does not carry the series outline as 3pt in EMU (38100); got:\n%s", part)
	}

	got := firstChartShape(t, roundTrip(t, p))
	lc, ok := got.GetPlotArea().GetType().(*LineChart)
	if !ok || len(lc.Series) != 1 {
		t.Fatalf("plot type after round trip = %T with %d series, want one *LineChart series",
			got.GetPlotArea().GetType(), len(lc.Series))
	}
	if lc.Series[0].Outline == nil {
		t.Fatal("series outline was lost on save")
	}
	if got, want := lc.Series[0].Outline.Width, 3; got != want {
		t.Errorf("series outline width = %d, want %d", got, want)
	}
	if got, want := argbToRGBA(lc.Series[0].Outline.Color), (color.RGBA{R: 0xFF, A: 0xFF}); got != want {
		t.Errorf("series outline colour = %v, want %v", got, want)
	}
}

// TestLineSeriesColourIsWrittenOnTheOutline checks where the colour goes. A line
// series is stroked, so PowerPoint reads its colour from <c:spPr><a:ln>; the
// writer used to emit a bare <a:solidFill> instead, which PowerPoint ignores.
// Our own reader tolerates that because it falls back to the fill, so only a
// structural assertion catches it.
func TestLineSeriesColourIsWrittenOnTheOutline(t *testing.T) {
	p := New()
	chart := p.GetActiveSlide().CreateChartShape()
	line := NewLineChart()
	ser := NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20})
	ser.SetFillColor(NewColor("FF0000"))
	line.AddSeries(ser)
	chart.GetPlotArea().SetType(line)

	part, ok := zipParts(t, writeToBytes(t, p))["ppt/charts/chart1.xml"]
	if !ok {
		t.Fatal("chart part was not written")
	}
	text := string(part)
	if !strings.Contains(text, `<c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>`) {
		t.Errorf("line series colour is not on the outline; got:\n%s", text)
	}
	if strings.Contains(text, `<c:spPr><a:solidFill>`) {
		t.Errorf("line series still emits a bare fill, which PowerPoint ignores for a stroked series; got:\n%s", text)
	}
}

// TestSeriesStrokeWidthScalesWithOutline is the rasteriser half: an outline
// width has to reach the canvas. The expression it replaced multiplied and
// divided by 12700 in the same term, so the width collapsed towards zero and
// every series line was clamped to the 1px floor whatever the document said.
// Counting ink rather than measuring a single row keeps this honest under
// anti-aliasing.
func TestSeriesStrokeWidthScalesWithOutline(t *testing.T) {
	fc := NewFontCache()

	inkFor := func(widthPt int) int {
		p := New()
		chart := p.GetActiveSlide().CreateChartShape()
		chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
		chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
		line := NewLineChart()
		ser := NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 10})
		ser.SetFillColor(NewColor("FF0000"))
		// No markers: they would contribute ink of their own and dilute the
		// difference the line width makes.
		ser.Marker = &SeriesMarker{Symbol: MarkerNone}
		ser.Outline = &SeriesOutline{Width: widthPt}
		line.AddSeries(ser)
		chart.GetPlotArea().SetType(line)
		chart.GetTitle().SetVisible(false)
		chart.GetLegend().Visible = false

		opts := DefaultRenderOptions()
		opts.Width = 960
		opts.FontCache = fc
		img, err := p.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		return countNearColorIn(img, img.Bounds(), color.RGBA{R: 0xFF, A: 0xFF}, 40)
	}

	thin := inkFor(1)
	thick := inkFor(4)

	if thin == 0 {
		t.Fatal("the series line drew nothing; the red-counting assertion is meaningless")
	}
	// 4pt at 960px wide is 4*12700*(960/9144000) ≈ 5px, against ≈1px for 1pt,
	// so the ratio is ~5. Requiring 3 leaves room for anti-aliased edges.
	if thick < 3*thin {
		t.Errorf("series ink = %d px at 4pt vs %d px at 1pt; the outline width is not reaching the canvas", thick, thin)
	}
}

// firstChartShape returns the chart on the presentation's first slide.
func firstChartShape(t *testing.T, pres *Presentation) *ChartShape {
	t.Helper()
	slides := pres.GetAllSlides()
	if len(slides) == 0 {
		t.Fatal("round trip produced no slides")
	}
	for _, sh := range slides[0].GetShapes() {
		if cs, ok := sh.(*ChartShape); ok {
			return cs
		}
	}
	t.Fatal("chart shape was dropped by the round trip")
	return nil
}
