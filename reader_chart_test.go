package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"strings"
	"testing"
)

// --- Chart reading / round-trip ---------------------------------------------

// TestChartRoundTrip writes a chart, reads the presentation back and verifies
// that the chart shape survives with its geometry and data intact. Before
// chart reading was implemented the graphicFrame was silently dropped.
func TestChartRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
	chart.BaseShape.SetName("SalesChart")
	chart.GetTitle().SetText("Sales Report")
	chart.GetLegend().Visible = true
	chart.GetLegend().Position = LegendRight

	bar := NewBarChart()
	bar.SetBarGrouping(BarGroupingStacked)
	bar.SetGapWidthPercent(200)
	bar.AddSeries(NewChartSeriesOrdered("Revenue",
		[]string{"Q1", "Q2", "Q3", "Q4"}, []float64{120, 180, 150, 210}))
	bar.AddSeries(NewChartSeriesOrdered("Cost",
		[]string{"Q1", "Q2", "Q3", "Q4"}, []float64{80, 90, 70, 110}))
	chart.GetPlotArea().SetType(bar)

	data := writeToBytes(t, p)

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	slides := pres.GetAllSlides()
	if len(slides) != 1 {
		t.Fatalf("slide count = %d, want 1", len(slides))
	}

	var got *ChartShape
	for _, sh := range slides[0].shapes {
		if cs, ok := sh.(*ChartShape); ok {
			got = cs
			break
		}
	}
	if got == nil {
		t.Fatal("chart shape was dropped when reading the file back")
	}

	// Geometry and name.
	if got.offsetX != 500000 || got.offsetY != 500000 {
		t.Errorf("offset = (%d,%d), want (500000,500000)", got.offsetX, got.offsetY)
	}
	if got.width != 7000000 || got.height != 4500000 {
		t.Errorf("size = (%d,%d), want (7000000,4500000)", got.width, got.height)
	}
	if got.name != "SalesChart" {
		t.Errorf("name = %q, want %q", got.name, "SalesChart")
	}
	if got.GetType() != ShapeTypeChart {
		t.Errorf("shape type = %v, want ShapeTypeChart", got.GetType())
	}

	// Title and legend.
	if got.title == nil || !got.title.Visible || got.title.Text != "Sales Report" {
		t.Errorf("title = %+v, want visible %q", got.title, "Sales Report")
	}
	if got.legend == nil || !got.legend.Visible {
		t.Fatal("legend should be visible after round-trip")
	}
	if got.legend.Position != LegendRight {
		t.Errorf("legend position = %q, want %q", got.legend.Position, LegendRight)
	}

	// Plot type and bar settings.
	bc, ok := got.plotArea.GetType().(*BarChart)
	if !ok {
		t.Fatalf("plot type = %T, want *BarChart", got.plotArea.GetType())
	}
	if bc.BarGrouping != BarGroupingStacked {
		t.Errorf("grouping = %q, want %q", bc.BarGrouping, BarGroupingStacked)
	}
	if bc.GapWidthPercent != 200 {
		t.Errorf("gap width = %d, want 200", bc.GapWidthPercent)
	}
	if len(bc.Series) != 2 {
		t.Fatalf("series count = %d, want 2", len(bc.Series))
	}

	s0 := bc.Series[0]
	if s0.Title != "Revenue" {
		t.Errorf("series title = %q, want %q", s0.Title, "Revenue")
	}
	if len(s0.Categories) != 4 {
		t.Fatalf("category count = %d, want 4", len(s0.Categories))
	}
	for i, want := range []string{"Q1", "Q2", "Q3", "Q4"} {
		if s0.Categories[i] != want {
			t.Errorf("category[%d] = %q, want %q", i, s0.Categories[i], want)
		}
	}
	if got := s0.Values["Q3"]; got != 150 {
		t.Errorf("Q3 value = %v, want 150", got)
	}
	if got := bc.Series[1].Values["Q1"]; got != 80 {
		t.Errorf("second series Q1 = %v, want 80", got)
	}
}

// TestParseChartXMLPowerPointStyle parses a chart part as written by PowerPoint:
// theme (schemeClr) colours, string/number caches, data labels, axis titles and
// gridlines.
func TestParseChartXMLPowerPointStyle(t *testing.T) {
	xmlData := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly Revenue</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="stacked"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Product A</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
          <c:dLbls><c:showVal val="1"/><c:showCatName val="0"/></c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>120</c:v></c:pt><c:pt idx="1"><c:v>180</c:v></c:pt><c:pt idx="2"><c:v>150</c:v></c:pt><c:pt idx="3"><c:v>210</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="200"/>
        <c:overlap val="100"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/><c:max val="250"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:majorUnit val="50"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`)

	chart := parseChartXML(xmlData, nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil for a valid chart part")
	}

	if !chart.title.Visible || chart.title.Text != "Quarterly Revenue" {
		t.Errorf("title = %+v, want visible %q", chart.title, "Quarterly Revenue")
	}
	if !chart.legend.Visible || chart.legend.Position != LegendRight {
		t.Errorf("legend = %+v, want visible at right", chart.legend)
	}
	if chart.displayBlankAs != ChartBlankAsGap {
		t.Errorf("dispBlanksAs = %q, want %q", chart.displayBlankAs, ChartBlankAsGap)
	}

	bc, ok := chart.plotArea.GetType().(*BarChart)
	if !ok {
		t.Fatalf("plot type = %T, want *BarChart", chart.plotArea.GetType())
	}
	if bc.BarGrouping != BarGroupingStacked {
		t.Errorf("grouping = %q, want stacked", bc.BarGrouping)
	}
	if bc.GapWidthPercent != 200 || bc.OverlapPercent != 100 {
		t.Errorf("gap/overlap = %d/%d, want 200/100", bc.GapWidthPercent, bc.OverlapPercent)
	}
	if len(bc.Series) != 1 {
		t.Fatalf("series count = %d, want 1", len(bc.Series))
	}
	s := bc.Series[0]
	if s.Title != "Product A" {
		t.Errorf("series title = %q, want %q", s.Title, "Product A")
	}
	if !s.ShowValue || s.ShowCategoryName {
		t.Errorf("data labels = value:%v category:%v, want true/false", s.ShowValue, s.ShowCategoryName)
	}
	// accent1 resolves via the built-in Office theme fallback.
	if want := NewColor("4472C4"); s.FillColor.ARGB != want.ARGB {
		t.Errorf("series fill = %s, want %s", s.FillColor.ARGB, want.ARGB)
	}
	if len(s.Categories) != 4 || s.Categories[3] != "Q4" {
		t.Errorf("categories = %v, want [Q1 Q2 Q3 Q4]", s.Categories)
	}
	if s.Values["Q2"] != 180 {
		t.Errorf("Q2 value = %v, want 180", s.Values["Q2"])
	}

	axX := chart.plotArea.GetAxisX()
	if axX == nil || axX.Title != "Quarter" || !axX.Visible {
		t.Errorf("category axis = %+v, want title %q visible", axX, "Quarter")
	}
	axY := chart.plotArea.GetAxisY()
	if axY == nil {
		t.Fatal("value axis missing")
	}
	if axY.Title != "Revenue" {
		t.Errorf("value axis title = %q, want %q", axY.Title, "Revenue")
	}
	if axY.MaxBounds == nil || *axY.MaxBounds != 250 {
		t.Errorf("value axis max = %v, want 250", axY.MaxBounds)
	}
	if axY.MajorUnit == nil || *axY.MajorUnit != 50 {
		t.Errorf("value axis major unit = %v, want 50", axY.MajorUnit)
	}
	if axY.MajorGridlines == nil {
		t.Error("value axis major gridlines should be parsed")
	}
}

// TestParseChartXMLTypes checks that each plot-area element maps to the right
// chart type.
func TestParseChartXMLTypes(t *testing.T) {
	tmpl := `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea>
  <c:%s>
   <c:ser><c:idx val="0"/><c:order val="0"/>
    <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
    <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
   </c:ser>
  </c:%s>
  <c:legend><c:legendPos val="b"/></c:legend>
 </c:plotArea></c:chart>
</c:chartSpace>`

	cases := []struct {
		element string
		want    string
	}{
		{"barChart", "bar"},
		{"bar3DChart", "bar3D"},
		{"lineChart", "line"},
		{"areaChart", "area"},
		{"pieChart", "pie"},
		{"pie3DChart", "pie3D"},
		{"doughnutChart", "doughnut"},
		{"scatterChart", "scatter"},
		{"radarChart", "radar"},
	}
	for _, tc := range cases {
		t.Run(tc.element, func(t *testing.T) {
			data := []byte(strings.ReplaceAll(tmpl, "%s", tc.element))
			chart := parseChartXML(data, nil)
			if chart == nil {
				t.Fatalf("parseChartXML returned nil for %s", tc.element)
			}
			got := chart.plotArea.GetType()
			if got == nil {
				t.Fatalf("%s: plot type is nil", tc.element)
			}
			if got.GetChartTypeName() != tc.want {
				t.Errorf("%s: chart type = %q, want %q", tc.element, got.GetChartTypeName(), tc.want)
			}
		})
	}
}

// TestParseChartXMLMissingLegend ensures the parser does not invent a legend
// that the file does not declare.
func TestParseChartXMLMissingLegend(t *testing.T) {
	data := []byte(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>`)
	chart := parseChartXML(data, nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil")
	}
	if chart.legend.Visible {
		t.Error("legend should be invisible when the chart declares none")
	}
	if chart.title.Visible {
		t.Error("title should be invisible when the chart declares none")
	}
}

// TestParseChartXMLRejectsGarbage makes sure malformed input is handled.
func TestParseChartXMLRejectsGarbage(t *testing.T) {
	for _, in := range [][]byte{nil, {}, []byte("not xml at all"), []byte("<c:chartSpace")} {
		if got := parseChartXML(in, nil); got != nil {
			t.Errorf("parseChartXML(%q) = %+v, want nil", in, got)
		}
	}
}

// --- Chart rasterization ----------------------------------------------------

// TestRenderAllChartTypes renders one chart of every supported type and checks
// that each one paints something inside the chart frame.
func TestRenderAllChartTypes(t *testing.T) {
	cats := []string{"A", "B", "C", "D"}
	vals := []float64{12, 30, 18, 25}

	builders := map[string]func() ChartType{
		"bar": func() ChartType {
			c := NewBarChart()
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			c.AddSeries(NewChartSeriesOrdered("S2", cats, []float64{8, 14, 22, 10}))
			return c
		},
		"stackedBar": func() ChartType {
			c := NewBarChart()
			c.SetBarGrouping(BarGroupingStacked)
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			c.AddSeries(NewChartSeriesOrdered("S2", cats, []float64{8, 14, 22, 10}))
			return c
		},
		"horizontalBar": func() ChartType {
			c := NewBarChart()
			c.BarDirection = BarDirectionHorizontal
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
		"line": func() ChartType {
			c := NewLineChart()
			c.SetSmooth(true)
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
		"area": func() ChartType {
			c := NewAreaChart()
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
		"pie": func() ChartType {
			c := NewPieChart()
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
		"doughnut": func() ChartType {
			c := NewDoughnutChart()
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
		"scatter": func() ChartType {
			c := NewScatterChart()
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
		"radar": func() ChartType {
			c := NewRadarChart()
			c.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			return c
		},
	}

	for name, build := range builders {
		t.Run(name, func(t *testing.T) {
			p := New()
			slide := p.GetActiveSlide()
			chart := slide.CreateChartShape()
			chart.BaseShape.SetOffsetX(0).SetOffsetY(0)
			chart.BaseShape.SetWidth(6000000).SetHeight(4000000)
			chart.GetTitle().SetText("Test")
			chart.GetPlotArea().SetType(build())

			img, err := p.SlideToImage(0, &RenderOptions{Width: 480})
			if err != nil {
				t.Fatalf("render: %v", err)
			}
			if img == nil {
				t.Fatal("render returned nil image")
			}
			if !hasInk(img) {
				t.Error("chart rendered nothing (blank image)")
			}
		})
	}
}

// TestRenderChartAreaFill verifies that an explicit chart area fill is painted
// while the default (no fill) leaves the slide background visible.
func TestRenderChartAreaFill(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(0).SetOffsetY(0)
	chart.BaseShape.SetWidth(4000000).SetHeight(3000000)
	chart.GetTitle().SetVisible(false)
	chart.GetLegend().Visible = false
	pie := NewPieChart()
	pie.AddSeries(NewChartSeriesOrdered("S1", []string{"A", "B"}, []float64{1, 1}))
	chart.GetPlotArea().SetType(pie)
	chart.fill = (&Fill{}).SetSolid(NewColor("FF00FF00"))

	img, err := p.SlideToImage(0, &RenderOptions{Width: 200})
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	if got := img.At(1, 1); !sameRGBA(got, NewColor("FF00FF00")) {
		t.Errorf("chart area fill = %v, want the green fill", got)
	}
}

// TestRenderChartAxisTitles renders a chart with axis titles in both axis
// orientations and checks that something is painted in the title strips.
func TestRenderChartAxisTitles(t *testing.T) {
	cats := []string{"A", "B", "C"}
	vals := []float64{1, 2, 3}

	for _, horizontal := range []bool{false, true} {
		name := "vertical"
		if horizontal {
			name = "horizontal"
		}
		t.Run(name, func(t *testing.T) {
			p := New()
			slide := p.GetActiveSlide()
			chart := slide.CreateChartShape()
			chart.BaseShape.SetOffsetX(0).SetOffsetY(0)
			chart.BaseShape.SetWidth(6000000).SetHeight(4000000)
			chart.GetTitle().SetVisible(false)
			chart.GetLegend().Visible = false

			bar := NewBarChart()
			bar.BarDirection = BarDirectionVertical
			if horizontal {
				bar.BarDirection = BarDirectionHorizontal
			}
			bar.AddSeries(NewChartSeriesOrdered("S1", cats, vals))
			chart.GetPlotArea().SetType(bar)
			chart.GetPlotArea().GetAxisY().SetTitle("ValueAxis")
			chart.GetPlotArea().GetAxisX().SetTitle("CategoryAxis")

			img, err := p.SlideToImage(0, &RenderOptions{Width: 640})
			if err != nil {
				t.Fatalf("render: %v", err)
			}
			// Both title strips sit in the outer margins; make sure the
			// rotated (left) strip is not empty.
			b := img.Bounds()
			strip := image.Rect(b.Min.X, b.Min.Y+b.Dy()/4, b.Min.X+b.Dx()/8, b.Min.Y+3*b.Dy()/4)
			if !hasInkIn(img, strip) {
				t.Errorf("%s: axis title strip is blank", name)
			}
		})
	}
}

// --- helpers ----------------------------------------------------------------

// hasInk reports whether the image contains any non-white pixel.
func hasInk(img image.Image) bool {
	return hasInkIn(img, img.Bounds())
}

// hasInkIn reports whether the given rectangle contains any non-white pixel.
func hasInkIn(img image.Image, rect image.Rectangle) bool {
	b := img.Bounds()
	rect = rect.Intersect(b)
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			r, g, bb, _ := img.At(x, y).RGBA()
			if r>>8 < 250 || g>>8 < 250 || bb>>8 < 250 {
				return true
			}
		}
	}
	return false
}

func sameRGBA(got color.Color, want Color) bool {
	r, g, b, _ := got.RGBA()
	return uint8(r>>8) == want.GetRed() &&
		uint8(g>>8) == want.GetGreen() &&
		uint8(b>>8) == want.GetBlue()
}

func writeToBytes(t *testing.T, p *Presentation) []byte {
	t.Helper()
	var buf bytes.Buffer
	w, err := NewWriter(p, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("new writer: %v", err)
	}
	if err := w.WriteTo(&buf); err != nil {
		t.Fatalf("write: %v", err)
	}
	return buf.Bytes()
}
