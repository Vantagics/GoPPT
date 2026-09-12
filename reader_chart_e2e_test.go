package gopresentation

import (
	"archive/zip"
	"fmt"
	"image"
	"os"
	"path/filepath"
	"testing"
)

// This file covers the "write -> read -> render" loop for chart slides as a
// whole. The unit-level tests in reader_chart_test.go exercise the chart part
// parser directly; these tests instead go through the public, file-based path
// that a preview tool actually uses:
//
//	writer -> .pptx on disk -> Open() -> GetShapes() -> SlideToImage()
//
// The failure this guards against is a slide that reads back with only its
// text shapes, so the rendered page shows the title and nothing else.

// chartSlidePresentation builds a one-slide presentation containing a titled
// bar chart plus a title text box, mimicking a generated "chart" slide.
func chartSlidePresentation() *Presentation {
	p := New()
	slide := p.GetActiveSlide()

	title := slide.CreateRichTextShape()
	title.BaseShape.SetOffsetX(500000).SetOffsetY(300000)
	title.BaseShape.SetWidth(8000000).SetHeight(600000)
	title.CreateTextRun("Quarterly Sales")

	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(1000000)
	chart.BaseShape.SetWidth(8000000).SetHeight(4500000)
	chart.BaseShape.SetName("SalesChart")
	chart.GetTitle().SetText("Revenue by Quarter")
	chart.GetTitle().SetVisible(true)
	chart.GetLegend().Visible = true
	chart.GetLegend().Position = LegendBottom

	bar := NewBarChart()
	bar.SetBarGrouping(BarGroupingClustered)
	bar.AddSeries(NewChartSeriesOrdered("Revenue",
		[]string{"Q1", "Q2", "Q3", "Q4"}, []float64{120, 180, 150, 210}))
	bar.AddSeries(NewChartSeriesOrdered("Cost",
		[]string{"Q1", "Q2", "Q3", "Q4"}, []float64{80, 90, 70, 110}))
	chart.GetPlotArea().SetType(bar)

	return p
}

// writeChartPPTX writes the presentation to a temp .pptx and returns its path.
func writeChartPPTX(t *testing.T, p *Presentation) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "chart.pptx")
	w, err := NewWriter(p, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("new writer: %v", err)
	}
	pw, ok := w.(*PPTXWriter)
	if !ok {
		t.Fatalf("writer type = %T, want *PPTXWriter", w)
	}
	if err := pw.Save(path); err != nil {
		t.Fatalf("save: %v", err)
	}
	return path
}

// TestChartRoundTripThroughFile reproduces the reported blocker end to end:
// the file must contain a chart part, Open() must surface a ChartShape, and the
// chart region must actually render ink.
func TestChartRoundTripThroughFile(t *testing.T) {
	path := writeChartPPTX(t, chartSlidePresentation())

	// 1. The writer must have emitted a native chart part.
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()
	var chartParts int
	for _, f := range zr.File {
		if len(f.Name) > len("ppt/charts/") && f.Name[:len("ppt/charts/")] == "ppt/charts/" {
			chartParts++
		}
	}
	if chartParts == 0 {
		t.Fatal("writer emitted no ppt/charts/chartN.xml part")
	}

	// 2. Open() must surface the chart as a real shape.
	pres, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	slides := pres.GetAllSlides()
	if len(slides) != 1 {
		t.Fatalf("slide count = %d, want 1", len(slides))
	}
	shapes := slides[0].GetShapes()

	var got *ChartShape
	texts := 0
	for _, sh := range shapes {
		switch s := sh.(type) {
		case *ChartShape:
			got = s
		case *RichTextShape:
			texts++
		}
	}
	if got == nil {
		types := make([]string, 0, len(shapes))
		for _, sh := range shapes {
			types = append(types, sh.GetType().String())
		}
		t.Fatalf("chart dropped by reader: GetShapes() = %v (only %d text shape(s))", types, texts)
	}
	if texts == 0 {
		t.Error("title text shape should also survive the round-trip")
	}
	if got.name != "SalesChart" {
		t.Errorf("chart name = %q, want %q", got.name, "SalesChart")
	}
	if got.width == 0 || got.height == 0 {
		t.Errorf("chart geometry not restored: %dx%d", got.width, got.height)
	}
	bar, ok := got.GetPlotArea().GetType().(*BarChart)
	if !ok {
		t.Fatalf("chart type = %T, want *BarChart", got.GetPlotArea().GetType())
	}
	if len(bar.Series) != 2 {
		t.Fatalf("series count = %d, want 2", len(bar.Series))
	}
	if v := bar.Series[0].Values["Q3"]; v != 150 {
		t.Errorf("series[0][Q3] = %v, want 150", v)
	}
	if v := bar.Series[1].Values["Q4"]; v != 110 {
		t.Errorf("series[1][Q4] = %v, want 110", v)
	}

	// 3. The chart must actually render. The title text alone is not enough —
	//    check the chart body region, which sits below the title text box.
	opts := DefaultRenderOptions()
	opts.Width = 960
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}
	if !hasInk(img) {
		t.Fatal("rendered slide is blank")
	}
	// The chart occupies roughly the lower two thirds of the slide. Restrict the
	// ink check to the middle band so a title-only render fails the test.
	b := img.Bounds()
	band := image.Rect(b.Min.X, b.Min.Y+b.Dy()*35/100, b.Max.X, b.Min.Y+b.Dy()*85/100)
	if !hasInkIn(img, band) {
		t.Fatal("chart region rendered blank: slide read back with text only")
	}
}

// TestChartRoundTripEveryTypeFile goes through the file API for each chart type
// so a regression in any single type is caught at the integration level too.
func TestChartRoundTripEveryTypeFile(t *testing.T) {
	cats := []string{"Q1", "Q2", "Q3", "Q4"}
	vals := []float64{10, 25, 18, 30}

	cases := []struct {
		name  string
		build func() ChartType
		want  string
	}{
		{"bar", func() ChartType {
			c := NewBarChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.BarChart"},
		{"bar3D", func() ChartType {
			c := NewBar3DChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.Bar3DChart"},
		{"line", func() ChartType {
			c := NewLineChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.LineChart"},
		{"area", func() ChartType {
			c := NewAreaChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.AreaChart"},
		{"pie", func() ChartType {
			c := NewPieChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.PieChart"},
		{"pie3D", func() ChartType {
			c := NewPie3DChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.Pie3DChart"},
		{"doughnut", func() ChartType {
			c := NewDoughnutChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.DoughnutChart"},
		{"scatter", func() ChartType {
			c := NewScatterChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.ScatterChart"},
		{"radar", func() ChartType {
			c := NewRadarChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}, "*gopresentation.RadarChart"},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			p := New()
			chart := p.GetActiveSlide().CreateChartShape()
			chart.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
			chart.BaseShape.SetWidth(8000000).SetHeight(5000000)
			chart.GetPlotArea().SetType(tc.build())

			path := writeChartPPTX(t, p)
			pres, err := Open(path)
			if err != nil {
				t.Fatalf("Open: %v", err)
			}
			shapes := pres.GetAllSlides()[0].GetShapes()
			var got *ChartShape
			for _, sh := range shapes {
				if cs, ok := sh.(*ChartShape); ok {
					got = cs
				}
			}
			if got == nil {
				t.Fatal("chart shape dropped on read-back")
			}
			if name := fmt.Sprintf("%T", got.GetPlotArea().GetType()); name != tc.want {
				t.Errorf("chart type = %s, want %s", name, tc.want)
			}
		})
	}
}

// TestOpenNonexistentAndGarbageFile verifies Open fails cleanly (error, no
// panic) for inputs that are not valid packages.
func TestOpenNonexistentAndGarbageFile(t *testing.T) {
	if _, err := Open(filepath.Join(t.TempDir(), "does-not-exist.pptx")); err == nil {
		t.Error("Open(missing file) = nil error, want error")
	}

	dir := t.TempDir()
	junk := filepath.Join(dir, "junk.pptx")
	if err := os.WriteFile(junk, []byte("this is definitely not a zip archive"), 0o644); err != nil {
		t.Fatalf("write junk: %v", err)
	}
	if _, err := Open(junk); err == nil {
		t.Error("Open(garbage) = nil error, want error")
	}
}
