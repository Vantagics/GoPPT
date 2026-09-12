package gopresentation

import (
	"bytes"
	"flag"
	"fmt"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// Golden-image regression tests.
//
// Run with -update to (re)generate the reference images:
//
//	go test -run Golden -update ./...
//
// These tests compare a rendered slide against a committed reference PNG. Font
// rasterization differs between operating systems and font versions, so the
// comparison is tolerant and the fixtures that need a specific kind of font are
// skipped when it is unavailable. The font-independent structural checks in
// renderer_regression_test.go are the ones that must hold everywhere.
//
// Regenerating goldens on a different machine is expected: the point of these
// images is to catch *layout* regressions (a shape moving, a chart losing its
// axis, a table losing its grid), not to pin exact glyph pixels.

var updateGolden = flag.Bool("update", false, "regenerate golden image files under testdata/render")

const goldenDir = "testdata/render"

// Tolerances. A per-channel allowance absorbs anti-aliasing differences; the
// ratio allowance absorbs a small number of pixels flipping near a glyph edge.
const (
	goldenChannelTolerance = 16
	goldenMaxDiffRatio     = 0.01
)

// goldenOptions returns render options with a fixed width and a shared font
// cache, so a golden render is as reproducible as the machine allows.
func goldenOptions(fc *FontCache) *RenderOptions {
	opts := DefaultRenderOptions()
	opts.Width = 640
	opts.FontCache = fc
	return opts
}

// requireAnyFont skips the test when the machine has no usable fonts, since
// every golden here draws at least some text.
func requireAnyFont(t *testing.T, fc *FontCache) {
	t.Helper()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine; golden images would be meaningless")
	}
}

// requireCJKFont skips the test when no CJK-capable font is installed.
func requireCJKFont(t *testing.T, fc *FontCache) {
	t.Helper()
	requireAnyFont(t, fc)
	for _, name := range defaultCJKFallbackChain {
		if fc.HasFont(name, false, false) {
			return
		}
	}
	t.Skip("no CJK-capable font installed; CJK goldens would just record tofu")
}

// goldenPath returns the reference image path for a case name.
func goldenPath(name string) string {
	return filepath.Join(goldenDir, name+".png")
}

// checkGolden renders the presentation's first slide and compares it against the
// stored reference, or writes the reference when -update is given.
func checkGolden(t *testing.T, name string, pres *Presentation, opts *RenderOptions) {
	t.Helper()

	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render %s: %v", name, err)
	}

	if *updateGolden {
		if err := os.MkdirAll(goldenDir, 0o755); err != nil {
			t.Fatalf("create golden dir: %v", err)
		}
		if err := writePNG(goldenPath(name), img); err != nil {
			t.Fatalf("write golden %s: %v", name, err)
		}
		t.Logf("updated golden %s", goldenPath(name))
		return
	}

	want, err := readPNG(goldenPath(name))
	if err != nil {
		if os.IsNotExist(err) {
			t.Fatalf("golden %s is missing; run `go test -run Golden -update ./...` to create it", goldenPath(name))
		}
		t.Fatalf("read golden %s: %v", goldenPath(name), err)
	}

	if want.Bounds() != img.Bounds() {
		t.Errorf("%s: rendered size %v, golden size %v", name, img.Bounds().Size(), want.Bounds().Size())
		return
	}

	diff, total, worst := countDiffPixels(want, img)
	ratio := float64(diff) / float64(total)
	if diff == 0 {
		return
	}

	actualPath := filepath.Join(goldenDir, name+".actual.png")
	_ = writePNG(actualPath, img)

	if ratio > goldenMaxDiffRatio {
		t.Errorf("%s: %d/%d pixels differ (%.3f%%, allowed %.3f%%, worst channel delta %d)\n"+
			"wrote %s for inspection; run `go test -run Golden -update ./...` to accept the new output",
			name, diff, total, ratio*100, goldenMaxDiffRatio*100, worst, actualPath)
		return
	}
	t.Logf("%s: %d/%d pixels differ (%.3f%%), within tolerance", name, diff, total, ratio*100)
}

// countDiffPixels counts pixels whose channels differ by more than the
// tolerance, returning the count, the total and the largest channel delta seen.
func countDiffPixels(a, b image.Image) (diff, total int, worst int) {
	bounds := a.Bounds()
	total = bounds.Dx() * bounds.Dy()
	for y := bounds.Min.Y; y < bounds.Max.Y; y++ {
		for x := bounds.Min.X; x < bounds.Max.X; x++ {
			ar, ag, ab, aa := a.At(x, y).RGBA()
			br, bg, bb, ba := b.At(x, y).RGBA()
			d := channelDelta(ar>>8, br>>8)
			if v := channelDelta(ag>>8, bg>>8); v > d {
				d = v
			}
			if v := channelDelta(ab>>8, bb>>8); v > d {
				d = v
			}
			if v := channelDelta(aa>>8, ba>>8); v > d {
				d = v
			}
			if d > worst {
				worst = d
			}
			if d > goldenChannelTolerance {
				diff++
			}
		}
	}
	return diff, total, worst
}

// channelDelta returns the absolute difference between two 8-bit channels.
func channelDelta(a, b uint32) int {
	d := int(a) - int(b)
	if d < 0 {
		return -d
	}
	return d
}

func readPNG(path string) (image.Image, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	img, err := png.Decode(f)
	if err != nil {
		return nil, fmt.Errorf("decode %s: %w", path, err)
	}
	return img, nil
}

func writePNG(path string, img image.Image) error {
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		return err
	}
	return os.WriteFile(path, buf.Bytes(), 0o644)
}

// --- fixtures ---------------------------------------------------------------

// goldenFixtureChart is a clustered column chart with title, legend and axes.
func goldenFixtureChart() *Presentation {
	p := New()
	slide := p.GetActiveSlide()
	chart := slide.CreateChartShape()
	chart.BaseShape.SetOffsetX(300000).SetOffsetY(300000)
	chart.BaseShape.SetWidth(8500000).SetHeight(5000000)
	chart.GetTitle().SetText("Quarterly Revenue")
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

// goldenFixtureTable is a 3x3 table with a filled header row and explicit
// grid borders, so the golden exercises cell fills, borders and text.
func goldenFixtureTable() *Presentation {
	p := New()
	slide := p.GetActiveSlide()
	tbl := slide.CreateTableShape(3, 3)
	tbl.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	tbl.SetWidth(8000000)
	tbl.SetHeight(4000000)

	header := []string{"Region", "Q1", "Q2"}
	rows := [][]string{
		{"North", "120", "180"},
		{"South", "90", "110"},
	}

	gridColor := NewColor("8EA9DB")
	headerFill := NewFill().SetSolid(NewColor("4472C4"))

	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			cell := tbl.GetCell(r, c)
			if cell == nil {
				continue
			}
			b := cell.GetBorders()
			// Border.Width is in points, not EMU.
			b.Top.SetSolidFill(gridColor).SetWidth(1)
			b.Bottom.SetSolidFill(gridColor).SetWidth(1)
			b.Left.SetSolidFill(gridColor).SetWidth(1)
			b.Right.SetSolidFill(gridColor).SetWidth(1)
			if r == 0 {
				cell.SetFill(headerFill)
			}
		}
	}
	for c, text := range header {
		if cell := tbl.GetCell(0, c); cell != nil {
			cell.SetText(text)
		}
	}
	for r, row := range rows {
		for c, text := range row {
			if cell := tbl.GetCell(r+1, c); cell != nil {
				cell.SetText(text)
			}
		}
	}
	return p
}

// goldenFixtureCJKWrap is a narrow text box holding a Chinese sentence with
// line-start-prohibited punctuation, so the golden records kinsoku behaviour.
func goldenFixtureCJKWrap() *Presentation {
	p := New()
	slide := p.GetActiveSlide()
	shape := slide.CreateRichTextShape()
	shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	shape.BaseShape.SetWidth(2600000)
	shape.BaseShape.SetHeight(3000000)
	run := shape.CreateTextRun("本季度营收同比增长百分之十八，主要来自华东与华南地区的新增客户（含两家制造业龙头）。" +
		"成本端，原材料价格回落，毛利率提升至百分之四十二。")
	run.GetFont().SetSize(18)
	return p
}

// goldenFixtureRotation is a wide bar rotated 90 degrees. It is centred and
// sized so the rotated result still fits inside the slide: content that leaves
// the slide is clipped, which would make the fixture measure the clip rather
// than the rotation.
func goldenFixtureRotation() *Presentation {
	p := New()
	slide := p.GetActiveSlide()
	sh := slide.CreateAutoShape()
	sh.SetAutoShapeType(AutoShapeRectangle)
	sh.BaseShape.SetOffsetX(4096000).SetOffsetY(2979000) // centred
	sh.BaseShape.SetWidth(4000000)
	sh.BaseShape.SetHeight(900000)
	sh.BaseShape.SetRotation(90)
	sh.SetSolidFill(NewColor("4472C4"))
	return p
}

// goldenFixtureShadow is a filled rectangle with a visible drop shadow.
func goldenFixtureShadow() *Presentation {
	p := New()
	slide := p.GetActiveSlide()
	sh := slide.CreateAutoShape()
	sh.SetAutoShapeType(AutoShapeRectangle)
	sh.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	sh.BaseShape.SetWidth(4000000).SetHeight(2500000)
	sh.SetSolidFill(NewColor("ED7D31"))

	shadow := NewShadow()
	shadow.Visible = true
	shadow.Direction = 45
	shadow.Distance = 12
	shadow.BlurRadius = 8
	shadow.Color = NewColor("000000")
	shadow.Alpha = 60
	sh.BaseShape.SetShadow(shadow)
	return p
}

// goldenFixtureUnsupported is a stand-in for a SmartArt diagram, the construct
// most likely to appear in a real deck that this library cannot draw. The golden
// records that the placeholder is visible, so a regression that makes it render
// blank (or invisible) fails loudly instead of quietly producing clean-looking
// previews with missing content.
func goldenFixtureUnsupported() *Presentation {
	p := New()
	u := NewUnsupportedShape("SmartArt diagram")
	u.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	u.BaseShape.SetWidth(4000000).SetHeight(2500000)
	u.SetContentType("http://schemas.openxmlformats.org/drawingml/2006/diagram")
	p.GetActiveSlide().AddShape(u)
	return p
}

// goldenFixtureUnsupportedSmall is the same placeholder at a size the label
// cannot fit into, so the golden pins the shrink-to-fit and clipping behaviour
// rather than only the roomy case. Two boxes of different sizes make it obvious
// from the image alone whether the label is adapting or overrunning.
func goldenFixtureUnsupportedSmall() *Presentation {
	p := New()
	for i, box := range [][4]int64{
		{1500000, 1500000, 1600000, 1200000}, // narrow: the label has to shrink
		{1500000, 3300000, 900000, 700000},   // tiny: the label has to be cropped
	} {
		u := NewUnsupportedShape("SmartArt diagram")
		u.BaseShape.SetOffsetX(box[0]).SetOffsetY(box[1])
		u.BaseShape.SetWidth(box[2]).SetHeight(box[3])
		u.SetContentType("http://schemas.openxmlformats.org/drawingml/2006/diagram")
		if i == 1 {
			u.SetReason("OLE object")
		}
		p.GetActiveSlide().AddShape(u)
	}
	return p
}

// --- cases ------------------------------------------------------------------

func TestGoldenUnsupportedPlaceholder(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)
	checkGolden(t, "unsupported_placeholder", goldenFixtureUnsupported(), goldenOptions(fc))
}

func TestGoldenUnsupportedPlaceholderSmall(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)
	checkGolden(t, "unsupported_placeholder_small", goldenFixtureUnsupportedSmall(), goldenOptions(fc))
}

func TestGoldenChartBar(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)
	checkGolden(t, "chart_bar", goldenFixtureChart(), goldenOptions(fc))
}

func TestGoldenTable(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)
	checkGolden(t, "table_basic", goldenFixtureTable(), goldenOptions(fc))
}

func TestGoldenCJKWrap(t *testing.T) {
	fc := NewFontCache()
	requireCJKFont(t, fc)
	checkGolden(t, "cjk_wrap", goldenFixtureCJKWrap(), goldenOptions(fc))
}

func TestGoldenRotation(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)
	checkGolden(t, "shape_rotation", goldenFixtureRotation(), goldenOptions(fc))
}

func TestGoldenShadow(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)
	checkGolden(t, "shape_shadow", goldenFixtureShadow(), goldenOptions(fc))
}

// TestGoldenEveryChartType records one image per chart type, so a regression in
// any single renderer shows up as a failing golden rather than silently.
func TestGoldenEveryChartType(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	cats := []string{"Q1", "Q2", "Q3", "Q4"}
	vals := []float64{10, 25, 18, 30}

	cases := []struct {
		name  string
		build func() ChartType
	}{
		{"chart_bar3d", func() ChartType {
			c := NewBar3DChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
		{"chart_line", func() ChartType {
			c := NewLineChart()
			c.SetSmooth(true)
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
		{"chart_area", func() ChartType {
			c := NewAreaChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
		{"chart_pie", func() ChartType {
			c := NewPieChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
		{"chart_pie3d", func() ChartType {
			// Deliberately different data from chart_pie: the renderer flattens
			// 3-D charts to 2-D, so reusing the same series would make this
			// golden byte-identical to chart_pie and prove nothing. Distinct
			// values pin that the 3-D path reads its own series.
			c := NewPie3DChart()
			c.AddSeries(NewChartSeriesOrdered("S3D", cats, []float64{22, 8, 31, 14}))
			return c
		}},
		{"chart_doughnut", func() ChartType {
			c := NewDoughnutChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
		{"chart_scatter", func() ChartType {
			c := NewScatterChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
		{"chart_radar", func() ChartType {
			c := NewRadarChart()
			c.AddSeries(NewChartSeriesOrdered("S", cats, vals))
			return c
		}},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			p := New()
			chart := p.GetActiveSlide().CreateChartShape()
			chart.BaseShape.SetOffsetX(300000).SetOffsetY(300000)
			chart.BaseShape.SetWidth(8500000).SetHeight(5000000)
			chart.GetLegend().Visible = true
			chart.GetPlotArea().SetType(tc.build())
			checkGolden(t, tc.name, p, goldenOptions(fc))
		})
	}
}
