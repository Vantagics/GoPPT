package gopresentation

import (
	"image"
	"strings"
	"testing"
)

// Chart text fonts, end to end.
//
// A chart's strings carry the same font model as slide text — a Latin face and
// an East Asian face — but all three layers of the chart path dropped it: the
// writer emitted only the size and bold flag, the reader never looked at
// <a:latin>/<a:ea>, and the rasteriser resolved chart text through the Latin
// path with no glyph-coverage check. The symptom was Chinese chart labels drawn
// as .notdef boxes while every layer reported success: the declared name
// resolves, so FontDiagnostics has nothing to complain about, and the chart
// goldens use ASCII strings, so they compared tofu against tofu quite happily.
// Each layer therefore gets an assertion here.

// fixtureChartWithFonts builds a bar chart whose chart title and both axis
// titles declare fonts, so one write exercises every <a:rPr> the chart writer
// emits.
func fixtureChartWithFonts() *Presentation {
	p := New()
	chart := p.GetActiveSlide().CreateChartShape()
	chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
	chart.BaseShape.SetWidth(7000000).SetHeight(4500000)

	title := chart.GetTitle()
	title.SetText("收入趋势").SetVisible(true)
	title.Font.SetName("Microsoft YaHei")
	title.Font.SetSize(18).SetBold(true)
	title.Font.NameEA = "SimSun"

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20}))
	chart.GetPlotArea().SetType(bar)
	chart.GetLegend().Visible = false

	for _, ax := range []*ChartAxis{chart.GetPlotArea().GetAxisX(), chart.GetPlotArea().GetAxisY()} {
		ax.SetTitle("季度")
		ax.Font.SetName("Microsoft YaHei")
		ax.Font.NameEA = "SimSun"
	}
	return p
}

// TestChartRunPropsCarryTypefaces covers the writer half. The chart writer knew
// the run's font — the model carries it — and serialised a size and a bold flag
// only, so the author's face was gone the next time the part was opened.
func TestChartRunPropsCarryTypefaces(t *testing.T) {
	part, ok := zipParts(t, writeToBytes(t, fixtureChartWithFonts()))["ppt/charts/chart1.xml"]
	if !ok {
		t.Fatal("chart part was not written")
	}
	text := string(part)

	if !strings.Contains(text, `<a:latin typeface="Microsoft YaHei"/>`) {
		t.Errorf("the chart title's Latin typeface is not in the part; got:\n%s", text)
	}
	if !strings.Contains(text,
		`<a:rPr lang="en-US" sz="1800" b="1"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:rPr>`) {
		t.Errorf("the chart title run does not carry its size, bold flag and both typefaces together; got:\n%s", text)
	}
	// The chart title plus the two axis titles.
	if got, want := strings.Count(text, `<a:ea typeface="SimSun"/>`), 3; got != want {
		t.Errorf("<a:ea typeface=\"SimSun\"/> appears %d times, want %d (chart title and both axes); got:\n%s",
			got, want, text)
	}
}

// chartTextFontXML is a minimal bar chart part whose title runs declare a Latin
// and an East Asian typeface, as PowerPoint writes them.
const chartTextFontXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1400" b="0"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:rPr>
              <a:t>收入趋势</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="SimSun"/></a:rPr><a:t>季度</a:t></a:r></a:p></c:rich></c:tx></c:title>
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
</c:chartSpace>`

// TestChartTextFontIsReadFromTypeface covers the reader half. Both typefaces
// live on the run's <a:rPr>, and reading the text without them left every chart
// string on the model default.
func TestChartTextFontIsReadFromTypeface(t *testing.T) {
	chart := parseChartXML([]byte(chartTextFontXML), nil)
	if chart == nil {
		t.Fatal("parseChartXML returned nil for a valid chart part")
	}

	title := chart.GetTitle()
	if title.Font == nil {
		t.Fatal("chart title has no font after parsing")
	}
	if got, want := title.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("chart title font = %q, want %q", got, want)
	}
	if got, want := title.Font.NameEA, "SimSun"; got != want {
		t.Errorf("chart title East Asian font = %q, want %q", got, want)
	}

	axX := chart.GetPlotArea().GetAxisX()
	if axX == nil {
		t.Fatal("category axis was not read")
	}
	if got, want := axX.Title, "季度"; got != want {
		t.Errorf("category axis title = %q, want %q", got, want)
	}
	if got, want := axX.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("category axis font = %q, want %q", got, want)
	}
	if got, want := axX.Font.NameEA, "SimSun"; got != want {
		t.Errorf("category axis East Asian font = %q, want %q", got, want)
	}
}

// TestChartTextFontSurvivesRoundTrip is the integration case: a chart written
// with a font must read back with it. Either half alone leaves this failing, and
// nothing else in the suite notices, because a chart that has lost its font
// still renders — just not with the right face.
func TestChartTextFontSurvivesRoundTrip(t *testing.T) {
	rt := firstChartShape(t, roundTrip(t, fixtureChartWithFonts()))

	title := rt.GetTitle()
	if title.Font == nil {
		t.Fatal("chart title has no font after the round trip")
	}
	if got, want := title.Text, "收入趋势"; got != want {
		t.Errorf("chart title text = %q, want %q", got, want)
	}
	if got, want := title.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("chart title font = %q, want %q", got, want)
	}
	if got, want := title.Font.NameEA, "SimSun"; got != want {
		t.Errorf("chart title East Asian font = %q, want %q", got, want)
	}

	axX := rt.GetPlotArea().GetAxisX()
	if axX == nil {
		t.Fatal("category axis was lost by the round trip")
	}
	if got, want := axX.Title, "季度"; got != want {
		t.Errorf("category axis title = %q, want %q", got, want)
	}
	if got, want := axX.Font.Name, "Microsoft YaHei"; got != want {
		t.Errorf("category axis font = %q, want %q", got, want)
	}
	if got, want := axX.Font.NameEA, "SimSun"; got != want {
		t.Errorf("category axis East Asian font = %q, want %q", got, want)
	}
}

// TestChartTitleCJKRendersWithoutTofu covers the rasteriser half.
//
// Two charts are built that are identical in every respect except the three
// characters of their title. Titles that are drawn with a face lacking those
// glyphs come out as three identical .notdef boxes each, so the two slides are
// pixel-identical; titles drawn with a face that has the glyphs differ. The
// test therefore needs no reference image and cannot be satisfied by a
// uniformly wrong render.
func TestChartTitleCJKRendersWithoutTofu(t *testing.T) {
	const textA, textB = "本季度", "一二三"
	sample := textA + textB

	fc := NewFontCache()
	latin := latinOnlyFont(fc, sample)
	if latin == "" {
		t.Skip("no installed Latin-only font; the tofu comparison needs one")
	}
	if coverageFont(fc, sample) == "" {
		t.Skip("no installed font covers the sample; a tofu comparison would be meaningless")
	}

	build := func(title string) *Presentation {
		p := New()
		chart := p.GetActiveSlide().CreateChartShape()
		chart.BaseShape.SetOffsetX(300000).SetOffsetY(300000)
		chart.BaseShape.SetWidth(8500000).SetHeight(5000000)
		chart.GetTitle().SetText(title).SetVisible(true)
		// A face that is installed, resolves, and cannot draw a single one of
		// the characters: exactly the case that used to render as tofu while
		// reporting a perfect font match.
		chart.GetTitle().Font.SetName(latin)
		chart.GetLegend().Visible = false

		bar := NewBarChart()
		bar.AddSeries(NewChartSeriesOrdered("S", []string{"Q1", "Q2"}, []float64{10, 20}))
		chart.GetPlotArea().SetType(bar)
		return p
	}

	render := func(p *Presentation) image.Image {
		t.Helper()
		opts := DefaultRenderOptions()
		opts.Width = 1600
		opts.FontCache = fc
		img, err := p.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		return img
	}

	presA, presB := build(textA), build(textB)
	imgA, imgB := render(presA), render(presB)

	// The title strip. Everything below it is identical between the two
	// slides, so including some of it costs nothing and keeps the band from
	// depending on font metrics.
	strip := emuRect(t, presA, 300000, 300000, 8500000, 5000000, 1600)
	strip.Max.Y = strip.Min.Y + 200

	inStrip := countDiffInRect(imgA, imgB, strip)
	if inStrip == 0 {
		t.Errorf("two different three-character titles (%q vs %q, both declared as %q) rendered identically "+
			"inside the title strip: the chart title is drawn as missing-glyph boxes", textA, textB, latin)
	}

	if whole := countDiffInRect(imgA, imgB, imgA.Bounds()); whole != inStrip {
		t.Errorf("%d of %d differing pixels lie outside the title strip; the comparison is measuring "+
			"something other than the title", whole-inStrip, whole)
	}
}

// countDiffInRect counts pixels inside rect whose channels differ by more than
// the golden tolerance.
func countDiffInRect(a, b image.Image, rect image.Rectangle) int {
	n := 0
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
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
			if d > goldenChannelTolerance {
				n++
			}
		}
	}
	return n
}

// latinOnlyFont returns an installed font that cannot draw any character of
// sample, or "" when the machine has none.
//
// The precondition asks the font cache directly and never touches the chart
// font-selection code, so reverting that code makes the test FAIL rather than
// SKIP and hide the regression.
func latinOnlyFont(fc *FontCache, sample string) string {
	candidates := []string{"Arial", "Calibri", "Segoe UI", "Tahoma", "Verdana", "DejaVu Sans", "Helvetica"}
	for _, name := range candidates {
		if !fc.HasFont(name, false, false) {
			continue
		}
		drawable := false
		for _, r := range sample {
			if fc.CoversRune(name, false, false, r) {
				drawable = true
				break
			}
		}
		if !drawable {
			return name
		}
	}
	return ""
}

// coverageFont returns the name of an installed font in the built-in East Asian
// fallback chain that can draw every character of sample, or "" when none can.
// Like latinOnlyFont it works strictly off the font cache.
func coverageFont(fc *FontCache, sample string) string {
	for _, name := range defaultCJKFallbackChain {
		if !fc.HasFont(name, false, false) {
			continue
		}
		all := true
		for _, r := range sample {
			if !fc.CoversRune(name, false, false, r) {
				all = false
				break
			}
		}
		if all {
			return name
		}
	}
	return ""
}
