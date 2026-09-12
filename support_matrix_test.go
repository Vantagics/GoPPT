package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"testing"
)

// This file substantiates rows of the shape support matrix documented in
// README.md. It covers the constructs whose support is easiest to lose without
// noticing: image cropping through a:srcRect, and charts nested inside a group.
// The other rows are covered by the chart round-trip tests, the golden tests and
// the unsupported-shape tests.

// twoToneImage returns an image that is red on its left half and blue on its
// right half, so a crop applied to it is visible as the loss of one colour.
func twoToneImage() *image.RGBA {
	const w, h = 64, 64
	src := image.NewRGBA(image.Rect(0, 0, w, h))
	red := color.RGBA{R: 255, A: 255}
	blue := color.RGBA{B: 255, A: 255}
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			c := red
			if x >= w/2 {
				c = blue
			}
			src.Set(x, y, c)
		}
	}
	return src
}

// twoTonePNGBytes encodes the fixture. The image is static, so an encode failure
// would be a programming error rather than a test condition; use twoTonePNG when
// a *testing.T is available so the failure is reported properly.
func twoTonePNGBytes() []byte {
	var buf bytes.Buffer
	if err := png.Encode(&buf, twoToneImage()); err != nil {
		panic("encode two-tone fixture: " + err.Error())
	}
	return buf.Bytes()
}

// twoTonePNG encodes twoToneImage as PNG.
func twoTonePNG(t *testing.T) []byte {
	t.Helper()
	return twoTonePNGBytes()
}

// TestSrcRectCropIsApplied checks the a:srcRect crop is honoured end to end. The
// crop values are in 1/1000 of a percent, so 50000 removes exactly half.
func TestSrcRectCropIsApplied(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)
	img := twoTonePNG(t)

	build := func(cropLeft int) *Presentation {
		p := New()
		d := NewDrawingShape()
		d.SetImageData(img, "image/png")
		d.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
		d.BaseShape.SetWidth(4000000).SetHeight(2500000)
		d.cropLeft = cropLeft
		p.GetActiveSlide().AddShape(d)
		return p
	}

	red := color.RGBA{R: 255, A: 255}
	blue := color.RGBA{B: 255, A: 255}

	uncropped := build(0)
	// Geometry comes from a real presentation: New() is 4:3, not 16:9.
	box := emuRect(t, uncropped, 2000000, 1500000, 4000000, 2500000, opts.Width)
	if box.Empty() {
		t.Fatal("could not derive the picture rectangle")
	}

	full, err := uncropped.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render uncropped: %v", err)
	}
	fullRed := countNearColorIn(full, box, red, 40)
	fullBlue := countNearColorIn(full, box, blue, 40)
	if fullRed == 0 || fullBlue == 0 {
		t.Fatalf("uncropped fixture is not two-tone: red=%d blue=%d in %v", fullRed, fullBlue, box)
	}

	// Crop away the entire red half, then stretch what remains across the frame.
	cropped, err := build(50000).SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render cropped: %v", err)
	}
	cropRed := countNearColorIn(cropped, box, red, 40)
	cropBlue := countNearColorIn(cropped, box, blue, 40)
	if cropBlue == 0 {
		t.Errorf("cropped picture lost the surviving (blue) half: blue=%d in %v", cropBlue, box)
	}
	if cropRed > fullRed/20 {
		t.Errorf("srcRect crop not applied: red=%d of the original %d survived in %v",
			cropRed, fullRed, box)
	}
}

// TestChartInsideGroupRenders covers the "chart nested inside a group" row: a
// chart in a group must reach the chart draw path, not be skipped because the
// group renderer only knows about a few shape kinds.
func TestChartInsideGroupRenders(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	p := New()
	chart := NewChartShape()
	chart.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	chart.BaseShape.SetWidth(4000000).SetHeight(2500000)
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S", []string{"A", "B", "C"}, []float64{10, 20, 30}))
	chart.GetPlotArea().SetType(bar)

	g := NewGroupShape()
	g.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	g.BaseShape.SetWidth(4000000).SetHeight(2500000)
	g.AddShape(chart)
	p.GetActiveSlide().AddShape(g)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	box := emuRect(t, p, 2000000, 1500000, 4000000, 2500000, opts.Width)
	if n := countInkIn(img, box); n == 0 {
		t.Errorf("chart inside a group drew nothing in %v; the group path dropped it", box)
	}
}

// TestPictureThatCannotBeDecodedIsVisible covers the row for an undecodable
// picture: it must leave a visible outline, not a blank hole.
func TestPictureThatCannotBeDecodedIsVisible(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)

	p := New()
	d := NewDrawingShape()
	// Not a valid image in any format the decoder knows.
	d.SetImageData([]byte("not an image at all"), "image/png")
	d.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
	d.BaseShape.SetWidth(4000000).SetHeight(2500000)
	p.GetActiveSlide().AddShape(d)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	box := emuRect(t, p, 2000000, 1500000, 4000000, 2500000, opts.Width)
	if n := countInkIn(img, box); n == 0 {
		t.Errorf("an undecodable picture rendered nothing in %v; it should leave a visible outline", box)
	}
}
