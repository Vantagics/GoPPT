package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"testing"
)

// This file pins the picture resampling semantics that r38 pinned against
// PowerPoint COM goldens: pictures are scaled with a Mitchell-Netravali
// B=C=1/3 filter (the kernel behind GDI+ HighQualityBicubic), corner-mapped
// onto the destination rectangle, with a:srcRect applied as a float source
// sub-rectangle rather than truncated to whole pixels first.

// quadImage is red / green on the top row and blue / white on the bottom row —
// four corners whose blends identify every tap of the resampler.
func quadImage() *image.RGBA {
	src := image.NewRGBA(image.Rect(0, 0, 2, 2))
	src.SetRGBA(0, 0, color.RGBA{R: 255, A: 255})
	src.SetRGBA(1, 0, color.RGBA{G: 255, A: 255})
	src.SetRGBA(0, 1, color.RGBA{B: 255, A: 255})
	src.SetRGBA(1, 1, color.RGBA{R: 255, G: 255, B: 255, A: 255})
	return src
}

// TestMitchellResampleIdentityAtSameSize: no crop and a 1:1 destination must
// be a pixel-exact copy — the fast path guards against the filter's edge
// taps (Mitchell's k(±1) = 1/18) smearing even an unscaled image.
func TestMitchellResampleIdentityAtSameSize(t *testing.T) {
	src := quadImage()
	dst := scaleImageMitchell(src, 2, 2, nil)
	for y := 0; y < 2; y++ {
		for x := 0; x < 2; x++ {
			want := src.RGBAAt(x, y)
			got := dst.RGBAAt(x, y)
			if got != want {
				t.Fatalf("identity resample changed (%d,%d): want %v got %v", x, y, want, got)
			}
		}
	}
}

// TestMitchellUpscale2xPinnedValues: 2×2 → 4×4 with expected values produced
// by the independent Python reference model that was validated against the
// PowerPoint golden renders (masked full-page rms 2.9 on the fingerprint
// texture of comparison-deck slide 1). Any change to the kernel constants,
// the corner mapping or the weight normalisation moves these numbers.
func TestMitchellUpscale2xPinnedValues(t *testing.T) {
	dst := scaleImageMitchell(quadImage(), 4, 4, nil)
	want := [4][4][3]uint8{
		{{228, 14, 14}, {128, 128, 14}, {27, 241, 14}, {6, 255, 14}},
		{{128, 14, 128}, {128, 128, 128}, {128, 241, 128}, {128, 255, 128}},
		{{27, 14, 241}, {128, 128, 241}, {228, 241, 241}, {249, 255, 241}},
		{{6, 14, 255}, {128, 128, 255}, {249, 241, 255}, {255, 255, 255}},
	}
	for y := 0; y < 4; y++ {
		for x := 0; x < 4; x++ {
			got := dst.RGBAAt(x, y)
			w := want[y][x]
			if got.R != w[0] || got.G != w[1] || got.B != w[2] {
				t.Fatalf("upscale (%d,%d): want rgb%v got rgb(%d,%d,%d)", x, y, w, got.R, got.G, got.B)
			}
		}
	}
}

// TestMitchellFractionalCropSamplesSubRectangle: a 30% top crop on a 4px tall
// image starts the sample at source y=1.2 — a fractional pixel edge. The
// filter sees rows 0..3 weighted continuously from that offset, giving
// 60/107/149; truncating the crop rect to whole pixels first (the pre-r38
// behaviour) resamples rows 1..3 instead and produces 50/100/150.
func TestMitchellFractionalCropSamplesSubRectangle(t *testing.T) {
	src := image.NewRGBA(image.Rect(0, 0, 4, 4))
	for y, v := range []uint8{0, 50, 100, 150} {
		for x := 0; x < 4; x++ {
			src.SetRGBA(x, y, color.RGBA{R: v, G: v, B: v, A: 255})
		}
	}
	crop := &[4]float64{0, 1.2, 4, 4} // top 30% removed: fractional edge
	dst := scaleImageMitchell(src, 4, 3, crop)
	want := []uint8{60, 107, 149}
	for y := 0; y < 3; y++ {
		got := dst.RGBAAt(1, y).R
		if got != want[y] {
			t.Fatalf("cropped row %d: want %d got %d", y, want[y], got)
		}
	}
}

// TestRenderPictureUpscaleIsFiltered wires the resampler into a real render:
// a 2×2 picture stretched onto a large frame must show blended interior
// pixels. Nearest-neighbour sampling (the draft path, or a regression to the
// old corner-truncated bilinear's blockiness) leaves only the four pure
// source colours.
func TestRenderPictureUpscaleIsFiltered(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)
	var buf bytes.Buffer
	if err := png.Encode(&buf, quadImage()); err != nil {
		t.Fatalf("encode fixture: %v", err)
	}

	p := New()
	d := NewDrawingShape()
	d.SetImageData(buf.Bytes(), "image/png")
	d.BaseShape.SetOffsetX(1000000).SetOffsetY(1000000).SetWidth(4000000).SetHeight(4000000)
	p.GetActiveSlide().AddShape(d)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rgba, ok := img.(*image.RGBA)
	if !ok {
		t.Fatalf("render is %T, want *image.RGBA", img)
	}
	box := emuRect(t, p, 1000000, 1000000, 4000000, 4000000, opts.Width)
	if box.Empty() {
		t.Fatal("could not derive the picture rectangle")
	}
	pure := map[color.RGBA]bool{
		{R: 255, A: 255}: true, {G: 255, A: 255}: true,
		{B: 255, A: 255}: true, {R: 255, G: 255, B: 255, A: 255}: true,
	}
	blended := 0
	for y := box.Min.Y + 2; y < box.Max.Y-2; y++ {
		for x := box.Min.X + 2; x < box.Max.X-2; x++ {
			if !pure[rgba.RGBAAt(x, y)] {
				blended++
			}
		}
	}
	// A filtered upscale blends everywhere except a handful of pixels; a
	// nearest-neighbour one is pure colour over 95% of the area.
	if blended < (box.Dx()-4)*(box.Dy()-4)/2 {
		t.Fatalf("interior looks nearest-neighbour: only %d/%d blended pixels", blended, (box.Dx()-4)*(box.Dy()-4))
	}
}

// TestRenderPictureDraftStaysNearest: Draft mode is the quality/speed trade —
// it must keep the cheap nearest-neighbour resampler, so its interior is pure
// source colour almost everywhere (the opposite contract of the filtered
// render above).
func TestRenderPictureDraftStaysNearest(t *testing.T) {
	fc := NewFontCache()
	opts := goldenOptions(fc)
	opts.Draft = true
	var buf bytes.Buffer
	if err := png.Encode(&buf, quadImage()); err != nil {
		t.Fatalf("encode fixture: %v", err)
	}

	p := New()
	d := NewDrawingShape()
	d.SetImageData(buf.Bytes(), "image/png")
	d.BaseShape.SetOffsetX(1000000).SetOffsetY(1000000).SetWidth(4000000).SetHeight(4000000)
	p.GetActiveSlide().AddShape(d)

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rgba, ok := img.(*image.RGBA)
	if !ok {
		t.Fatalf("render is %T, want *image.RGBA", img)
	}
	box := emuRect(t, p, 1000000, 1000000, 4000000, 4000000, opts.Width)
	if box.Empty() {
		t.Fatal("could not derive the picture rectangle")
	}
	pure := map[color.RGBA]bool{
		{R: 255, A: 255}: true, {G: 255, A: 255}: true,
		{B: 255, A: 255}: true, {R: 255, G: 255, B: 255, A: 255}: true,
	}
	total := 0
	blended := 0
	for y := box.Min.Y + 2; y < box.Max.Y-2; y++ {
		for x := box.Min.X + 2; x < box.Max.X-2; x++ {
			total++
			if !pure[rgba.RGBAAt(x, y)] {
				blended++
			}
		}
	}
	if blended > total/10 {
		t.Fatalf("draft render is filtering: %d/%d blended pixels", blended, total)
	}
}
