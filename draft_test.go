package gopresentation

import (
	"image"
	"image/color"
	"testing"
)

// This file covers RenderOptions.Draft, the low-fidelity mode intended for batch
// previews. The tests check the two things that matter: draft must actually skip
// the expensive effects, and it must not quietly change the image in ways that
// make a preview misleading.

// shadowedPair returns two presentations with identical geometry, one with a
// drop shadow and one without. Comparing them isolates the cost of the shadow.
func shadowedPair() (with, without *Presentation) {
	build := func(addShadow bool) *Presentation {
		p := New()
		sh := p.GetActiveSlide().CreateAutoShape()
		sh.SetAutoShapeType(AutoShapeRectangle)
		sh.BaseShape.SetOffsetX(2000000).SetOffsetY(1500000)
		sh.BaseShape.SetWidth(4000000).SetHeight(2500000)
		sh.SetSolidFill(NewColor("ED7D31"))
		if addShadow {
			sd := NewShadow()
			sd.Visible = true
			sd.Direction = 45
			sd.Distance = 14
			sd.BlurRadius = 10
			sd.Color = NewColor("000000")
			sd.Alpha = 70
			sh.BaseShape.SetShadow(sd)
		}
		return p
	}
	return build(true), build(false)
}

// TestDraftSkipsShadows is the core draft-mode test. With shadows disabled, the
// shadowed and unshadowed slides must render identically; in full mode they must
// differ, which proves the option is actually wired through and that the
// equality below is not vacuous.
func TestDraftSkipsShadows(t *testing.T) {
	fc := NewFontCache()
	with, without := shadowedPair()

	fullOpts := goldenOptions(fc)
	withShadow, err := with.SlideToImage(0, fullOpts)
	if err != nil {
		t.Fatalf("render full with shadow: %v", err)
	}
	plain, err := without.SlideToImage(0, fullOpts)
	if err != nil {
		t.Fatalf("render full without shadow: %v", err)
	}
	if diff, _, _ := countDiffPixels(withShadow, plain); diff == 0 {
		t.Fatal("full render draws no shadow; the fixture cannot distinguish draft from full")
	}

	draftOpts := goldenOptions(fc)
	draftOpts.Draft = true
	draftShadow, err := with.SlideToImage(0, draftOpts)
	if err != nil {
		t.Fatalf("render draft with shadow: %v", err)
	}
	draftPlain, err := without.SlideToImage(0, draftOpts)
	if err != nil {
		t.Fatalf("render draft without shadow: %v", err)
	}
	if diff, total, worst := countDiffPixels(draftShadow, draftPlain); diff != 0 {
		t.Errorf("draft render still drew a shadow: %d/%d pixels differ (worst channel %d)",
			diff, total, worst)
	}
}

// TestDraftSkipsAntiAliasing checks the second saving: draft snaps partial
// coverage, so the palette of a rendered line contains fewer intermediate
// colours than in full mode. It counts distinct greys along a diagonal border
// rather than asserting exact pixel values, which would be brittle.
func TestDraftSkipsAntiAliasing(t *testing.T) {
	fc := NewFontCache()

	build := func(draft bool) *image.RGBA {
		p := New()
		sh := p.GetActiveSlide().CreateAutoShape()
		sh.SetAutoShapeType(AutoShapeEllipse)
		sh.BaseShape.SetOffsetX(2500000).SetOffsetY(1800000)
		sh.BaseShape.SetWidth(3800000).SetHeight(2600000)
		sh.SetSolidFill(NewColor("255000"))
		b := NewBorder()
		b.SetSolidFill(NewColor("000000"))
		b.SetWidth(2)
		sh.BaseShape.SetBorder(b)

		opts := goldenOptions(fc)
		opts.Draft = draft
		img, err := p.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render (draft=%v): %v", draft, err)
		}
		return img.(*image.RGBA)
	}

	fullTones := distinctTones(build(false))
	draftTones := distinctTones(build(true))
	if draftTones >= fullTones {
		t.Errorf("draft produced %d distinct tones, full produced %d; draft should soften fewer edges",
			draftTones, fullTones)
	}
}

// distinctTones counts how many distinct non-background grey levels appear in an
// image. Anti-aliasing creates many; a hard-edged render creates few.
func distinctTones(img *image.RGBA) int {
	seen := make(map[uint32]struct{})
	b := img.Bounds()
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			r, g, bl, a := img.At(x, y).RGBA()
			if a>>8 < 8 {
				continue
			}
			r8, g8, b8 := r>>8, g>>8, bl>>8
			if r8 > 250 && g8 > 250 && b8 > 250 {
				continue // white background
			}
			// Only grey-ish pixels matter for measuring edge softening.
			if channelDelta(r8, g8) > 12 || channelDelta(g8, b8) > 12 {
				continue
			}
			seen[r8<<16|g8<<8|b8] = struct{}{}
		}
	}
	return len(seen)
}

// TestDraftStillRendersContent guards the failure that would make draft useless:
// it must still draw the slide. A preview that is fast because it is empty is
// worse than no preview.
func TestDraftStillRendersContent(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	pres := chartSlidePresentation()
	opts := goldenOptions(fc)
	opts.Draft = true

	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render draft: %v", err)
	}

	// Title text box.
	title := emuRect(t, pres, 500000, 300000, 8000000, 600000, opts.Width)
	if n := countInkIn(img, title); n == 0 {
		t.Errorf("draft render drew no text in the title area %v", title)
	}
	// Chart body, which must not collapse to a blank frame.
	chartBox := emuRect(t, pres, 500000, 1000000, 8000000, 4500000, opts.Width)
	if n := countInkIn(img, chartBox); n == 0 {
		t.Errorf("draft render drew no chart in %v", chartBox)
	}
}

// TestDraftKeepsImageSize checks the option does not silently alter the output
// geometry, so callers can switch a batch between draft and full without
// changing anything else.
func TestDraftKeepsImageSize(t *testing.T) {
	fc := NewFontCache()
	pres := chartSlidePresentation()

	fullOpts := goldenOptions(fc)
	full, err := pres.SlideToImage(0, fullOpts)
	if err != nil {
		t.Fatalf("render full: %v", err)
	}
	draftOpts := goldenOptions(fc)
	draftOpts.Draft = true
	draft, err := pres.SlideToImage(0, draftOpts)
	if err != nil {
		t.Fatalf("render draft: %v", err)
	}
	if full.Bounds() != draft.Bounds() {
		t.Errorf("draft bounds = %v, full bounds = %v; they must match", draft.Bounds(), full.Bounds())
	}
}

// TestDraftPropagatesToSubRenderer checks the flag survives the temporary
// renderers used for rotated and grouped content. Those copies are exactly where
// a configuration field gets dropped, so each must be checked explicitly.
func TestDraftPropagatesToSubRenderer(t *testing.T) {
	r := &renderer{draft: true}
	if sub := r.subRenderer(image.NewRGBA(image.Rect(0, 0, 4, 4))); !sub.draft {
		t.Error("subRenderer dropped the draft flag; rotated/grouped content would render in full quality")
	}
}

// TestScaleImageNearestSamplesWithoutInterpolation pins the sampler's behaviour:
// each destination pixel takes exactly one source pixel, so a checkerboard stays
// a checkerboard instead of averaging towards grey.
func TestScaleImageNearestSamplesWithoutInterpolation(t *testing.T) {
	src := image.NewRGBA(image.Rect(0, 0, 2, 2))
	black := color.RGBA{A: 255}
	white := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	src.Set(0, 0, black)
	src.Set(1, 0, white)
	src.Set(0, 1, white)
	src.Set(1, 1, black)

	up := scaleImageNearest(src, 4, 4)
	// Sampling dx*srcW/dstW maps destination columns 0,1 -> 0 and 2,3 -> 1.
	if got := up.RGBAAt(0, 0); got.R > 8 || got.G > 8 || got.B > 8 {
		t.Errorf("top-left block = %v, want black (no interpolation)", got)
	}
	if got := up.RGBAAt(3, 0); got.R < 247 || got.G < 247 || got.B < 247 {
		t.Errorf("top-right block = %v, want white (no interpolation)", got)
	}
	if got := up.RGBAAt(0, 3); got.R < 247 {
		t.Errorf("bottom-left block = %v, want white", got)
	}

	down := scaleImageNearest(src, 1, 1)
	if got := down.RGBAAt(0, 0); got.R > 8 {
		t.Errorf("downsampled pixel = %v, want the top-left sample (black)", got)
	}

	// Degenerate sizes must not panic or allocate a nonsense image.
	if got := scaleImageNearest(src, 0, 10); got.Bounds().Empty() && got.Bounds().Dx() != 0 {
		t.Errorf("zero-width scale returned %v", got.Bounds())
	}
	if got := scaleImageNearest(image.NewRGBA(image.Rectangle{}), 8, 8); got.Bounds().Dx() != 8 {
		t.Errorf("scaling an empty source returned %v, want an 8x8 canvas", got.Bounds())
	}
}

// TestScaleImageNearestHandlesNonRGBASources covers the normalisation path.
// JPEG decodes to *image.YCbCr and alpha PNGs to *image.NRGBA, so sampling only
// works correctly if those are converted first — and converting inside the
// per-pixel loop is what made an earlier version of this function slower than
// the bilinear filter it replaces.
func TestScaleImageNearestHandlesNonRGBASources(t *testing.T) {
	const w, h = 8, 8

	nrgba := image.NewNRGBA(image.Rect(0, 0, w, h))
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			if x < w/2 {
				nrgba.SetNRGBA(x, y, color.NRGBA{R: 255, A: 255})
			} else {
				nrgba.SetNRGBA(x, y, color.NRGBA{B: 255, A: 255})
			}
		}
	}

	for _, src := range []image.Image{
		nrgba,
		image.NewYCbCr(image.Rect(0, 0, w, h), image.YCbCrSubsampleRatio444),
		image.NewGray(image.Rect(0, 0, w, h)),
	} {
		got := scaleImageNearest(src, 40, 40)
		if got.Bounds().Dx() != 40 || got.Bounds().Dy() != 40 {
			t.Errorf("%T: size = %v, want 40x40", src, got.Bounds())
		}
		// Every destination pixel must have been written from a source pixel, so
		// the alpha channel must be fully opaque for these sources.
		if _, _, _, a := got.At(20, 20).RGBA(); a>>8 != 255 {
			t.Errorf("%T: centre pixel alpha = %d, want 255 (the normalisation lost data)", src, a>>8)
		}
	}
}

// BenchmarkPreviewQuality measures the two modes across the workload types a
// preview pipeline meets, because draft mode's benefit is very uneven: it does
// nothing for glyph-heavy slides (text rasterisation cannot be switched off) and
// a lot for shadow-heavy or image-heavy ones. The low-resolution variant shows
// the combination that actually matters in practice.
func BenchmarkPreviewQuality(b *testing.B) {
	fc := NewFontCache()

	imageHeavy := func() *Presentation {
		p := New()
		d := NewDrawingShape()
		d.SetImageData(twoTonePNGBytes(), "image/png")
		d.BaseShape.SetOffsetX(0).SetOffsetY(0)
		d.BaseShape.SetWidth(9144000).SetHeight(6858000)
		p.GetActiveSlide().AddShape(d)
		return p
	}

	withShadow, _ := shadowedPair()

	benches := []struct {
		name string
		pres *Presentation
		opts func() *RenderOptions
	}{
		{
			// Text and charts dominate: draft has little to remove here.
			name: "textchart_full",
			pres: chartSlidePresentation(),
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 960; return o },
		},
		{
			name: "textchart_draft",
			pres: chartSlidePresentation(),
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 960; o.Draft = true; return o },
		},
		{
			name: "textchart_draft_lowres",
			pres: chartSlidePresentation(),
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 480; o.Draft = true; return o },
		},
		{
			name: "shadow_full",
			pres: withShadow,
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 960; return o },
		},
		{
			name: "shadow_draft",
			pres: withShadow,
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 960; o.Draft = true; return o },
		},
		{
			name: "image_full",
			pres: imageHeavy(),
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 960; return o },
		},
		{
			name: "image_draft",
			pres: imageHeavy(),
			opts: func() *RenderOptions { o := goldenOptions(fc); o.Width = 960; o.Draft = true; return o },
		},
	}

	for _, bc := range benches {
		b.Run(bc.name, func(b *testing.B) {
			opts := bc.opts()
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if _, err := bc.pres.SlideToImage(0, opts); err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}
