package gopresentation

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

// deckWithSlides builds a presentation of n text slides.
func deckWithSlides(n int) *Presentation {
	p := New()
	for i := 1; i < n; i++ {
		p.CreateSlide()
	}
	for i, slide := range p.GetAllSlides() {
		shape := slide.CreateRichTextShape()
		shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
		shape.BaseShape.SetWidth(8000000).SetHeight(2000000)
		shape.CreateTextRun(fmt.Sprintf("Slide %d", i+1))
	}
	return p
}

// BenchmarkSaveSlidesAsImagesFontCache quantifies what sharing one FontCache
// across a call is worth.
//
// "shared" is the shipped behaviour: the call builds one cache and every slide
// reuses it. "per-slide" reproduces the earlier behaviour, rendering each slide
// through options that carry no cache, so every slide scans the font
// directories and keeps its own copy of every parsed font.
//
// Measured on an AMD Ryzen 7 8745HS (Windows, 8 slides, font scan cold in every
// iteration, so the scan dominates the figures):
//
//	shared     163 ms/op   528 MB/op    12530 allocs/op
//	per-slide 1350 ms/op  4171 MB/op    96378 allocs/op
//
// The saving is per-slide work that used to be repeated: one directory scan and
// one full set of parsed fonts instead of one per slide.
func BenchmarkSaveSlidesAsImagesFontCache(b *testing.B) {
	const slides = 8

	p := deckWithSlides(slides)
	dir := b.TempDir()

	run := func(b *testing.B, shared bool) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			opts := DefaultRenderOptions()
			opts.Width = 200
			if shared {
				if err := p.SaveSlidesAsImages(filepath.Join(dir, "shared_%d.png"), opts); err != nil {
					b.Fatal(err)
				}
				continue
			}
			for j := range p.GetAllSlides() {
				name := filepath.Join(dir, fmt.Sprintf("solo_%d_%d.png", i, j+1))
				if err := p.SaveSlideAsImage(j, name, opts); err != nil {
					b.Fatal(err)
				}
			}
		}
	}

	b.Run("shared", func(b *testing.B) { run(b, true) })
	b.Run("per-slide", func(b *testing.B) { run(b, false) })
}

// Contracts of the public render pipeline: what RenderOptions the library is
// allowed to touch, what it does with a degenerate slide size, and the
// equivalence between rendering a whole deck and rendering slide by slide.
//
// These are the properties a caller relies on when they share one RenderOptions
// across several renders, or drive a preview pool from a single configuration.

// TestSlideToImageDoesNotMutateOptions guards the ownership rule for
// RenderOptions: the struct belongs to the caller, so filling in defaults must
// happen on a copy. Mutating it would be an invisible side effect and would race
// when two renders are handed the same options.
func TestSlideToImageDoesNotMutateOptions(t *testing.T) {
	p := textSlideWithFont(bogusFontName, "Hello world")

	opts := DefaultRenderOptions()
	opts.Width = 0 // deliberately unset, so the render has a default to fill in
	opts.DPI = 0
	opts.JPEGQuality = 0

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}
	if img.Bounds().Dx() != 960 {
		t.Errorf("rendered width = %d, want the documented default 960", img.Bounds().Dx())
	}

	if opts.Width != 0 {
		t.Errorf("Width was written back into the caller's options: %d, want 0", opts.Width)
	}
	if opts.DPI != 0 {
		t.Errorf("DPI was written back into the caller's options: %v, want 0", opts.DPI)
	}
	if opts.JPEGQuality != 0 {
		t.Errorf("JPEGQuality was written back into the caller's options: %d, want 0", opts.JPEGQuality)
	}
	if opts.FontCache != nil {
		t.Error("FontCache was written back into the caller's options")
	}
}

// TestSlideToImageRejectsInvalidLayout covers a slide size the library cannot
// turn into pixels. All three are reachable: the reader takes the width and
// height straight from the package, and SetLayout accepts any struct, including
// nil. Each must come back as an ordinary error, not as a recovered panic from
// deep inside the rasterizer, which reports the symptom rather than the cause.
func TestSlideToImageRejectsInvalidLayout(t *testing.T) {
	layouts := map[string]*DocumentLayout{
		"zero width":  {CX: 0, CY: 6858000},
		"zero height": {CX: 9144000, CY: 0},
		"negative":    {CX: -9144000, CY: 6858000},
		"nil":         nil,
	}
	for name, layout := range layouts {
		t.Run(name, func(t *testing.T) {
			p := New()
			p.SetLayout(layout)

			_, err := p.SlideToImage(0, DefaultRenderOptions())
			if err == nil {
				t.Fatal("a degenerate slide size rendered without an error")
			}
			if v, isPanic := ErrIsPanic(err); isPanic {
				t.Errorf("degenerate slide size surfaced as a recovered panic (%v) instead of an error: %v", v, err)
			}
		})
	}
}

// TestSaveSlidesAsImagesMatchesPerSlideRendering checks the sharing that
// SaveSlidesAsImages does internally is not observable in the output. It reuses
// one FontCache across every slide; a cache that leaked per-slide state would
// make a deck render differently from the same slides rendered one at a time.
func TestSaveSlidesAsImagesMatchesPerSlideRendering(t *testing.T) {
	const slides = 4

	build := func() *Presentation {
		p := New()
		for i := 1; i < slides; i++ {
			p.CreateSlide()
		}
		for i, slide := range p.GetAllSlides() {
			shape := slide.CreateRichTextShape()
			shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
			shape.BaseShape.SetWidth(8000000).SetHeight(2000000)
			shape.CreateTextRun(fmt.Sprintf("Slide %d", i+1))
		}
		return p
	}

	dir := t.TempDir()
	sharedPattern := filepath.Join(dir, "shared_%d.png")
	soloPattern := filepath.Join(dir, "solo_%d.png")

	p := build()
	all := DefaultRenderOptions()
	all.Width = 320
	if err := p.SaveSlidesAsImages(sharedPattern, all); err != nil {
		t.Fatalf("SaveSlidesAsImages: %v", err)
	}

	solo := DefaultRenderOptions()
	solo.Width = 320
	for i := range p.GetAllSlides() {
		if err := p.SaveSlideAsImage(i, fmt.Sprintf(soloPattern, i+1), solo); err != nil {
			t.Fatalf("SaveSlideAsImage(%d): %v", i, err)
		}
	}

	for i := 1; i <= slides; i++ {
		got, err := os.ReadFile(fmt.Sprintf(sharedPattern, i))
		if err != nil {
			t.Fatalf("read shared slide %d: %v", i, err)
		}
		want, err := os.ReadFile(fmt.Sprintf(soloPattern, i))
		if err != nil {
			t.Fatalf("read solo slide %d: %v", i, err)
		}
		if !bytes.Equal(got, want) {
			t.Errorf("slide %d differs between a shared font cache and per-slide caches: %d vs %d bytes",
				i, len(got), len(want))
		}
	}
}
