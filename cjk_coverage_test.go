package gopresentation

import (
	"image"
	"strings"
	"testing"
)

// East Asian font selection.
//
// The bug these tests pin down: a document whose generator copies one
// font-family list into both <a:latin> and <a:ea> declares a Latin face as its
// East Asian font. The name resolves, so font resolution looked satisfied and
// the diagnostics reported a perfect match, while every Chinese character was
// drawn as that font's .notdef box. Nothing in the golden images caught it
// because none of them declared a Latin font for Chinese text.

// cjkSample is the string the tests render. It is several distinct characters,
// so a missing-glyph box per character cannot coincidentally match real glyphs.
const cjkSample = "组织的驱动系统"

// latinFallbackNames returns the font names in the built-in chain that are not
// in the CJK chain, i.e. the ones expected to lack Chinese glyphs. Deriving it
// from the production lists keeps the test honest if either list changes.
func latinFallbackNames() []string {
	cjk := make(map[string]bool, len(defaultCJKFallbackChain))
	for _, n := range defaultCJKFallbackChain {
		cjk[strings.ToLower(n)] = true
	}
	var out []string
	for _, n := range defaultFontFallbackChain {
		if !cjk[strings.ToLower(n)] {
			out = append(out, n)
		}
	}
	return out
}

// installedLatinOnlyFont returns an installed font that has no glyph for 组, or
// "" when the machine has none.
func installedLatinOnlyFont(fc *FontCache) string {
	for _, name := range latinFallbackNames() {
		if fc.HasFont(name, false, false) && !fc.CoversRune(name, false, false, '组') {
			return name
		}
	}
	return ""
}

// cjkCoveringFont returns the first font in the CJK chain that has a glyph for
// every East Asian character of sample, or "" when none does.
//
// The loop deliberately re-derives the answer from glyph coverage instead of
// calling the renderer, so the expectation it produces is independent of the
// code under test.
func cjkCoveringFont(fc *FontCache, sample string) string {
	for _, name := range defaultCJKFallbackChain {
		tested := false
		covers := true
		for _, r := range sample {
			if !isCJK(r) {
				continue
			}
			tested = true
			if !fc.CoversRune(name, false, false, r) {
				covers = false
				break
			}
		}
		if tested && covers {
			return name
		}
	}
	return ""
}

// cjkSlide builds a one-slide presentation holding cjkSample, with the Latin
// and East Asian font names set independently so a caller can declare a Latin
// face as the East Asian one.
func cjkSlide(latin, eastAsian string) *Presentation {
	p := New()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	shape.BaseShape.SetWidth(9000000).SetHeight(2000000)
	run := shape.CreateTextRun(cjkSample)
	run.GetFont().SetName(latin).SetSize(28)
	run.GetFont().NameEA = eastAsian
	return p
}

// TestCoversRuneAnswersGlyphCoverage tests the primitive the fallback decision
// rests on. A font name that resolves says nothing about whether the font has
// the glyph, which is the distinction tofu hides behind.
func TestCoversRuneAnswersGlyphCoverage(t *testing.T) {
	fc := NewFontCache()

	latin := installedLatinOnlyFont(fc)
	if latin == "" {
		t.Skip("no Latin-only font installed; nothing to check absence of coverage against")
	}
	if fc.CoversRune(latin, false, false, '组') {
		t.Errorf("CoversRune(%q, 组) = true, but %q has no Chinese glyphs; "+
			"a Latin face claiming coverage is how tofu gets drawn", latin, latin)
	}
	// The same font must still cover its own script, so the check is not simply
	// reporting false for everything.
	if !fc.CoversRune(latin, false, false, 'A') {
		t.Errorf("CoversRune(%q, A) = false, but every installed font has Latin glyphs", latin)
	}

	cjk := cjkCoveringFont(fc, cjkSample)
	if cjk == "" {
		t.Skip("no CJK-capable font installed; only the Latin half of this test is meaningful")
	}
	if !fc.CoversRune(cjk, false, false, '组') {
		t.Errorf("CoversRune(%q, 组) = false, but %q was selected for covering %q", cjk, cjk, cjkSample)
	}
}

// TestCJKDeclaredAsLatinStillDrawsGlyphs is the end-to-end regression test.
//
// Two presentations hold identical Chinese text. One declares the Latin-only
// font as its East Asian font, exactly as the generator that exposed this bug
// did; the other declares the East Asian font it will actually be drawn with.
// The rendered ink has to be the same, because the same characters must end up
// as the same glyphs. With name-based font selection the first one draws a
// missing-glyph box per character and the two counts diverge.
func TestCJKDeclaredAsLatinStillDrawsGlyphs(t *testing.T) {
	fc := NewFontCache()
	latin := installedLatinOnlyFont(fc)
	if latin == "" {
		t.Skip("no Latin-only font installed; the mis-declaration cannot be reproduced")
	}
	cjk := cjkCoveringFont(fc, cjkSample)
	if cjk == "" {
		t.Skip("no CJK-capable font installed; there is no correct rendering to compare against")
	}

	const width = 640
	render := func(nameEA string) (int, []FontUsage, *image.RGBA) {
		t.Helper()
		pres := cjkSlide("Arial", nameEA)
		diag := NewFontDiagnostics()
		opts := DefaultRenderOptions()
		opts.Width = width
		opts.FontCache = fc
		opts.FontDiagnostics = diag
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render with ea=%q: %v", nameEA, err)
		}
		rect := emuRect(t, pres, 400000, 400000, 9000000, 2000000, width)
		return countInkIn(img, rect), diag.Substituted(), toRGBA(t, img)
	}

	misdeclared, misdeclaredSubs, misdeclaredImg := render(latin)
	correct, correctSubs, correctImg := render(cjk)

	if misdeclared == 0 {
		t.Fatalf("declaring %q as the East Asian font produced no ink at all", latin)
	}
	if correct == 0 {
		t.Fatalf("declaring %q as the East Asian font produced no ink at all", cjk)
	}
	if misdeclared != correct {
		t.Errorf("Chinese text declared as %q inked %d pixels, declared as %q inked %d: "+
			"the Latin face must not be used for East Asian text",
			latin, misdeclared, cjk, correct)
	}
	if diff, _ := countDifferingInk(misdeclaredImg, correctImg); diff != 0 {
		t.Errorf("%d pixels differ between the mis-declared and correctly declared renders; "+
			"the same characters should become the same glyphs", diff)
	}

	// Silent tofu is the failure mode being guarded against, so the substitution
	// has to be reported. Without this the only symptom is a preview full of
	// boxes and a diagnostics summary claiming every request was satisfied.
	if len(misdeclaredSubs) == 0 {
		t.Errorf("no font substitution reported for ea=%q; silent tofu is what diagnostics exist to prevent", latin)
	}
	for _, u := range misdeclaredSubs {
		if u.Requested != latin {
			continue
		}
		if u.Kind != FontSubstituted {
			t.Errorf("replacing %q recorded kind %v, want %v", latin, u.Kind, FontSubstituted)
		}
		if u.Used != cjk {
			t.Errorf("replacing %q recorded used=%q, want %q", latin, u.Used, cjk)
		}
	}
	for _, u := range correctSubs {
		if u.Requested == cjk {
			t.Errorf("declaring %q as the East Asian font was reported as a substitution: %v", cjk, u)
		}
	}
}

// TestCoversRuneIgnoresNonCJKSample pins the deliberate refusal to vouch for a
// font that was never tested: a sample with no East Asian characters must not
// make the CJK path claim a match.
func TestCoversRuneIgnoresNonCJKSample(t *testing.T) {
	fc := NewFontCache()
	latin := installedLatinOnlyFont(fc)
	if latin == "" {
		t.Skip("no Latin-only font installed")
	}
	// coversCJK reads fontCache and nothing else, so a bare renderer is enough
	// here; no rendering happens and no other field is consulted.
	r := &renderer{fontCache: fc, cjkFallback: defaultCJKFallbackChain}
	if r.coversCJK(latin, false, false, "plain latin text") {
		t.Errorf("coversCJK(%q, %q) = true although the sample has no East Asian characters",
			latin, "plain latin text")
	}
}

// toRGBA re-reads an image through a concrete *image.RGBA so pixel access in the
// comparison helpers is O(1) rather than an interface call per pixel.
func toRGBA(t *testing.T, img image.Image) *image.RGBA {
	t.Helper()
	if src, ok := img.(*image.RGBA); ok {
		return src
	}
	b := img.Bounds()
	out := image.NewRGBA(b)
	for y := b.Min.Y; y < b.Max.Y; y++ {
		for x := b.Min.X; x < b.Max.X; x++ {
			out.Set(x, y, img.At(x, y))
		}
	}
	return out
}

// countDifferingInk returns how many pixels differ between two renders and the
// worst per-channel delta seen.
func countDifferingInk(a, b *image.RGBA) (n int, worst int) {
	if a.Bounds() != b.Bounds() {
		return a.Bounds().Dx()*a.Bounds().Dy() + b.Bounds().Dx()*b.Bounds().Dy(), 255
	}
	bb := a.Bounds()
	for y := bb.Min.Y; y < bb.Max.Y; y++ {
		for x := bb.Min.X; x < bb.Max.X; x++ {
			i := a.PixOffset(x, y)
			j := b.PixOffset(x, y)
			d := 0
			for k := 0; k < 4; k++ {
				if v := channelDelta(uint32(a.Pix[i+k]), uint32(b.Pix[j+k])); v > d {
					d = v
				}
			}
			if d > 0 {
				n++
			}
			if d > worst {
				worst = d
			}
		}
	}
	return n, worst
}
