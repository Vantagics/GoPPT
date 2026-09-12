package gopresentation

import (
	"testing"
)

// Enclosed symbols in CJK typography.
//
// The bug these tests pin down: ①②③ are Enclosed Alphanumerics (U+2460–U+24FF),
// which isCJK did not cover. Because every font-selection decision is gated on
// isCJK, a character outside that set never reached the coverage check at all —
// it went straight to the Latin face, and Arial drew its .notdef box. The deck
// that exposed it declares ea="Arial", so the circled digits in its numbered
// list rendered as tofu while the font diagnostics reported a perfect match.
//
// The fix has two halves, and both need a test:
//
//  1. isCJK classifies the enclosed symbol blocks, so those characters reach the
//     coverage check (TestIsCJKClassifiesEnclosedSymbols,
//     TestEnclosedDigitsReachTheCJKFace).
//  2. Face selection no longer surrenders when no font covers *every* character,
//     because a single uncoverable symbol in a run must not drag the Chinese
//     characters that ARE drawable down with it
//     (TestUncoverableSymbolKeepsTheRestOfTheRunDrawable).

// enclosedDigits are the characters that exposed the bug. They are visually
// near-identical, which is what makes the ink test below work: a .notdef box is
// byte-identical for each of them, while real glyphs differ.
const enclosedDigits = "①②③"

// renderInk renders text with the given Latin and East Asian font names and
// counts the ink inside the run's rectangle.
func renderInk(t *testing.T, fc *FontCache, text, latin, eastAsian string, width int) int {
	t.Helper()
	pres := New()
	shape := pres.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	shape.BaseShape.SetWidth(6000000).SetHeight(1500000)
	run := shape.CreateTextRun(text)
	run.GetFont().SetName(latin).SetSize(40)
	run.GetFont().NameEA = eastAsian

	opts := DefaultRenderOptions()
	opts.Width = width
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render %q with ea=%q: %v", text, eastAsian, err)
	}
	rect := emuRect(t, pres, 400000, 400000, 6000000, 1500000, width)
	return countInkIn(img, rect)
}

// TestIsCJKClassifiesEnclosedSymbols pins the classification itself, including
// the boundaries either side of each added block, so a later edit that widens a
// range too far is caught as well as one that narrows it.
func TestIsCJKClassifiesEnclosedSymbols(t *testing.T) {
	cases := []struct {
		r    rune
		want bool
		why  string
	}{
		{0x245F, false, "just below Enclosed Alphanumerics"},
		{0x2460, true, "① first of Enclosed Alphanumerics"},
		{0x2461, true, "②"},
		{0x2473, true, "⑳ last circled digit"},
		{0x24B6, true, "Ⓐ still inside the block"},
		{0x24FF, true, "last of Enclosed Alphanumerics"},
		{0x2500, false, "just above the block"},
		{0x3105, false, "Bopomofo ㄅ is in none of the tables isCJK tests"},
		{0x3200, true, "㈀ first of Enclosed CJK Letters and Months"},
		{0x3220, true, "㈠"},
		{0x3251, true, "㉑"},
		{0x32FF, true, "last of Enclosed CJK Letters and Months"},
		{0x3300, true, "first of CJK Compatibility"},
		{0x33A1, true, "㎡"},
		{0x33FF, true, "last of CJK Compatibility"},
		{0x4DC0, false, "Yijing Hexagram Symbols sit just past Han Extension A"},
		{0x2FFF, false, "just below CJK Symbols and Punctuation"},
		{0x3400, true, "CJK Extension A start is Han"},
		{0x4E2D, true, "中 is Han"},
		{0xFF01, true, "Fullwidth Forms"},
		{0x0041, false, "A is Latin"},
		{0x2776, false, "❶ is Dingbats, not a CJK block"},
	}
	for _, c := range cases {
		if got := isCJK(c.r); got != c.want {
			t.Errorf("isCJK(U+%04X) = %v, want %v (%s)", c.r, got, c.want, c.why)
		}
	}
}

// fontCoveringAll returns the first font of the CJK fallback chain whose glyph
// table covers every character of sample, or "" when none does.
//
// Unlike cjkCoveringFont this never consults isCJK: it asks the font about each
// character directly. That matters for these tests. A precondition built on
// cjkCoveringFont shares the bug under test — with the classification reverted
// it finds no covering font and the test skips, so a regression shows up as a
// skip rather than a failure. Asking the font keeps the precondition true
// independently of the code being checked.
func fontCoveringAll(fc *FontCache, sample string) string {
	for _, name := range defaultCJKFallbackChain {
		if !fc.HasFont(name, false, false) {
			continue
		}
		covers := true
		for _, r := range sample {
			if !fc.CoversRune(name, false, false, r) {
				covers = false
				break
			}
		}
		if covers {
			return name
		}
	}
	return ""
}

// TestEnclosedDigitsReachTheCJKFace is the core regression test, stated as the
// property that actually matters: the circled digits must be drawn with a face
// that has glyphs for them, not with a Latin face that will box them.
//
// Before the fix resolveCJKFace returned ("", nil) here, because the sample's
// CJK-character count was zero and coversCJK refuses to vouch for a font it
// never tested. That nil is what sent the text to Arial.
func TestEnclosedDigitsReachTheCJKFace(t *testing.T) {
	fc := NewFontCache()
	if fontCoveringAll(fc, enclosedDigits) == "" {
		t.Skip("no installed CJK font covers the enclosed digits; nothing to select")
	}

	r := &renderer{fontCache: fc, cjkFallback: defaultCJKFallbackChain}
	f := &Font{Name: "Arial"}
	f.NameEA = "Arial" // the mis-declaration the deck ships with

	face, used := r.resolveCJKFace(f, 40, false, enclosedDigits)
	if face == nil || used == "" {
		t.Fatalf("resolveCJKFace(%q) found no face; the digits would be drawn by the Latin face "+
			"and render as .notdef boxes", enclosedDigits)
	}
	for _, ch := range enclosedDigits {
		if !fc.CoversRune(used, false, false, ch) {
			t.Errorf("resolveCJKFace chose %q, which has no glyph for %s (U+%04X)", used, string(ch), ch)
		}
	}
}

// TestEnclosedDigitsAreNotDrawnAsBoxes is the end-to-end pixel test.
//
// A .notdef box is the same shape for every character, so the ink count for ①,
// ② and ③ is identical when they are boxed. Real glyphs differ, because the
// digits inside the circles differ. Requiring the three counts not to be
// uniformly equal therefore distinguishes "drawn" from "boxed" without
// hard-coding any font's metrics — which is what makes it survive a machine
// with different fonts installed.
func TestEnclosedDigitsAreNotDrawnAsBoxes(t *testing.T) {
	fc := NewFontCache()
	if fontCoveringAll(fc, enclosedDigits) == "" {
		t.Skip("no installed CJK font covers the enclosed digits; the tofu cannot be avoided")
	}

	var ink []int
	for _, ch := range enclosedDigits {
		ink = append(ink, renderInk(t, fc, string(ch), "Arial", "Arial", 640))
	}
	for i, n := range ink {
		if n == 0 {
			t.Fatalf("%s rendered no ink at all", string([]rune(enclosedDigits)[i]))
		}
	}
	if ink[0] == ink[1] && ink[1] == ink[2] {
		t.Errorf("①②③ all inked exactly %d pixels: a .notdef box is identical for every "+
			"character, so uniform counts mean they were boxed rather than drawn", ink[0])
	}
}

// TestUncoverableSymbolKeepsTheRestOfTheRunDrawable guards the hazard that
// widening isCJK introduced.
//
// coversCJK asks "does this font cover every CJK character in the sample", so
// putting one symbol that nothing can draw into a run would answer no for every
// candidate. Resolving that to nil would send the whole run — Chinese
// characters included — to the Latin face, turning text that used to render
// correctly into tofu. Best-effort selection must keep the drawable characters.
//
// The chain is pinned to a single known font so the case is constructed rather
// than hoped for: which enclosed symbols this machine can draw is a property of
// the installed fonts, and a test that depends on finding an undrawable one
// silently skips wherever the fonts are good enough.
func TestUncoverableSymbolKeepsTheRestOfTheRunDrawable(t *testing.T) {
	fc := NewFontCache()
	drawable := fontCoveringAll(fc, cjkSample)
	if drawable == "" {
		t.Skip("no CJK-capable font installed")
	}
	// An enclosed symbol that this font cannot draw, so full coverage of the
	// mixed sample is impossible whichever candidate is considered.
	uncoverable := rune(0)
	for _, cand := range []rune{0x24B6, 0x24EA, 0x32A4, 0x24B7, 0x24B8, 0x24D0} {
		if !isCJK(cand) {
			t.Fatalf("U+%04X is not classified as CJK; the premise of this test is stale", cand)
		}
		if !fc.CoversRune(drawable, false, false, cand) {
			uncoverable = cand
			break
		}
	}
	if uncoverable == 0 {
		t.Skipf("%q covers every sampled enclosed symbol; cannot construct the case", drawable)
	}

	r := &renderer{fontCache: fc, cjkFallback: []string{drawable}}
	f := &Font{Name: "Arial"}
	f.NameEA = "Arial"

	mixed := cjkSample + string(uncoverable)
	face, used := r.resolveCJKFace(f, 40, false, mixed)
	if face == nil || used == "" {
		t.Fatalf("resolveCJKFace(%q) gave up because U+%04X is uncoverable; "+
			"that would draw the Chinese characters with the Latin face", mixed, uncoverable)
	}
	if used != drawable {
		t.Errorf("best-effort selection chose %q, want %q: the font that covers the most "+
			"of the sample is the one that keeps the drawable characters visible", used, drawable)
	}
	for _, ch := range cjkSample {
		if !fc.CoversRune(used, false, false, ch) {
			t.Errorf("best-effort face %q has no glyph for %s, so text that used to render "+
				"correctly would become tofu", used, string(ch))
		}
	}
}
