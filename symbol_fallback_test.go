package gopresentation

import (
	"testing"

	"golang.org/x/image/font"
)

// Pictographs (emoji, dingbats, the miscellaneous symbol blocks) in slide text.
//
// The bug these tests pin down: an emoji is drawn from a symbol face, and no
// face the renderer could choose had a glyph for one. The face a document
// declares for body text is a text face, and the CJK fallback chain is made of
// text faces too, so every emoji reached the declared face, found no glyph, and
// was drawn as its .notdef box — the hollow rectangle that reads as a white
// square on the slide.
//
// The failure was silent in exactly the way the ①②③ tofu was: the requested
// font name resolves, so the font diagnostics reported "all requests satisfied
// exactly" while six white squares sat on the slide. The deck that exposed it is
// a bulleted list whose twelve bullets each lead with an emoji.
//
// The fix has four parts, and each one has a test here:
//
//  1. isSymbolRune classifies the pictograph blocks, so those characters reach a
//     coverage check at all (TestIsSymbolRune..., TestClassOf...).
//  2. A third fallback chain supplies font faces that actually carry emoji, and
//     the class is resolved by glyph coverage rather than by name
//     (TestSymbolFaceDrawsEveryEmojiInTheRun).
//  3. The run splitter cuts on the class, so one run can draw its emoji from a
//     symbol face and its Chinese from a CJK face
//     (TestSplitRunByClassSeparatesEmojiFromChinese).
//  4. End to end, the emoji are not boxes (TestEmojiAreNotDrawnAsBoxes), and
//     the substitution is reported instead of hidden
//     (TestNoteSymbolFontReportsTheSubstitution).

// emojiSample is the string the resolution tests render. The characters are
// visually unlike each other, which is what makes the ink test below work: a
// .notdef box is byte-identical for every one of them, while real glyphs differ.
const emojiSample = "💙🐱🏆"

// fontCoveringAllIn returns the first font of a chain whose glyph table covers
// every character of sample, or "" when none does.
//
// It asks the font about each character directly rather than calling the
// renderer, so the expectation it produces is independent of the code under
// test — a precondition that shares the bug would turn a regression into a skip.
func fontCoveringAllIn(fc *FontCache, chain []string, sample string) string {
	for _, name := range chain {
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

// installedFaceLacking returns an installed face that has no glyph for any
// character of sample, preferring the face the real deck declares. It returns ""
// when every installed face can draw the sample.
//
// The pixel test needs a declared face that boxes the sample, because that is
// the case being fixed: with such a face in place, identical ink counts for
// different emoji can only mean they were all drawn as the same .notdef box.
func installedFaceLacking(fc *FontCache, sample string) string {
	names := append([]string{"Microsoft YaHei"}, defaultFontFallbackChain...)
	for _, name := range names {
		if !fc.HasFont(name, false, false) {
			continue
		}
		lacks := true
		for _, r := range sample {
			if fc.CoversRune(name, false, false, r) {
				lacks = false
				break
			}
		}
		if lacks {
			return name
		}
	}
	return ""
}

// TestIsSymbolRuneClassifiesPictographs pins the classification itself,
// including the boundaries either side of each block, so an edit that widens a
// range too far is caught as well as one that narrows it.
func TestIsSymbolRuneClassifiesPictographs(t *testing.T) {
	cases := []struct {
		r    rune
		want bool
		why  string
	}{
		{0x1EFFF, false, "just below the pictograph planes"},
		{0x1F000, true, "🀀 first Mahjong tile"},
		{0x1F431, true, "🐱"},
		{0x1F499, true, "💙"},
		{0x1F9F6, true, "🧶 Symbols and Pictographs Extended-A"},
		{0x1FAFF, true, "last of the Extended-A emoji block"},
		{0x1FB00, false, "Symbols for Legacy Computing are not emoji"},
		{0x25FF, false, "just below Miscellaneous Symbols"},
		{0x2600, true, "☀ first of Miscellaneous Symbols"},
		{0x263A, true, "☺"},
		{0x2764, true, "❤ is Dingbats, written as an emoji"},
		{0x27BF, true, "last of Dingbats"},
		{0x27C0, false, "Miscellaneous Mathematical Symbols-A sit past Dingbats"},
		{0x2AFF, false, "just below Miscellaneous Symbols and Arrows"},
		{0x2B00, true, "⬀ fist of the arrow block"},
		{0x2BFF, true, "last of Miscellaneous Symbols and Arrows"},
		{0x22FF, false, "just below Miscellaneous Technical"},
		{0x2300, true, "⌀ fist of Miscellaneous Technical"},
		{0x23F0, true, "⏰ is written as an emoji"},
		{0x2460, false, "① is CJK typography, not a pictograph"},
		{0x33A1, false, "㎡ is CJK typography"},
		{0x4E2D, false, "中 is Han"},
		{0x0041, false, "A is Latin"},
		{0xFE0F, false, "a variation selector carries no glyph of its own"},
	}
	for _, c := range cases {
		if got := isSymbolRune(c.r); got != c.want {
			t.Errorf("isSymbolRune(U+%04X) = %v, want %v (%s)", c.r, got, c.want, c.why)
		}
	}
}

// TestClassOfRoutesEmojiCJKAndTextApart pins the dispatch every face-selection
// decision hangs off. Getting a class wrong is not a cosmetic problem: the class
// decides which fallback chain a character is offered to, and an emoji sent to
// the CJK chain finds no glyph there either.
func TestClassOfRoutesEmojiCJKAndTextApart(t *testing.T) {
	cases := []struct {
		r    rune
		want textClass
		why  string
	}{
		{'A', textClassLatin, "ordinary text belongs to the declared face"},
		{'中', textClassCJK, "Han belongs to a CJK face"},
		{0x2460, textClassCJK, "① is CJK typography, so it goes to the CJK chain"},
		{0x1F431, textClassSymbol, "🐱 needs a symbol face"},
		{0x2600, textClassSymbol, "☀ needs a symbol face"},
		{0xFE0F, textClassLatin, "a variation selector has no class of its own"},
	}
	for _, c := range cases {
		if got := classOf(c.r); got != c.want {
			t.Errorf("classOf(U+%04X) = %d, want %d (%s)", c.r, got, c.want, c.why)
		}
	}
}

// TestSymbolFaceDrawsEveryEmojiInTheRun is the resolution test, stated as the
// property that matters: the face chosen for a run's emoji must have glyphs for
// them.
//
// Before the fix there was no such resolution at all — the emoji were not a
// class, so nothing asked the question, and the declared face drew its .notdef
// box for each of them.
func TestSymbolFaceDrawsEveryEmojiInTheRun(t *testing.T) {
	fc := NewFontCache()
	if fontCoveringAllIn(fc, defaultSymbolFallbackChain, emojiSample) == "" {
		t.Skip("no installed symbol face covers the sample emoji; the boxes cannot be avoided")
	}

	r := &renderer{fontCache: fc, symbolFallback: defaultSymbolFallbackChain}
	f := &Font{Name: "Microsoft YaHei"}
	f.NameEA = "Microsoft YaHei" // the declaration the deck ships with

	face, used := r.resolveSymbolFace(f, 40, false, emojiSample)
	if face == nil || used == "" {
		t.Fatalf("resolveSymbolFace(%q) found no face; the emoji would be drawn by the declared "+
			"text face and render as .notdef boxes", emojiSample)
	}
	for _, ch := range emojiSample {
		if !fc.CoversRune(used, false, false, ch) {
			t.Errorf("resolveSymbolFace chose %q, which has no glyph for %s (U+%04X)",
				used, string(ch), ch)
		}
	}
}

// TestSplitRunByClassSeparatesEmojiFromChinese is the structural test: a run
// written the way the deck writes it — an emoji leading Chinese text — must come
// apart into one segment drawn by a symbol face and one drawn by a CJK face.
//
// It builds the runs through buildParaTextRuns, so it exercises the real path
// including the class resolution, rather than a hand-built approximation of it.
func TestSplitRunByClassSeparatesEmojiFromChinese(t *testing.T) {
	fc := NewFontCache()
	symbolFont := fontCoveringAllIn(fc, defaultSymbolFallbackChain, "💙")
	cjkFont := fontCoveringAllIn(fc, defaultCJKFallbackChain, cjkSample)
	if symbolFont == "" || cjkFont == "" {
		t.Skip("this machine lacks either a symbol face or a CJK face; the split cannot happen")
	}

	const text = "💙" + cjkSample
	pres := New()
	shape := pres.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetWidth(9000000).SetHeight(2000000)
	run := shape.CreateTextRun(text)
	run.GetFont().SetName("Microsoft YaHei").SetSize(24)
	run.GetFont().NameEA = "Microsoft YaHei"

	layout := pres.GetLayout()
	if layout == nil || layout.CX <= 0 {
		t.Fatal("presentation has no usable slide size")
	}
	r := &renderer{
		fontCache:      fc,
		scaleX:         float64(1200) / float64(layout.CX),
		dpi:            96,
		fontFallback:   mergeFallbackChain(nil, defaultFontFallbackChain),
		cjkFallback:    mergeFallbackChain(nil, defaultCJKFallbackChain),
		symbolFallback: mergeFallbackChain(nil, defaultSymbolFallbackChain),
	}

	runs := r.buildParaTextRuns(shape.GetActiveParagraph().GetElements())
	if len(runs) != 2 {
		t.Fatalf("run %q split into %d segments, want 2 (the emoji, then the Chinese); "+
			"one segment means the emoji is drawn with the CJK face or the Chinese with the "+
			"symbol face", text, len(runs))
	}
	if runs[0].text != "💙" || runs[1].text != cjkSample {
		t.Fatalf("segments = %q, %q; want %q, %q", runs[0].text, runs[1].text, "💙", cjkSample)
	}
	if runs[0].face == nil || runs[1].face == nil {
		t.Fatal("a segment has no face")
	}
	if runs[0].face == runs[1].face {
		t.Errorf("both segments share one face (%v); no single face can draw both the emoji "+
			"and the Chinese characters", runs[0].face)
	}
	// The Chinese characters are drawn as before; only the emoji changed face.
	for _, ch := range cjkSample {
		if !fc.CoversRune(cjkFont, false, false, ch) {
			t.Errorf("the CJK font %q has no glyph for %s", cjkFont, string(ch))
		}
	}
	if runs[1].face != facesFor(t, r, run.GetFont(), cjkSample, textClassCJK) {
		t.Errorf("the Chinese segment is not drawn with the CJK face the renderer resolves")
	}
}

// facesFor resolves the face the renderer would use for one class of a sample,
// so a test can compare a segment's face against it by identity. FontCache hands
// out one face per (name, size, style), so identity is stable within a render.
func facesFor(t *testing.T, r *renderer, f *Font, sample string, class textClass) font.Face {
	t.Helper()
	face, _ := r.resolveClassFace(f, r.fontSizePixels(f), false, sample, class)
	if face == nil {
		t.Fatalf("no face for class %d of %q", class, sample)
	}
	return face
}

// TestEmojiAreNotDrawnAsBoxes is the end-to-end pixel test, and the one that
// would have caught the deck.
//
// A .notdef box is the same shape for every character, so three different emoji
// drawn as boxes ink an identical number of pixels. Real glyphs differ, because
// the characters differ. Requiring the counts not to be uniformly equal
// therefore distinguishes "drawn" from "boxed" without hard-coding any font's
// metrics — which is what lets it survive a machine with different fonts.
//
// The case is constructed rather than hoped for: a declared face that boxes the
// sample (the situation the deck is in) and a symbol face that can draw it (what
// the fix supplies). Without the first, a machine whose text face happens to
// carry emoji would pass this test even with the fix reverted.
func TestEmojiAreNotDrawnAsBoxes(t *testing.T) {
	fc := NewFontCache()
	if fontCoveringAllIn(fc, defaultSymbolFallbackChain, emojiSample) == "" {
		t.Skip("no installed symbol face covers the sample emoji; the boxes cannot be avoided")
	}
	declared := installedFaceLacking(fc, emojiSample)
	if declared == "" {
		t.Skip("every installed face draws the sample emoji; the boxed case cannot be constructed")
	}

	var ink []int
	for _, ch := range emojiSample {
		ink = append(ink, renderInk(t, fc, string(ch), declared, declared, 640))
	}
	for i, n := range ink {
		if n == 0 {
			t.Fatalf("%s rendered no ink at all", string([]rune(emojiSample)[i]))
		}
	}
	if ink[0] == ink[1] && ink[1] == ink[2] {
		t.Errorf("%s all inked exactly %d pixels while declared in %q: a .notdef box is "+
			"identical for every character, so uniform counts mean they were boxed rather "+
			"than drawn", emojiSample, ink[0], declared)
	}
}

// TestNoteSymbolFontReportsTheSubstitution pins the new visibility. A silent
// substitution at a *coverage* boundary is the worst kind: the name resolves, so
// nothing looks wrong, and the diagnostics that exist precisely to expose tofu
// have to say something or a reader has no way to tell why the slide differs
// from what PowerPoint draws.
func TestNoteSymbolFontReportsTheSubstitution(t *testing.T) {
	diag := NewFontDiagnostics()
	r := &renderer{fontDiag: diag}
	f := &Font{Name: "Microsoft YaHei"}
	f.NameEA = "Microsoft YaHei"

	r.noteSymbolFont(f, "Segoe UI Emoji")

	usages := diag.Usages()
	if len(usages) != 1 {
		t.Fatalf("recorded %d usages, want exactly 1: %v", len(usages), usages)
	}
	if usages[0].Requested != "Microsoft YaHei" || usages[0].Used != "Segoe UI Emoji" ||
		usages[0].Kind != FontSubstituted {
		t.Errorf("usage = %+v, want Microsoft YaHei substituted by Segoe UI Emoji", usages[0])
	}

	// A face the document declared itself is not a substitution, and an emoji
	// nothing can draw is a miss rather than a substitution.
	exact := NewFontDiagnostics()
	r2 := &renderer{fontDiag: exact}
	f2 := &Font{Name: "Segoe UI Emoji"}
	r2.noteSymbolFont(f2, "Segoe UI Emoji")
	if len(exact.Usages()) != 0 {
		t.Errorf("a declared symbol face was reported as substituted: %v", exact.Usages())
	}

	missing := NewFontDiagnostics()
	r3 := &renderer{fontDiag: missing}
	r3.noteSymbolFont(f, "")
	got := missing.Usages()
	if len(got) != 1 || got[0].Kind != FontMissing {
		t.Errorf("usages = %v, want one missing entry for an undrawable emoji", got)
	}
}
