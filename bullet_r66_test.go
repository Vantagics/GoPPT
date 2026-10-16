package gopresentation

import (
	"image"
	"testing"
)

// Two rules PowerPoint applies to bullets that the renderer did not, both found
// by rendering real decks rather than by reading the spec.
//
// 1. A character bullet is ONE glyph. <a:buChar> is a single character in
//    PowerPoint's model, and everything after the first is dropped. The COM
//    probe pinned which half that is: char="•••" exported as one dot (so it is
//    not a de-duplication of repeats) and char="○•" exported as the circle (so
//    the survivor is the first character). customGeo slide47 declares
//    buChar char="••" on every list paragraph of its cached SmartArt drawing,
//    so every bulleted line of that shape drew two dots against PowerPoint's
//    one.
//
// 2. A paragraph with no elements gets no bullet. PowerPoint still gives the
//    line its box — sized by endParaRPr — but draws nothing on it. The same
//    cached drawing carries one such paragraph, which we drew as an orphan dot
//    under the column. A break counts as content: the probe paragraph whose
//    only elements were a break and a run kept its bullet, so the test is the
//    element list being empty, not the text being empty.

// bulletCharPresentation builds one shape whose single paragraph carries the
// given character bullet, with wrapping off so the bullet prefix is not pushed
// onto a line of its own by a narrow default box.
func bulletCharPresentation(char string) *Presentation {
	p := New()
	layout := p.GetLayout()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.SetOffsetX(0).SetOffsetY(0)
	if layout != nil {
		shape.SetWidth(layout.CX)
	}
	shape.SetHeight(1000000)
	shape.SetWordWrap(false)
	shape.CreateTextRun("item")
	para := shape.GetParagraphs()[0]
	para.SetBullet(NewBullet().SetCharBullet(char))
	return p
}

// TestCharacterBulletKeepsOnlyTheFirstCharacter is the headline for rule 1.
// The prefix is the string that reaches the line, so asserting on it asserts on
// the visible bullet.
func TestCharacterBulletKeepsOnlyTheFirstCharacter(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	cases := []struct {
		name string
		char string
		want string
	}{
		{"one character is unchanged", "•", "• "},
		{"two characters keep the first", "••", "• "},
		{"three characters keep the first", "•••", "• "},
		{"a different first character survives", "○•", "○ "},
		{"the arrow is the first character", "→•", "→ "},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := bulletPrefix(t, fc, NewBullet().SetCharBullet(tc.char), 0).text
			if got != tc.want {
				t.Errorf("prefix for buChar %q = %q, want %q", tc.char, got, tc.want)
			}
		})
	}
}

// TestCharacterBulletDropsExtraGlyphsOnTheCanvas is the pixel half of rule 1:
// the two-dot bullet has to ink the same pixels as the one-dot bullet, which is
// what "one glyph" means once it is drawn rather than measured.
func TestCharacterBulletDropsExtraGlyphsOnTheCanvas(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	one, err := bulletCharPresentation("•").SlideToImage(0, goldenOptions(fc))
	if err != nil {
		t.Fatalf("render single: %v", err)
	}
	two, err := bulletCharPresentation("••").SlideToImage(0, goldenOptions(fc))
	if err != nil {
		t.Fatalf("render double: %v", err)
	}

	bandsOne := inkRowBands(one)
	bandsTwo := inkRowBands(two)
	if len(bandsOne) != 1 || len(bandsTwo) != 1 {
		t.Fatalf("expected one text line each, got %d and %d", len(bandsOne), len(bandsTwo))
	}
	// Not just the same band height: the second dot of "••" lands to the right
	// of the first, so it shifts the ink's right edge and its pixel count.
	if bandsOne[0].Dx() != bandsTwo[0].Dx() {
		t.Errorf("a two-dot bullet inked %d columns, a one-dot bullet %d — the extra glyph reached the canvas",
			bandsTwo[0].Dx(), bandsOne[0].Dx())
	}
	if !sameInk(one, bandsOne[0], two, bandsTwo[0]) {
		t.Error("a two-dot bullet drew different pixels from a one-dot bullet")
	}
}

// TestEmptyParagraphDrawsNoBullet is rule 2's headline: the shape holds one
// real item and one paragraph with no elements at all, so the render has to
// show exactly one line of ink. Before the fix the empty paragraph contributed
// a second band holding nothing but the bullet.
func TestEmptyParagraphDrawsNoBullet(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	p := New()
	layout := p.GetLayout()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.SetOffsetX(0).SetOffsetY(0)
	if layout != nil {
		shape.SetWidth(layout.CX)
	}
	shape.SetHeight(3000000)
	shape.SetWordWrap(false)
	shape.CreateTextRun("item")
	shape.GetParagraphs()[0].SetBullet(NewBullet().SetCharBullet("•"))
	// No CreateTextRun and no CreateBreak: this is the paragraph the cached
	// SmartArt drawing of customGeo slide47 carries, endParaRPr and all.
	empty := shape.CreateParagraph()
	empty.SetBullet(NewBullet().SetCharBullet("•"))

	img, err := p.SlideToImage(0, goldenOptions(fc))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	bands := inkRowBands(img)
	if len(bands) != 1 {
		t.Fatalf("found %d text lines in a one-item list; the empty paragraph inked something", len(bands))
	}
}

// TestParagraphWithOnlyABreakStillGetsItsBullet is the control that keeps rule 2
// narrow. A break is content, so a paragraph holding one must keep its bullet —
// without this, "no text" would be the test and the break would lose it.
func TestParagraphWithOnlyABreakStillGetsItsBullet(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	p := New()
	layout := p.GetLayout()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.SetOffsetX(0).SetOffsetY(0)
	if layout != nil {
		shape.SetWidth(layout.CX)
	}
	shape.SetHeight(3000000)
	shape.SetWordWrap(false)
	para := shape.GetParagraphs()[0]
	para.SetBullet(NewBullet().SetCharBullet("•"))
	para.CreateBreak()
	para.CreateTextRun("item")

	img, err := p.SlideToImage(0, goldenOptions(fc))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// The break opens a second line, so two bands is the fixture working; the
	// assertion is that ink reaches the left margin of the second one, where
	// only the bullet can be.
	bands := inkRowBands(img)
	if len(bands) != 2 {
		t.Fatalf("found %d text lines in a one-break paragraph; the fixture cannot see the bullet", len(bands))
	}
	if !bandHasInkAtLeft(img, bands[1]) {
		t.Error("the line opened by a break carries no bullet at its left edge")
	}
}

// TestEmptyParagraphStillOccupiesItsLineBox guards the other half of rule 2:
// dropping the bullet must not drop the line, AND the empty line must keep the
// height endParaRPr gives it. A shape anchored to the bottom therefore has to
// leave a gap that scales with endParaRPrSize. That scaling is exactly what the
// bug (an orphan bullet run sizing the line) destroys: a bullet run pins the
// line to a fixed bullet height and ignores endParaRPr entirely, so the gap
// stops moving when the size doubles.
func TestEmptyParagraphStillOccupiesItsLineBox(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	// build returns a bottom-anchored one-item shape with a trailing empty
	// paragraph. When size > 0 the empty paragraph carries that endParaRPrSize,
	// which is the only thing that should size its (bullet-less) line.
	build := func(size int) *Presentation {
		p := New()
		layout := p.GetLayout()
		shape := p.GetActiveSlide().CreateRichTextShape()
		shape.SetOffsetX(0).SetOffsetY(0)
		if layout != nil {
			shape.SetWidth(layout.CX)
		}
		shape.SetHeight(3000000)
		shape.SetWordWrap(false)
		// Bottom-anchored so the last line sits against the bottom inset: the
		// first line's top edge then moves up by exactly the empty line's
		// height when the paragraph is added, which is what makes that height
		// observable. Top-anchored text hides it entirely.
		shape.SetTextAnchor(TextAnchorBottom)
		shape.CreateTextRun("item")
		shape.GetParagraphs()[0].SetBullet(NewBullet().SetCharBullet("•"))
		if size > 0 {
			empty := shape.CreateParagraph()
			empty.SetBullet(NewBullet().SetCharBullet("•"))
			// endParaRPr is what sizes the empty line; without it the line is a
			// fixed 14px and the comparison below says nothing about endParaRPr.
			// The field is package-private and the writer does not carry it, so
			// it is set directly — this is a render-side assertion.
			empty.endParaRPrSize = size
		}
		return p
	}
	topOf := func(size int) int {
		img, err := build(size).SlideToImage(0, goldenOptions(fc))
		if err != nil {
			t.Fatalf("render size %d: %v", size, err)
		}
		return inkRowBands(img)[0].Min.Y
	}

	// No empty paragraph: the baseline against which the pushes are measured.
	topNone := topOf(0)
	// Two empty sizes: doubling endParaRPrSize must push the text up further.
	topSmall := topOf(1800)
	topLarge := topOf(3600)

	pushSmall := topNone - topSmall
	pushLarge := topNone - topLarge
	// The empty paragraph must occupy space at all.
	if pushSmall <= 0 {
		t.Fatalf("an empty paragraph added no height: top %d with it, %d without", topSmall, topNone)
	}
	// Doubling endParaRPrSize must move the text up by a further visible amount.
	// This is the part the orphan-bullet bug kills: with a bullet run sizing the
	// line, topSmall == topLarge and this extra push is ~0.
	extra := pushLarge - pushSmall
	if extra < 15 {
		t.Errorf("doubling endParaRPrSize moved the text only %dpx (small push %d, large push %d); the empty line height is pinned, not sized by endParaRPr",
			extra, pushSmall, pushLarge)
	}
}

// sameInk reports whether the two bands hold the same pixels.
func sameInk(a image.Image, ra image.Rectangle, b image.Image, rb image.Rectangle) bool {
	if ra.Dx() != rb.Dx() || ra.Dy() != rb.Dy() {
		return false
	}
	for y := 0; y < ra.Dy(); y++ {
		for x := 0; x < ra.Dx(); x++ {
			r1, g1, b1, a1 := a.At(ra.Min.X+x, ra.Min.Y+y).RGBA()
			r2, g2, b2, a2 := b.At(rb.Min.X+x, rb.Min.Y+y).RGBA()
			if r1 != r2 || g1 != g2 || b1 != b2 || a1 != a2 {
				return false
			}
		}
	}
	return true
}

// bandHasInkAtLeft reports whether the band's leftmost columns hold ink, which
// is where a bullet sits.
func bandHasInkAtLeft(img image.Image, band image.Rectangle) bool {
	for x := band.Min.X; x < band.Min.X+band.Dx()/4 && x < band.Max.X; x++ {
		for y := band.Min.Y; y < band.Max.Y; y++ {
			r, g, b, a := img.At(x, y).RGBA()
			if a>>8 >= 8 && (r>>8 < 250 || g>>8 < 250 || b>>8 < 250) {
				return true
			}
		}
	}
	return false
}
