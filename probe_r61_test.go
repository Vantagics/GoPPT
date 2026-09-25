package gopresentation

import (
	"os"
	"testing"
)

// Round 61 probe deck: pins the vertical-layout float laws together.
//
// Round 59 proved PowerPoint accumulates the 1.2 × size line pitch as a
// float and rounds each line's position, but landing that alone regressed
// the comparison decks because the anchor (ctr/b) needs the block height —
// and nobody pinned whether that total is the rounded sum of rounded lines
// or the round of the float sum. This deck separates the two questions:
//
//	s1  anchor=top, 20×16pt   — pitch accumulation (round(k·42.67px))
//	s2  anchor=ctr, 13×16pt   — totalH law from the first line's offset
//	s3  anchor=bot, 13×16pt   — the same from the other end
//	s4  anchor=ctr + spcBef   — do spaces ride the float accumulation too
//	s5  anchor=top + spcPct   — 125% lines, float pitch scaled
//	s6  anchor=top + empties  — empty paragraphs' 1.2×endParaRPr lines
//
// Boxes are placed on round pixel coordinates (×5715 EMU/px at 160dpi) so
// the box origin itself never contributes a fraction; the first text line's
// offset from the box then exposes the anchor arithmetic directly.
func TestZZProbe61Build(t *testing.T) {
	const pxEMU = 5715 // 1 px at 160 dpi
	px := func(v int64) int64 { return v * pxEMU }

	addBox := func(s *Slide, x, y, w, h int64, anchor TextAnchorType) *RichTextShape {
		box := s.AddTextBox()
		box.SetOffsetX(px(x)).SetOffsetY(px(y)).SetWidth(px(w)).SetHeight(px(h))
		box.SetTextAnchor(anchor)
		return box
	}
	// linePara adds one short one-line paragraph at 16pt Calibri (named, so
	// the GDI win metrics load — the round 59 lesson). The first call must
	// reuse the shape's initial paragraph, or every box grows a stray empty
	// line at the top and every measurement shifts by one pitch.
	linePara := func(box *RichTextShape, text string) {
		var p *Paragraph
		if init := box.GetParagraphs(); len(init) == 1 && len(init[0].GetElements()) == 0 && init[0].endParaRPrSize == 0 {
			p = init[0]
		} else {
			p = box.CreateParagraph()
		}
		if text == "" {
			return
		}
		r := p.CreateTextRun(text)
		f := NewFont()
		f.SetName("Calibri").SetSize(16)
		r.SetFont(f)
	}

	pres := New()

	// s1: top anchor, 20 lines, float pitch 42.67px.
	s1 := pres.CreateSlide()
	b1 := addBox(s1, 200, 100, 700, 1000, TextAnchorTop)
	for i := 1; i <= 20; i++ {
		linePara(b1, "L"+string(rune('A'+i/10))+string(rune('A'+i%10)))
	}

	// s2/s3: ctr and bottom anchors, 13 lines (float total 554.67px) in a
	// 700px box: current per-line rounding sums 559 and centres at +70,
	// a float total of 555 centres at +72 — 2px apart on every line.
	for _, tc := range []struct {
		anchor TextAnchorType
	}{{TextAnchorMiddle}, {TextAnchorBottom}} {
		s := pres.CreateSlide()
		b := addBox(s, 200, 100, 700, 700, tc.anchor)
		for i := 1; i <= 13; i++ {
			linePara(b, "L"+string(rune('A'+i/10))+string(rune('A'+i%10)))
		}
	}

	// s4: centre anchor with a 20pt space before every paragraph after the
	// first — 9 spaces of 44.44px and 10 pitches of 42.67px, both fractional.
	s4 := pres.CreateSlide()
	b4 := addBox(s4, 200, 100, 700, 900, TextAnchorMiddle)
	for i := 1; i <= 10; i++ {
		var p *Paragraph
		if init := b4.GetParagraphs(); i == 1 && len(init) == 1 && len(init[0].GetElements()) == 0 {
			p = init[0]
		} else {
			p = b4.CreateParagraph()
			p.SetSpaceBefore(2000)
		}
		r := p.CreateTextRun("L" + string(rune('A'+i/10)) + string(rune('A'+i%10)))
		f := NewFont()
		f.SetName("Calibri").SetSize(16)
		r.SetFont(f)
	}

	// s5: top anchor, 125% line spacing — the scaled pitch 53.33px also
	// accumulates a third of a pixel per line.
	s5 := pres.CreateSlide()
	b5 := addBox(s5, 200, 50, 700, 700, TextAnchorTop)
	for i := 1; i <= 12; i++ {
		var p *Paragraph
		if init := b5.GetParagraphs(); i == 1 && len(init) == 1 && len(init[0].GetElements()) == 0 {
			p = init[0]
		} else {
			p = b5.CreateParagraph()
		}
		p.SetLineSpacing(-125000)
		r := p.CreateTextRun("L" + string(rune('A'+i/10)) + string(rune('A'+i%10)))
		f := NewFont()
		f.SetName("Calibri").SetSize(16)
		r.SetFont(f)
	}

	// s6: top anchor with two empty paragraphs (16pt endParaRPr) — their
	// 1.2 × 16pt lines must ride the same float accumulation.
	s6 := pres.CreateSlide()
	b6 := addBox(s6, 200, 100, 700, 700, TextAnchorTop)
	for i := 1; i <= 3; i++ {
		linePara(b6, "L"+string(rune('A'+i)))
	}
	for i := 0; i < 2; i++ {
		p := b6.CreateParagraph()
		p.endParaRPrSize = 1600
	}
	for i := 4; i <= 7; i++ {
		linePara(b6, "L"+string(rune('A'+i)))
	}

	// s7: centre anchor, ONE 44pt line in a 120px box — the slide-11 title
	// configuration exactly. Multi-line probes cannot see the block-height
	// law because the fractional parts cancel; a single centred line pins
	// whether the anchor block is N×pitch (117.33→117) or
	// (N−1)×pitch + ascent+descent (119.36→119): the two models put the
	// first baseline 1px apart here.
	mkLine := func(box *RichTextShape, size int, text string) {
		p := box.GetParagraphs()[0]
		r := p.CreateTextRun(text)
		f := NewFont()
		f.SetName("Calibri").SetSize(size)
		r.SetFont(f)
	}
	s7 := pres.CreateSlide()
	b7 := addBox(s7, 200, 100, 700, 120, TextAnchorMiddle)
	mkLine(b7, 44, "TITLE")

	// s8: the same from the bottom anchor.
	s8 := pres.CreateSlide()
	b8 := addBox(s8, 200, 100, 700, 120, TextAnchorBottom)
	mkLine(b8, 44, "TITLE")

	// s9: two 44pt lines centred in a 240px box — checks whether the extra
	// descent+ascent rides on the last line only.
	s9 := pres.CreateSlide()
	b9 := addBox(s9, 200, 100, 700, 240, TextAnchorMiddle)
	mkLine(b9, 44, "TITLE")
	mkLine(b9, 44, "TITLE")

	// s10: one 16pt line centred in a 60px box — the small-size counterpart.
	s10 := pres.CreateSlide()
	b10 := addBox(s10, 200, 100, 700, 60, TextAnchorMiddle)
	mkLine(b10, 16, "tiny")

	out := os.Getenv("ZZ_PROBE61_OUT")
	if out == "" {
		t.Skip("set ZZ_PROBE61_OUT to build the probe deck")
	}
	if err := pres.Save(out); err != nil {
		t.Fatalf("save: %v", err)
	}
	t.Logf("probe deck written to %s", out)
}

// TestZZProbe61C pins the centre-anchor rounding staircase: box heights rise
// one pixel at a time so the centre offset crosses every rounding boundary.
//
//	s1..s20  one 44pt Calibri line, box h = 120..139 px
//	s21..s40 thirteen 16pt Calibri lines, box h = 570..589 px
//
// The first band top of each slide reads back the offset PowerPoint chose;
// comparing the two staircases separates the block-height law (N×pitch vs
// (N−1)×pitch+ascent+descent) from the half-slack rounding (floor vs round).
func TestZZProbe61C(t *testing.T) {
	const pxEMU = 5715
	px := func(v int64) int64 { return v * pxEMU }

	pres := New()
	for i := 0; i < 20; i++ {
		s := pres.CreateSlide()
		box := s.AddTextBox()
		box.SetOffsetX(px(200)).SetOffsetY(px(100)).SetWidth(px(700)).SetHeight(px(int64(120 + i)))
		box.SetTextAnchor(TextAnchorMiddle)
		p := box.GetParagraphs()[0]
		r := p.CreateTextRun("TITLE")
		f := NewFont()
		f.SetName("Calibri").SetSize(44)
		r.SetFont(f)
	}
	for i := 0; i < 20; i++ {
		s := pres.CreateSlide()
		box := s.AddTextBox()
		box.SetOffsetX(px(200)).SetOffsetY(px(100)).SetWidth(px(700)).SetHeight(px(int64(570 + i)))
		box.SetTextAnchor(TextAnchorMiddle)
		first := true
		for j := 0; j < 13; j++ {
			var p *Paragraph
			if first {
				p = box.GetParagraphs()[0]
				first = false
			} else {
				p = box.CreateParagraph()
			}
			r := p.CreateTextRun("L" + string(rune('A'+j/10)) + string(rune('A'+j%10)))
			f := NewFont()
			f.SetName("Calibri").SetSize(16)
			r.SetFont(f)
		}
	}

	out := os.Getenv("ZZ_PROBE61C_OUT")
	if out == "" {
		t.Skip("set ZZ_PROBE61C_OUT to build the probe deck")
	}
	if err := pres.Save(out); err != nil {
		t.Fatalf("save: %v", err)
	}
	t.Logf("probe deck written to %s", out)
}

// TestZZProbe61D: a 16pt single-line centre staircase (box h = 56..75) —
// the small-size counterpart of the first C group. It separates the
// single-line block law (floor(N×pitch) vs ascent+descent) at a size where
// the two candidates differ by 2px.
func TestZZProbe61D(t *testing.T) {
	const pxEMU = 5715
	px := func(v int64) int64 { return v * pxEMU }

	pres := New()
	for i := 0; i < 20; i++ {
		s := pres.CreateSlide()
		box := s.AddTextBox()
		box.SetOffsetX(px(200)).SetOffsetY(px(100)).SetWidth(px(700)).SetHeight(px(int64(56 + i)))
		box.SetTextAnchor(TextAnchorMiddle)
		p := box.GetParagraphs()[0]
		r := p.CreateTextRun("Lx")
		f := NewFont()
		f.SetName("Calibri").SetSize(16)
		r.SetFont(f)
	}

	out := os.Getenv("ZZ_PROBE61D_OUT")
	if out == "" {
		t.Skip("set ZZ_PROBE61D_OUT to build the probe deck")
	}
	if err := pres.Save(out); err != nil {
		t.Fatalf("save: %v", err)
	}
	t.Logf("probe deck written to %s", out)
}

// TestZZProbe61B pins the spcPct law: one ten-line paragraph per percentage
// (110 / 125 / 150 / 175 / 200 %), so the COM export yields both the scaled
// pitch (the band-gap cycle) and the first baseline's phase for each.
func TestZZProbe61B(t *testing.T) {
	const pxEMU = 5715
	px := func(v int64) int64 { return v * pxEMU }

	pres := New()
	for _, pct := range []int{-110000, -125000, -150000, -175000, -200000} {
		s := pres.CreateSlide()
		box := s.AddTextBox()
		box.SetOffsetX(px(150)).SetOffsetY(px(50)).SetWidth(px(400)).SetHeight(px(1100))
		box.SetTextAnchor(TextAnchorTop)
		first := true
		for i := 0; i < 10; i++ {
			var p *Paragraph
			if first {
				p = box.GetParagraphs()[0]
				first = false
			} else {
				p = box.CreateParagraph()
			}
			p.SetLineSpacing(pct)
			r := p.CreateTextRun("LP" + string(rune('A'+i)))
			f := NewFont()
			f.SetName("Calibri").SetSize(16)
			r.SetFont(f)
		}
	}

	out := os.Getenv("ZZ_PROBE61B_OUT")
	if out == "" {
		t.Skip("set ZZ_PROBE61B_OUT to build the probe deck")
	}
	if err := pres.Save(out); err != nil {
		t.Fatalf("save: %v", err)
	}
	t.Logf("probe deck written to %s", out)
}

// TestZZProbe61E pins the horizontal advance law. Each slide places one long
// string at 14pt in a known style (regular / bold / bold-italic) plus word
// markers, so the COM export's per-word pen positions can be read off the
// gold bitmaps and compared against candidate advance models.
func TestZZProbe61E(t *testing.T) {
	const pxEMU = 5715
	px := func(v int64) int64 { return v * pxEMU }

	pres := New()
	s := pres.CreateSlide()
	box := s.AddTextBox()
	box.SetOffsetX(px(60)).SetOffsetY(px(100)).SetWidth(px(1480)).SetHeight(px(300))
	p := box.GetParagraphs()[0]
	r := p.CreateTextRun("Human Genome Sequencing Using Unchained Base Reads on Self-Assembling DNA Nanoarrays")
	f := NewFont()
	f.SetName("Calibri").SetSize(14)
	r.SetFont(f)

	s2 := pres.CreateSlide()
	box2 := s2.AddTextBox()
	box2.SetOffsetX(px(60)).SetOffsetY(px(100)).SetWidth(px(1480)).SetHeight(px(300))
	p2 := box2.GetParagraphs()[0]
	r2 := p2.CreateTextRun("Human Genome Sequencing Using Unchained Base Reads on Self-Assembling DNA Nanoarrays")
	f2 := NewFont()
	f2.SetName("Calibri").SetSize(14).SetBold(true)
	r2.SetFont(f2)

	s3 := pres.CreateSlide()
	box3 := s3.AddTextBox()
	box3.SetOffsetX(px(60)).SetOffsetY(px(100)).SetWidth(px(1480)).SetHeight(px(300))
	p3 := box3.GetParagraphs()[0]
	r3 := p3.CreateTextRun("Human Genome Sequencing Using Unchained Base Reads on Self-Assembling DNA Nanoarrays")
	f3 := NewFont()
	f3.SetName("Calibri").SetSize(14).SetBold(true).SetItalic(true)
	r3.SetFont(f3)

	out := os.Getenv("ZZ_PROBE61E_OUT")
	if out == "" {
		t.Skip("set ZZ_PROBE61E_OUT to build the probe deck")
	}
	if err := pres.Save(out); err != nil {
		t.Fatalf("save: %v", err)
	}
	t.Logf("probe deck written to %s", out)
}
