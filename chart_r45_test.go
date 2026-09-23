package gopresentation

// The r45 text-break and wide-stroke semantics tests.
//
// Pinned against COM exports of deck 00022693 slide19 (the "Naïve
// Out-of-Core Multigrid" body) and slide23 (the Temporal Blocking diagram):
//
//   - An <a:br> with no runs before it forms a blank line whose advance is
//     1.2 × the break's own <a:rPr> size — COM variants: sz=6000 on the br
//     grew the blank by exactly the 1.2 × 60pt line, deleting the br
//     removed it, and the empty run before it was irrelevant. The old code
//     collapsed the blank to a 14px floor, holding the paragraph a full
//     line too high.
//   - Wide strokes (≥6px) must fill solid: the parallel-Wu passes composite
//     per-pass and leave hatching on diagonals — slide23's 6pt connector
//     arrows rendered as stripes where PowerPoint drew solid blue.

import (
	"bytes"
	"testing"
)

const breakSizeSlide = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="TB"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="6400800" cy="3657600"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>
    <a:p>
      <a:br><a:rPr lang="en-US" sz="4000"/></a:br>
      <a:r><a:rPr lang="en-US" sz="1800"/><a:t>After the break</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>`

// TestReaderBreakCarriesRPrSize: the <a:rPr> inside an <a:br> is the blank
// line's font, not a text run's — the reader must record its sz on the
// break element.
func TestReaderBreakCarriesRPrSize(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(breakSizeSlide))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	var br *BreakElement
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		rt, ok := sh.(*RichTextShape)
		if !ok {
			continue
		}
		for _, para := range rt.paragraphs {
			for _, el := range para.elements {
				if b, ok := el.(*BreakElement); ok {
					br = b
				}
			}
		}
	}
	if br == nil {
		t.Fatal("break element dropped")
	}
	if br.rprSize != 4000 {
		t.Errorf("break rPrSize = %d, want 4000", br.rprSize)
	}
}

// TestBlankBreakLineTakesSpace is the end-to-end guard: a paragraph opened
// by an <a:br> renders its first text one full break-font line down (the
// blank line holds 1.2 × the br size), not one 14px stub down.
func TestBlankBreakLineTakesSpace(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	p := New()
	slide := p.GetActiveSlide()
	tb := slide.CreateRichTextShape()
	tb.BaseShape.SetOffsetX(914400).SetOffsetY(914400)
	tb.BaseShape.SetWidth(6400800).SetHeight(3657600)
	para := tb.paragraphs[0]
	para.CreateBreak().SetRPrSize(4000) // 40pt blank line
	run := para.CreateTextRun("After the break")
	f := NewFont()
	f.Size = 18
	run.SetFont(f)

	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = fc
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// First ink row inside the shape (inset top = 457200 EMU = 80px at
	// 160dpi; the blank 40pt line adds 1.2 × 40 × 2.222 ≈ 107px).
	first := -1
	for y := 60; y < img.Bounds().Dy(); y++ {
		cnt := 0
		for x := 0; x < img.Bounds().Dx(); x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			if int(r>>8) < 110 && int(g>>8) < 110 && int(b>>8) < 110 {
				cnt++
			}
		}
		if cnt >= 3 {
			first = y
			break
		}
	}
	if first < 0 {
		t.Fatal("no text ink rendered")
	}
	// Shape top 160px + inset 8px + blank 40pt line (1.2×40×2.222 ≈ 107)
	// + ~12px glyph offset ≈ 287. The old 14px floor put the text at ~194.
	if first < 240 || first > 330 {
		t.Errorf("first text at y=%d, want ~287 (shape top + inset + one blank 40pt line) — got %d",
			first, first)
	}
}

// TestWideDiagonalLineSolid: a 6pt diagonal connector fills solid — the
// parallel-Wu rasterizer left hatching (light gaps) across the band, which
// PowerPoint does not draw. Asserted through a perpendicular cut: the dark
// run must be one contiguous band, not stripes.
func TestWideDiagonalLineSolid(t *testing.T) {
	fc := NewFontCache()
	if pickInstalledFont(fc) == "" {
		t.Skip("no fonts installed on this machine")
	}
	p := New()
	slide := p.GetActiveSlide()
	ln := slide.CreateLineShape()
	ln.BaseShape.SetOffsetX(2000000).SetOffsetY(2000000)
	ln.BaseShape.SetWidth(3657600).SetHeight(3657600) // 45° diagonal
	ln.SetLineWidth(6)                                // 6pt → ~13px
	ln.SetLineColor(ColorBlack)

	opts := DefaultRenderOptions()
	opts.Width = 1600
	opts.FontCache = fc
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}

	// Strict core threshold: a solid fill gives a contiguous run of pure
	// black across the band; the legacy striped rasterizer leaves seams at
	// 60-125 luminance between the parallel Wu passes, so only isolated
	// pixels pass < 48. (At < 128 the seams read as dark and the cut looks
	// solid — measured on the 45° profile.)
	dark := func(x, y int) bool {
		r, g, b, _ := img.At(x, y).RGBA()
		return int(r>>8) < 48 && int(g>>8) < 48 && int(b>>8) < 48
	}
	var scale = 1600.0 / 9144000.0
	ox := int(2000000.0*scale + 0.5)
	oy := ox
	size := int(3657600.0*scale + 0.5)
	worst := 1 << 30
	for i := size / 4; i < 3*size/4; i += 7 {
		cx, cy := ox+i, oy+i
		best := 0
		cur := 0
		for tt := -16; tt <= 16; tt++ {
			if dark(cx+tt, cy-tt) {
				cur++
				if cur > best {
					best = cur
				}
			} else {
				cur = 0
			}
		}
		if best < worst {
			worst = best
		}
	}
	// A solid 13px band cut perpendicular at 45° spans a ~9px pure-black
	// core; the striped legacy path keeps at most 1-2 isolated core pixels
	// per cut.
	t.Logf("worst contiguous dark run: %d", worst)
	if worst < 7 {
		t.Errorf("wide diagonal line has gaps: worst contiguous dark run %d px across the band, want >= 7 (solid)", worst)
	}
}
