package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// slide47's SmartArt body paragraphs declare <a:spcAft><a:spcPct val="15000"/>
// (15%). The reader used to drop the spcPct inside spcAft, so every body
// column's trailing/inner paragraph gaps collapsed to zero and the block
// undershot PowerPoint's export by ~9px per boundary (the COM staircase
// 0,0,0,-9,-9,-17,... measured on col2). The renderer resolves the
// percentage against the full 1.2× line — NOT the bare font size the
// declared spcBef percentage uses — pinned at 17pt: 15% × 1.2 × 17pt =
// 3.06pt = 306 hundredths; the bare model predicts 255 and is ruled out.
func TestDeclaredSpcAftPctResolvesAgainstLine(t *testing.T) {
	r := &renderer{}
	para := NewParagraph()
	run := para.CreateTextRun("H")
	run.GetFont().SetName("Calibri").SetSize(17)
	para.spaceAfterPct = 15000
	// 17pt × 1.2 × 15% = 3.06pt = 306 hundredths of a point.
	if got := r.paraSpaceAfter(para); got != 306 {
		t.Fatalf("paraSpaceAfter = %d, want 306 (15%% × 1.2 × 17pt)", got)
	}
	// The bare-font model (no 1.2 line factor) would give 255 — that is the
	// wrong resolution and the COM measurement rules it out.
	if got := r.paraSpaceAfter(para); got == 255 {
		t.Fatalf("paraSpaceAfter = 255, the bare-font model; want the 1.2× line model (306)")
	}
	// A 32pt run resolves the same 15% as 32 × 1.2 × 15% = 5.76pt = 576;
	// the bare model would give 480.
	big := NewParagraph()
	br := big.CreateTextRun("H")
	br.GetFont().SetName("Calibri").SetSize(32)
	big.spaceAfterPct = 15000
	if got := r.paraSpaceAfter(big); got != 576 {
		t.Fatalf("paraSpaceAfter on 32pt = %d, want 576 (15%% × 1.2 × 32pt)", got)
	}
	// An explicit spcPts wins and the pct is left unused.
	para.spaceAfter = 1200
	if got := r.paraSpaceAfter(para); got != 1200 {
		t.Fatalf("paraSpaceAfter with both = %d, want the declared pts 1200", got)
	}
	// Nothing declared resolves to zero.
	plain := NewParagraph()
	if got := r.paraSpaceAfter(plain); got != 0 {
		t.Fatalf("paraSpaceAfter on a plain paragraph = %d, want 0", got)
	}
}

// The reader must store a declared <a:spcAft><a:spcPct> raw on the
// paragraph — that storage is what the renderer's layout-time resolution
// consumes, so a reader drop silently flattens every declared percentage
// to zero (the slide47 undershoot). This is the read+write round trip
// mirroring TestDeclaredSpcPctIsReadFromPPr for the after-spacing half.
func TestDeclaredSpcAftPctIsReadFromPPr(t *testing.T) {
	slide := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:pPr algn="l"><a:spcAft><a:spcPct val="15000"/></a:spcAft></a:pPr><a:r><a:rPr lang="en-US" sz="1700"/><a:t>PCT</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(slide))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	rt, ok := firstShapeOfType[*RichTextShape](pres)
	if !ok {
		t.Fatal("no rich text shape found")
	}
	paras := rt.GetParagraphs()
	if len(paras) != 1 {
		t.Fatalf("got %d paragraphs, want 1", len(paras))
	}
	if paras[0].spaceAfterPct != 15000 {
		t.Fatalf("paragraph spaceAfterPct = %d, want 15000; the declared percentage was dropped on the read", paras[0].spaceAfterPct)
	}
	// The write half: the declaration survives as a spcPct, not flattened.
	out := string(zipParts(t, writeToBytes(t, pres))["ppt/slides/slide1.xml"])
	if !strings.Contains(out, `<a:spcAft><a:spcPct val="15000"/></a:spcAft>`) {
		t.Errorf("written slide lost the spcPct declaration:\n%s", out)
	}
}

// Render level: a block whose paragraphs declare a spcAft percentage is
// measurably taller than the same block without it — the resolved gap
// reaches the layout. This bites the draw/measure wiring (renderer.go
// lines 7401 and 7562), independent of the reader: the percentage is set
// on the paragraph directly. A reverted wiring leaves the gap at zero and
// the two renders come out identical (delta 0).
func TestDeclaredSpcAftPctAddsBlockHeight(t *testing.T) {
	fc := NewFontCache()
	blockBottom := func(pct int) int {
		pres := New()
		shape := pres.GetActiveSlide().CreateRichTextShape()
		shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
		shape.BaseShape.SetWidth(6000000).SetHeight(4000000)
		shape.autoFit = AutoFitNormal
		for k := 0; k < 2; k++ {
			var p *Paragraph
			if k == 0 {
				p = shape.GetParagraphs()[0]
			} else {
				p = shape.CreateParagraph()
			}
			run := p.CreateTextRun("ROW")
			run.GetFont().SetName("Arial").SetSize(17)
			run.GetFont().Color = NewColor("000000")
			p.spaceAfterPct = pct
		}
		opts := DefaultRenderOptions()
		opts.Width = 640
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		rect := emuRect(t, pres, 400000, 400000, 6000000, 4000000, 640)
		top, bottom := -1, -1
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			for x := rect.Min.X; x < rect.Max.X; x++ {
				r, g, b, a := img.At(x, y).RGBA()
				if a != 0 && (r+g+b)/3 < 20000 {
					if top < 0 {
						top = y
					}
					bottom = y
					break
				}
			}
		}
		if top < 0 {
			t.Fatal("no ink rendered")
		}
		return bottom
	}
	// The internal spcAft gap (between the two rows) pushes the second row's
	// ink down ~3px at 640px width; the trailing gap adds no ink. Either way
	// the block with the percentage is strictly taller, and clearly less than
	// a full text line (~19px), which would mean a broken factor.
	control := blockBottom(0)
	withPct := blockBottom(15000)
	delta := withPct - control
	if delta <= 0 {
		t.Fatalf("block with spcAft pct bottom = %d, control = %d, want strictly taller (gap reached layout)", withPct, control)
	}
	if delta > 12 {
		t.Fatalf("block height delta = %dpx, want under a full text line (~19px); the spcAft factor is wrong", delta)
	}
}
