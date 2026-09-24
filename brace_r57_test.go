package gopresentation

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"testing"
)

// slide34 of the comparison deck caps a bullet list with a rightBrace and an
// orange run-shadowed caption. Three things were wrong, all pinned here:
//   - the 2pt brace stroke went through drawLineAA's parallel-Wu path, where
//     the middle columns took two 50% passes and over-composited to 75% ink
//     (washed-out orange instead of solid accent);
//   - the brace cast no shadow at all (the shadow switch skipped every
//     non-rect, non-ellipse shape, though a brace is pure outline and
//     PowerPoint blurs the pen trace);
//   - the run shadow blurred with the full blurPx instead of the shape law's
//     blurPx/2, spreading twice as wide as the COM gold.

const r57BraceBoxX = 500000
const r57BraceBoxY = 500000
const r57BraceBoxW = 300000
const r57BraceBoxH = 2000000

func r57Render(t *testing.T, body string) (image.Image, int, int, int, int) {
	t.Helper()
	return r57RenderWidth(t, body, 1920)
}

func r57RenderWidth(t *testing.T, body string, rw int) (image.Image, int, int, int, int) {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(body))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = rw
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	x0 := r57BraceBoxX * rw / 9144000
	y0 := r57BraceBoxY * rw / 9144000
	w := r57BraceBoxW * rw / 9144000
	h := r57BraceBoxH * rw / 9144000
	return img, x0, y0, w, h
}

// TestBraceSpineSolid: the 2pt spine must be SOLID black across its width.
// The old parallel-Wu path left the middle columns at 75% ink (rgb 64) and
// the run at half-coverage on both flanks.
func TestBraceSpineSolid(t *testing.T) {
	// 1440px puts the 2pt pen at 4px — an EVEN width, the band where the
	// old parallel-Wu passes over-composited to 75%. At 5px (odd) the old
	// path was accidentally solid and the probe would not bite.
	img, x0, y0, w, h := r57RenderWidth(t, braceShape("rightBrace", 26110, 50000), 1440)
	spineX := x0 + w/2
	midY := y0 + h/4
	if midY >= y0+h {
		t.Fatalf("bad probe row")
	}
	dark := 0
	darkest := 255
	for x := spineX - 6; x <= spineX+6; x++ {
		c := color.NRGBAModel.Convert(img.At(x, midY)).(color.NRGBA)
		if c.R < 250 { // not background
			dark++
			if int(c.R) < darkest {
				darkest = int(c.R)
			}
		}
	}
	if dark < 4 {
		t.Fatalf("spine ink run = %d columns, want >=4 (the stroke all but vanished)", dark)
	}
	if darkest > 40 {
		t.Fatalf("darkest spine pixel = %d, want pure black (the 75%% over-composite is back)", darkest)
	}
}

// TestBraceShadowRendered: a brace with an explicit outerShdw must show the
// blurred grey trace right of the spine — the shadow switch used to skip
// every non-rect/non-ellipse shape, so the brace glow was missing entirely.
func TestBraceShadowRendered(t *testing.T) {
	body := fmt.Sprintf(`
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>
    <a:prstGeom prst="rightBrace"><a:avLst><a:gd name="adj1" fmla="val 26110"/><a:gd name="adj2" fmla="val 50000"/></a:avLst></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
        <a:prstClr val="black"><a:alpha val="40000"/></a:prstClr>
      </a:outerShdw>
    </a:effectLst>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`, r57BraceBoxX, r57BraceBoxY, r57BraceBoxW, r57BraceBoxH)
	img, x0, y0, w, h := r57Render(t, body)
	_ = y0
	spineX := x0 + w/2
	penPx := 2 * 12700 * 1920 / 9144000 // 2pt in px at this scale
	// Probe right of the stroke (stroke half-width + AA). Rows stay in the
	// pure-spine band — the apex hooks reach further right and their own AA
	// would fake grey hits even without a shadow.
	found := 0
	for y := y0 + h/4; y < y0+h/2-40; y += 2 {
		for x := spineX + penPx + 3; x <= spineX+penPx+20; x++ {
			c := color.NRGBAModel.Convert(img.At(x, y)).(color.NRGBA)
			grey := abs(int(c.R)-int(c.G)) < 14 && abs(int(c.G)-int(c.B)) < 14
			if grey && c.R > 100 && c.R < 246 {
				found++
			}
		}
	}
	if found < 10 {
		t.Fatalf("grey shadow pixels right of the spine = %d, want a blurred trace (the brace cast no shadow)", found)
	}
}

// TestRunShadowBlurHalfLaw: a run shadow with blurRad=76200 (6pt) blurs with
// box radius blurPx/2 like every shape shadow. The old text path used the
// full blurPx, spreading the halo twice as wide as PowerPoint's.
func TestRunShadowBlurHalfLaw(t *testing.T) {
	body := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
    <a:p><a:r>
      <a:rPr lang="en-US" sz="4000" dirty="0">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
        <a:effectLst>
          <a:outerShdw blurRad="76200" dist="0" dir="2700000" algn="tl">
            <a:srgbClr val="000000"><a:alpha val="100000"/></a:srgbClr>
          </a:outerShdw>
        </a:effectLst>
      </a:rPr>
      <a:t>Hl</a:t>
    </a:r></a:p>
  </p:txBody>
</p:sp>`
	img, _, _, _, _ := r57Render(t, body)
	b := img.Bounds()
	// Find the rightmost dark glyph pixel on the stem row (mid-height of the
	// caps), then measure the shadow ink a fixed distance to its right.
	midY := 0
	best := 0
	for y := b.Min.Y; y < b.Max.Y; y++ {
		cnt := 0
		for x := b.Min.X; x < b.Max.X; x++ {
			c := color.NRGBAModel.Convert(img.At(x, y)).(color.NRGBA)
			if c.R < 100 {
				cnt++
			}
		}
		if cnt > best {
			best = cnt
			midY = y
		}
	}
	if best == 0 {
		t.Fatal("no glyphs rendered")
	}
	rightEdge := 0
	for x := b.Max.X - 1; x > b.Min.X; x-- {
		c := color.NRGBAModel.Convert(img.At(x, midY)).(color.NRGBA)
		if c.R < 100 {
			rightEdge = x
			break
		}
	}
	// blurRad 6pt at this scale = 16px; the half-law (radius 8) zeroes the
	// halo ~20px out, the full-radius bug keeps ink past 30px (measured
	// profile: half-law 255 by d=21, full-law still ~180-240 at d=24-32).
	near := 0
	for x := rightEdge + 3; x <= rightEdge+10 && x < b.Max.X; x++ {
		c := color.NRGBAModel.Convert(img.At(x, midY)).(color.NRGBA)
		if c.R < 235 {
			near++
		}
	}
	if near == 0 {
		t.Fatalf("no shadow ink 3-10px right of the stem (blur collapsed)")
	}
	far := 0
	for x := rightEdge + 24; x <= rightEdge+32 && x < b.Max.X; x++ {
		c := color.NRGBAModel.Convert(img.At(x, midY)).(color.NRGBA)
		if c.R < 250 {
			far++
		}
	}
	if far != 0 {
		t.Fatalf("shadow ink still present 24-32px out on %d columns: the halo spreads at the full blurPx instead of half", far)
	}
}
