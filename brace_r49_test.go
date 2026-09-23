package gopresentation

import (
	"bytes"
	"fmt"
	"image"
	"math"
	"testing"
)

// The brace presets follow the ECMA-376 preset path: quarter-ellipse hooks
// from the box corners to the centre spine at x=w/2, and a middle apex made
// of two quarter ellipses meeting in a cusp at the outer edge — (w, y3) for
// rightBrace, (0, y3) for leftBrace — with y1 = ss·adj1/100000 the hook
// radius (pinned so the spine never inverts) and y3 = h·adj2/100000 the apex
// centre.

const braceBoxX = 500000
const braceBoxY = 500000
const braceBoxW = 300000
const braceBoxH = 2000000

func braceShape(prst string, adj1, adj2 int) string {
	av := "<a:avLst/>"
	if adj1 >= 0 {
		av = fmt.Sprintf(`<a:avLst><a:gd name="adj1" fmla="val %d"/><a:gd name="adj2" fmla="val %d"/></a:avLst>`, adj1, adj2)
	}
	return fmt.Sprintf(`
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>
    <a:prstGeom prst="%s">%s</a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
  <p:style>
    <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
  </p:style>
</p:sp>`, braceBoxX, braceBoxY, braceBoxW, braceBoxH, prst, av)
}

func renderBraceShape(t *testing.T, body string) (img image.Image, x0, y0, wpx, hpx int) {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(body))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = 640
	opts.FontCache = NewFontCache()
	img, err = pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rect := emuRect(t, pres, braceBoxX, braceBoxY, braceBoxW, braceBoxH, opts.Width)
	return img, rect.Min.X, rect.Min.Y, rect.Dx(), rect.Dy()
}

func braceInk(img image.Image) func(int, int) bool {
	return func(x, y int) bool {
		r, g, b, a := img.At(x, y).RGBA()
		return a != 0 && (r+g+b)/3 < 20000
	}
}

// A rightBrace with slide19's adjustments (adj1=37626, adj2=54714) strokes
// the centre spine, reaches a cusp at the right edge, and keeps both corner
// hooks — with the interior off the path empty.
func TestRightBraceAdjValuesStrokePresetPath(t *testing.T) {
	img, x0, y0, wpx, hpx := renderBraceShape(t, braceShape("rightBrace", 37626, 54714))
	ink := braceInk(img)
	if !anyInkAround(img, x0+wpx/2, y0+hpx/4, ink) {
		t.Errorf("no ink on the centre spine at quarter height")
	}
	if !anyInkAround(img, x0+wpx, y0+hpx*54714/100000, ink) {
		t.Errorf("no ink at the apex cusp (x=%d, y=%d)", x0+wpx, y0+hpx*54714/100000)
	}
	if !anyInkAround(img, x0, y0, ink) {
		t.Errorf("no hook ink at the top-left corner")
	}
	if !anyInkAround(img, x0, y0+hpx, ink) {
		t.Errorf("no hook ink at the bottom-left corner")
	}
	probe := x0 + wpx - 2
	for dx := -2; dx <= 2; dx++ {
		for dy := -2; dy <= 2; dy++ {
			if ink(probe+dx, y0+hpx/4+dy) {
				t.Fatalf("ink at (x=%d, y=%d): the brace was filled or boxed", probe, y0+hpx/4)
			}
		}
	}
}

// A leftBrace is the mirror: cusp at the LEFT edge, spine at the centre,
// hooks at the right-hand corners.
func TestLeftBraceMirrorsTheRightBrace(t *testing.T) {
	img, x0, y0, wpx, hpx := renderBraceShape(t, braceShape("leftBrace", -1, -1))
	ink := braceInk(img)
	if !anyInkAround(img, x0+wpx/2, y0+hpx/4, ink) {
		t.Errorf("no ink on the centre spine at quarter height")
	}
	if !anyInkAround(img, x0, y0+hpx/2, ink) {
		t.Errorf("no ink at the apex cusp (x=%d, y=%d)", x0, y0+hpx/2)
	}
	if !anyInkAround(img, x0+wpx, y0, ink) {
		t.Errorf("no hook ink at the top-right corner")
	}
	if !anyInkAround(img, x0+wpx, y0+hpx, ink) {
		t.Errorf("no hook ink at the bottom-right corner")
	}
	probe := x0 + 2
	for dx := -2; dx <= 2; dx++ {
		for dy := -2; dy <= 2; dy++ {
			if ink(probe+dx, y0+hpx/4+dy) {
				t.Fatalf("ink at (x=%d, y=%d): the brace was filled or boxed", probe, y0+hpx/4)
			}
		}
	}
}

// adj1 is pinned to maxAdj1 = min(1−adj2, adj2)/2·h/ss: an out-of-range value
// must not invert the spine or push the hooks past the apex.
func TestBraceAdj1PinnedToMaxAdj1(t *testing.T) {
	// adj2=50000 -> maxAdj1 = 25000·h/ss = 25000·(2000000/300000) ≈ 166667;
	// ask for 300000 and the shape must still render sanely.
	img, x0, y0, wpx, hpx := renderBraceShape(t, braceShape("rightBrace", 300000, 50000))
	ink := braceInk(img)
	ss := math.Min(float64(wpx), float64(hpx))
	y1 := ss * math.Min(25000*float64(hpx)/ss, 300000) / 100000.0
	if y1 >= float64(hpx)/2 {
		t.Fatalf("pin law violated: y1=%.1f not below half height", y1)
	}
	if !anyInkAround(img, x0+wpx/2, y0+hpx/4, ink) {
		t.Errorf("no ink on the centre spine at quarter height")
	}
	if !anyInkAround(img, x0+wpx, y0+hpx/2, ink) {
		t.Errorf("no ink at the apex cusp")
	}
}

// adj2 is pinned to 0..100000 and y3 rides it: adj2=90000 puts the apex cusp
// at 90% height, not mid-height.
func TestBraceAdj2MovesTheApex(t *testing.T) {
	img, x0, y0, wpx, hpx := renderBraceShape(t, braceShape("rightBrace", 26110, 90000))
	ink := braceInk(img)
	if !anyInkAround(img, x0+wpx, y0+hpx*90/100, ink) {
		t.Errorf("no ink at the apex cusp at 90%% height (x=%d, y=%d)", x0+wpx, y0+hpx*90/100)
	}
	// and the old mid-height cusp position is empty — no apex left there.
	for dx := -2; dx <= 2; dx++ {
		for dy := -2; dy <= 2; dy++ {
			if ink(x0+wpx+dx, y0+hpx/2+dy) {
				t.Fatalf("ink at mid-height right edge: the apex did not move with adj2")
			}
		}
	}
}
