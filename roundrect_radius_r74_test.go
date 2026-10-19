package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"testing"
)

// roundRect corner radius must follow the OOXML preset: x1 = ss*a/100000 with
// a = pin(0, adj, 50000). The old renderer divided by 200000, halving every
// rounded corner. customGeo slide10 node 2 is a 392px roundRect with adj=10000;
// PowerPoint COM draws a ~39px corner, the old code drew ~19px. The fix is the
// single source of truth roundedRectRadius(); all four render sites call it.

func TestRoundedRectRadiusDivisor(t *testing.T) {
	// adj=10000 on a 392x392 node (customGeo slide10 node 2 @1600px):
	// OOXML radius = 392 * 10000 / 100000 = 39px.
	const w, h, adj = 392, 392, 10000
	got := roundedRectRadius(w, h, map[string]int{"adj": adj})
	want := minInt(w, h) * adj / 100000
	if got != want {
		t.Fatalf("roundedRectRadius(392,392,adj=10000) = %d, want %d (OOXML x1 = ss*a/100000)", got, want)
	}
	if got != 39 {
		t.Fatalf("corner must be 39px for adj=10000 on a 392px node, got %d", got)
	}
}

func TestRoundedRectRadiusDefault(t *testing.T) {
	// No adj override -> default 16667 -> min*0.16667.
	const w, h = 600, 400
	got := roundedRectRadius(w, h, nil)
	want := minInt(w, h) * 16667 / 100000
	if got != want {
		t.Fatalf("default roundedRectRadius = %d, want %d", got, want)
	}
}

func TestRoundedRectCornerMatchesCOM(t *testing.T) {
	// Build a slide with one red roundRect (adj=10000), render it, and sample a
	// point in the corner that is BACKGROUND for the correct (large) radius but
	// FILL for the old half-radius. A point 7px from the left and 5px from the
	// top edge, on a 350x262px (at 800px width) roundRect, sits inside the old
	// 13px corner arc but outside the correct 26px arc — so it is red when the
	// divisor is wrong and white when it is right.
	shape := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="RR"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="400000" y="400000"/><a:ext cx="4000000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(shape))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read package: %v", err)
	}
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("shape count = %d, want 1", len(shapes))
	}
	as, ok := shapes[0].(*AutoShape)
	if !ok {
		t.Fatalf("shape type %T, want *AutoShape", shapes[0])
	}
	if as.shapeType != AutoShapeRoundedRect {
		t.Fatalf("shapeType = %v, want AutoShapeRoundedRect", as.shapeType)
	}
	if as.adjustValues["adj"] != 10000 {
		t.Fatalf("parsed adj = %d, want 10000", as.adjustValues["adj"])
	}

	opts := DefaultRenderOptions()
	opts.Width = 800
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rect := emuRect(t, pres, 400000, 400000, 4000000, 3000000, opts.Width)
	at := func(x, y int) color.RGBA {
		r, g, b, a := img.At(x, y).RGBA()
		return color.RGBA{uint8(r >> 8), uint8(g >> 8), uint8(b >> 8), uint8(a >> 8)}
	}
	// Slide background, well away from the shape.
	bg := at(2, 2)
	// Corner discriminator: 7px from the shape's left edge, 5px from its top.
	cx := rect.Min.X + 7
	cy := rect.Min.Y + 5
	if !rect.Overlaps(image.Rect(cx, cy, cx+1, cy+1)) {
		t.Fatalf("discriminator (%d,%d) outside shape rect %v", cx, cy, rect)
	}
	corner := at(cx, cy)
	// Correct radius: the corner is cut away, so it matches the slide background.
	if absDiff(int(corner.R), int(bg.R)) > 8 || absDiff(int(corner.G), int(bg.G)) > 8 || absDiff(int(corner.B), int(bg.B)) > 8 {
		t.Fatalf("corner (%d,%d) = rgb(%d,%d,%d), bg = rgb(%d,%d,%d): the corner radius is too small (old /200000 divisor); PowerPoint COM cuts this point away",
			cx, cy, corner.R, corner.G, corner.B, bg.R, bg.G, bg.B)
	}
	// Sanity: the shape centre is filled red.
	centre := at(rect.Min.X+rect.Dx()/2, rect.Min.Y+rect.Dy()/2)
	if centre.R < 200 || centre.G > 60 || centre.B > 60 {
		t.Fatalf("shape centre = rgb(%d,%d,%d), want saturated red", centre.R, centre.G, centre.B)
	}
}
