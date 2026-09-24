package gopresentation

import (
	"bytes"
	"fmt"
	"image/color"
	"testing"
)

// A no-fill ellipse whose <a:effectLst> carries a white (bg1) outer shadow.
// PowerPoint casts the shadow of the stroke itself — the white glow around a
// stroked callout oval — so the renderer must rasterise the ring silhouette
// instead of skipping the shadow like it does for other non-rectangles.
const ellipseShadowShape = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="O"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="400000"/><a:ext cx="2400000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" rotWithShape="0">
        <a:schemeClr val="bg1"><a:alpha val="75000"/></a:schemeClr>
      </a:outerShdw>
    </a:effectLst>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`

func readEllipseShadowSlide(t *testing.T, shapes string) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(shapes))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func TestNoFillEllipseShadowParsed(t *testing.T) {
	pres := readEllipseShadowSlide(t, ellipseShadowShape)
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("shape count = %d, want 1", len(shapes))
	}
	s, ok := shapes[0].(*AutoShape)
	if !ok {
		t.Fatalf("shape type %T, want an AutoShape", shapes[0])
	}
	if s.shadow == nil || !s.shadow.Visible {
		t.Fatalf("the ellipse's outerShdw was not attached to the shape")
	}
	// The alpha case bakes the opacity into the ARGB prefix as well (75% →
	// 0xBF); the RGB half must still be bg1 white.
	if s.shadow.Color.ARGB[2:] != "FFFFFF" {
		t.Fatalf("shadow color = %s, want bg1 white", s.shadow.Color.ARGB)
	}
	if s.shadow.Alpha != 75 {
		t.Fatalf("shadow alpha = %d, want 75", s.shadow.Alpha)
	}
}

func TestNoFillEllipseShadowRendersRingGlow(t *testing.T) {
	// The slide background is black so the white glow is measurable.
	withBg := `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
		ellipseShadowShape
	pres := readEllipseShadowSlide(t, withBg)
	opts := DefaultRenderOptions()
	opts.Width = 1280
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rect := emuRect(t, pres, 500000, 400000, 2400000, 1200000, opts.Width)
	at := func(x, y int) color.RGBA {
		r, g, b, a := img.At(x, y).RGBA()
		return color.RGBA{uint8(r >> 8), uint8(g >> 8), uint8(b >> 8), uint8(a >> 8)}
	}
	cx := rect.Min.X + rect.Dx()/2
	// The ring silhouette keeps the ellipse interior dark: a disc silhouette
	// would smear white across the centre through the blur.
	centre := at(cx, rect.Min.Y+rect.Dy()/2)
	if centre.R > 20 {
		t.Fatalf("ellipse centre rgb(%d,%d,%d) is bright — the shadow used a disc, not the stroke ring",
			centre.R, centre.G, centre.B)
	}
	// Just outside the bottom edge the blurred ring glows.
	below := at(cx, rect.Max.Y+3)
	if below.R < 20 || below.G < 20 || below.B < 20 {
		t.Fatalf("below the ellipse rgb(%d,%d,%d) shows no white glow", below.R, below.G, below.B)
	}
	// The shadow offset (dist 38100 EMU at 45° shifts the glow down-right):
	// the bottom edge glows brighter than the top edge.
	above := at(cx, rect.Min.Y-3)
	if below.R <= above.R {
		t.Fatalf("glow below (%d) is not brighter than above (%d) — the shadow offset is lost",
			below.R, above.R)
	}
}

func TestEllipseWithoutEffectHasNoGlow(t *testing.T) {
	plain := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="O"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="400000"/><a:ext cx="2400000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
	withBg := `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` + plain
	pres := readEllipseShadowSlide(t, withBg)
	opts := DefaultRenderOptions()
	opts.Width = 1280
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rect := emuRect(t, pres, 500000, 400000, 2400000, 1200000, opts.Width)
	r, g, b, _ := img.At(rect.Min.X+rect.Dx()/2, rect.Max.Y+4).RGBA()
	if r>>8 > 8 || g>>8 > 8 || b>>8 > 8 {
		t.Fatalf("no-effect ellipse glows below the edge: rgb(%d,%d,%d)", r>>8, g>>8, b>>8)
	}
}

func TestLnWidthInheritsFromLnRef(t *testing.T) {
	// An explicit <a:ln> without w keeps the style reference's width: the
	// default theme's lnStyleLst idx=2 is 12700 EMU (1pt). A declared w must
	// win over the reference.
	const tmpl = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="O"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="400000"/><a:ext cx="2400000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln%s><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
</p:sp>`
	pres := readEllipseShadowSlide(t, fmt.Sprintf(tmpl, ""))
	shapes := pres.GetAllSlides()[0].GetShapes()
	s, ok := shapes[0].(*AutoShape)
	if !ok {
		t.Fatalf("shape type %T, want an AutoShape", shapes[0])
	}
	if s.border == nil || s.border.Width != 1 {
		got := 0
		if s.border != nil {
			got = s.border.Width
		}
		t.Fatalf("inherited border width = %dpt, want 1pt from lnRef idx=2 (12700 EMU)", got)
	}
	pres = readEllipseShadowSlide(t, fmt.Sprintf(tmpl, ` w="25400"`))
	shapes = pres.GetAllSlides()[0].GetShapes()
	s, ok = shapes[0].(*AutoShape)
	if !ok {
		t.Fatalf("shape type %T, want an AutoShape", shapes[0])
	}
	if s.border == nil || s.border.Width != 2 {
		got := 0
		if s.border != nil {
			got = s.border.Width
		}
		t.Fatalf("declared border width = %dpt, want 2pt from the ln's own w", got)
	}
}
