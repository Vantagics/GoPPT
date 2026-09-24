package gopresentation

import (
	"image"
	"image/color"
	"testing"
)

// colorAtNRGBA reads a pixel in straight-alpha form.
func colorAtNRGBA(t *testing.T, img image.Image, px, py int) color.NRGBA {
	t.Helper()
	return color.NRGBAModel.Convert(img.At(px, py)).(color.NRGBA)
}

// renderOpts914 gives 1 px per 10000 EMU.
func renderOpts914(t *testing.T) *RenderOptions {
	t.Helper()
	opts := DefaultRenderOptions()
	opts.Width = 914
	opts.FontCache = NewFontCache()
	return opts
}

// The comparison deck's slide 32 stacks three photos in one frame and gives
// the top one prstGeom="rtTriangle": PowerPoint clips the tone-mapped image
// to the bottom-left triangle (right angle at the bottom-left corner,
// hypotenuse from the top-left to the bottom-right corner), so the diagonal
// seam between the two exposures stays visible. The picFramePoints table only
// knew snip2DiagRect and ellipse, so the photo was pasted as the full
// rectangle and the seam vanished.

// TestRtTriangleFrameClipsThePicture: a triangle-framed photo leaves the
// bounding-box region above the hypotenuse as untouched background.
func TestRtTriangleFrameClipsThePicture(t *testing.T) {
	img := renderRedPic(t, `
<p:pic>
  <p:nvPicPr><p:cNvPr id="2" name="Tri"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId101"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
    <a:prstGeom prst="rtTriangle"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`)
	// Box: (91,91)-(274,274) at 1px/10000 EMU. The hypotenuse of a square
	// rtTriangle runs corner to corner; stay well clear of it for AA.
	// Top-right region — above the hypotenuse — must stay background.
	c := colorAtNRGBA(t, img, 265, 105)
	if c.R != 255 || c.G != 255 || c.B != 255 {
		t.Errorf("above-hypotenuse pixel = rgb(%d,%d,%d), want untouched background (the image spilled outside the triangle)", c.R, c.G, c.B)
	}
	// Bottom-left corner — the right-angle side — stays the image.
	c = colorAtNRGBA(t, img, 100, 265)
	if c.R < 200 || c.G > 100 {
		t.Errorf("right-angle corner pixel = rgb(%d,%d,%d), want the red image", c.R, c.G, c.B)
	}
	// Just inside the bottom-right corner, below the hypotenuse: image.
	c = colorAtNRGBA(t, img, 265, 270)
	if c.R < 200 || c.G > 100 {
		t.Errorf("below-hypotenuse pixel = rgb(%d,%d,%d), want the red image", c.R, c.G, c.B)
	}
}

// TestRtTriangleShapeOutline pins the same vertex order on the AutoShape side
// (the shared rtTrianglePoints helper feeds both): solid black fill, corners
// probed the same way.
func TestRtTriangleShapeOutline(t *testing.T) {
	pres := readEllipseShadowSlide(t, `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Tri"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
    <a:prstGeom prst="rtTriangle"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`)
	img, err := pres.SlideToImage(0, renderOpts914(t))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// Box (91,91)-(274,274): above the hypotenuse stays white, the
	// right-angle corner below stays black.
	c := colorAtNRGBA(t, img, 265, 105)
	if c.R != 255 || c.G != 255 || c.B != 255 {
		t.Errorf("above-hypotenuse pixel = rgb(%d,%d,%d), want background", c.R, c.G, c.B)
	}
	c = colorAtNRGBA(t, img, 100, 265)
	if c.R > 60 || c.G > 60 || c.B > 60 {
		t.Errorf("right-angle corner pixel = rgb(%d,%d,%d), want the black fill", c.R, c.G, c.B)
	}
}
