package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"testing"
)

// renderXMLSlide builds a one-slide deck from a raw shape XML fragment and
// renders it, returning the image.
func renderXMLSlide(t *testing.T, shapeXML string, width int) image.Image {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(shapeXML))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = width
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	return img
}

// Rotated and flipped shapes composite through an offscreen buffer whose
// pixels are alpha-premultiplied (blendPixel multiplies by alpha when writing
// over the transparent canvas, as does image/draw). The composite back onto
// the slide must unpremultiply before its straight-alpha blend, or the alpha
// is applied twice: a 50%-white fill rendered as 75% grey (slide39's "Window"
// bar). These tests pin single-alpha blending through both paths.

func TestRotatedAlphaFillBlendsOnce(t *testing.T) {
	// A 90°-rotated rect filled with red at 50% alpha over an opaque black
	// base rect. Straight blending: red@50% over black = 50% red (127,0,0).
	// The double-alpha bug produced 25% red (64,0,0).
	shape := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Base"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="3240000" y="1828800"/><a:ext cx="2664000" cy="3200400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Tint"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="5400000"><a:off x="3543320" y="2128800"/><a:ext cx="2063360" cy="1600200"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
	img := renderXMLSlide(t, shape, 914)
	// The rotated rect's centre is its unrotated centre.
	cx, cy := (3543320+2063360/2)*914/9144000, (2128800+1600200/2)*914/9144000
	c := color.NRGBAModel.Convert(img.At(cx, cy)).(color.NRGBA)
	if c.R < 110 || c.G > 20 || c.B > 20 {
		t.Errorf("rotated 50%% red over black: got {%d %d %d}, want ~(127,0,0) — the alpha was applied twice", c.R, c.G, c.B)
	}
}

func TestFlippedAlphaFillBlendsOnce(t *testing.T) {
	// Same pin through the flip-only composite path (flipH, no rotation).
	shape := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Base"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="3240000" y="1828800"/><a:ext cx="2664000" cy="3200400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Tint"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm flipH="1"><a:off x="3543320" y="2128800"/><a:ext cx="2063360" cy="1600200"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
	img := renderXMLSlide(t, shape, 914)
	cx, cy := (3543320+2063360/2)*914/9144000, (2128800+1600200/2)*914/9144000
	c := color.NRGBAModel.Convert(img.At(cx, cy)).(color.NRGBA)
	if c.R < 110 || c.G > 20 || c.B > 20 {
		t.Errorf("flipped 50%% red over black: got {%d %d %d}, want ~(127,0,0) — the alpha was applied twice", c.R, c.G, c.B)
	}
}

func TestRotatedAlphaFillOverWhiteStaysWhite(t *testing.T) {
	// The slide39 pin: 50% white over the white slide must stay pure white
	// (255), not the 191 grey the double-alpha bug painted.
	shape := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="16200000"><a:off x="3543320" y="2128800"/><a:ext cx="2063360" cy="400110"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="50000"/></a:srgbClr></a:solidFill>
    <a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
	img := renderXMLSlide(t, shape, 914)
	cx, cy := (3543320+2063360/2)*914/9144000, (2128800+400110/2)*914/9144000
	c := color.NRGBAModel.Convert(img.At(cx+10, cy)).(color.NRGBA)
	if c.R != 255 || c.G != 255 || c.B != 255 {
		t.Errorf("rotated 50%% white over white: got {%d %d %d}, want pure white — the alpha was applied twice", c.R, c.G, c.B)
	}
}
