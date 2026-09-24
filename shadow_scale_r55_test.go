package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"strings"
	"testing"
)

// The HDR callout ovals carry <a:outerShdw ... sx="105000" sy="105000"
// algn="tl">: the shadow silhouette is 5% larger than the shape and anchored
// at its top-left, so the white glow grows right and down (profile-verified
// against the COM export). The reader must keep those attributes and the
// writer must re-emit them, or a saved deck loses the halo's asymmetry.
const shadowScaleShape = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="O"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="400000"/><a:ext cx="2400000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" sx="105000" sy="105000" algn="tl" rotWithShape="0">
        <a:schemeClr val="bg1"><a:alpha val="75000"/></a:schemeClr>
      </a:outerShdw>
    </a:effectLst>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`

func TestShadowScaleAndAlignParsed(t *testing.T) {
	pres := readEllipseShadowSlide(t, shadowScaleShape)
	shapes := pres.GetAllSlides()[0].GetShapes()
	s, ok := shapes[0].(*AutoShape)
	if !ok {
		t.Fatalf("shape type %T, want an AutoShape", shapes[0])
	}
	if s.shadow == nil || !s.shadow.Visible {
		t.Fatalf("the outerShdw was not attached")
	}
	if s.shadow.ScaleX != 105 || s.shadow.ScaleY != 105 {
		t.Fatalf("shadow scale = %d/%d, want 105/105", s.shadow.ScaleX, s.shadow.ScaleY)
	}
	if s.shadow.Algn != "tl" {
		t.Fatalf("shadow algn = %q, want tl", s.shadow.Algn)
	}
}

func TestShapeShadowRoundTripsThroughWriter(t *testing.T) {
	pres := readEllipseShadowSlide(t, shadowScaleShape)
	w, err := NewWriter(pres, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("new writer: %v", err)
	}
	var buf bytes.Buffer
	if err := w.WriteTo(&buf); err != nil {
		t.Fatalf("write: %v", err)
	}
	parts := zipParts(t, buf.Bytes())
	out := string(parts["ppt/slides/slide1.xml"])
	if !strings.Contains(out, "<a:effectLst>") {
		t.Fatal("the saved slide lost the shape's effectLst entirely")
	}
	for _, want := range []string{`sx="105000"`, `sy="105000"`, `algn="tl"`} {
		if !strings.Contains(out, want) {
			t.Fatalf("the saved slide lost %s:\n%s", want, out)
		}
	}
}

// The ellipse stroke used to measure pixel distance to the ellipse with the
// SMALLER radius, which fattened the left/right vertex arcs by rx/ry (8px vs
// the COM gold's 4px on a 2pt ring). The horizontal run across the left
// vertex and the vertical run across the top vertex must now both match the
// stroke width within AA tolerance.
func TestEllipseStrokeWidthUniformAroundRing(t *testing.T) {
	pres := readEllipseShadowSlide(t, `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="O"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="400000"/><a:ext cx="2400000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`)
	img, err := pres.SlideToImage(0, &RenderOptions{Width: 1920})
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	b := img.Bounds()
	dark := func(px, py int) bool {
		c := color.RGBAModel.Convert(img.At(px, py)).(color.RGBA)
		return int(c.R)+int(c.G)+int(c.B) < 150
	}
	// locate the ring: dark pixels of the oval on a white canvas
	var minX, maxX, minY, maxY int
	minX, minY = 1<<30, 1<<30
	for py := b.Min.Y; py < b.Max.Y; py++ {
		for px := b.Min.X; px < b.Max.X; px++ {
			if dark(px, py) {
				if px < minX {
					minX = px
				}
				if px > maxX {
					maxX = px
				}
				if py < minY {
					minY = py
				}
				if py > maxY {
					maxY = py
				}
			}
		}
	}
	if minX > maxX {
		t.Fatal("nothing rendered")
	}
	runH := 0 // horizontal run across the left vertex (vertical tangent)
	cy := (minY + maxY) / 2
	for px := minX; dark(px, cy) && px <= maxX; px++ {
		runH++
	}
	runV := 0 // vertical run across the top vertex (horizontal tangent)
	cx := (minX + maxX) / 2
	for py := minY; dark(cx, py) && py <= maxY; py++ {
		runV++
	}
	if runH == 0 || runV == 0 {
		t.Fatalf("ring not found: runH=%d runV=%d", runH, runV)
	}
	// 2pt at the default render scale ~4-5px; the old bug made runH ~2x runV.
	if runH > runV+2 {
		t.Fatalf("left-vertex stroke %dpx vs top-vertex %dpx: the ellipse metric still fattens the side arcs", runH, runV)
	}
	if runH < 2 || runH > 9 {
		t.Fatalf("stroke run %dpx out of the expected 2pt band", runH)
	}
	_ = image.Point{}
}
