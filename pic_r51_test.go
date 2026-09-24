package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"strings"
	"testing"
)

// softEdgeSlide carries the comparison deck's slide 40 cylinder declaration:
// an ellipse-framed photo whose spPr effectLst feathers the edge over
// rad="112500".
const softEdgeSlide = `
<p:pic>
  <p:nvPicPr><p:cNvPr id="2" name="Cyl"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId101" cstate="print"/>
    <a:srcRect/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr bwMode="auto">
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:ln><a:noFill/></a:ln>
    <a:effectLst><a:softEdge rad="914400"/></a:effectLst>
  </p:spPr>
</p:pic>`

// TestReaderParsesSoftEdge: the feather radius survives the read.
func TestReaderParsesSoftEdge(t *testing.T) {
	pres := readPicFixture(t, softEdgeSlide)
	d := findDrawing(pres.GetAllSlides()[0].GetShapes())
	if d == nil {
		t.Fatal("the picture was not read as a DrawingShape")
	}
	if got := d.GetSoftEdge(); got != 914400 {
		t.Errorf("softEdge rad = %d, want 914500-1 (the declared 914400)", got)
	}
	if d.presetGeom != "ellipse" {
		t.Errorf("presetGeom = %q, want ellipse", d.presetGeom)
	}
}

// TestWriterRoundTripsSoftEdge: the feather rides back out in the pic's
// spPr effectLst.
func TestWriterRoundTripsSoftEdge(t *testing.T) {
	pres := readPicFixture(t, softEdgeSlide)
	parts := zipParts(t, writeToBytes(t, pres))
	xml := string(parts["ppt/slides/slide1.xml"])
	if !strings.Contains(xml, `<a:softEdge rad="914400"/>`) {
		t.Errorf("written slide XML lacks the softEdge element:\n%s", xml)
	}
}

// renderRedPic renders a deck whose single picture is a solid red image in
// the given frame, on the white slide background.
func renderRedPic(t *testing.T, picXML string) image.Image {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	var buf bytes.Buffer
	src := image.NewRGBA(image.Rect(0, 0, 4, 4))
	for y := 0; y < 4; y++ {
		for x := 0; x < 4; x++ {
			src.SetRGBA(x, y, color.RGBA{R: 255, A: 255})
		}
	}
	if err := png.Encode(&buf, src); err != nil {
		t.Fatalf("encode fixture: %v", err)
	}
	parts["ppt/media/image1.png"] = buf.Bytes()
	rels := string(parts["ppt/slides/_rels/slide1.xml.rels"])
	rels = strings.Replace(rels, "</Relationships>",
		`<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`, 1)
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(rels)
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(picXML))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = 914 // 1 px per 10000 EMU
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	return img
}

// TestEllipseFrameClipsThePicture: an ellipse-framed photo shows the
// background at its bounding-box corners — the image used to be pasted as
// the full rectangle, spilling well outside the oval.
func TestEllipseFrameClipsThePicture(t *testing.T) {
	img := renderRedPic(t, `
<p:pic>
  <p:nvPicPr><p:cNvPr id="2" name="Oval"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId101"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`)
	// Box: (91,91)-(274,274) at 1px/10000 EMU. Corner inside the box but
	// outside the oval must stay background.
	c := color.NRGBAModel.Convert(img.At(100, 100)).(color.NRGBA)
	if c.R != 255 || c.G != 255 || c.B != 255 {
		t.Errorf("box corner pixel = rgb(%d,%d,%d), want untouched background (the image spilled outside the ellipse)", c.R, c.G, c.B)
	}
	// Centre stays the image.
	c = color.NRGBAModel.Convert(img.At(182, 182)).(color.NRGBA)
	if c.R < 200 || c.G > 100 {
		t.Errorf("centre pixel = rgb(%d,%d,%d), want the red image", c.R, c.G, c.B)
	}
}

// TestSoftEdgeFeathersTheFrameEdge: the render sentinel. With rad equal to
// half the box the centre stays fully opaque while the rim fades to the
// background — the picture used to be cut out hard at the frame line.
func TestSoftEdgeFeathersTheFrameEdge(t *testing.T) {
	img := renderRedPic(t, softEdgeSlide)
	// Box: (91,91)-(274,274), rad 914400 EMU = 91.4px = half the box. The
	// centre sits a full radius in: fully opaque.
	c := color.NRGBAModel.Convert(img.At(182, 182)).(color.NRGBA)
	if c.R < 200 || c.G > 100 {
		t.Errorf("centre pixel = rgb(%d,%d,%d), want opaque red (a full radius in, the feather is done)", c.R, c.G, c.B)
	}
	// Ten pixels in from the left edge: d/rad ~0.11, smoothstep ~0.03 —
	// essentially background. (The red image keeps R=255 over white's 255,
	// so the fade shows in G/B.)
	c = color.NRGBAModel.Convert(img.At(102, 182)).(color.NRGBA)
	if c.G < 200 {
		t.Errorf("near-edge pixel = rgb(%d,%d,%d), want the feather to have faded the image out", c.R, c.G, c.B)
	}
	// Quarter radius in: d/rad = 0.25. The COM-calibrated smoothstep puts
	// alpha at 0.156 (G≈215); a linear ramp would give 0.25 (G≈191) — the
	// pin separates the two.
	c = color.NRGBAModel.Convert(img.At(114, 182)).(color.NRGBA)
	if c.G < 205 {
		t.Errorf("quarter-radius pixel = rgb(%d,%d,%d), want smoothstep alpha ~0.16 (G≈215)", c.R, c.G, c.B)
	}
}
