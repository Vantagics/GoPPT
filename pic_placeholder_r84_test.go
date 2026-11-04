package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// A <p:pic> can carry a <p:ph>: a picture placeholder, whose geometry the
// inheritance ladder answers from the layout while its own spPr is empty.
// The corpus deck whose only shape is exactly that rendered 0x0 — invisible,
// 26% of the slide against Office's export — until the pic joined the ladder.
func TestPicPlaceholderInheritsLayoutGeometry(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slideLayouts/slideLayout1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Content 1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="1524000" y="1143000"/><a:ext cx="5486400" cy="3886200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`)
	// The baked-in package does not carry a layout rel for the base slide 1;
	// the slide below points at this layout through its rels part.
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`)
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="6" name=""/><p:cNvPicPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId1"/><a:stretch/></p:blipFill>
      <p:spPr/>
    </p:pic>
  </p:spTree></p:cSld>
</p:sld>`)
	// A 2x2 red PNG is the fill; its visible span tells the geometry apart.
	img := image.NewRGBA(image.Rect(0, 0, 2, 2))
	for y := 0; y < 2; y++ {
		for x := 0; x < 2; x++ {
			img.Set(x, y, color.RGBA{R: 200, G: 30, B: 10, A: 255})
		}
	}
	var pb bytes.Buffer
	png.Encode(&pb, img)
	parts["ppt/media/image1.png"] = pb.Bytes()
	pkg := buildZip(t, parts)

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("slide holds %d shapes, want the one picture placeholder", len(shapes))
	}
	ds, ok := shapes[0].(*DrawingShape)
	if !ok {
		t.Fatalf("shape is %T, want a DrawingShape", shapes[0])
	}
	if ds.width == 0 || ds.height == 0 {
		t.Fatalf("picture placeholder kept zero geometry: off=(%d,%d) ext=(%d,%d)",
			ds.offsetX, ds.offsetY, ds.width, ds.height)
	}
	if ds.offsetX != 1524000 || ds.offsetY != 1143000 || ds.width != 5486400 || ds.height != 3886200 {
		t.Errorf("inherited geometry off=(%d,%d) ext=(%d,%d), want the layout's",
			ds.offsetX, ds.offsetY, ds.width, ds.height)
	}

	opts := DefaultRenderOptions()
	opts.Width = 960
	path := filepath.Join(t.TempDir(), "slide.png")
	if err := pres.SaveSlideAsImage(0, path, opts); err != nil {
		t.Fatalf("render: %v", err)
	}
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	decoded, err := png.Decode(f)
	if err != nil {
		t.Fatal(err)
	}
	// 960px over a 9144000 EMU slide: the picture's left edge lands at x=160.
	r, g, b, _ := decoded.At(170, 200).RGBA()
	if r>>8 != 200 || g>>8 != 30 || b>>8 != 10 {
		t.Errorf("picture fill pixel = (%d,%d,%d), want the red fill drawn in place", r>>8, g>>8, b>>8)
	}
}
