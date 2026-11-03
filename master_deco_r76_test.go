package gopresentation

import (
	"bytes"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// Master and layout decoration shapes — the bands, logos and footers every
// slide of a deck shares.
//
// The reader used to take a layout's pictures and its text shapes but dropped
// a plain filled rectangle, which is what a colour band is: the corpus deck
// whose master carries one <a:sp> with a B9252E solid fill rendered with the
// band missing entirely, a fifth of the slide wrong against Office's export.
// The master's non-placeholder shapes were not read at all.
//
// Decoration is also a stacking order contract: master shapes at the bottom,
// then the layout's, then the slide's own — prepend-only, never a re-sort.

// decoPackage assembles a one-slide package whose master and layout each
// carry one non-placeholder band plus one placeholder, and whose slide has a
// single text shape.
func decoPackage(t *testing.T) []byte {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="MasterBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="6248400"/><a:ext cx="9144000" cy="609600"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="B9252E"/></a:solidFill>
        <a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1325562"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
</p:sldMaster>`)
	parts["ppt/slideLayouts/slideLayout1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="LayoutBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="5524560"/><a:ext cx="9144000" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="2E3861"/></a:solidFill>
        <a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`)
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="TextBox 3"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>probe</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`)
	return buildZip(t, parts)
}

// TestMasterAndLayoutDecorationsStackUnderTheSlide reads the package above
// and pins the order: master band, layout band, then the slide's own shape,
// with neither rung's placeholders leaking in.
func TestMasterAndLayoutDecorationsStackUnderTheSlide(t *testing.T) {
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(decoPackage(t)), int64(len(decoPackage(t))))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 3 {
		t.Fatalf("slide holds %d shapes, want master band + layout band + its own text box", len(shapes))
	}
	for i, s := range shapes[:2] {
		if _, ok := s.(*PlaceholderShape); ok {
			t.Errorf("shape %d is a placeholder: a master/layout placeholder must not render as decoration", i)
		}
	}
	if _, ok := shapes[2].(*RichTextShape); !ok {
		t.Fatalf("last shape is %T, want the slide's own text box", shapes[2])
	}
}

// TestMasterBandIsDrawn renders the same package and checks the pixels the
// two bands occupy: the master's red at the slide's foot, the layout's navy
// directly above it. The reading half of this round passed for a year before
// anyone noticed the bands were never on the canvas.
func TestMasterBandIsDrawn(t *testing.T) {
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(decoPackage(t)), int64(len(decoPackage(t))))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = 800
	path := filepath.Join(t.TempDir(), "slide.png")
	if err := pres.SaveSlideAsImage(0, path, opts); err != nil {
		t.Fatalf("render: %v", err)
	}
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	img, err := png.Decode(f)
	if err != nil {
		t.Fatal(err)
	}
	w := img.Bounds().Dx()
	h := img.Bounds().Dy()
	at := func(x, y int) (int, int, int) {
		r, g, b, _ := img.At(x, y).RGBA()
		return int(r >> 8), int(g >> 8), int(b >> 8)
	}
	// Slide is 10in x 7.5in at 800px wide: the master band spans y≈547-600
	// (6248400+609600 EMU) and the layout band y≈483-523 — sample middles.
	mr, mg, mb := at(w/2, h-27)
	if mr != 0xB9 || mg != 0x25 || mb != 0x2E {
		t.Errorf("master band pixel = (%d,%d,%d), want B9252E", mr, mg, mb)
	}
	lr, lg, lb := at(w/2, h-117)
	if lr != 0x2E || lg != 0x38 || lb != 0x61 {
		t.Errorf("layout band pixel = (%d,%d,%d), want 2E3861", lr, lg, lb)
	}
}
