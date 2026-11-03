package gopresentation

import (
	"bytes"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// The master's <p:clrMap> redefines the mapped colour names. bg1, tx1, bg2,
// tx2 are not theme slots — a master can point bg1 at dk2, and two corpus
// decks whose whole point is exactly that rendered with the unmapped default:
// master-bg-color (bg1→dk2) came out white instead of its dark blue, and
// onemaster-twolayouts (layout bg2→dk2) painted every slide grey instead of
// navy, both against Office's export at 100%/50% slide mismatch.
func TestMasterClrMapRemapsBackgroundColour(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    </p:spTree></p:cSld>
  <p:clrMap bg1="dk2" tx1="dk1" bg2="lt1" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
</p:sldMaster>`)
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="TextBox 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>probe</a:t></a:r></a:p></p:txBody>
</p:sp>`))
	pkg := buildZip(t, parts)

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = 100
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
	r, g, b, _ := img.At(50, 10).RGBA()
	// The master's background is schemeClr bg1, remapped to dk2 — a colour
	// nowhere near the white the unmapped name would have given.
	if r>>8 > 0x70 || g>>8 > 0x70 || b>>8 > 0x70 {
		t.Errorf("background = (%d,%d,%d), want the dark bed of the remapped bg1", r>>8, g>>8, b>>8)
	}
}
