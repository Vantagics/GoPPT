package gopresentation

import (
	"bytes"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

// A layout's clrMapOvr can carry <a:overrideClrMapping>, rebinding the colour
// names a second time under the master's own map. The treck-theme corpus deck
// has one: its section-header layout sends bg2 at dk2, so slides that use it
// put a dark background where the master's mapping would have made it light.
// Unread, the slide rendered in the master's colours — Office's own export
// disagreed with nearly the whole slide.
func TestLayoutClrMapOverrideRebindsColourNames(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	// The master maps bg2 to the light slot; the layout overrides it back to
	// the dark one, and the slide fills its whole canvas with bg2.
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
</p:sldMaster>`)
	parts["ppt/slideLayouts/slideLayout1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>
</p:sldLayout>`)
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg2"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`)
	pkg := buildZip(t, parts)

	resolved := make(map[string]string)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	// Sanity for the read: the slide's own background is bg2, which the
	// layout's override pointed at the dark slot — the render below pins it.
	if bg := pres.GetAllSlides()[0].GetBackground(); bg == nil {
		t.Fatal("no background resolved")
	}
	_ = resolved

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
	r, g, b, _ := img.At(50, 30).RGBA()
	// The default theme's lt2 is a light fill and dk2 is dark: without the
	// layout's override the slide paints light. The override wins.
	if r>>8 > 0x70 || g>>8 > 0x70 || b>>8 > 0x70 {
		t.Errorf("background = (%d,%d,%d), want the dark bed of the overridden bg2", r>>8, g>>8, b>>8)
	}
}
