package gopresentation

import (
	"bytes"
	"image"
	"strings"
	"testing"
)

// masterBgPackage assembles the New() package with its slide master replaced
// by one whose <p:cSld> opens with the given <p:bg> body (or none), plus an
// optional image part wired through the master's own rels. The default layout
// and slide carry no <p:bg>, so the background the reader lands on is exactly
// the rung under test.
func masterBgPackage(t *testing.T, bgBody string, withImage bool) []byte {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	master := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>` + bgBody + `<p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
</p:sldMaster>`
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(master)
	if withImage {
		parts["ppt/media/image1.png"] = twoTonePNG(t)
		rels := string(parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"])
		imgRel := `<Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`
		rels = strings.Replace(rels, "</Relationships>", imgRel+"</Relationships>", 1)
		parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = []byte(rels)
	}
	return buildZip(t, parts)
}

func readMasterBgSlide(t *testing.T, pkg []byte) *Presentation {
	t.Helper()
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

// TestMasterSolidBackgroundInheritsToSlide covers the missing rung of the
// background ladder: a deck that paints its <p:bg> only on the slide master
// used to render white, because only the layout rung was ever read. The
// assertion is a render-level pixel count, not just the model field.
func TestMasterSolidBackgroundInheritsToSlide(t *testing.T) {
	pkg := masterBgPackage(t,
		`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`, false)
	pres := readMasterBgSlide(t, pkg)

	imgs, err := pres.SlidesToImages(goldenOptions(NewFontCache()))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	if len(imgs) != 1 {
		t.Fatalf("rendered %d slides, want 1", len(imgs))
	}
	want := argbToRGBA(NewColor("FF1F4E79"))
	if n := countNearColorIn(imgs[0], imgs[0].Bounds(), want, 8); n == 0 {
		t.Error("no pixel of the master background colour FF1F4E79: the master background rung was not drawn")
	}
}

// TestLayoutRungStillBeatsMaster pins the ladder order: a layout background
// must win over the master one, exactly as it already won over nothing.
func TestLayoutRungStillBeatsMaster(t *testing.T) {
	pkg := masterBgPackage(t,
		`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`, false)
	parts := zipParts(t, pkg)
	layout := string(parts["ppt/slideLayouts/slideLayout1.xml"])
	i := strings.Index(layout, "<p:cSld")
	if i < 0 {
		t.Fatal("the default layout has no <p:cSld> to patch")
	}
	j := strings.Index(layout[i:], ">")
	layout = layout[:i+j+1] + `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="7F1D1D"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` + layout[i+j+1:]
	parts["ppt/slideLayouts/slideLayout1.xml"] = []byte(layout)
	pres := readMasterBgSlide(t, buildZip(t, parts))

	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("get slide: %v", err)
	}
	if slide.background == nil {
		t.Fatal("slide background stayed nil with a layout background present")
	}
	if got := slide.background.Color.ARGB; got != "FF7F1D1D" {
		t.Errorf("slide background = %s, want FF7F1D1D from the layout rung", got)
	}
}

// TestMasterBackgroundImageIsPrepended covers the picture form of the same
// rung: a blipFill <p:bg> on the master becomes the full-slide drawing,
// resolved through the MASTER's rels (the layout rels handed to readMaster
// would 404 the image part).
func TestMasterBackgroundImageIsPrepended(t *testing.T) {
	pkg := masterBgPackage(t,
		`<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`, true)
	pres := readMasterBgSlide(t, pkg)

	imgs, err := pres.SlidesToImages(goldenOptions(NewFontCache()))
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	if len(imgs) != 1 {
		t.Fatalf("rendered %d slides, want 1", len(imgs))
	}
	img := imgs[0]
	b := img.Bounds()
	blue := argbToRGBA(NewColor("FF0000FF"))
	red := argbToRGBA(NewColor("FFFF0000"))
	left := image.Rect(b.Min.X, b.Min.Y, b.Min.X+b.Dx()/4, b.Min.Y+b.Dy())
	right := image.Rect(b.Min.X+b.Dx()*3/4, b.Min.Y, b.Max.X, b.Max.Y)
	if n := countNearColorIn(img, left, red, 8); n == 0 {
		t.Error("no red in the left quarter: the master background picture was not drawn")
	}
	if n := countNearColorIn(img, right, blue, 8); n == 0 {
		t.Error("no blue in the right quarter: the master background picture was not drawn")
	}
}

// TestBgRefResolvesThroughThemeFillStyle pins the bgRef form: idx 1001 names
// the FIRST body of the theme's fillStyleLst with the schemeClr inside as
// phClr. The New() theme's accent1 is 4472C4 and its first fill style is a
// solid phClr, so bgRef idx="1001" on accent1 must land on FF4472C4 rather
// than falling through to white. (Both comparison decks carry exactly this
// markup on their masters, with bg1 — the coincidence that hid the gap.)
func TestBgRefResolvesThroughThemeFillStyle(t *testing.T) {
	pkg := masterBgPackage(t,
		`<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>`, false)
	pres := readMasterBgSlide(t, pkg)

	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("get slide: %v", err)
	}
	if slide.background == nil {
		t.Fatal("bgRef idx=1001 on accent1 resolved to no fill; want the theme's first fill style")
	}
	if got := slide.background.Color.ARGB; got != "FF4472C4" {
		t.Errorf("bgRef background = %s, want FF4472C4 (theme fillStyleLst[0] with accent1 as phClr)", got)
	}
}

// TestNoBackgroundAnywhereStillRendersWhite is the negative-direction guard:
// a master with no <p:bg> at all must leave the slide white, proving the new
// rung cannot fire on its own.
func TestNoBackgroundAnywhereStillRendersWhite(t *testing.T) {
	pkg := masterBgPackage(t, "", false)
	pres := readMasterBgSlide(t, pkg)

	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("get slide: %v", err)
	}
	if slide.background != nil {
		t.Errorf("slide background = %+v, want nil when no rung declares one", slide.background)
	}
}
