package gopresentation

import (
	"bytes"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// One theme per master.
//
// A deck can carry any number of masters, and Office writes a theme part per
// master — the reader used to parse theme1.xml once for the whole file, so
// every slide on master2..N inherited master1's palette and rendered in the
// wrong colours outright. The corpus probe is a deck with seven masters whose
// slides averaged 60% wrong against Office's export; two masters are enough
// to pin the rule.

const twoMasterThemeXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="probe">
  <a:themeElements>
    <a:clrScheme name="probe"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2><a:lt2><a:srgbClr val="F1F1F1"/></a:lt2>
      <a:accent1><a:srgbClr val="%ACCENT1%"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="366092"/></a:hlink><a:folHlink><a:srgbClr val="366092"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`

// twoMasterSlideXML is a slide filling one rect with accent1. The extra
// attribute is empty for slide1 and showMasterSp="0" for slide2.
func twoMasterSlideXML(extra string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ` + extra + ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Filler"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
        <a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`
}

func twoMasterLayoutXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="LayoutBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="6148800"/><a:ext cx="9144000" cy="304800"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
        <a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`
}

func twoMasterMasterXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="MasterBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="6553200"/><a:ext cx="9144000" cy="304800"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        <a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
</p:sldMaster>`
}

func twoMasterMasterRels(themeTarget string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/` + themeTarget + `"/>
</Relationships>`
}

func twoMasterLayoutRels(masterName string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/` + masterName + `"/>
</Relationships>`
}

func twoMasterSlideRels(layoutName string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/` + layoutName + `"/>
</Relationships>`
}

// twoMasterPackage assembles a deck with two masters, each naming its own
// theme: master1's accent1 is red, master2's is blue. Slide1 fills its canvas
// with accent1 through master1; slide2 does the same through master2 and sets
// showMasterSp="0", asking that the decoration above it stay off the canvas.
func twoMasterPackage(t *testing.T) []byte {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/presentation.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
    <p:sldMasterId id="2147483649" r:id="rId2"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId3"/>
    <p:sldId id="257" r:id="rId4"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>`)
	parts["ppt/_rels/presentation.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`)

	theme1 := strings.Replace(twoMasterThemeXML, "%ACCENT1%", "E00000", 1)
	theme2 := strings.Replace(twoMasterThemeXML, "%ACCENT1%", "0000FF", 1)
	parts["ppt/theme/theme1.xml"] = []byte(theme1)
	parts["ppt/theme/theme2.xml"] = []byte(theme2)

	for i, name := range []string{"slideMaster1.xml", "slideMaster2.xml"} {
		parts["ppt/slideMasters/"+name] = []byte(twoMasterMasterXML())
		parts["ppt/slideMasters/_rels/"+name+".rels"] = []byte(twoMasterMasterRels("theme" + masterIndexName(i+1) + ".xml"))
	}
	for i, name := range []string{"slideLayout1.xml", "slideLayout2.xml"} {
		parts["ppt/slideLayouts/"+name] = []byte(twoMasterLayoutXML())
		parts["ppt/slideLayouts/_rels/"+name+".rels"] = []byte(twoMasterLayoutRels("slideMaster" + masterIndexName(i+1) + ".xml"))
	}
	parts["ppt/slides/slide1.xml"] = []byte(twoMasterSlideXML(""))
	parts["ppt/slides/slide2.xml"] = []byte(twoMasterSlideXML(`showMasterSp="0"`))
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(twoMasterSlideRels("slideLayout1.xml"))
	parts["ppt/slides/_rels/slide2.xml.rels"] = []byte(twoMasterSlideRels("slideLayout2.xml"))
	return buildZip(t, parts)
}

func masterIndexName(n int) string {
	return string(rune('0' + n))
}

// TestEachSlideUsesItsOwnMastersTheme renders both slides and checks the
// canvas colour: slide1 must come out red (master1's accent1), slide2 blue
// (master2's) — not both master1's.
func TestEachSlideUsesItsOwnMastersTheme(t *testing.T) {
	pkg := twoMasterPackage(t)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	opts := DefaultRenderOptions()
	opts.Width = 80
	want := []string{"E00000", "0000FF"}
	for i := 0; i < 2; i++ {
		path := filepath.Join(t.TempDir(), "slide.png")
		if err := pres.SaveSlideAsImage(i, path, opts); err != nil {
			t.Fatalf("render slide %d: %v", i+1, err)
		}
		f, err := os.Open(path)
		if err != nil {
			t.Fatal(err)
		}
		img, err := png.Decode(f)
		f.Close()
		if err != nil {
			t.Fatal(err)
		}
		r, g, b, _ := img.At(40, 10).RGBA()
		got := canvasColor(r, g, b)
		if got != want[i] {
			t.Errorf("slide %d canvas = %s, want %s (its own master's accent1)", i+1, got, want[i])
		}
	}
}

func canvasColor(r, g, b uint32) string {
	const hexdigits = "0123456789ABCDEF"
	hex := func(v uint32) byte { return hexdigits[(v>>8)/16] }
	hex2 := func(v uint32) byte { return hexdigits[(v>>8)%16] }
	return string([]byte{hex(r), hex2(r), hex(g), hex2(g), hex(b), hex2(b)})
}

// TestSlideShowMasterSpSilencesDecoration checks the second half: slide2's
// showMasterSp="0" keeps its layout's green band (and any master decoration)
// off its shape list, while slide1 — same layout family, no flag — carries
// both decoration rungs. Suppression is a reading decision, not a drawing
// one: the shapes never enter the slide's list.
func TestSlideShowMasterSpSilencesDecoration(t *testing.T) {
	pkg := twoMasterPackage(t)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	slides := pres.GetAllSlides()
	if n := len(slides[0].GetShapes()); n != 3 {
		t.Errorf("slide1 holds %d shapes, want master band + layout band + its own filler", n)
	}
	if n := len(slides[1].GetShapes()); n != 1 {
		t.Errorf("slide2 holds %d shapes, want its own filler only — showMasterSp=\"0\" silences the decoration above it", n)
	}
}
