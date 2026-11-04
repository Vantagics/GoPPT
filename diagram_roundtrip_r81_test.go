package gopresentation

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

// A SmartArt frame survives a read-write round trip.
//
// The reader renders a diagram from PowerPoint's cached drawing part; before
// this round the writer then flattened it to a plain group and dropped the
// frame and all four auxiliary parts, so what went in as a diagram came out
// as shapes a SmartArt editor could never pick up again. The stash fixes
// that: the parts are put back verbatim and the slide re-emits the original
// graphicFrame. The corpus probe is fill-color-list — a LibreOffice-made
// diagram with no cached drawing at all — whose Office-export roundtrip gap
// was 47.6% and dropped under 8% once the frame was preserved.

// diagramPartsPackage builds a one-slide deck whose slide carries a SmartArt
// frame pointing at four on-disk diagram parts, the way a real package does.
func diagramPartsPackage(t *testing.T) []byte {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="5" name="Diagram 4"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="1266498" y="625577"/><a:ext cx="6611007" cy="4401731"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
        <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>`)
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
</Relationships>`)
	for _, name := range []string{"data1", "layout1", "quickStyle1", "colors1"} {
		parts["ppt/diagrams/"+name+".xml"] = []byte("<" + name + " data/>")
	}
	return buildZip(t, parts)
}

// TestSmartArtRoundTripKeepsTheFrame reads the package above, writes it back,
// and checks the written package: the diagram parts are byte-identical and
// the slide still carries the graphicFrame with the five relationship ids.
func TestSmartArtRoundTripKeepsTheFrame(t *testing.T) {
	pkg := diagramPartsPackage(t)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	var buf bytes.Buffer
	w, err := NewWriter(pres, WriterPowerPoint2007)
	if err != nil {
		t.Fatal(err)
	}
	if err := w.(*PPTXWriter).WriteTo(&buf); err != nil {
		t.Fatalf("write: %v", err)
	}

	again, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	have := make(map[string]string)
	for _, f := range again.File {
		b := new(bytes.Buffer)
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		b.ReadFrom(rc)
		rc.Close()
		have[f.Name] = b.String()
	}
	for _, name := range []string{"ppt/diagrams/data1.xml", "ppt/diagrams/layout1.xml", "ppt/diagrams/quickStyle1.xml", "ppt/diagrams/colors1.xml"} {
		got, ok := have[name]
		if !ok {
			t.Errorf("round trip lost %s", name)
		} else if got != "<"+strings.TrimSuffix(name[strings.LastIndex(name, "/")+1:], ".xml")+" data/>" {
			t.Errorf("%s content changed: %q", name, got)
		}
	}
	slide := have["ppt/slides/slide1.xml"]
	if !strings.Contains(slide, "dgm:relIds") {
		t.Errorf("the saved slide has no graphicFrame relIds:\n%s", slide)
	}
	rels := have["ppt/slides/_rels/slide1.xml.rels"]
	for _, relType := range []string{"diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors"} {
		if !strings.Contains(rels, relType) {
			t.Errorf("the saved slide rels lost the %s relationship", relType)
		}
	}
}
