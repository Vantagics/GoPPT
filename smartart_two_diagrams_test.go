package gopresentation

import (
	"strings"
	"testing"
)

// A slide can carry several diagram graphicFrames — customGeo's slide14 lays
// two SmartArt diagrams side by side. Each frame must render ITS OWN cached
// drawing: the frame's <dgm:relIds r:dm> names the data part (dataN.xml) and
// PowerPoint numbers the cached drawing to match (drawingN.xml). The first
// implementation ignored r:dm entirely and handed every frame the first
// diagramDrawing relationship it found, so both frames drew the same diagram
// and the second one vanished. These tests pin the pairing.

var twoDiagramSlideXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="` + nsDrawingML + `" xmlns:r="` + nsOfficeDocRels + `" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
` + diagramFrameWithDMXML(100000, 200000, 2000000, 1000000, "rId10") + `
` + diagramFrameWithDMXML(2600000, 200000, 2000000, 1000000, "rId20") + `
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`

// diagramFrameWithDMXML is diagramFrameXML whose relIds carry a given r:dm.
func diagramFrameWithDMXML(x, y, cx, cy int, dm string) string {
	return `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="3" name="Diagram"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="` + itoa(x) + `" y="` + itoa(y) + `"/>
    <a:ext cx="` + itoa(cx) + `" cy="` + itoa(cy) + `"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="` + nsDiagram + `">
      <dgm:relIds xmlns:dgm="` + nsDiagram + `" xmlns:r="` + nsOfficeDocRels + `" r:dm="` + dm + `" r:lo="rId91" r:qs="rId92" r:cs="rId93"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`
}

// gammaDrawingXML is a second, distinguishable diagram: one shape saying Gamma.
func gammaDrawingXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="` + nsDsp + `" xmlns:a="` + nsDrawingML + `">
  <dsp:spTree>
    <dsp:nvGrpSpPr>
      <dsp:cNvPr id="0" name=""/>
      <dsp:cNvGrpSpPr/>
    </dsp:nvGrpSpPr>
    <dsp:grpSpPr/>
    <dsp:sp modelId="{33333333-3333-3333-3333-333333333333}">
      <dsp:nvSpPr>
        <dsp:cNvPr id="0" name="gamma"/>
        <dsp:cNvSpPr/>
      </dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
      </dsp:spPr>
      <dsp:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Gamma</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>`
}

// twoDiagramDeck builds a one-slide package with two diagram frames whose
// data parts (data1/data2) pair with two cached drawings (drawing1/drawing2).
func twoDiagramDeck(t *testing.T) []byte {
	t.Helper()
	pres := New()
	pres.AddSlide(newSlide())
	parts := zipParts(t, writeToBytes(t, pres))
	parts["ppt/slides/slide1.xml"] = []byte(twoDiagramSlideXML)
	parts["ppt/diagrams/drawing1.xml"] = []byte(diagramDrawingXML())
	parts["ppt/diagrams/drawing2.xml"] = []byte(gammaDrawingXML())

	ct := string(parts["[Content_Types].xml"])
	ct = strings.Replace(ct, "</Types>",
		`  <Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>`+"\n"+
			`  <Override PartName="/ppt/diagrams/drawing2.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>`+"\n</Types>", 1)
	parts["[Content_Types].xml"] = []byte(ct)

	relPath := "ppt/slides/_rels/slide1.xml.rels"
	rel := string(parts[relPath])
	rel = strings.Replace(rel, "</Relationships>",
		`  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>`+"\n"+
			`  <Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/>`+"\n"+
			`  <Relationship Id="rId11" Type="`+relDrawingType+`" Target="../diagrams/drawing1.xml"/>`+"\n"+
			`  <Relationship Id="rId21" Type="`+relDrawingType+`" Target="../diagrams/drawing2.xml"/>`+"\n</Relationships>", 1)
	parts[relPath] = []byte(rel)
	return buildZip(t, parts)
}

// diagramGroupsOf returns the diagram GroupShapes on a slide keyed by the
// frame offset X they were given.
func diagramGroupsOf(slide *Slide) map[int64]*GroupShape {
	out := map[int64]*GroupShape{}
	for _, shape := range slide.GetShapes() {
		if g, ok := shape.(*GroupShape); ok && g.name == "Diagram" {
			out[g.offsetX] = g
		}
	}
	return out
}

func groupText(g *GroupShape) string {
	var parts []string
	for _, child := range g.GetShapes() {
		parts = appendShapeText(parts, child)
	}
	return strings.Join(parts, "\n")
}

func TestTwoDiagramFramesRenderTheirOwnDrawings(t *testing.T) {
	pres := readDiagramDeck(t, twoDiagramDeck(t))
	slide := pres.Slides()[0]

	groups := diagramGroupsOf(slide)
	if len(groups) != 2 {
		t.Fatalf("slide has %d diagram groups, want 2 (one per graphicFrame)", len(groups))
	}
	left, ok := groups[100000]
	if !ok {
		t.Fatalf("no diagram group at the first frame's offset 100000; got offsets %v", groups)
	}
	right, ok := groups[2600000]
	if !ok {
		t.Fatalf("no diagram group at the second frame's offset 2600000; got offsets %v", groups)
	}
	// The pairing direction is the whole point: rId10→data1.xml pairs with
	// drawing1.xml (Alpha/Beta), rId20→data2.xml with drawing2.xml (Gamma).
	if got := groupText(left); !strings.Contains(got, "Alpha") {
		t.Errorf("frame r:dm=rId10 (data1) drew %q, want drawing1's Alpha", got)
	}
	if got := groupText(right); !strings.Contains(got, "Gamma") {
		t.Errorf("frame r:dm=rId20 (data2) drew %q, want drawing2's Gamma — a frame that renders the wrong diagram's drawing means the r:dm pairing was ignored", got)
	}
	// And the model still answers for both diagrams.
	if text := slide.ExtractText(); !strings.Contains(text, "Gamma") || !strings.Contains(text, "Alpha") {
		t.Errorf("ExtractText() = %q, want both diagrams' text", text)
	}
}

func TestDrawingPartForDataRelPairsByNumber(t *testing.T) {
	rels := []xmlRelForRead{
		{ID: "rId3", Target: "../diagrams/data8.xml"},
		{ID: "rId8", Target: "../diagrams/data9.xml"},
	}
	if got := drawingPartForDataRel(rels, "rId3"); got != "drawing8.xml" {
		t.Errorf("rId3 → %q, want drawing8.xml (the drawing carries its data part's number)", got)
	}
	if got := drawingPartForDataRel(rels, "rId8"); got != "drawing9.xml" {
		t.Errorf("rId8 → %q, want drawing9.xml", got)
	}
	if got := drawingPartForDataRel(rels, "rId404"); got != "" {
		t.Errorf("unknown rel → %q, want empty (caller falls back to first drawing)", got)
	}
	if got := drawingPartForDataRel(rels, ""); got != "" {
		t.Errorf("empty rel id → %q, want empty", got)
	}
}
