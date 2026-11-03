package gopresentation

import (
	"bytes"
	"testing"
)

// An explicit "0" on a run property is a statement, not an absence.
//
// The inheritance ladder applied the master bodyStyle's bold whenever it was
// set, and a plain bool cannot tell "the run said b=\"0\"" from "the run said
// nothing" — so a run that explicitly declined bold came out bold anyway.
// The corpus case is a placeholder whose master bodyStyle carries b="1"
// while its runs each carry b="0": PowerPoint draws regular weight, the
// renderer drew bold and wrapped the line differently as a consequence.

// boldLadderPackage builds the r62-style one-slide package whose master
// bodyStyle declares bold, with the given run markup in the body placeholder.
func boldLadderPackage(t *testing.T, runMarkup string) []byte {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1325562"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
  <p:txStyles>
    <p:bodyStyle><a:lvl1pPr><a:defRPr b="1" sz="3200"/></a:lvl1pPr></p:bodyStyle>
  </p:txStyles>
</p:sldMaster>`)
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p>` + runMarkup + `</a:p></p:txBody>
</p:sp>`))
	return buildZip(t, parts)
}

func bodyFontOfFirstRun(t *testing.T, pkg []byte) *Font {
	t.Helper()
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		if ph, ok := sh.(*PlaceholderShape); ok && ph.phIdx == 1 {
			if len(ph.paragraphs) == 0 || len(ph.paragraphs[0].elements) == 0 {
				t.Fatal("body placeholder has no run to inspect")
			}
			tr, ok := ph.paragraphs[0].elements[0].(*TextRun)
			if !ok || tr.font == nil {
				t.Fatalf("first element is %T with font nil, want a TextRun with a font", ph.paragraphs[0].elements[0])
			}
			return tr.font
		}
	}
	t.Fatal("idx=1 body placeholder not found")
	return nil
}

// TestExplicitBoldZeroBeatsMasterBold: the run said b="0", the master said
// b="1" — the run wins.
func TestExplicitBoldZeroBeatsMasterBold(t *testing.T) {
	pkg := boldLadderPackage(t, `<a:r><a:rPr lang="en-US" sz="2000" b="0"/><a:t>probe</a:t></a:r>`)
	if f := bodyFontOfFirstRun(t, pkg); f.Bold {
		t.Errorf("run with explicit b=\"0\" came out bold: the master bodyStyle overrode a stated property")
	}
}

// TestUnstatedBoldStillInherits guards the ladder's working half: a run that
// says nothing takes the master's bold.
func TestUnstatedBoldStillInherits(t *testing.T) {
	pkg := boldLadderPackage(t, `<a:r><a:rPr lang="en-US"/><a:t>probe</a:t></a:r>`)
	if f := bodyFontOfFirstRun(t, pkg); !f.Bold {
		t.Errorf("run without a b attribute did not inherit the master bodyStyle's bold")
	}
}
