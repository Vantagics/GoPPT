package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// layoutBulletPackage assembles a one-slide package in the shape of the
// corpus's "Title and Content" layouts: the master's bodyStyle declares a
// character bullet at every level, the layout's body placeholder carries the
// given <a:lstStyle> body, and the slide's body placeholder has one run and
// declares no bullet of its own. What the paragraph ends up holding is
// exactly the rung of the ladder that answered.
func layoutBulletPackage(t *testing.T, layoutLstStyle string) []byte {
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
    <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="3200"/></a:lvl1pPr>
      <a:lvl2pPr><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2800"/></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`)
	parts["ppt/slideLayouts/slideLayout1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/>` + layoutLstStyle + `<a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`)
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Bullet ladder probe</a:t></a:r></a:p></p:txBody>
</p:sp>`))
	return buildZip(t, parts)
}

func readLayoutBulletParagraph(t *testing.T, pkg []byte) *Paragraph {
	t.Helper()
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		if ph, ok := sh.(*PlaceholderShape); ok && ph.phIdx == 1 {
			if len(ph.paragraphs) == 0 {
				t.Fatal("body placeholder has no paragraphs")
			}
			return ph.paragraphs[0]
		}
	}
	t.Fatal("idx=1 body placeholder not found")
	return nil
}

// TestLayoutBuNoneSuppressesMasterBullet is the corpus-common case: the
// layout's body placeholder declares <a:buNone/> at level 1 while the master's
// bodyStyle carries a buChar for the same level. The paragraph that declares
// nothing must end up with NONE, not the master's "•" — PowerPoint's "Title
// and Content" layouts silence exactly this way.
func TestLayoutBuNoneSuppressesMasterBullet(t *testing.T) {
	pkg := layoutBulletPackage(t, `<a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400" b="1"/></a:lvl1pPr></a:lstStyle>`)
	para := readLayoutBulletParagraph(t, pkg)
	if para.bullet == nil {
		t.Fatal("paragraph bullet stayed nil: neither the layout's buNone nor the master's buChar was applied")
	}
	if para.bullet.Type != BulletTypeNone {
		t.Errorf("paragraph bullet type = %v, want BulletTypeNone — the layout's <a:buNone/> must silence the master's buChar", para.bullet.Type)
	}
}

// TestLayoutBuCharBeatsMasterBullet: a layout that declares its own buChar
// wins over the master rung below it, with the buFont coming along.
func TestLayoutBuCharBeatsMasterBullet(t *testing.T) {
	pkg := layoutBulletPackage(t, `<a:lstStyle><a:lvl1pPr marL="228600" indent="-228600"><a:buFont typeface="Courier New"/><a:buChar char="▪"/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>`)
	para := readLayoutBulletParagraph(t, pkg)
	if para.bullet == nil {
		t.Fatal("paragraph bullet stayed nil: the layout's buChar was not applied")
	}
	if para.bullet.Type != BulletTypeChar || para.bullet.Style != "▪" {
		t.Errorf("paragraph bullet = %+v, want char bullet %q from the layout rung", para.bullet, "▪")
	}
	if para.bullet.Font != "Courier New" {
		t.Errorf("bullet font = %q, want Courier New from the layout's buFont", para.bullet.Font)
	}
}

// TestNoLayoutBulletFallsThroughToMaster is the regression guard for the rung
// order: a layout whose lstStyle declares no bullet at all must leave the
// paragraph to the master's bodyStyle, exactly as it did before the layout
// rung existed.
func TestNoLayoutBulletFallsThroughToMaster(t *testing.T) {
	pkg := layoutBulletPackage(t, `<a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:defRPr sz="2400" b="1"/></a:lvl1pPr></a:lstStyle>`)
	para := readLayoutBulletParagraph(t, pkg)
	if para.bullet == nil {
		t.Fatal("paragraph bullet stayed nil: the master bodyStyle buChar did not fall through")
	}
	if para.bullet.Type != BulletTypeChar || para.bullet.Style != "•" {
		t.Errorf("paragraph bullet = %+v, want the master's char bullet %q", para.bullet, "•")
	}
	if para.bullet.Font != "Arial" {
		t.Errorf("bullet font = %q, want Arial from the master's buFont", para.bullet.Font)
	}
}

// TestLayoutBulletLevelTwoOnlyLeavesLevelOneToMaster: a declaration at lvl2
// must not leak onto the level-0 paragraph — the per-level table is what
// keeps the rungs honest.
func TestLayoutBulletLevelTwoOnlyLeavesLevelOneToMaster(t *testing.T) {
	pkg := layoutBulletPackage(t, `<a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr><a:lvl2pPr marL="457200" indent="-457200"><a:buFont typeface="Arial"/><a:buChar char="–"/><a:defRPr sz="2000"/></a:lvl2pPr></a:lstStyle>`)
	para := readLayoutBulletParagraph(t, pkg)
	if para.bullet == nil || para.bullet.Type != BulletTypeNone {
		t.Fatalf("level-0 paragraph bullet = %+v, want the lvl1pPr buNone", para.bullet)
	}
}

// TestSlideDeclaredBulletBeatsLayoutRung is the reverse-direction guard: a
// paragraph that declares its own bullet keeps it — the layout rung only
// fills paragraphs that said nothing.
func TestSlideDeclaredBulletBeatsLayoutRung(t *testing.T) {
	pkg := layoutBulletPackage(t, `<a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>`)
	// Patch the slide's paragraph to declare its own bullet.
	parts := zipParts(t, pkg)
	slide := string(parts["ppt/slides/slide1.xml"])
	declared := `<a:p><a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="Ø"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Bullet ladder probe</a:t></a:r></a:p>`
	slide = strings.Replace(slide, `<a:p><a:r><a:rPr lang="en-US"/><a:t>Bullet ladder probe</a:t></a:r></a:p>`, declared, 1)
	parts["ppt/slides/slide1.xml"] = []byte(slide)
	para := readLayoutBulletParagraph(t, buildZip(t, parts))
	if para.bullet == nil {
		t.Fatal("paragraph bullet = nil; the slide's own declaration was dropped")
	}
	if para.bullet.Type != BulletTypeChar || para.bullet.Style != "Ø" {
		t.Errorf("paragraph bullet = %+v, want the slide's own char bullet %q — the layout rung must not override a declaration", para.bullet, "Ø")
	}
}
