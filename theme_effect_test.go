package gopresentation

import (
	"strings"
	"testing"
)

// The effectRef → theme effectStyleLst tests.
//
// An <a:effectRef idx="N"> does not describe a shadow — it names the Nth entry
// of the theme's <a:effectStyleLst>, whose bodies carry an <a:outerShdw>.
// Ignoring it left every themed shape without the soft grey fade PowerPoint
// draws under it (slide04 of the comparison deck: a flat cut-off where the
// export shows an 8px gaussian fade).

// themeShadowEffects replaces the empty effect styles with Office-style ones:
// style 2 carries a downward outer shadow, black at 38% alpha.
const themeShadowEffects = `<a:effectStyle><a:effectLst/></a:effectStyle>` +
	`<a:effectStyle><a:effectLst>` +
	`<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>` +
	`</a:effectLst></a:effectStyle>` +
	`<a:effectStyle><a:effectLst/></a:effectStyle>`

func themeWithShadowEffects() string {
	return strings.Replace(themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles),
		`<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>`,
		`<a:effectStyleLst>`+themeShadowEffects+`</a:effectStyleLst>`, 1)
}

// styledShapeEffect is styledShape with a caller-chosen effectRef.
func styledShapeEffect(effectRef string) string {
	shape := styledShape("", "")
	return strings.Replace(shape,
		`<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>`,
		effectRef, 1)
}

// TestEffectRefResolvesThemeShadow: an effectRef on a shape whose spPr has no
// effectLst inherits the theme style's outer shadow.
func TestEffectRefResolvesThemeShadow(t *testing.T) {
	pres := themeTestRead(t, themeWithShadowEffects(),
		styledShapeEffect(`<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>`), "")
	sh := pres.GetAllSlides()[0].GetShapes()
	if len(sh) == 0 {
		t.Fatal("the slide has no shapes")
	}
	rt, ok := sh[0].(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", sh[0])
	}
	if rt.shadow == nil || !rt.shadow.Visible {
		t.Fatalf("effectRef idx=2 resolved to %+v, want a visible shadow", rt.shadow)
	}
	// The body's geometry, converted the way the reader converts an explicit
	// one: EMU→points for distance and blur, 60000ths→degrees for direction.
	if rt.shadow.Direction != 90 {
		t.Errorf("shadow direction = %d°, want 90 (dir=5400000)", rt.shadow.Direction)
	}
	if rt.shadow.Distance != 1 {
		t.Errorf("shadow distance = %dpt, want 1 (dist=20000 EMU)", rt.shadow.Distance)
	}
	if rt.shadow.BlurRadius != 3 {
		t.Errorf("shadow blur = %dpt, want 3 (blurRad=40000 EMU)", rt.shadow.BlurRadius)
	}
	if rt.shadow.Alpha != 38 {
		t.Errorf("shadow alpha = %d%%, want 38", rt.shadow.Alpha)
	}
	if rt.shadow.Color.ARGB != "FF000000" {
		t.Errorf("shadow colour = %s, want FF000000", rt.shadow.Color.ARGB)
	}
}

// TestExplicitShadowBeatsEffectRef: the shape's own <a:effectLst> outranks the
// theme style, the same way an explicit fill outranks the fillRef.
func TestExplicitShadowBeatsEffectRef(t *testing.T) {
	shape := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Both"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:effectLst><a:outerShdw blurRad="12700" dist="38100" dir="2700000"><a:srgbClr val="FF0000"/></a:outerShdw></a:effectLst>
  </p:spPr>
  <p:style>
    <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
    <a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
    <a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
	pres := themeTestRead(t, themeWithShadowEffects(), shape, "")
	sh := pres.GetAllSlides()[0].GetShapes()
	if len(sh) == 0 {
		t.Fatal("the slide has no shapes")
	}
	rt, ok := sh[0].(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", sh[0])
	}
	if rt.shadow == nil {
		t.Fatal("explicit effectLst produced no shadow")
	}
	if rt.shadow.Direction != 45 {
		t.Errorf("shadow direction = %d°, want 45 from the explicit effectLst (dir=2700000)", rt.shadow.Direction)
	}
	if rt.shadow.Distance != 3 {
		t.Errorf("shadow distance = %dpt, want 3 (dist=38100 EMU)", rt.shadow.Distance)
	}
	if rt.shadow.Color.ARGB != "FFFF0000" {
		t.Errorf("shadow colour = %s, want FFFF0000", rt.shadow.Color.ARGB)
	}
}

// TestEffectRefIdx0MeansNoShadow: idx="0" is "no theme effect" — a shape must
// not inherit style 1's shadow (if it had one) by falling off the index.
func TestEffectRefIdx0MeansNoShadow(t *testing.T) {
	pres := themeTestRead(t, themeWithShadowEffects(),
		styledShapeEffect(`<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>`), "")
	sh := pres.GetAllSlides()[0].GetShapes()
	rt := sh[0].(*RichTextShape)
	if rt.shadow != nil && rt.shadow.Visible {
		t.Errorf("effectRef idx=0 resolved to a visible shadow %+v, want none", rt.shadow)
	}
}
