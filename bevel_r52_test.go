package gopresentation

import (
	"bytes"
	"image/color"
	"testing"
)

// The theme's third effect style carries <a:scene3d><a:sp3d><a:bevelT> next
// to its outer shadow. An <a:effectRef idx="3"> resolves the bevel along with
// the shadow, and the renderer lights the face and shades the rim with it —
// the puff a themed gradient box has in PowerPoint that a flat fill lacks.

// bevelTestThemeOverride rebuilds the themeTestTheme layout with a bevelled
// third effect style instead of the empty one.
var bevelTestThemeOverride = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="` + nsDrawingML + `" name="Test">
  <a:themeElements>
    <a:clrScheme name="Test">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Test">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Test">
      <a:fillStyleLst>` + themeOfficeFillStyles + `</a:fillStyleLst>
      <a:lnStyleLst>` + themeOfficeLnStyles + `</a:lnStyleLst>
      <a:effectStyleLst>` +
		`<a:effectStyle><a:effectLst/></a:effectStyle>` +
		`<a:effectStyle><a:effectLst/></a:effectStyle>` +
		`<a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst>` +
		`<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>` +
		`<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle>` +
		`</a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`

// A styled rect: fillRef idx=1 (solid phClr = accent1), effectRef idx=3.
const bevelTestShape = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
  <p:style>
    <a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
</p:sp>`

func readBevelShape(t *testing.T) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(bevelTestShape))
	parts["ppt/theme/theme1.xml"] = []byte(bevelTestThemeOverride)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func TestEffectRefResolvesThemeBevel(t *testing.T) {
	pres := readBevelShape(t)
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("shape count = %d, want 1", len(shapes))
	}
	var w, h int64
	switch s := shapes[0].(type) {
	case *AutoShape:
		w, h = s.GetBevelWidth(), s.GetBevelHeight()
	case *RichTextShape:
		w, h = s.GetBevelWidth(), s.GetBevelHeight()
	default:
		t.Fatalf("shape type %T, want an AutoShape or RichTextShape", shapes[0])
	}
	if w != 63500 || h != 25400 {
		t.Fatalf("resolved bevel = %dx%d EMU, want 63500x25400", w, h)
	}
	// The shadow rides along on the same reference.
	var shadowOK bool
	switch s := shapes[0].(type) {
	case *AutoShape:
		shadowOK = s.shadow != nil && s.shadow.Visible
	case *RichTextShape:
		shadowOK = s.shadow != nil && s.shadow.Visible
	}
	if !shadowOK {
		t.Fatalf("effectRef idx=3 must also resolve the style's outer shadow")
	}
}

func TestBevelShadingLightsFaceAndRims(t *testing.T) {
	pres := readBevelShape(t)
	opts := DefaultRenderOptions()
	opts.Width = 640
	opts.FontCache = NewFontCache()
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rect := emuRect(t, pres, 500000, 500000, 3000000, 1500000, opts.Width)
	at := func(x, y int) color.RGBA {
		r, g, b, a := img.At(x, y).RGBA()
		return color.RGBA{uint8(r >> 8), uint8(g >> 8), uint8(b >> 8), uint8(a >> 8)}
	}
	// accent1 4F81BD = (79,129,189); the lit face is fill ×1.19 in linear
	// light — R: lin(79/255)=0.078 ×1.19 → sRGB 86, B → 202.
	c := at(rect.Min.X+rect.Dx()/2, rect.Min.Y+rect.Dy()/2)
	if c.R < 83 || c.R > 90 || c.B < 197 || c.B > 207 {
		t.Fatalf("face centre = rgb(%d,%d,%d), want the fill lifted to about (86,140,203)", c.R, c.G, c.B)
	}
	// One row inside the top edge: the specular crest — brighter than the
	// face, B pushed toward 255.
	crest := at(rect.Min.X+rect.Dx()/2, rect.Min.Y+1)
	if crest.R <= c.R+10 {
		t.Fatalf("top crest rgb(%d,%d,%d) is not brighter than the face rgb(%d,%d,%d)",
			crest.R, crest.G, crest.B, c.R, c.G, c.B)
	}
	// The last drawn row (Max.Y is exclusive): shaded — darker than the face.
	shade := at(rect.Min.X+rect.Dx()/2, rect.Max.Y)
	if shade.R >= c.R-8 {
		t.Fatalf("bottom rim rgb(%d,%d,%d) is not darker than the face rgb(%d,%d,%d)",
			shade.R, shade.G, shade.B, c.R, c.G, c.B)
	}
}

func TestNoBevelWithoutEffectRef(t *testing.T) {
	// The same shape with effectRef idx=0 keeps the flat fill — the bevel
	// comes only from a style reference that carries one.
	flat := `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>
  <p:style>
    <a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
</p:sp>`
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(flat))
	parts["ppt/theme/theme1.xml"] = []byte(bevelTestThemeOverride)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("shape count = %d, want 1", len(shapes))
	}
	if s, ok := shapes[0].(*AutoShape); ok {
		if s.GetBevelWidth() != 0 {
			t.Fatalf("effectRef idx=0 resolved a bevel %d", s.GetBevelWidth())
		}
		return
	}
	if s, ok := shapes[0].(*RichTextShape); ok {
		if s.GetBevelWidth() != 0 {
			t.Fatalf("effectRef idx=0 resolved a bevel %d", s.GetBevelWidth())
		}
		return
	}
	t.Fatalf("shape type %T", shapes[0])
}
