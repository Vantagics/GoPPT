package gopresentation

import (
	"archive/zip"
	"bytes"
	"image"
	"image/color"
	"strings"
	"testing"
)

// An OLE object frame (Equation.3 and friends) has no picture of its own in
// the slide XML: PowerPoint draws the preview metafile stored in the legacy
// VML drawing part, addressed by the <p:oleObj spid>. The reader must resolve
// that chain — slide rels → vmlDrawingN.vml → v:shape[@id=spid] →
// v:imagedata@o:relid → VML part rels → media part — and model the frame as
// the picture that preview is.

const oleFrameShape = `
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="9" name="Object 12"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="3840163" y="2173288"/><a:ext cx="1265237" cy="531812"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
      <p:oleObj spid="_x0000_s96259" name="Equation" r:id="rId5" imgW="482400" imgH="203040" progId="Equation.3">
        <p:embed/>
      </p:oleObj>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`

const oleVMLDrawing = `<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="94"/></o:shapelayout>
 <v:shape id="_x0000_s96259" type="#_x0000_t75" style='position:absolute;
  left:302.375pt;top:171.125pt;width:99.625pt;height:41.875pt'>
  <v:imagedata o:relid="rId2" o:title=""/>
 </v:shape>
</xml>`

const oleVMLRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image42.wmf"/></Relationships>`

const oleSlideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
	`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
	`<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>` +
	`</Relationships>`

func readOLEDeck(t *testing.T, wmf []byte) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(oleFrameShape))
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(oleSlideRels)
	parts["ppt/drawings/vmlDrawing1.vml"] = []byte(oleVMLDrawing)
	parts["ppt/drawings/_rels/vmlDrawing1.vml.rels"] = []byte(oleVMLRels)
	parts["ppt/media/image42.wmf"] = wmf
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func TestOLEFrameRendersThroughVMLPreview(t *testing.T) {
	wmf := wmfWithWindowExt(100, 80)
	pres := readOLEDeck(t, wmf)
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("shape count = %d, want 1 (the OLE frame as its preview picture)", len(shapes))
	}
	ds, ok := shapes[0].(*DrawingShape)
	if ok {
		if !bytes.Equal(ds.data, wmf) {
			t.Fatalf("preview bytes = %d long, want the %d bytes of the VML-referenced metafile", len(ds.data), len(wmf))
		}
		if ds.mimeType != "image/x-wmf" {
			t.Fatalf("mimeType = %q, want image/x-wmf", ds.mimeType)
		}
		if ds.offsetX != 3840163 || ds.offsetY != 2173288 || ds.width != 1265237 || ds.height != 531812 {
			t.Fatalf("frame rect = (%d,%d %dx%d), want the graphicFrame xfrm",
				ds.offsetX, ds.offsetY, ds.width, ds.height)
		}
		return
	}
	if _, isUn := shapes[0].(*UnsupportedShape); isUn {
		t.Fatalf("the OLE frame stayed an UnsupportedShape: the VML imagedata chain was not followed")
	}
	t.Fatalf("shape type %T, want a *DrawingShape", shapes[0])
}

func TestOLEWithoutVMLMatchStaysUnsupported(t *testing.T) {
	// A spid the VML part does not name must not silently become an empty
	// picture; the visible stand-in keeps the gap reportable.
	broken := strings.Replace(oleVMLDrawing, "_x0000_s96259", "_x0000_s99999", 1)
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(oleFrameShape))
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(oleSlideRels)
	parts["ppt/drawings/vmlDrawing1.vml"] = []byte(broken)
	parts["ppt/drawings/_rels/vmlDrawing1.vml.rels"] = []byte(oleVMLRels)
	parts["ppt/media/image42.wmf"] = wmfWithWindowExt(100, 80)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 1 {
		t.Fatalf("shape count = %d, want 1", len(shapes))
	}
	if _, ok := shapes[0].(*UnsupportedShape); !ok {
		t.Fatalf("shape type %T, want an UnsupportedShape when no VML shape matches the spid", shapes[0])
	}
}

// buildSymbolEquationWMF plays the record sequence an Equation.3 preview
// uses: CreateFontIndirect(Symbol, charset 2) + SelectObject + ExtTextOut of
// one glyph byte.
func buildSymbolEquationWMF(glyph byte) []byte {
	u16 := func(v int) []byte { return []byte{byte(v), byte(v >> 8)} }
	rec := func(function int, body []byte) []byte {
		words := (6 + len(body) + 1) / 2
		out := []byte{byte(words), byte(words >> 8), 0, 0}
		out = append(out, u16(function)...)
		out = append(out, body...)
		for len(out) < words*2 {
			out = append(out, 0)
		}
		return out
	}
	var out []byte
	out = append(out, 0x01, 0x00, 0x09, 0x00)
	out = append(out, make([]byte, 14)...)
	out = append(out, rec(0x020C, append(u16(80), u16(100)...))...) // SetWindowExt y,x

	logfont := make([]byte, 50)
	logfont[0] = 0xF0 // lfHeight = -16 (em size)
	logfont[1] = 0xFF
	logfont[13] = 2 // lfCharSet = SYMBOL_CHARSET
	copy(logfont[18:], "Symbol\x00")
	out = append(out, rec(0x02FB, logfont)...)
	out = append(out, rec(0x012D, u16(0))...) // SelectObject 0

	body := append(u16(10), u16(10)...) // y, x
	body = append(body, u16(1)...)      // count
	body = append(body, u16(0)...)      // options
	body = append(body, glyph)
	out = append(out, rec(0x0A32, body)...)
	return out
}

func TestWMFSymbolCharsetRendersFromPUAMapping(t *testing.T) {
	// Symbol's cmap only publishes its glyphs in the U+F000 private-use
	// range (probed directly: rune 0xD1 misses, 0xF0D1 hits). A byte 0xB4 —
	// the multiply sign × in the Symbol encoding — must map there: without
	// the mapping the byte stays U+00B4, which the Symbol face has no glyph
	// for at all, and the equation loses its operators. The canvas must
	// therefore carry a solid glyph's worth of ink.
	img := decodeMetafileBitmap(buildSymbolEquationWMF(0xB4), NewFontCache())
	if img == nil {
		t.Fatal("the symbol metafile rendered nothing")
	}
	ink := 0
	b := img.Bounds()
	for py := b.Min.Y; py < b.Max.Y; py++ {
		for px := b.Min.X; px < b.Max.X; px++ {
			c := color.RGBAModel.Convert(img.At(px, py)).(color.RGBA)
			if c.A > 0 && c.R < 128 {
				ink++
			}
		}
	}
	if ink < 100 {
		t.Fatalf("only %d ink pixels on the canvas: the SYMBOL_CHARSET byte never reached a glyph", ink)
	}
	if b.Dx() != 400 || b.Dy() != 320 {
		t.Fatalf("canvas %dx%d, want 400x320 (4x of the 100x80 window extent)", b.Dx(), b.Dy())
	}
	_ = image.Point{} // keep the image import honest if assertions change
}

func TestWriterEmitsOLEPreviewAsWMFPicture(t *testing.T) {
	pres := readOLEDeck(t, wmfWithWindowExt(100, 80))
	var buf bytes.Buffer
	w, err := NewWriter(pres, WriterPowerPoint2007)
	if err != nil {
		t.Fatalf("new writer: %v", err)
	}
	if err := w.WriteTo(&buf); err != nil {
		t.Fatalf("write: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	var wmfName, slideXML, slideRels, ctypes string
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		var b bytes.Buffer
		if _, err := b.ReadFrom(rc); err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		rc.Close()
		switch {
		case strings.HasSuffix(f.Name, ".wmf"):
			wmfName = f.Name
		case f.Name == "ppt/slides/slide1.xml":
			slideXML = b.String()
		case f.Name == "ppt/slides/_rels/slide1.xml.rels":
			slideRels = b.String()
		case f.Name == "[Content_Types].xml":
			ctypes = b.String()
		}
	}
	if wmfName == "" {
		t.Fatal("no .wmf media part in the written package")
	}
	if !strings.Contains(ctypes, `Extension="wmf"`) {
		t.Fatal("[Content_Types].xml lacks a wmf default; PowerPoint would reject the part")
	}
	if !strings.Contains(slideXML, "<p:pic>") {
		t.Fatal("the OLE frame was not written back as its preview picture")
	}
	if !strings.Contains(slideRels, ".wmf") {
		t.Fatal("the slide rels do not reference the wmf part")
	}
}
