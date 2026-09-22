package gopresentation

import (
	"bytes"
	"image"
	"testing"
)

// The VML ink annotation tests.
//
// PowerPoint stores pen-annotated ink in legacy VML drawing parts related from
// the slide's .rels but referenced by no slide XML element. The reader used to
// ignore them, so slides 11-14 of the fidelity deck drew only the base shapes
// where PowerPoint shows a full hand-drawn diagram.

// inkTestVML is a miniature vmlDrawing part: two subpath strokes in one shape
// (an arrow and a squiggle), plus a filled shape the reader must not model as
// a stroked connector.
const inkTestVML = `<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="3"/></o:shapelayout>
 <v:shape id="_x0000_s3074" style='position:absolute;left:151pt;
   top:226.5pt;width:99.75pt;height:28.125pt' coordorigin="5326,7989"
   coordsize="3518,995" path="m6730,8318v6,,12,-1,18,-1em5635,7998v-44,31,-69,60,-105,102r-200,300r-100,300e"
   filled="f" strokecolor="#0070c0" strokeweight="1.5pt">
   <v:stroke endcap="round"/>
   <o:ink i="AAAA" annotation="t"/>
 </v:shape>
 <v:shape id="_x0000_s3099" style='position:absolute;left:10pt;top:10pt;width:40pt;height:20pt'
   coordorigin="0,0" coordsize="800,400" path="m0,0l800,400e"
   strokecolor="#000000" strokeweight="1pt">
 </v:shape>
</xml>`

const inkTestSlideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
	`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
	`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
	`</Relationships>`

// inkTestRead reads a one-slide package carrying the VML part above.
func inkTestRead(t *testing.T) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(""))
	parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(inkTestSlideRels)
	parts["ppt/drawings/vmlDrawing1.vml"] = []byte(inkTestVML)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

// TestReaderParsesInkAnnotation is the headline: a slide whose rels relate a
// VML drawing part must surface its shapes, positioned in EMU, with the
// declared stroke weight and colour, and the path normalised to 0..coordsize.
func TestReaderParsesInkAnnotation(t *testing.T) {
	pres := inkTestRead(t)
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) != 2 {
		t.Fatalf("the slide has %d shapes, want 2 (stroke ink + filled ink)", len(shapes))
	}

	line, ok := shapes[0].(*LineShape)
	if !ok {
		t.Fatalf("first shape is %T, want *LineShape", shapes[0])
	}
	// 151pt * 12700 EMU/pt = 1917700; 226.5pt = 2876550.
	if line.GetOffsetX() != 1917700 || line.GetOffsetY() != 2876550 {
		t.Errorf("ink offset %d,%d EMU, want 1917700,2876550", line.GetOffsetX(), line.GetOffsetY())
	}
	// 99.75pt = 1266825 EMU; 28.125pt = 357187.5 -> rounds to 357188.
	if line.GetWidth() != 1266825 || line.GetHeight() != 357188 {
		t.Errorf("ink size %d,%d EMU, want 1266825,357187", line.GetWidth(), line.GetHeight())
	}
	if line.GetLineWidthEMU() != 19050 {
		t.Errorf("stroke width %d EMU, want 19050 (1.5pt)", line.GetLineWidthEMU())
	}
	if got := line.lineColor.ARGB; got != "FF0070C0" {
		t.Errorf("stroke colour %s, want FF0070C0", got)
	}
	cp := line.customPath
	if cp == nil {
		t.Fatal("the ink shape carries no custom path")
	}
	if cp.Width != 3518 || cp.Height != 995 {
		t.Errorf("path space %dx%d, want 3518x995", cp.Width, cp.Height)
	}
	// Two moveTo subpaths, path coordinates minus coordorigin.
	var moves int
	for _, cmd := range cp.Commands {
		if cmd.Type == "moveTo" {
			moves++
			if moves == 1 && (cmd.Pts[0].X != 6730-5326 || cmd.Pts[0].Y != 8318-7989) {
				t.Errorf("first subpath starts at %d,%d, want %d,%d",
					cmd.Pts[0].X, cmd.Pts[0].Y, 6730-5326, 8318-7989)
			}
		}
	}
	if moves != 2 {
		t.Errorf("the path has %d moveTo subpaths, want 2", moves)
	}

	if _, ok := shapes[1].(*UnsupportedShape); !ok {
		t.Errorf("second shape is %T, want *UnsupportedShape (filled ink)", shapes[1])
	}
}

// TestVMLPathEmptyFieldsAreZero pins the tokenizer quirk that silently
// distorted every stroke the first time: "v6,,12,-1,18,-1" is six numbers
// with dy1=0, not five.
func TestVMLPathEmptyFieldsAreZero(t *testing.T) {
	cmds := parseVMLPath("m6730,8318v6,,12,-1,18,-1e", 0, 0)
	if len(cmds) != 2 {
		t.Fatalf("got %d commands, want moveTo + cubicBezTo", len(cmds))
	}
	bez := cmds[1]
	if bez.Type != "cubicBezTo" || len(bez.Pts) != 3 {
		t.Fatalf("second command is %v, want cubicBezTo with 3 points", bez)
	}
	// Relative cubic, SVG semantics: all three points offset from the current
	// point (6730,8318): c1=(6736,8318) c2=(6742,8317) end=(6748,8317).
	want := [][2]int64{{6736, 8318}, {6742, 8317}, {6748, 8317}}
	for i, p := range bez.Pts {
		if p.X != want[i][0] || p.Y != want[i][1] {
			t.Errorf("point %d is %d,%d, want %d,%d", i, p.X, p.Y, want[i][0], want[i][1])
		}
	}
}

// TestInkSubpathsSplitAtMoveTo is the renderer half: a path whose subpaths are
// separated by moveTo must stroke as independent polylines. Concatenating
// them draws a bridge segment between strokes PowerPoint never does.
func TestInkSubpathsSplitAtMoveTo(t *testing.T) {
	cp := &CustomGeomPath{
		Width: 100, Height: 100,
		Commands: []PathCommand{
			{Type: "moveTo", Pts: []PathPoint{{X: 0, Y: 0}}},
			{Type: "lnTo", Pts: []PathPoint{{X: 100, Y: 0}}},
			{Type: "moveTo", Pts: []PathPoint{{X: 0, Y: 100}}},
			{Type: "lnTo", Pts: []PathPoint{{X: 100, Y: 100}}},
		},
	}
	subs := (&renderer{}).customPathToPixelSubpaths(cp, 0, 0, 100, 100)
	if len(subs) != 2 {
		t.Fatalf("got %d subpaths, want 2", len(subs))
	}
	for i, sub := range subs {
		if len(sub) != 2 {
			t.Fatalf("subpath %d has %d points, want 2", i, len(sub))
		}
	}
	// No bridging: the second subpath must start at its own moveTo, not where
	// the first one ended.
	if subs[1][0].y != 100 {
		t.Errorf("second subpath starts at y=%.0f, want 100 (the first subpath's end would bridge)", subs[1][0].y)
	}
}

// TestInkRendersBlue is the end-to-end half: the ink must actually reach the
// pixels, in the declared colour, at the declared position.
func TestInkRendersBlue(t *testing.T) {
	pres := inkTestRead(t)
	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 720 // 1pt = 1px on the 720x540pt slide
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// Both strokes' boxes: the first subpath is a short blip near
	// (191, 236); the second sweeps from (160, 227) down-left to ~(148, 256).
	box := image.Rect(140, 220, 251, 262)
	blue := 0
	b := img.Bounds()
	for y := box.Min.Y; y < box.Max.Y; y++ {
		for x := box.Min.X; x < box.Max.X; x++ {
			if x < b.Min.X || x >= b.Max.X || y < b.Min.Y || y >= b.Max.Y {
				continue
			}
			r, g, bl, _ := img.At(x, y).RGBA()
			if bl > 150 && bl > r+40 && bl > g+30 {
				blue++
			}
		}
	}
	if blue < 15 {
		t.Errorf("only %d blue-ish pixels in the ink's box; the annotation did not render", blue)
	}
}

// TestWriterWritesInkAsCustGeom is the writer half of the three-half rule:
// saving a deck that carries ink must serialise the strokes as custom
// geometry so PowerPoint re-renders them after a round trip.
func TestWriterWritesInkAsCustGeom(t *testing.T) {
	pres := inkTestRead(t)
	var buf bytes.Buffer
	if err := pres.WriteTo(&buf); err != nil {
		t.Fatalf("write: %v", err)
	}
	parts := zipParts(t, buf.Bytes())
	slide1 := string(parts["ppt/slides/slide1.xml"])
	if !bytes.Contains([]byte(slide1), []byte("<a:custGeom>")) {
		t.Fatal("slide1.xml has no custGeom; the ink strokes were dropped on save")
	}
	if n := bytes.Count([]byte(slide1), []byte("<a:moveTo>")); n < 2 {
		t.Errorf("slide1.xml has %d moveTo, want at least 2 (the ink's subpaths)", n)
	}
}
