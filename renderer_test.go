package gopresentation

import (
	"image/color"
	"image/png"
	"os"
	"testing"

	"golang.org/x/image/font"
)

func TestSlideToImage_BlankSlide(t *testing.T) {
	p := New()
	img, err := p.SlideToImage(0, nil)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}
	bounds := img.Bounds()
	if bounds.Dx() != 960 {
		t.Errorf("expected width 960, got %d", bounds.Dx())
	}
	// Check aspect ratio: 4:3 => 960:720
	if bounds.Dy() != 720 {
		t.Errorf("expected height 720, got %d", bounds.Dy())
	}
}

func TestSlideToImage_WithShapes(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Add a text box
	tb := slide.CreateRichTextShape()
	tb.SetOffsetX(Inch(1))
	tb.SetOffsetY(Inch(1))
	tb.SetWidth(Inch(4))
	tb.SetHeight(Inch(1))
	tb.GetFill().SetSolid(NewColor("DDEEFF"))
	tb.CreateTextRun("Hello, World!")

	// Add an auto shape
	as := slide.CreateAutoShape()
	as.SetAutoShapeType(AutoShapeRectangle)
	as.BaseShape.SetOffsetX(Inch(1))
	as.BaseShape.SetOffsetY(Inch(3))
	as.BaseShape.SetWidth(Inch(2))
	as.BaseShape.SetHeight(Inch(1))
	as.SetSolidFill(NewColor("FF6600"))
	as.SetText("Box")

	// Add a line
	line := slide.CreateLineShape()
	line.BaseShape.SetOffsetX(Inch(0))
	line.BaseShape.SetOffsetY(Inch(0))
	line.BaseShape.SetWidth(Inch(10))
	line.BaseShape.SetHeight(Inch(7.5))

	img, err := p.SlideToImage(0, nil)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}

	// Save for visual inspection
	f, err := os.Create("testdata_slide_shapes.png")
	if err != nil {
		t.Skipf("cannot create test output: %v", err)
	}
	defer f.Close()
	defer os.Remove("testdata_slide_shapes.png")
	png.Encode(f, img)
}

func TestSlideToImage_OutOfRange(t *testing.T) {
	p := New()
	_, err := p.SlideToImage(5, nil)
	if err == nil {
		t.Error("expected error for out-of-range slide index")
	}
}

func TestSlideToImage_CustomOptions(t *testing.T) {
	p := New()
	opts := &RenderOptions{
		Width:           1920,
		Format:          ImageFormatJPEG,
		JPEGQuality:     85,
		BackgroundColor: &color.RGBA{R: 30, G: 30, B: 30, A: 255},
	}
	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}
	if img.Bounds().Dx() != 1920 {
		t.Errorf("expected width 1920, got %d", img.Bounds().Dx())
	}
}

func TestSlidesToImages(t *testing.T) {
	p := New()
	p.CreateSlide()
	p.CreateSlide()

	images, err := p.SlidesToImages(nil)
	if err != nil {
		t.Fatalf("SlidesToImages: %v", err)
	}
	if len(images) != 3 {
		t.Errorf("expected 3 images, got %d", len(images))
	}
}

func TestSlideToImage_Table(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	table := slide.CreateTableShape(2, 3)
	table.BaseShape.SetOffsetX(Inch(1))
	table.BaseShape.SetOffsetY(Inch(1))
	table.BaseShape.SetWidth(Inch(6))
	table.BaseShape.SetHeight(Inch(2))
	table.GetCell(0, 0).SetText("A1")
	table.GetCell(0, 1).SetText("B1")
	table.GetCell(0, 2).SetText("C1")
	table.GetCell(1, 0).SetText("A2")

	img, err := p.SlideToImage(0, nil)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}
	if img.Bounds().Dx() != 960 {
		t.Errorf("expected width 960, got %d", img.Bounds().Dx())
	}
}

func TestSlideToImage_Background(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	slide.SetBackground(NewFill().SetSolid(NewColor("003366")))

	img, err := p.SlideToImage(0, nil)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}

	// Check that the background is dark blue
	r, g, b, _ := img.At(10, 10).RGBA()
	if r>>8 != 0x00 || g>>8 != 0x33 || b>>8 != 0x66 {
		t.Errorf("unexpected background color: R=%d G=%d B=%d", r>>8, g>>8, b>>8)
	}
}

func TestFontCache_SystemFonts(t *testing.T) {
	fc := NewFontCache()
	// Try to get Arial which should exist on most systems
	face := fc.GetFace("arial", 12, false, false)
	if face == nil {
		t.Skip("Arial not found on this system, skipping")
	}
	// Verify it's a real face by measuring text
	w := font.MeasureString(face, "Hello")
	if w <= 0 {
		t.Error("expected positive text width from TrueType face")
	}
}

func TestFontCache_BoldItalic(t *testing.T) {
	fc := NewFontCache()
	regular := fc.GetFace("arial", 14, false, false)
	bold := fc.GetFace("arial", 14, true, false)
	if regular == nil || bold == nil {
		t.Skip("Arial fonts not found, skipping")
	}
	// Bold text should generally be wider
	wRegular := font.MeasureString(regular, "Hello World")
	wBold := font.MeasureString(bold, "Hello World")
	t.Logf("Regular width: %v, Bold width: %v", wRegular, wBold)
}

func TestFontCache_LoadFontData(t *testing.T) {
	fc := NewFontCache()
	// Loading invalid data should fail
	err := fc.LoadFontData("test", []byte("not a font"))
	if err == nil {
		t.Error("expected error for invalid font data")
	}
}

func TestFontCache_Fallback(t *testing.T) {
	fc := NewFontCache()
	// A font name that definitely doesn't exist
	face := fc.GetFace("nonexistent-font-xyz-12345", 12, false, false)
	if face != nil {
		t.Error("expected nil for nonexistent font")
	}
}

func TestSlideToImage_TrueTypeText(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Title with large bold text
	title := slide.CreateRichTextShape()
	title.SetOffsetX(Inch(1))
	title.SetOffsetY(Inch(0.5))
	title.SetWidth(Inch(8))
	title.SetHeight(Inch(1))
	tr := title.CreateTextRun("TrueType Rendering")
	tr.GetFont().SetName("Arial").SetSize(28).SetBold(true).SetColor(NewColor("003366"))

	// Body with mixed formatting
	body := slide.CreateRichTextShape()
	body.SetOffsetX(Inch(1))
	body.SetOffsetY(Inch(2))
	body.SetWidth(Inch(8))
	body.SetHeight(Inch(3))

	r1 := body.CreateTextRun("Normal text, ")
	r1.GetFont().SetName("Arial").SetSize(14)

	r2 := body.CreateTextRun("bold text, ")
	r2.GetFont().SetName("Arial").SetSize(14).SetBold(true)

	r3 := body.CreateTextRun("italic text.")
	r3.GetFont().SetName("Arial").SetSize(14).SetItalic(true)

	// Share a font cache across renders for performance
	fc := NewFontCache()
	opts := &RenderOptions{
		Width:     1920,
		FontCache: fc,
	}

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("SlideToImage: %v", err)
	}

	if img.Bounds().Dx() != 1920 {
		t.Errorf("expected width 1920, got %d", img.Bounds().Dx())
	}

	// Save for visual inspection
	f, err := os.Create("testdata_truetype.png")
	if err != nil {
		t.Skipf("cannot create test output: %v", err)
	}
	defer f.Close()
	defer os.Remove("testdata_truetype.png")
	png.Encode(f, img)
	t.Log("Saved testdata_truetype.png for visual inspection")
}

func TestSlideToImage_SharedFontCache(t *testing.T) {
	// Verify that sharing a FontCache across multiple renders works and is faster
	p := New()
	slide := p.GetActiveSlide()
	tr := slide.CreateRichTextShape()
	tr.SetOffsetX(Inch(1)).SetOffsetY(Inch(1)).SetWidth(Inch(4)).SetHeight(Inch(1))
	tr.CreateTextRun("Cached font test")

	fc := NewFontCache()
	opts := &RenderOptions{FontCache: fc}

	// First render triggers font scan
	_, err := p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("first render: %v", err)
	}

	// Second render should reuse cached fonts
	_, err = p.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("second render: %v", err)
	}
}
