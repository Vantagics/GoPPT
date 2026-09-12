package gopresentation_test

import (
	"fmt"

	ppt "github.com/VantageDataChat/GoPPT"
)

// These examples are compiled and run as part of `go test`, so the snippets in
// the README cannot drift away from the real API without failing the build.

// Example_presentation creates a presentation and renders a slide to an image.
func Example_presentation() {
	p := ppt.New()
	slide := p.GetActiveSlide()
	title := slide.CreateRichTextShape()
	title.CreateTextRun("Quarterly report")

	opts := ppt.DefaultRenderOptions()
	opts.Width = 320 // height follows the slide ratio

	img, err := p.SlideToImage(0, opts)
	if err != nil {
		fmt.Println("render failed:", err)
		return
	}

	fmt.Println("slides:", len(p.GetAllSlides()))
	fmt.Println("shapes:", len(slide.GetShapes()))
	fmt.Printf("image: %dx%d\n", img.Bounds().Dx(), img.Bounds().Dy())

	// Output:
	// slides: 1
	// shapes: 1
	// image: 320x240
}

// ExamplePresentation_UnsupportedShapes shows how to find content that could not
// be rendered, so a preview run can report it rather than silently omitting it.
// The reader creates these shapes itself; one is built by hand here to keep the
// example self-contained.
func ExamplePresentation_UnsupportedShapes() {
	p := ppt.New()
	diagram := ppt.NewUnsupportedShape("SmartArt diagram")
	diagram.BaseShape.SetWidth(4000000)
	diagram.BaseShape.SetHeight(2500000)
	p.GetActiveSlide().AddShape(diagram)

	for _, s := range p.UnsupportedShapes() {
		fmt.Println(s.Label())
	}

	// Output:
	// Unsupported: SmartArt diagram
}

// ExamplePresentation_SlidesToImages renders every slide at draft quality, the
// pattern to use for a batch preview of a whole deck.
func ExamplePresentation_SlidesToImages() {
	p := ppt.New()
	p.GetActiveSlide().CreateRichTextShape().CreateTextRun("First")
	p.CreateSlide().CreateRichTextShape().CreateTextRun("Second")

	opts := ppt.DefaultRenderOptions()
	opts.Width = 480
	opts.Draft = true
	// Share one cache across renders: constructing it scans font directories.
	opts.FontCache = ppt.NewFontCache()

	imgs, err := p.SlidesToImages(opts)
	if err != nil {
		fmt.Println("render failed:", err)
		return
	}

	fmt.Println("rendered:", len(imgs))
	fmt.Printf("first: %dx%d\n", imgs[0].Bounds().Dx(), imgs[0].Bounds().Dy())

	// Output:
	// rendered: 2
	// first: 480x360
}
