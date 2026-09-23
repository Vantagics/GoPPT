package gopresentation

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/draw"
	_ "image/gif"
	"image/jpeg"
	"image/png"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"unicode"

	"golang.org/x/image/font"
	"golang.org/x/image/font/basicfont"
	"golang.org/x/image/math/fixed"
	_ "golang.org/x/image/tiff"
	"golang.org/x/text/encoding/simplifiedchinese"
)

// ImageFormat represents the output image format.
type ImageFormat int

const (
	ImageFormatPNG ImageFormat = iota
	ImageFormatJPEG
)

// RenderOptions configures slide-to-image rendering.
type RenderOptions struct {
	// Width is the output image width in pixels. Height is calculated from slide aspect ratio.
	// Default: 960
	Width int
	// Format is the output image format (PNG or JPEG).
	Format ImageFormat
	// JPEGQuality is the JPEG quality (1-100). Default: 90.
	JPEGQuality int
	// BackgroundColor overrides the slide background. Nil means use slide background or white.
	BackgroundColor *color.RGBA
	// DPI is the rendering DPI for font sizing. Default: 96.
	DPI float64
	// FontDirs specifies additional directories to search for TrueType/OpenType fonts.
	// System font directories are always searched automatically.
	FontDirs []string
	// FontCache allows sharing a pre-configured FontCache across multiple renders.
	// If nil, a new FontCache is created using FontDirs.
	//
	// Creating a FontCache scans the font directories, so when rendering many
	// slides or many files, build one cache and reuse it here. SlidesToImages
	// already reuses the cache across the slides of a single call.
	FontCache *FontCache
	// FontFallback lists font names tried, in order, when a document requests a
	// font that is not installed. These are tried before the built-in fallback
	// chain. Setting e.g. []string{"Noto Sans CJK SC"} pins CJK text to a known
	// installed font instead of depending on the built-in order.
	FontFallback []string
	// FontDiagnostics, when non-nil, collects every font request that was
	// substituted or not found at all during the render. Use it to tell a
	// broken document apart from a machine that lacks the required font.
	FontDiagnostics *FontDiagnostics
	// OnFontFallback, when non-nil, is called for each font request that could
	// not be satisfied exactly. It may be called concurrently.
	OnFontFallback func(FontUsage)
	// OverlayOpacityScale scales the opacity of semi-transparent shape fills.
	// Value between 0.0 and 1.0. Default 0 means use 1.0 (no change).
	// Set to e.g. 0.5 to halve the opacity of overlays, making dark backgrounds brighter.
	OverlayOpacityScale float64
	// Draft trades fidelity for speed, for batch previews whose purpose is to
	// show roughly what a slide looks like rather than to judge it pixel by
	// pixel. It skips anti-aliasing on lines and ellipses, skips shadows
	// entirely, and scales images with nearest-neighbour sampling instead of
	// bilinear.
	//
	// Glyph rasterisation is still anti-aliased, because that cannot be turned
	// off without replacing the text renderer; text stays readable, which is the
	// point of a preview. Combine Draft with a small Width (e.g. 480) for the
	// largest saving.
	Draft bool
}

// DefaultRenderOptions returns default rendering options.
func DefaultRenderOptions() *RenderOptions {
	return &RenderOptions{
		Width:       960,
		Format:      ImageFormatPNG,
		JPEGQuality: 90,
		DPI:         96,
	}
}

// SlideToImage renders a single slide to an image.
//
// SlideToImage never panics: malformed shape data that trips up the rasterizer
// is reported as a *PanicError.
func (p *Presentation) SlideToImage(slideIndex int, opts *RenderOptions) (img image.Image, err error) {
	defer recoverToError(&err, "Presentation.SlideToImage")
	if p == nil {
		return nil, fmt.Errorf("presentation is nil")
	}
	if slideIndex < 0 || slideIndex >= len(p.slides) {
		return nil, fmt.Errorf("slide index %d out of range (0-%d)", slideIndex, len(p.slides)-1)
	}
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	// Fill in defaults on a copy. Options belong to the caller, so writing a
	// derived value back into them would be an unexpected side effect, and it
	// would race with a second render handed the same struct. SlidesToImages
	// relies on this too.
	local := *opts
	if local.Width <= 0 {
		local.Width = 960
	}
	opts = &local

	slide := p.slides[slideIndex]
	layout := p.layout
	if layout == nil {
		return nil, fmt.Errorf("presentation has no document layout")
	}
	if layout.CX <= 0 || layout.CY <= 0 {
		return nil, fmt.Errorf("presentation has an invalid slide size (%d x %d EMU)", layout.CX, layout.CY)
	}

	slideW := float64(layout.CX)
	slideH := float64(layout.CY)
	imgW := opts.Width
	imgH := int(float64(imgW) * slideH / slideW)
	// An extreme slide aspect ratio can round the height away; a zero-height
	// canvas is not a usable image, so keep at least one row.
	if imgH < 1 {
		imgH = 1
	}

	scaleX := float64(imgW) / slideW
	scaleY := float64(imgH) / slideH

	canvas := image.NewRGBA(image.Rect(0, 0, imgW, imgH))

	fc := opts.FontCache
	if fc == nil {
		fc = NewFontCache(opts.FontDirs...)
	}
	dpi := opts.DPI
	if dpi <= 0 {
		dpi = 96
	}

	r := &renderer{
		img:                 canvas,
		scaleX:              scaleX,
		scaleY:              scaleY,
		fontCache:           fc,
		dpi:                 dpi,
		overlayOpacityScale: opts.OverlayOpacityScale,
		fontFallback:        mergeFallbackChain(opts.FontFallback, defaultFontFallbackChain),
		cjkFallback:         mergeFallbackChain(opts.FontFallback, defaultCJKFallbackChain),
		symbolFallback:      mergeFallbackChain(opts.FontFallback, defaultSymbolFallbackChain),
		fontDiag:            opts.FontDiagnostics,
		onFontFallback:      opts.OnFontFallback,
		draft:               opts.Draft,
		slideNumber:         slideIndex + 1,
	}

	// Fill background
	bgColor := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	drawn := false
	if opts.BackgroundColor != nil {
		bgColor = *opts.BackgroundColor
	} else if slide.background != nil {
		switch slide.background.Type {
		case FillSolid:
			bgColor = argbToRGBA(slide.background.Color)
		case FillGradientLinear:
			r.fillGradientLinear(canvas.Bounds(), slide.background)
			drawn = true
		case FillGradientPath:
			r.fillGradientPath(canvas.Bounds(), slide.background)
			drawn = true
		}
	}
	if !drawn {
		r.fillRectFast(canvas.Bounds(), bgColor)
	}

	// Render shapes in their original XML order (z-order).
	// Shapes that appear earlier in the spTree are behind shapes that appear later,
	// matching PowerPoint's rendering behavior.
	for _, shape := range slide.shapes {
		r.renderShape(shape)
	}

	return canvas, nil
}

// SlidesToImages renders all slides to images.
//
// The FontCache in opts is reused across slides; if it is nil one is created
// once per call, so pass a shared cache when rendering repeatedly. opts itself
// is never modified.
func (p *Presentation) SlidesToImages(opts *RenderOptions) (imgs []image.Image, err error) {
	defer recoverToError(&err, "Presentation.SlidesToImages")
	if p == nil {
		return nil, fmt.Errorf("presentation is nil")
	}
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	// Work on a copy rather than writing the derived cache back into the
	// caller's struct: options are caller-owned data, and mutating them here
	// would race if the same options were passed to two presentations at once.
	local := *opts
	if local.FontCache == nil {
		local.FontCache = NewFontCache(local.FontDirs...)
	}
	imgs = make([]image.Image, len(p.slides))
	for i := range p.slides {
		img, slideErr := p.SlideToImage(i, &local)
		if slideErr != nil {
			return nil, fmt.Errorf("slide %d: %w", i, slideErr)
		}
		imgs[i] = img
	}
	return imgs, nil
}

// SaveSlideAsImage renders a slide and saves it to a file.
func (p *Presentation) SaveSlideAsImage(slideIndex int, path string, opts *RenderOptions) (err error) {
	defer recoverToError(&err, "Presentation.SaveSlideAsImage")
	img, err := p.SlideToImage(slideIndex, opts)
	if err != nil {
		return err
	}
	return saveImage(img, path, opts)
}

// SaveSlidesAsImages renders all slides and saves them to files.
// The pattern should contain %d for the slide number (1-based), e.g. "slide_%d.png".
//
// As in SlidesToImages, the FontCache in opts is shared across the slides and
// opts itself is never modified.
func (p *Presentation) SaveSlidesAsImages(pattern string, opts *RenderOptions) (err error) {
	defer recoverToError(&err, "Presentation.SaveSlidesAsImages")
	if p == nil {
		return fmt.Errorf("presentation is nil")
	}
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	// Build the font cache once for the whole call. Creating one scans the font
	// directories, so leaving it to each slide would repeat that scan — and keep
	// a second copy of every parsed font — once per slide.
	local := *opts
	if local.FontCache == nil {
		local.FontCache = NewFontCache(local.FontDirs...)
	}
	for i := range p.slides {
		path := fmt.Sprintf(pattern, i+1)
		if err := p.SaveSlideAsImage(i, path, &local); err != nil {
			return fmt.Errorf("slide %d: %w", i+1, err)
		}
	}
	return nil
}

func saveImage(img image.Image, path string, opts *RenderOptions) error {
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	dir := filepath.Dir(path)
	if dir != "" && dir != "." {
		if err := os.MkdirAll(dir, 0750); err != nil {
			return fmt.Errorf("create directory: %w", err)
		}
	}
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		return fmt.Errorf("create file: %w", err)
	}
	var encodeErr error
	switch opts.Format {
	case ImageFormatJPEG:
		quality := opts.JPEGQuality
		if quality <= 0 || quality > 100 {
			quality = 90
		}
		encodeErr = jpeg.Encode(f, img, &jpeg.Options{Quality: quality})
	default:
		encodeErr = png.Encode(f, img)
	}
	closeErr := f.Close()
	if encodeErr != nil {
		return encodeErr
	}
	return closeErr
}

// --- renderer core ---

type renderer struct {
	img                 *image.RGBA
	scaleX              float64
	scaleY              float64
	fontCache           *FontCache
	dpi                 float64
	overlayOpacityScale float64 // 0 means 1.0 (no change)
	fontScale           float64 // normAutofit font scale factor (0 or 1.0 = no scaling)
	lnSpcReduction      float64 // normAutofit lnSpcReduction as a fraction (0 = none, 0.10 = lines 10% shorter)
	fontFallback        []string
	cjkFallback         []string
	symbolFallback      []string
	fontDiag            *FontDiagnostics
	onFontFallback      func(FontUsage)
	// draft skips anti-aliasing, shadows and image smoothing. See
	// RenderOptions.Draft.
	draft bool
	// slideNumber is the 1-based number of the slide being rendered. A
	// <a:fld type="slidenum"> run evaluates to it at draw time; the cached
	// <a:t> in the file is whatever the last save saw.
	slideNumber int
}

// subRenderer returns a renderer that shares this renderer's configuration but
// draws into a different image. Always use this rather than building a renderer
// literal: a literal silently drops any configuration field added later.
func (r *renderer) subRenderer(img *image.RGBA) *renderer {
	clone := *r
	clone.img = img
	return &clone
}

func (r *renderer) renderShape(shape Shape) {
	// <p:cNvPr hidden="1">: PowerPoint keeps the shape in the file but never
	// draws it — including whole groups, which hides their children too.
	// Rendering it would paint hidden whiteout rectangles over real content.
	if shape.base().hidden {
		return
	}
	switch s := shape.(type) {
	case *RichTextShape:
		r.renderRichText(s)
	case *PlaceholderShape:
		r.renderRichText(&s.RichTextShape)
	case *DrawingShape:
		r.renderDrawing(s)
	case *AutoShape:
		r.renderAutoShape(s)
	case *LineShape:
		r.renderLine(s)
	case *TableShape:
		r.renderTable(s)
	case *ChartShape:
		r.renderChart(s)
	case *GroupShape:
		r.renderGroup(s)
	case *UnsupportedShape:
		// A construct the reader could not represent. Draw a visible
		// placeholder so the gap shows up in a preview instead of a blank area.
		r.renderUnsupported(s)
	}
}

func (r *renderer) emuToPixelX(emu int64) int { return int(math.Round(float64(emu) * r.scaleX)) }
func (r *renderer) emuToPixelY(emu int64) int { return int(math.Round(float64(emu) * r.scaleY)) }

// hundredthPtToPixelY converts hundredths of a point (from spcPts) to pixels.
// spcPts values are in 1/100 of a point, e.g. 1200 = 12pt.
// 1 point = 12700 EMU, so 1/100 point = 127 EMU.
func (r *renderer) hundredthPtToPixelY(val int) int {
	emu := float64(val) * 127.0
	// Round, not truncate: a 12pt space-before is 26.67px, and truncation
	// shaves a pixel off every gap on the slide — the loss accumulates
	// down the text block (slide35 drifted 6px by its last line).
	return int(emu*r.scaleY + 0.5)
}

func argbToRGBA(c Color) color.RGBA {
	return color.RGBA{R: c.GetRed(), G: c.GetGreen(), B: c.GetBlue(), A: c.GetAlpha()}
}

// drawSrc adapts one of this renderer's colours for use with image/draw and
// font.Drawer.
//
// Colour values here carry *straight* alpha: blendPixel and its relatives
// multiply by c.A themselves, so the channels have to be the un-multiplied
// ones. Go, however, defines color.RGBA as alpha-premultiplied, and both
// image/draw and font.Drawer trust that definition. Handing them a
// straight-alpha value whose A is below 255 therefore violates the invariant —
// the channels exceed the alpha, the fixed-point blend overflows the byte it
// stores into, and the result wraps to something arbitrary rather than to the
// intended colour. In practice a white glyph at 12% opacity came out *darker*
// than the navy it was drawn on.
//
// color.NRGBA is the straight-alpha type, and its RGBA method premultiplies
// correctly, so it is the right source for those APIs. Opaque colours behave
// identically either way, which is why this only surfaced once a run carried
// <a:alpha>.
func drawSrc(c color.RGBA) color.Color {
	return color.NRGBA{R: c.R, G: c.G, B: c.B, A: c.A}
}

// --- Pixel operations (performance-critical) ---

// blendPixel alpha-blends color c over the existing pixel at (x, y).
// Uses direct Pix slice access for performance.
func (r *renderer) blendPixel(x, y int, c color.RGBA) {
	b := r.img.Bounds()
	if x < b.Min.X || x >= b.Max.X || y < b.Min.Y || y >= b.Max.Y {
		return
	}
	if c.A == 0 {
		return
	}
	off := (y-b.Min.Y)*r.img.Stride + (x-b.Min.X)*4
	pix := r.img.Pix
	if c.A == 255 {
		pix[off] = c.R
		pix[off+1] = c.G
		pix[off+2] = c.B
		pix[off+3] = 255
		return
	}
	a := uint32(c.A)
	ia := 255 - a
	pix[off] = uint8((uint32(c.R)*a + uint32(pix[off])*ia) / 255)
	pix[off+1] = uint8((uint32(c.G)*a + uint32(pix[off+1])*ia) / 255)
	pix[off+2] = uint8((uint32(c.B)*a + uint32(pix[off+2])*ia) / 255)
	pix[off+3] = uint8(uint32(pix[off+3]) + (255-uint32(pix[off+3]))*a/255)
}

// blendPixelF blends with fractional coverage (0.0–1.0) for anti-aliasing.
func (r *renderer) blendPixelF(x, y int, c color.RGBA, coverage float64) {
	if r.draft {
		// Draft mode snaps partial coverage to fully on or fully off. Anti-
		// aliasing costs a per-pixel alpha computation and blend; thresholding
		// keeps the shape's core while removing the soft edges.
		if coverage < 0.5 {
			return
		}
		r.blendPixel(x, y, c)
		return
	}
	if coverage <= 0 {
		return
	}
	if coverage >= 1.0 {
		r.blendPixel(x, y, c)
		return
	}
	r.blendPixel(x, y, color.RGBA{R: c.R, G: c.G, B: c.B, A: uint8(float64(c.A) * coverage)})
}

// fillRectFast fills a rectangle with an opaque color using draw.Draw.
func (r *renderer) fillRectFast(rect image.Rectangle, c color.RGBA) {
	// drawSrc even for the opaque case: Uniform.RGBA reports sa == 0xffff for
	// an opaque colour whatever its concrete type, so the fast fill path is
	// still taken, and a translucent caller gets the right answer.
	draw.Draw(r.img, rect, &image.Uniform{drawSrc(c)}, image.Point{}, draw.Over)
}

// fillRectBlend fills a rectangle with alpha blending, using row-based direct Pix access.
func (r *renderer) fillRectBlend(rect image.Rectangle, c color.RGBA) {
	b := r.img.Bounds()
	rect = rect.Intersect(b)
	if rect.Empty() {
		return
	}
	if c.A == 0 {
		return
	}
	if c.A == 255 {
		r.fillRectFast(rect, c)
		return
	}
	a := uint32(c.A)
	ia := 255 - a
	cr, cg, cb := uint32(c.R)*a, uint32(c.G)*a, uint32(c.B)*a
	pix := r.img.Pix
	stride := r.img.Stride
	minX := rect.Min.X - b.Min.X
	minY := rect.Min.Y - b.Min.Y
	w := rect.Dx()
	for dy := 0; dy < rect.Dy(); dy++ {
		off := (minY+dy)*stride + minX*4
		for dx := 0; dx < w; dx++ {
			pix[off] = uint8((cr + uint32(pix[off])*ia) / 255)
			pix[off+1] = uint8((cg + uint32(pix[off+1])*ia) / 255)
			pix[off+2] = uint8((cb + uint32(pix[off+2])*ia) / 255)
			pix[off+3] = uint8(uint32(pix[off+3]) + (255-uint32(pix[off+3]))*a/255)
			off += 4
		}
	}
}

// --- Rotation & flip support ---

func rotatedBounds(cx, cy float64, w, h int, angleDeg int) image.Rectangle {
	rad := float64(angleDeg) * math.Pi / 180.0
	cos := math.Abs(math.Cos(rad))
	sin := math.Abs(math.Sin(rad))
	fw, fh := float64(w), float64(h)
	newW := fw*cos + fh*sin
	newH := fw*sin + fh*cos
	return image.Rect(
		int(cx-newW/2), int(cy-newH/2),
		int(cx+newW/2)+1, int(cy+newH/2)+1,
	)
}

// rotateAndComposite rotates src (sw x sh) by angleDeg and composites it into
// dst at (dx, dy) fitting into a dw x dh area. Used for vertical text where
// the text is drawn into a buffer with swapped dimensions then rotated back.
func rotateAndComposite(dst *image.RGBA, src *image.RGBA, dx, dy, dw, dh, angleDeg int) {
	sw := src.Bounds().Dx()
	sh := src.Bounds().Dy()
	if sw <= 0 || sh <= 0 || dw <= 0 || dh <= 0 {
		return
	}
	rad := float64(angleDeg) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	// Center of source
	scx := float64(sw) / 2
	scy := float64(sh) / 2
	// Center of destination area
	dcx := float64(dx) + float64(dw)/2
	dcy := float64(dy) + float64(dh)/2

	dstBounds := dst.Bounds()
	minDY := maxInt(dy, dstBounds.Min.Y)
	maxDY := minInt(dy+dh, dstBounds.Max.Y)
	minDX := maxInt(dx, dstBounds.Min.X)
	maxDX := minInt(dx+dw, dstBounds.Max.X)

	for py := minDY; py < maxDY; py++ {
		ry := float64(py) - dcy
		for px := minDX; px < maxDX; px++ {
			rx := float64(px) - dcx
			// Inverse rotation to find source pixel
			sx := rx*cosA + ry*sinA + scx
			sy := -rx*sinA + ry*cosA + scy
			ix, iy := int(sx), int(sy)
			if ix >= 0 && ix < sw && iy >= 0 && iy < sh {
				sOff := iy*src.Stride + ix*4
				a := src.Pix[sOff+3]
				if a == 0 {
					continue
				}
				dOff := py*dst.Stride + px*4
				if a == 255 || dst.Pix[dOff+3] == 0 {
					copy(dst.Pix[dOff:dOff+4], src.Pix[sOff:sOff+4])
				} else {
					// Alpha blend
					sa := uint32(a)
					da := uint32(dst.Pix[dOff+3])
					outA := sa + da*(255-sa)/255
					if outA > 0 {
						dst.Pix[dOff] = uint8((uint32(src.Pix[sOff])*sa + uint32(dst.Pix[dOff])*(255-sa)) / 255)
						dst.Pix[dOff+1] = uint8((uint32(src.Pix[sOff+1])*sa + uint32(dst.Pix[dOff+1])*(255-sa)) / 255)
						dst.Pix[dOff+2] = uint8((uint32(src.Pix[sOff+2])*sa + uint32(dst.Pix[dOff+2])*(255-sa)) / 255)
						dst.Pix[dOff+3] = uint8(outA)
					}
				}
			}
		}
	}
}

func (r *renderer) renderRotated(x, y, w, h, rotation int, flipH, flipV bool, drawFn func(tmp *renderer)) {
	r.renderRotatedExpanded(x, y, w, h, h, rotation, flipH, flipV, drawFn)
}

// renderRotatedExpanded is like renderRotated but uses bufH for the temp buffer
// height, allowing text to overflow the shape bounds without being clipped.
// The rotation center remains at the center of the original shape (w × h).
func (r *renderer) renderRotatedExpanded(x, y, w, h, bufH, rotation int, flipH, flipV bool, drawFn func(tmp *renderer)) {
	if w <= 0 || h <= 0 {
		return
	}
	if bufH < h {
		bufH = h
	}
	tmp := image.NewRGBA(image.Rect(0, 0, w, bufH))
	tmpR := r.subRenderer(tmp)
	drawFn(tmpR)

	if rotation == 0 && !flipH && !flipV {
		draw.Draw(r.img, image.Rect(x, y, x+w, y+bufH), tmp, image.Point{}, draw.Over)
		return
	}

	// Handle flip-only case (no rotation)
	if rotation == 0 {
		for py := 0; py < bufH; py++ {
			sy := py
			if flipV {
				sy = bufH - 1 - py
			}
			for px := 0; px < w; px++ {
				sx := px
				if flipH {
					sx = w - 1 - px
				}
				sOff := sy*tmp.Stride + sx*4
				if tmp.Pix[sOff+3] > 0 {
					r.blendPixel(x+px, y+py, color.RGBA{
						R: tmp.Pix[sOff], G: tmp.Pix[sOff+1],
						B: tmp.Pix[sOff+2], A: tmp.Pix[sOff+3],
					})
				}
			}
		}
		return
	}

	// OOXML transform order: rotate first, then flip.
	// We combine both into a single inverse mapping from destination to source.
	// OOXML rot is clockwise; in screen coords (Y-down) the standard rotation
	// matrix [cos,-sin;sin,cos] already rotates clockwise for positive angles.
	// The inverse mapping (dest→src) for a clockwise rotation by θ is:
	//   sx = fx*cos(θ) + fy*sin(θ)
	//   sy = -fx*sin(θ) + fy*cos(θ)
	rad := float64(rotation) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	cx := float64(w) / 2
	cy := float64(h) / 2
	destCX := float64(x) + cx
	destCY := float64(y) + cy

	bounds := rotatedBounds(destCX, destCY, w, bufH, rotation)
	imgBounds := r.img.Bounds()
	minDY := maxInt(bounds.Min.Y, imgBounds.Min.Y)
	maxDY := minInt(bounds.Max.Y, imgBounds.Max.Y)
	minDX := maxInt(bounds.Min.X, imgBounds.Min.X)
	maxDX := minInt(bounds.Max.X, imgBounds.Max.X)

	for dy := minDY; dy < maxDY; dy++ {
		ry := float64(dy) - destCY
		for dx := minDX; dx < maxDX; dx++ {
			rx := float64(dx) - destCX
			// OOXML forward transform order: flip first, then rotate.
			// Inverse: un-rotate first, then un-flip.
			// Step 1: un-rotate (inverse rotation)
			ux := rx*cosA + ry*sinA
			uy := -rx*sinA + ry*cosA
			// Step 2: un-flip
			if flipH {
				ux = -ux
			}
			if flipV {
				uy = -uy
			}
			sx := ux + cx
			sy := uy + cy
			ix, iy := int(sx), int(sy)
			if ix >= 0 && ix < w && iy >= 0 && iy < bufH {
				sOff := iy*tmp.Stride + ix*4
				if tmp.Pix[sOff+3] > 0 {
					r.blendPixel(dx, dy, color.RGBA{
						R: tmp.Pix[sOff], G: tmp.Pix[sOff+1],
						B: tmp.Pix[sOff+2], A: tmp.Pix[sOff+3],
					})
				}
			}
		}
	}
}

func (r *renderer) renderGroup(g *GroupShape) {
	// Transform child coordinates from child space (chOff/chExt) to group space (off/ext)
	if g.childExtX > 0 && g.childExtY > 0 {
		for _, gs := range g.shapes {
			bs := gs.base()
			origX := bs.offsetX
			origY := bs.offsetY
			origW := bs.width
			origH := bs.height
			bs.offsetX = g.offsetX + (origX-g.childOffX)*g.width/g.childExtX
			bs.offsetY = g.offsetY + (origY-g.childOffY)*g.height/g.childExtY
			bs.width = origW * g.width / g.childExtX
			bs.height = origH * g.height / g.childExtY
			defer func(s Shape, ox, oy, ow, oh int64) {
				b := s.base()
				b.offsetX = ox
				b.offsetY = oy
				b.width = ow
				b.height = oh
			}(gs, origX, origY, origW, origH)
		}
	}

	rotation := g.GetRotation()
	flipH := g.GetFlipHorizontal()
	flipV := g.GetFlipVertical()
	if rotation == 0 && !flipH && !flipV {
		for _, gs := range g.shapes {
			r.renderShape(gs)
		}
		return
	}
	x := r.emuToPixelX(g.offsetX)
	y := r.emuToPixelY(g.offsetY)
	w := r.emuToPixelX(g.width)
	h := r.emuToPixelY(g.height)
	r.renderRotated(x, y, w, h, rotation, flipH, flipV, func(tmp *renderer) {
		// Shift children to render relative to (0,0) in the temp buffer.
		// Children have absolute slide coordinates; subtract group origin.
		for _, gs := range g.shapes {
			bs := gs.base()
			bs.offsetX -= g.offsetX
			bs.offsetY -= g.offsetY
		}
		defer func() {
			for _, gs := range g.shapes {
				bs := gs.base()
				bs.offsetX += g.offsetX
				bs.offsetY += g.offsetY
			}
		}()
		for _, gs := range g.shapes {
			tmp.renderShape(gs)
		}
	})
}

// --- Shape rendering ---

func (r *renderer) renderRichText(s *RichTextShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	// Apply normAutofit font scale
	prevFontScale := r.fontScale
	if s.fontScale > 0 && s.fontScale != 100000 {
		r.fontScale = float64(s.fontScale) / 100000.0
	}
	// Apply normAutofit line-spacing reduction: each line advance shrinks by
	// the recorded percentage, which is what compresses an overflowing text
	// body back into its shape in PowerPoint.
	prevLnSpcReduction := r.lnSpcReduction
	if s.lnSpcReduction > 0 && s.lnSpcReduction != 100000 {
		r.lnSpcReduction = float64(s.lnSpcReduction) / 100000.0
	}
	defer func() {
		r.fontScale = prevFontScale
		r.lnSpcReduction = prevLnSpcReduction
	}()

	// Text insets (padding). PowerPoint defaults: lIns=91440, rIns=91440, tIns=45720, bIns=45720
	lIns, rIns, tIns, bIns := int64(91440), int64(91440), int64(45720), int64(45720)
	if s.insetsSet {
		lIns, rIns, tIns, bIns = s.insetLeft, s.insetRight, s.insetTop, s.insetBottom
	}
	pxL := r.emuToPixelX(lIns)
	pxR := r.emuToPixelX(rIns)
	pxT := r.emuToPixelY(tIns)
	pxB := r.emuToPixelY(bIns)

	// Clamp default insets when they consume too much of the shape dimensions.
	// This happens for small shapes inside nested groups where group coordinate
	// transforms scale shape dimensions but insets remain absolute EMU values.
	if !s.insetsSet {
		maxInsetH := int(float64(h) * 0.35)
		maxInsetW := int(float64(w) * 0.35)
		if pxT+pxB > maxInsetH {
			scale := float64(maxInsetH) / float64(pxT+pxB)
			pxT = int(float64(pxT) * scale)
			pxB = int(float64(pxB) * scale)
		}
		if pxL+pxR > maxInsetW {
			scale := float64(maxInsetW) / float64(pxL+pxR)
			pxL = int(float64(pxL) * scale)
			pxR = int(float64(pxR) * scale)
		}
	}

	// Vertical text direction adds implicit rotation
	vertRotation := 0
	if s.textDirection == "vert" || s.textDirection == "eaVert" || s.textDirection == "wordArtVert" {
		vertRotation = 270
	} else if s.textDirection == "vert270" {
		vertRotation = 90
	}

	// Estimate total text height to detect overflow.
	// PowerPoint does not clip text to the text box boundary, so we must
	// expand the rendering buffer when text overflows.
	tw := w - pxL - pxR
	th := h - pxT - pxB
	if tw < 1 {
		tw = w
	}
	if th < 1 {
		th = h
	}

	// spAutoFit: shape resizes to fit text. When the shape has word-wrap
	// enabled, PowerPoint expands the shape vertically while keeping the
	// width fixed. We cannot resize the shape at render time, but we
	// should still honour word-wrap so text wraps within the available
	// width instead of overflowing horizontally and overlapping adjacent
	// shapes. Only disable word-wrap when the original shape had it off
	// (rare case where the box expands horizontally).
	wordWrap := s.wordWrap

	// When default insets are used and text overflows, progressively reduce
	// insets to make room. Font metric differences between systems can cause
	// text to be slightly larger than the original authoring environment
	// expected, so shrinking insets first avoids unnecessary text overflow.
	if !s.insetsSet {
		textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
		if textH > th && th > 0 && (pxT+pxB) > 0 {
			needed := textH - th
			avail := pxT + pxB
			if needed >= avail {
				pxT = 0
				pxB = 0
			} else {
				scale := float64(avail-needed) / float64(avail)
				pxT = int(float64(pxT) * scale)
				pxB = int(float64(pxB) * scale)
			}
			th = h - pxT - pxB
			if th < 1 {
				th = h
			}
		}
	}

	// Auto-shrink text when normAutofit is set without an explicit fontScale.
	// PowerPoint dynamically calculates the scale to fit text within the box.
	shouldAutoShrink := false
	if s.autoFit == AutoFitNormal && (s.fontScale == 0 || s.fontScale == 100000) {
		shouldAutoShrink = true
	}
	// For spAutoFit (AutoFitShape), PowerPoint resizes the shape to fit text.
	// Since we cannot resize the shape at render time, apply a conservative
	// shrink with a high floor. The overflow is typically caused by font
	// metric differences between Go and PowerPoint rather than text that
	// genuinely needs a much larger shape. Using a high floor (0.92)
	// preserves text readability while keeping text roughly within bounds.
	isAutoFitShape := false
	if s.autoFit == AutoFitShape && (s.fontScale == 0 || s.fontScale == 100000) && th > 0 {
		textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
		if textH > th {
			shouldAutoShrink = true
			isAutoFitShape = true
		}
	}
	// For AutoFitNone, PowerPoint does NOT shrink text — it lets text overflow
	// the shape boundary. Do not apply any auto-shrink here; instead let the
	// text overflow naturally (the renderer already handles overflow by
	// expanding the buffer height).
	isAutoFitNone := s.autoFit == AutoFitNone
	if shouldAutoShrink {
		textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
		if textH > th && th > 0 {
			// Binary search for the right scale factor.
			// For AutoFitNone, use a high floor (0.85) since PowerPoint does
			// not shrink at all — we only compensate for font metric differences.
			// For AutoFitShape, use a high floor (0.92) since PowerPoint
			// resizes the shape rather than shrinking text.
			lo, hi := 0.3, 1.0
			if isAutoFitNone {
				lo = 0.85
			}
			if isAutoFitShape {
				lo = 0.92
			}
			for i := 0; i < 15; i++ {
				mid := (lo + hi) / 2
				r.fontScale = mid
				mh := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
				if mh > th {
					hi = mid
				} else {
					lo = mid
				}
			}
			r.fontScale = lo
		}
	}

	// Horizontal overflow detection: Go's font metrics can produce wider text
	// than PowerPoint's DirectWrite renderer, causing text to overflow the
	// right edge of the text box. When word-wrap is enabled and any wrapped
	// line still exceeds the text area width, shrink the font to fit.
	// Apply the same 3% tolerance used by wrapRunLine so that lines allowed
	// by the wrapping tolerance don't falsely trigger horizontal shrinking.
	if wordWrap && tw > 0 {
		hTolerance := tw * 103 / 100 // 3% tolerance matching wrapRunLine
		maxLW := r.measureMaxLineWidth(s.paragraphs, tw, wordWrap)
		if maxLW > hTolerance {
			// Binary search for a scale that fits horizontally.
			// For AutoFitNone, use a higher floor to avoid over-shrinking.
			// Use the current fontScale as hi (may already be reduced by
			// vertical shrink). Ensure the combined vertical+horizontal
			// shrink doesn't go below a reasonable floor.
			lo, hi := 0.5, r.fontScale
			if isAutoFitNone && lo < 0.85 {
				lo = 0.85
			}
			if hi <= 0 {
				hi = 1.0
			}
			// Prevent compound shrinking from going too low: if vertical
			// shrink already reduced the scale, raise the floor so the
			// combined effect doesn't over-shrink text. But never raise
			// the floor above hi (the current scale from vertical shrink).
			compoundFloor := hi * 0.85
			if lo < compoundFloor && compoundFloor <= hi {
				lo = compoundFloor
			}
			if lo > hi {
				lo = hi
			}
			for i := 0; i < 12; i++ {
				mid := (lo + hi) / 2
				r.fontScale = mid
				mw := r.measureMaxLineWidth(s.paragraphs, tw, wordWrap)
				if mw > hTolerance {
					hi = mid
				} else {
					lo = mid
				}
			}
			r.fontScale = lo
		}
	}

	textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
	// Extra height needed beyond the shape box
	overflowH := 0
	if textH+pxT+pxB > h {
		overflowH = textH + pxT + pxB - h
	}
	// Use expanded height for the temp buffer when rotated
	bufH := h + overflowH

	// skipText is used to split geometry and text rendering when flip is set.
	// PowerPoint flips shape geometry but keeps text readable (un-flipped).
	skipText := false

	drawContent := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		rect := image.Rect(ox, oy, ox+w, oy+h)

		// Shadow BEFORE fill (so shadow appears behind)
		if s.shadow != nil && s.shadow.Visible {
			tr.renderShadow(s.shadow, rect)
		}
		if s.customPath != nil {
			tr.renderCustomPathFill(s.customPath, s.fill, ox, oy, w, h)
		} else {
			tr.renderFill(s.fill, rect)
		}
		if s.border != nil && s.border.Style != BorderNone {
			pw := maxInt(int(float64(maxInt(s.border.Width, 1))*12700.0*tr.scaleX), 1)
			if s.customPath != nil {
				// Draw border along the custom geometry path
				pts := tr.customPathToPixelPoints(s.customPath, ox, oy, w, h)
				bc := argbToRGBA(s.border.Color)
				if len(pts) >= 2 {
					if s.border.Style == BorderDash || s.border.Style == BorderDot {
						tr.drawDashedPolylineAA(pts, bc, pw, s.border.Style)
					} else {
						for i := 1; i < len(pts); i++ {
							tr.drawLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), bc, pw)
						}
					}
					// Draw arrowheads at the ends of the custom path
					intPts := make([][2]int, len(pts))
					for i, p := range pts {
						intPts[i] = [2]int{int(p.x), int(p.y)}
					}
					if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
						tr.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, bc, pw, s.headEnd)
					}
					if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
						last := intPts[len(intPts)-1]
						tr.drawArrowOnPath(last[0], last[1], intPts, bc, pw, s.tailEnd)
					}
				}
			} else {
				tr.drawRectBorder(rect, argbToRGBA(s.border.Color), pw, s.border.Style)
			}
		} else if s.customPath != nil && (s.headEnd != nil || s.tailEnd != nil) {
			// No visible border but has arrowheads — still need to draw them along the path
			pts := tr.customPathToPixelPoints(s.customPath, ox, oy, w, h)
			if len(pts) >= 2 {
				pw := maxInt(int(tr.scaleX*12700.0), 1)
				bc := color.RGBA{A: 255} // default black
				if s.border != nil {
					bc = argbToRGBA(s.border.Color)
				}
				intPts := make([][2]int, len(pts))
				for i, p := range pts {
					intPts[i] = [2]int{int(p.x), int(p.y)}
				}
				if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
					tr.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, bc, pw, s.headEnd)
				}
				if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
					last := intPts[len(intPts)-1]
					tr.drawArrowOnPath(last[0], last[1], intPts, bc, pw, s.tailEnd)
				}
			}
		}

		// Text area with insets applied; use bufH to allow overflow
		tx := ox + pxL
		ty := oy + pxT
		drawTH := bufH - pxT - pxB
		if drawTH < th {
			drawTH = th
		}
		// For middle-anchored text that overflows the inset area, center
		// relative to the full shape height so text doesn't appear shifted
		// down by the top inset. PowerPoint centers overflowing text within
		// the shape bounds, not the inset-reduced text body.
		if s.textAnchor == TextAnchorMiddle && overflowH > 0 {
			ty = oy
			drawTH = h
		}
		// For bottom-anchored text, keep the original text area height so
		// that drawParagraphs computes startY = ty + th - totalH, which
		// places text above the shape when it overflows. Expanding drawTH
		// by overflowH would cancel out the upward offset.
		if s.textAnchor == TextAnchorBottom && drawTH > th {
			drawTH = th
		}

		if !skipText {
			if vertRotation != 0 {
				// For vertical text, draw into a rotated buffer with swapped dimensions.
				vtw, vth := drawTH, tw // text area: width=drawTH, height=tw (before rotation)
				if vtw > 0 && vth > 0 {
					tmp := image.NewRGBA(image.Rect(0, 0, vtw, vth))
					tmpR := tr.subRenderer(tmp)
					tmpR.drawParagraphs(s.paragraphs, 0, 0, vtw, vth, s.textAnchor, wordWrap)
					rotateAndComposite(tr.img, tmp, tx, ty, tw, drawTH, vertRotation)
				}
			} else {
				tr.drawParagraphs(s.paragraphs, tx, ty, tw, drawTH, s.textAnchor, wordWrap)
			}
		}
	}

	// When flip is set, PowerPoint flips the shape geometry (fill/border)
	// but keeps text readable (un-flipped). We achieve this by rendering
	// geometry with flip, then compositing text separately without flip.
	if (flipH || flipV) && len(s.paragraphs) > 0 {
		// Phase 1: render geometry only (with flip)
		skipText = true
		r.renderRotatedExpanded(x, y, w, h, bufH, rotation, flipH, flipV, drawContent)
		// Phase 2: render text only (rotation only, no flip)
		skipText = false
		textOnly := func(tr *renderer) {
			ox, oy := x, y
			if tr != r {
				ox, oy = 0, 0
			}
			tx := ox + pxL
			ty := oy + pxT
			drawTH := bufH - pxT - pxB
			if drawTH < th {
				drawTH = th
			}
			if s.textAnchor == TextAnchorMiddle && overflowH > 0 {
				ty = oy
				drawTH = h
			}
			if s.textAnchor == TextAnchorBottom && drawTH > th {
				drawTH = th
			}
			if vertRotation != 0 {
				vtw, vth := drawTH, tw
				if vtw > 0 && vth > 0 {
					tmp := image.NewRGBA(image.Rect(0, 0, vtw, vth))
					tmpR := tr.subRenderer(tmp)
					tmpR.drawParagraphs(s.paragraphs, 0, 0, vtw, vth, s.textAnchor, wordWrap)
					rotateAndComposite(tr.img, tmp, tx, ty, tw, drawTH, vertRotation)
				}
			} else {
				tr.drawParagraphs(s.paragraphs, tx, ty, tw, drawTH, s.textAnchor, wordWrap)
			}
		}
		if rotation != 0 {
			r.renderRotatedExpanded(x, y, w, h, bufH, rotation, false, false, textOnly)
		} else {
			textOnly(r)
		}
	} else if rotation != 0 {
		r.renderRotatedExpanded(x, y, w, h, bufH, rotation, false, false, drawContent)
	} else {
		drawContent(r)
	}
}

// looksLikeSVG reports whether picture bytes are an SVG document. The test is
// content-based because a DrawingShape carries bytes, not the part's content
// type, so the file extension that named the part is no longer available here.
func looksLikeSVG(data []byte) bool {
	s := data
	if i := bytes.IndexByte(s, '<'); i >= 0 {
		s = s[i:]
	} else {
		return false
	}
	if bytes.HasPrefix(s, []byte("<svg")) {
		return true
	}
	return bytes.HasPrefix(s, []byte("<?xml")) && bytes.Contains(s, []byte("<svg"))
}

// renderPicturePlaceholder marks a picture that could not be rasterised.
//
// It reuses the unsupported-shape treatment — amber, dashed, labelled — because
// this is the same kind of event: part of the document is not in the preview.
// Drawing nothing, or a neutral grey frame, hides that, and a grey frame in
// particular reads as a picture that legitimately has a grey border.
func (r *renderer) renderPicturePlaceholder(x, y, w, h int, label string) {
	if w <= 0 || h <= 0 {
		return
	}
	rect := image.Rect(x, y, x+w, y+h)
	r.fillRectBlend(rect, unsupportedFill)
	r.drawRectBorder(rect, unsupportedBorder, placeholderBorderWidth(w, h), BorderDash)
	r.drawPlaceholderLabel(label, rect)
}

func (r *renderer) renderDrawing(s *DrawingShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	imgData := s.data
	if len(imgData) == 0 && s.path != "" {
		if data, err := os.ReadFile(s.path); err == nil {
			imgData = data
		}
	}
	if len(imgData) == 0 {
		return
	}

	srcImg, _, err := image.Decode(bytes.NewReader(imgData))
	if err != nil {
		// Try to extract bitmap from WMF/EMF metafiles
		if extracted := decodeMetafileBitmap(imgData, r.fontCache); extracted != nil {
			srcImg = extracted
			err = nil
		}
	}
	// SVG needs its own rasteriser; without it, every icon and connector line a
	// design pipeline emits as SVG disappears from the preview.
	pictureLabel := "Undecodable image"
	if err != nil && looksLikeSVG(imgData) {
		if svgImg, svgErr := renderSVG(imgData, w, h); svgErr == nil {
			srcImg, err = svgImg, nil
		} else {
			pictureLabel = "SVG: " + strings.TrimPrefix(svgErr.Error(), "svg: ")
		}
	}
	if err != nil {
		r.renderPicturePlaceholder(x, y, w, h, pictureLabel)
		return
	}

	// <a:clrChange> recolours source pixels before anything else sees them:
	// photos stacked on top of each other knock their background out with it
	// and rely on the transparency surviving the crop and the scale.
	if len(s.recolors) > 0 {
		srcImg = applyColorReplaces(srcImg, s.recolors)
	}

	// <a:srcRect> crop, kept as a float sub-rectangle of the source image.
	// The resampler samples it directly (fractional pixel edges included),
	// which is what PowerPoint's pipeline does; truncating to whole pixels
	// first shifts the content by up to a pixel after scaling.
	var cropF *[4]float64
	if s.cropLeft > 0 || s.cropTop > 0 || s.cropRight > 0 || s.cropBottom > 0 {
		b := srcImg.Bounds()
		iw, ih := float64(b.Dx()), float64(b.Dy())
		cropF = &[4]float64{
			iw * float64(s.cropLeft) / 100000.0,
			ih * float64(s.cropTop) / 100000.0,
			iw - iw*float64(s.cropRight)/100000.0,
			ih - ih*float64(s.cropBottom)/100000.0,
		}
		if (*cropF)[2]-(*cropF)[0] < 1 || (*cropF)[3]-(*cropF)[1] < 1 {
			cropF = nil // degenerate crop: draw the whole image
		}
	}

	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	// The frame outline the picture is clipped to: nil for the plain
	// rectangle (including "" — most pictures carry no prstGeom at all).
	framePts := r.picFramePoints(s.presetGeom, x, y, w, h)

	drawImg := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		pts := framePts
		if tr != r && pts != nil {
			// The rotated renderer works in unrotated space from the origin.
			pts = r.picFramePoints(s.presetGeom, 0, 0, w, h)
		}
		// Shadow first — PowerPoint composites it behind the picture.
		if s.shadow != nil && s.shadow.Visible {
			if pts != nil {
				tr.renderShadowPolygon(s.shadow, pts)
			} else {
				tr.renderShadow(s.shadow, image.Rect(ox, oy, ox+w, oy+h))
			}
		}
		scaledImg := tr.scaleForRender(srcImg, w, h, cropF)
		// <a:lum bright contrast> adjusts the picture's own pixels. It is
		// applied here rather than to the decoded image because it is a
		// per-channel point operation: scaling first costs one pass over the
		// destination instead of one over a 2560x1920 photograph.
		if s.lumBright != 0 || s.lumContrast != 0 {
			applyLumAdjust(scaledImg, s.lumBright, s.lumContrast)
		}
		// Apply alphaModFix opacity if set (value is in 1/1000 of a percent, e.g. 5000 = 5%)
		if s.alpha > 0 && s.alpha < 100000 {
			alphaScale := float64(s.alpha) / 100000.0
			bounds := scaledImg.Bounds()
			for py := bounds.Min.Y; py < bounds.Max.Y; py++ {
				for px := bounds.Min.X; px < bounds.Max.X; px++ {
					c := scaledImg.RGBAAt(px, py)
					// Scale all channels (premultiplied alpha format)
					c.R = uint8(float64(c.R) * alphaScale)
					c.G = uint8(float64(c.G) * alphaScale)
					c.B = uint8(float64(c.B) * alphaScale)
					c.A = uint8(float64(c.A) * alphaScale)
					scaledImg.SetRGBA(px, py, c)
				}
			}
		}
		// <a:duotone> repaints every pixel on the line between its two
		// colours at the pixel's luma — the recolour that turns a plain
		// screenshot into a framed, toned exhibit.
		if s.hasDuotone {
			applyDuotone(scaledImg, s.duotoneA, s.duotoneB)
		}
		if pts != nil {
			// pts are in the same space as the destination rect: absolute for
			// the unrotated renderer, origin-based for a rotated one (where
			// ox,oy are 0). Either way the mask is local, so shift by -ox,-oy.
			local := polygonAlphaMask(shiftPts(pts, -float64(ox), -float64(oy)), w, h)
			draw.DrawMask(tr.img, image.Rect(ox, oy, ox+w, oy+h), scaledImg, image.Point{}, local, image.Point{}, draw.Over)
		} else {
			draw.Draw(tr.img, image.Rect(ox, oy, ox+w, oy+h), scaledImg, image.Point{}, draw.Over)
		}
		// Frame line — drawn on the outline, like a shape's border.
		if s.border != nil && s.border.Style != BorderNone {
			bc := argbToRGBA(s.border.Color)
			pw := maxInt(int(float64(maxInt(s.border.Width, 1))*12700.0*tr.scaleX), 1)
			if pts != nil {
				tr.drawPolygon(pts, bc, pw)
			} else {
				tr.drawRectBorder(image.Rect(ox, oy, ox+w, oy+h), bc, pw, BorderSolid)
			}
		}
	}

	if rotation != 0 || flipH || flipV {
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawImg)
	} else {
		drawImg(r)
	}
}

// applyColorReplaces applies a picture's <a:clrChange> rules to its decoded
// pixels, in the order the file wrote them, first match wins per pixel.
//
// The comparison is on non-premultiplied colour: a rule names the pixel's
// colour as stored, not as alpha-scaled, and the replacement may well change
// the alpha (a knock-out sets it to zero), which would be lost on premultiplied
// values.
func applyColorReplaces(src image.Image, rules []colorReplace) image.Image {
	type rule struct {
		fr, fg, fb uint8
		tr, tg, tb uint8
		ta         float64 // -1: keep the source pixel's alpha
	}
	parsed := make([]rule, 0, len(rules))
	for _, cr := range rules {
		from, okFrom := hex6(cr.From)
		to, okTo := hex6(cr.To)
		if !okFrom || !okTo {
			continue
		}
		ta := -1.0
		if cr.ToAlpha >= 0 {
			ta = float64(cr.ToAlpha) / 100000.0
		}
		parsed = append(parsed, rule{from[0], from[1], from[2], to[0], to[1], to[2], ta})
	}
	if len(parsed) == 0 {
		return src
	}
	b := src.Bounds()
	dst := image.NewNRGBA(image.Rect(0, 0, b.Dx(), b.Dy()))
	draw.Draw(dst, dst.Bounds(), src, b.Min, draw.Src)
	for py := 0; py < b.Dy(); py++ {
		for px := 0; px < b.Dx(); px++ {
			i := dst.PixOffset(px, py)
			r, g, bl, a := dst.Pix[i], dst.Pix[i+1], dst.Pix[i+2], dst.Pix[i+3]
			for _, u := range parsed {
				if r == u.fr && g == u.fg && bl == u.fb {
					r, g, bl = u.tr, u.tg, u.tb
					if u.ta >= 0 {
						a = uint8(u.ta*255.0 + 0.5)
					}
					break
				}
			}
			dst.Pix[i], dst.Pix[i+1], dst.Pix[i+2], dst.Pix[i+3] = r, g, bl, a
		}
	}
	return dst
}

// hex6 parses six hex digits into an RGB triple. It is the form both
// <a:srgbClr val="..."> and the recolour rules on a drawing use.
func hex6(s string) ([3]uint8, bool) {
	var out [3]uint8
	if len(s) != 6 {
		return out, false
	}
	for i := 0; i < 3; i++ {
		hi := hexVal(s[i*2])
		lo := hexVal(s[i*2+1])
		if hi < 0 || lo < 0 {
			return out, false
		}
		out[i] = uint8(hi<<4 | lo)
	}
	return out, true
}

func (r *renderer) renderAutoShape(s *AutoShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	// Apply normAutofit font scale
	prevFontScale := r.fontScale
	if s.fontScale > 0 && s.fontScale != 100000 {
		r.fontScale = float64(s.fontScale) / 100000.0
	}
	// Same for the normAutofit line-spacing reduction (renderRichText's
	// mirror — the two must not drift).
	prevLnSpcReduction := r.lnSpcReduction
	if s.lnSpcReduction > 0 && s.lnSpcReduction != 100000 {
		r.lnSpcReduction = float64(s.lnSpcReduction) / 100000.0
	}
	defer func() {
		r.fontScale = prevFontScale
		r.lnSpcReduction = prevLnSpcReduction
	}()

	// Vertical text direction
	vertRotation := 0
	if s.textDirection == "vert" || s.textDirection == "eaVert" || s.textDirection == "wordArtVert" {
		vertRotation = 270
	} else if s.textDirection == "vert270" {
		vertRotation = 90
	}

	// The shape's own <a:bodyPr wrap>, not a hard-coded true: a deck that says
	// wrap="none" must not have its shape text wrapped for it.
	wordWrap := s.wordWrap

	drawContent := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		rect := image.Rect(ox, oy, ox+w, oy+h)
		if s.shadow != nil && s.shadow.Visible {
			switch s.shapeType {
			case AutoShapeRoundedRect, AutoShapeCallout1:
				sRadius := minInt(w, h) * 16667 / 100000
				if s.adjustValues != nil {
					if adj, ok := s.adjustValues["adj"]; ok {
						sRadius = minInt(w, h) * adj / 200000
					}
					if adj, ok := s.adjustValues["adj3"]; ok && s.shapeType == AutoShapeCallout1 {
						sRadius = int(math.Min(float64(w), float64(h)) * float64(adj) / 100000.0)
					}
				}
				tr.renderShadowRounded(s.shadow, rect, sRadius)
			case AutoShapeRectangle, "":
				tr.renderShadow(s.shadow, rect)
			default:
				// For non-rectangular shapes (arrows, triangles, ellipses, etc.),
				// skip the rectangular shadow — it would fill the entire
				// bounding box and look like a gray background.
			}
		}
		tr.renderAutoShapeFill(s, ox, oy, w, h)
		tr.renderAutoShapeBorder(s, ox, oy, w, h)
		// Arc shapes are stroke-only; if no explicit border was set, draw
		// the arc with a default black stroke so it remains visible.
		if s.shapeType == AutoShapeArc && (s.border == nil || s.border.Style == BorderNone) {
			defPw := maxInt(int(tr.scaleX*12700.0), 1)
			defC := color.RGBA{A: 255}
			tr.renderArcBorder(s, ox, oy, w, h, defC, defPw)
		}
		if len(s.paragraphs) > 0 {
			// Compute text area with insets
			lIns, rIns, tIns, bIns := int64(91440), int64(91440), int64(45720), int64(45720)
			if s.insetsSet {
				lIns, rIns, tIns, bIns = s.insetLeft, s.insetRight, s.insetTop, s.insetBottom
			}
			pxL := r.emuToPixelX(lIns)
			pxR := r.emuToPixelX(rIns)
			pxT := r.emuToPixelY(tIns)
			pxB := r.emuToPixelY(bIns)

			// Clamp default insets when they consume too much of the shape dimensions.
			if !s.insetsSet {
				maxInsetH := int(float64(h) * 0.35)
				maxInsetW := int(float64(w) * 0.35)
				if pxT+pxB > maxInsetH {
					scale := float64(maxInsetH) / float64(pxT+pxB)
					pxT = int(float64(pxT) * scale)
					pxB = int(float64(pxB) * scale)
				}
				if pxL+pxR > maxInsetW {
					scale := float64(maxInsetW) / float64(pxL+pxR)
					pxL = int(float64(pxL) * scale)
					pxR = int(float64(pxR) * scale)
				}
			}

			tx, ty, tw, th := ox+pxL, oy+pxT, w-pxL-pxR, h-pxT-pxB

			// For ellipses, further constrain text to the inscribed rectangle
			// The inscribed rect of an ellipse insets by factor (1 - 1/√2) ≈ 0.2929
			if s.shapeType == AutoShapeEllipse {
				insetX := int(float64(w) * 0.1464) // half of 0.2929
				insetY := int(float64(h) * 0.1464)
				etx := ox + insetX
				ety := oy + insetY
				etw := w - 2*insetX
				eth := h - 2*insetY
				// Use the tighter of explicit insets vs ellipse inscribed rect
				if etx > tx {
					tx = etx
				}
				if ety > ty {
					ty = ety
				}
				if etx+etw < ox+pxL+tw {
					tw = etx + etw - tx
				}
				if ety+eth < oy+pxT+th {
					th = ety + eth - ty
				}
			}

			if tw < 1 {
				tw = w
			}
			if th < 1 {
				th = h
			}

			// When default insets are used and text overflows, reduce insets
			// to make room. This handles font metric differences between systems.
			if !s.insetsSet {
				textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
				if textH > th && th > 0 && (pxT+pxB) > 0 {
					needed := textH - th
					avail := pxT + pxB
					if needed >= avail {
						pxT = 0
						pxB = 0
					} else {
						sc := float64(avail-needed) / float64(avail)
						pxT = int(float64(pxT) * sc)
						pxB = int(float64(pxB) * sc)
					}
					tx = ox + pxL
					ty = oy + pxT
					th = h - pxT - pxB
					if th < 1 {
						th = h
					}
				}
			}

			// Auto-shrink when text overflows the full shape height —
			// CJK font metrics in Go are often larger than PowerPoint's.
			// Use a conservative floor to avoid making text too small.
			if s.fontScale == 0 || s.fontScale == 100000 {
				atextH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
				if atextH > h && h > 0 && atextH > th && th > 0 {
					lo, hi := 0.65, 1.0
					for i := 0; i < 10; i++ {
						mid := (lo + hi) / 2
						r.fontScale = mid
						mh := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
						if mh > th {
							hi = mid
						} else {
							lo = mid
						}
					}
					r.fontScale = lo
				}
			}

			// Horizontal overflow: shrink font when wrapped lines still
			// exceed the text area width due to font metric differences.
			// Apply the same 3% tolerance used by wrapRunLine.
			if tw > 0 && (s.fontScale == 0 || s.fontScale == 100000) {
				hTol := tw * 103 / 100
				maxLW := r.measureMaxLineWidth(s.paragraphs, tw, wordWrap)
				if maxLW > hTol {
					lo, hi := 0.5, r.fontScale
					if hi <= 0 {
						hi = 1.0
					}
					for i := 0; i < 12; i++ {
						mid := (lo + hi) / 2
						r.fontScale = mid
						mw := r.measureMaxLineWidth(s.paragraphs, tw, wordWrap)
						if mw > hTol {
							hi = mid
						} else {
							lo = mid
						}
					}
					r.fontScale = lo
				}
			}

			if vertRotation != 0 {
				vtw, vth := th, tw
				if vtw > 0 && vth > 0 {
					tmp := image.NewRGBA(image.Rect(0, 0, vtw, vth))
					tmpR := tr.subRenderer(tmp)
					tmpR.drawParagraphs(s.paragraphs, 0, 0, vtw, vth, s.textAnchor, wordWrap)
					rotateAndComposite(tr.img, tmp, tx, ty, tw, th, vertRotation)
				}
			} else {
				tr.drawParagraphs(s.paragraphs, tx, ty, tw, th, s.textAnchor, wordWrap)
			}
		} else if s.text != "" {
			tr.drawStringCentered(s.text, tr.getFace(NewFont()), color.RGBA{A: 255}, rect)
		}
	}

	// For rtTriangle with 90/270 rotation, OOXML ext gives the rotated
	// bounding box size. Draw the mirror-image triangle in the buffer so
	// that after rotation the filled area covers the correct half.
	// With correct clockwise rotation, the standard triangle vertices
	// need to be adjusted for the swapped buffer dimensions.
	needsRtTriSwap := s.shapeType == AutoShapeRtTriangle &&
		(rotation == 90 || rotation == 270)

	if needsRtTriSwap {
		drawSwapped := func(tr *renderer) {
			if s.fill != nil && s.fill.Type != FillNone {
				fc := argbToRGBA(s.fill.Color)
				fc = tr.scaleAlpha(fc)
				// Draw mirror-image triangle that, after correct clockwise
				// rotation, produces the expected right-triangle orientation.
				pts := []fpoint{
					{0, 0},
					{0, float64(h)},
					{float64(w), float64(h)},
				}
				tr.fillPolygon(pts, fc)
			}
		}
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawSwapped)
	} else if (flipH || flipV) && len(s.paragraphs) > 0 {
		// PowerPoint flips shape geometry but keeps text readable (un-flipped).
		// Phase 1: render geometry only (fill + border) with flip applied.
		drawGeomOnly := func(tr *renderer) {
			ox, oy := x, y
			if tr != r {
				ox, oy = 0, 0
			}
			rect := image.Rect(ox, oy, ox+w, oy+h)
			if s.shadow != nil && s.shadow.Visible {
				switch s.shapeType {
				case AutoShapeRoundedRect, AutoShapeCallout1:
					sRadius := minInt(w, h) * 16667 / 100000
					if s.adjustValues != nil {
						if adj, ok := s.adjustValues["adj"]; ok {
							sRadius = minInt(w, h) * adj / 200000
						}
						if adj, ok := s.adjustValues["adj3"]; ok && s.shapeType == AutoShapeCallout1 {
							sRadius = int(math.Min(float64(w), float64(h)) * float64(adj) / 100000.0)
						}
					}
					tr.renderShadowRounded(s.shadow, rect, sRadius)
				case AutoShapeRectangle, "":
					tr.renderShadow(s.shadow, rect)
				default:
				}
			}
			tr.renderAutoShapeFill(s, ox, oy, w, h)
			tr.renderAutoShapeBorder(s, ox, oy, w, h)
			if s.shapeType == AutoShapeArc && (s.border == nil || s.border.Style == BorderNone) {
				defPw := maxInt(int(tr.scaleX*12700.0), 1)
				defC := color.RGBA{A: 255}
				tr.renderArcBorder(s, ox, oy, w, h, defC, defPw)
			}
		}
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawGeomOnly)

		// Phase 2: render text only (rotation only, no flip) so text stays readable.
		drawTextOnly := func(tr *renderer) {
			ox, oy := x, y
			if tr != r {
				ox, oy = 0, 0
			}
			lIns, rIns, tIns, bIns := int64(91440), int64(91440), int64(45720), int64(45720)
			if s.insetsSet {
				lIns, rIns, tIns, bIns = s.insetLeft, s.insetRight, s.insetTop, s.insetBottom
			}
			pxL := r.emuToPixelX(lIns)
			pxR := r.emuToPixelX(rIns)
			pxT := r.emuToPixelY(tIns)
			pxB := r.emuToPixelY(bIns)
			if !s.insetsSet {
				maxInsetH := int(float64(h) * 0.35)
				maxInsetW := int(float64(w) * 0.35)
				if pxT+pxB > maxInsetH {
					scale := float64(maxInsetH) / float64(pxT+pxB)
					pxT = int(float64(pxT) * scale)
					pxB = int(float64(pxB) * scale)
				}
				if pxL+pxR > maxInsetW {
					scale := float64(maxInsetW) / float64(pxL+pxR)
					pxL = int(float64(pxL) * scale)
					pxR = int(float64(pxR) * scale)
				}
			}
			tx, ty, tw, th := ox+pxL, oy+pxT, w-pxL-pxR, h-pxT-pxB
			if s.shapeType == AutoShapeEllipse {
				insetX := int(float64(w) * 0.1464)
				insetY := int(float64(h) * 0.1464)
				etx := ox + insetX
				ety := oy + insetY
				etw := w - 2*insetX
				eth := h - 2*insetY
				if etx > tx {
					tx = etx
				}
				if ety > ty {
					ty = ety
				}
				if etx+etw < ox+pxL+tw {
					tw = etx + etw - tx
				}
				if ety+eth < oy+pxT+th {
					th = ety + eth - ty
				}
			}
			if tw < 1 {
				tw = w
			}
			if th < 1 {
				th = h
			}
			if !s.insetsSet {
				textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
				if textH > th && th > 0 && (pxT+pxB) > 0 {
					needed := textH - th
					avail := pxT + pxB
					if needed >= avail {
						pxT = 0
						pxB = 0
					} else {
						sc := float64(avail-needed) / float64(avail)
						pxT = int(float64(pxT) * sc)
						pxB = int(float64(pxB) * sc)
					}
					tx = ox + pxL
					ty = oy + pxT
					th = h - pxT - pxB
					if th < 1 {
						th = h
					}
				}
			}
			// Auto-shrink when text overflows
			if s.fontScale == 0 || s.fontScale == 100000 {
				atextH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
				if atextH > h && h > 0 && atextH > th && th > 0 {
					lo, hi := 0.65, 1.0
					for i := 0; i < 10; i++ {
						mid := (lo + hi) / 2
						r.fontScale = mid
						mh := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
						if mh > th {
							hi = mid
						} else {
							lo = mid
						}
					}
					r.fontScale = lo
				}
				// Horizontal overflow — apply 3% tolerance matching wrapRunLine
				hTol := tw * 103 / 100
				maxLW := r.measureMaxLineWidth(s.paragraphs, tw, wordWrap)
				if maxLW > hTol && tw > 0 {
					lo, hi := 0.5, r.fontScale
					if hi <= 0 {
						hi = 1.0
					}
					for i := 0; i < 12; i++ {
						mid := (lo + hi) / 2
						r.fontScale = mid
						mw := r.measureMaxLineWidth(s.paragraphs, tw, wordWrap)
						if mw > hTol {
							hi = mid
						} else {
							lo = mid
						}
					}
					r.fontScale = lo
				}
			}
			if vertRotation != 0 {
				vtw, vth := th, tw
				if vtw > 0 && vth > 0 {
					tmp := image.NewRGBA(image.Rect(0, 0, vtw, vth))
					tmpR := tr.subRenderer(tmp)
					tmpR.drawParagraphs(s.paragraphs, 0, 0, vtw, vth, s.textAnchor, wordWrap)
					rotateAndComposite(tr.img, tmp, tx, ty, tw, th, vertRotation)
				}
			} else {
				tr.drawParagraphs(s.paragraphs, tx, ty, tw, th, s.textAnchor, wordWrap)
			}
		}
		if rotation != 0 {
			r.renderRotated(x, y, w, h, rotation, false, false, drawTextOnly)
		} else {
			drawTextOnly(r)
		}
	} else if rotation != 0 || flipH || flipV {
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawContent)
	} else {
		drawContent(r)
	}
}

func (r *renderer) renderAutoShapeFill(s *AutoShape, x, y, w, h int) {
	if s.fill == nil || s.fill.Type == FillNone {
		return
	}
	fc := argbToRGBA(s.fill.Color)
	fc = r.scaleAlpha(fc)
	rect := image.Rect(x, y, x+w, y+h)

	switch s.shapeType {
	case AutoShapeEllipse:
		if s.fill.Type == FillSolid {
			r.fillEllipseAA(x, y, w, h, fc)
		} else {
			r.fillGradientLinear(rect, s.fill)
		}
	case AutoShapeRoundedRect:
		radius := minInt(w, h) * 16667 / 100000
		if s.adjustValues != nil {
			if adj, ok := s.adjustValues["adj"]; ok {
				radius = minInt(w, h) * adj / 200000
			}
		}
		if s.fill.Type == FillSolid {
			r.fillRoundedRect(x, y, w, h, radius, fc)
		} else {
			r.fillGradientLinear(rect, s.fill)
		}
	case AutoShapeTriangle:
		r.fillTriangle(x, y, w, h, fc)
	case AutoShapeDiamond:
		r.fillDiamond(x, y, w, h, fc)
	case AutoShapeHexagon:
		r.fillHexagon(x, y, w, h, fc)
	case AutoShapeFlowchartPreparation:
		r.fillFlowChartPreparation(x, y, w, h, fc)
	case AutoShapePentagon:
		r.fillPentagon(x, y, w, h, fc)
	case AutoShapeArrowRight:
		r.fillArrowRight(x, y, w, h, fc)
	case AutoShapeArrowLeft:
		r.fillArrowLeft(x, y, w, h, fc)
	case AutoShapeArrowUp:
		r.fillArrowUp(x, y, w, h, fc)
	case AutoShapeArrowDown:
		r.fillArrowDown(x, y, w, h, fc)
	case AutoShapeStar5:
		r.fillStar(x, y, w, h, 5, fc)
	case AutoShapeStar4:
		r.fillStar(x, y, w, h, 4, fc)
	case AutoShapeHeart:
		r.fillHeart(x, y, w, h, fc)
	case AutoShapePlus:
		r.fillPlus(x, y, w, h, fc)
	case AutoShapeChevron:
		r.fillChevron(x, y, w, h, fc)
	case AutoShapeParallelogram:
		r.fillParallelogram(x, y, w, h, fc)
	case AutoShapeLeftRightArrow:
		r.fillLeftRightArrow(x, y, w, h, fc)
	case AutoShapeRtTriangle:
		r.fillRtTriangle(x, y, w, h, fc)
	case AutoShapeHomePlate:
		r.fillHomePlate(x, y, w, h, fc)
	case AutoShapeCallout1:
		r.fillWedgeRoundRectCallout(x, y, w, h, fc, s.adjustValues)
	case AutoShapeSnip2SameRect:
		r.fillSnip2SameRect(x, y, w, h, fc, s.adjustValues)
	case AutoShapeUturnArrow:
		r.fillUturnArrow(x, y, w, h, fc, s.adjustValues)
	case AutoShapeBentArrow:
		r.fillBentArrow(x, y, w, h, fc, s.adjustValues)
	case AutoShapeBentUpArrow:
		pts := r.bentUpArrowPoints(x, y, w, h, s.adjustValues)
		if s.fill.Type == FillSolid {
			r.fillPolygon(pts, fc)
		} else {
			r.fillPolygonGradient(pts, image.Rect(x, y, x+w, y+h), s.fill)
		}
	case AutoShapeSnip2DiagRect:
		pts := r.snip2DiagRectPoints(x, y, w, h, s.adjustValues)
		if s.fill.Type == FillSolid {
			r.fillPolygon(pts, fc)
		} else {
			r.fillPolygonGradient(pts, image.Rect(x, y, x+w, y+h), s.fill)
		}
	case AutoShapeArc:
		// Arc preset geometry has no fill by default (it's just a stroke).
		// Skip fill for arc shapes.
	case AutoShapeCan, AutoShapeFlowChartMagneticDisk:
		r.fillCan(x, y, w, h, fc, s.adjustValues)
	default:
		r.renderFill(s.fill, rect)
	}
}

func (r *renderer) renderAutoShapeBorder(s *AutoShape, x, y, w, h int) {
	if s.border == nil || s.border.Style == BorderNone {
		return
	}
	bc := argbToRGBA(s.border.Color)
	pw := maxInt(int(float64(maxInt(s.border.Width, 1))*12700.0*r.scaleX), 1)

	switch s.shapeType {
	case AutoShapeCan, AutoShapeFlowChartMagneticDisk:
		r.drawCan(x, y, w, h, bc, pw, s.adjustValues)
	case AutoShapeEllipse:
		r.drawEllipseAA(x, y, w, h, bc, pw)
	case AutoShapeRoundedRect:
		radius := minInt(w, h) * 16667 / 100000
		if s.adjustValues != nil {
			if adj, ok := s.adjustValues["adj"]; ok {
				radius = minInt(w, h) * adj / 200000
			}
		}
		r.drawRoundedRect(x, y, w, h, radius, bc, pw)
	case AutoShapeTriangle:
		r.drawTriangle(x, y, w, h, bc, pw)
	case AutoShapeDiamond:
		r.drawDiamond(x, y, w, h, bc, pw)
	case AutoShapeFlowchartPreparation:
		pts := flowChartPreparationPoints(x, y, w, h)
		r.drawPolygon(pts, bc, pw)
	case AutoShapeChevron:
		notch := w / 4
		pts := []fpoint{
			{float64(x), float64(y)},
			{float64(x + w - notch), float64(y)},
			{float64(x + w), float64(y + h/2)},
			{float64(x + w - notch), float64(y + h)},
			{float64(x), float64(y + h)},
			{float64(x + notch), float64(y + h/2)},
		}
		r.drawPolygon(pts, bc, pw)
	case AutoShapeParallelogram:
		offset := w / 4
		pts := []fpoint{
			{float64(x + offset), float64(y)},
			{float64(x + w), float64(y)},
			{float64(x + w - offset), float64(y + h)},
			{float64(x), float64(y + h)},
		}
		r.drawPolygon(pts, bc, pw)
	case AutoShapeBentUpArrow:
		pts := r.bentUpArrowPoints(x, y, w, h, s.adjustValues)
		r.drawPolygon(pts, bc, pw)
	case AutoShapeSnip2DiagRect:
		pts := r.snip2DiagRectPoints(x, y, w, h, s.adjustValues)
		r.drawPolygon(pts, bc, pw)
	case AutoShapeBentArrow:
		// Draw border following the bentArrow shape outline
		adj1v, adj2v, adj3v, adj4v := 25000, 25000, 25000, 43750
		if s.adjustValues != nil {
			if v, ok := s.adjustValues["adj1"]; ok {
				adj1v = v
			}
			if v, ok := s.adjustValues["adj2"]; ok {
				adj2v = v
			}
			if v, ok := s.adjustValues["adj3"]; ok {
				adj3v = v
			}
			if v, ok := s.adjustValues["adj4"]; ok {
				adj4v = v
			}
		}
		fx, fy := float64(x), float64(y)
		fw, fh := float64(w), float64(h)
		shaftW := fw * float64(adj1v) / 100000.0
		headExtra := fw * float64(adj2v) / 100000.0
		headLen := fw * float64(adj3v) / 100000.0
		bendYf := fy + fh*float64(adj4v)/100000.0
		tipX := fx + fw
		arrowCenterY := bendYf - shaftW/2
		arrowBaseX := tipX - headLen
		arrowTop := arrowCenterY - shaftW/2 - headExtra
		arrowBot := arrowCenterY + shaftW/2 + headExtra
		cornerR := shaftW * 0.85
		if cornerR < 1 {
			cornerR = 1
		}
		bpts := []fpoint{{fx, fy + fh}}
		// Outer corner arc
		outerR := cornerR
		maxOR := math.Min(bendYf-shaftW-fy, fw*0.3)
		if outerR > maxOR && maxOR > 0 {
			outerR = maxOR
		}
		ocx := fx + outerR
		ocy := bendYf - shaftW + outerR
		bpts = append(bpts, fpoint{fx, ocy})
		for i := 0; i <= 12; i++ {
			t := float64(i) / 12.0
			a := math.Pi + t*math.Pi/2.0
			bpts = append(bpts, fpoint{ocx + outerR*math.Cos(a), ocy + outerR*math.Sin(a)})
		}
		bpts = append(bpts,
			fpoint{arrowBaseX, bendYf - shaftW},
			fpoint{arrowBaseX, arrowTop},
			fpoint{tipX, arrowCenterY},
			fpoint{arrowBaseX, arrowBot},
			fpoint{arrowBaseX, bendYf},
		)
		// Inner corner arc
		innerR := cornerR
		maxIR := math.Min(fh-fh*float64(adj4v)/100000.0, shaftW*0.9)
		if innerR > maxIR && maxIR > 0 {
			innerR = maxIR
		}
		icx := fx + shaftW + innerR
		icy := bendYf + innerR
		bpts = append(bpts, fpoint{icx, bendYf})
		for i := 0; i <= 12; i++ {
			t := float64(i) / 12.0
			a := math.Pi/2.0 + t*math.Pi/2.0
			bpts = append(bpts, fpoint{icx + innerR*math.Cos(a), icy - innerR*math.Sin(a)})
		}
		bpts = append(bpts, fpoint{fx + shaftW, fy + fh})
		r.drawPolygon(bpts, bc, pw)
	case AutoShapeRtTriangle:
		pts := []fpoint{
			{float64(x), float64(y + h)},
			{float64(x), float64(y)},
			{float64(x + w), float64(y + h)},
		}
		r.drawPolygon(pts, bc, pw)
	case AutoShapeSnip2SameRect:
		pts := r.snip2SameRectPoints(x, y, w, h, s.adjustValues)
		r.drawPolygon(pts, bc, pw)
	case AutoShapeCallout1:
		r.drawWedgeRoundRectCalloutBorder(x, y, w, h, bc, pw, s.adjustValues)
	case AutoShapeArc:
		r.renderArcBorder(s, x, y, w, h, bc, pw)
	case AutoShapeRightBrace, AutoShapeLeftBrace:
		r.drawBrace(s.shapeType, x, y, w, h, bc, pw, s.adjustValues)
	default:
		r.drawRectBorder(image.Rect(x, y, x+w, y+h), bc, pw, s.border.Style)
	}
}

// drawBrace strokes a rightBrace/leftBrace preset: a corner hook curving into
// a centre spine and mirroring back out at the bottom. OOXML geometry —
// x1 = adj1·min(w,h)/100000 (hook depth and roundness), the spine sits at
// w−x1 (right) or x1 (left), y1 = adj2·(h/2)/100000 down from the top edge,
// y2 = h−y1. Each hook is one cubic Bezier; the OOXML path declares fill
// none, so a brace is outline-only by default.
func (r *renderer) drawBrace(t AutoShapeType, x, y, w, h int, c color.RGBA, pw int, adj map[string]int) {
	adj1, adj2 := 8333, 50000
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1 = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2 = v
		}
	}
	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)
	x1 := math.Min(fw, fh) * float64(adj1) / 100000.0
	y1 := fh / 2 * float64(adj2) / 100000.0
	y2 := fh - y1
	if t == AutoShapeRightBrace {
		x2 := fw - x1
		r.drawCubicBezierAA(fx, fy, fx+x1, fy, fx+x2, fy+y1-x1, fx+x2, fy+y1, c, pw)
		r.drawLineThick(int(math.Round(fx+x2)), int(math.Round(fy+y1)), int(math.Round(fx+x2)), int(math.Round(fy+y2)), c, pw)
		r.drawCubicBezierAA(fx+x2, fy+y2, fx+x2, fy+y2+x1, fx+x1, fy+fh, fx, fy+fh, c, pw)
	} else {
		r.drawCubicBezierAA(fx+fw, fy, fx+fw-x1, fy, fx+x1, fy+y1-x1, fx+x1, fy+y1, c, pw)
		r.drawLineThick(int(math.Round(fx+x1)), int(math.Round(fy+y1)), int(math.Round(fx+x1)), int(math.Round(fy+y2)), c, pw)
		r.drawCubicBezierAA(fx+x1, fy+y2, fx+x1, fy+y2+x1, fx+fw-x1, fy+fh, fx+fw, fy+fh, c, pw)
	}
}

// renderArcBorder draws an arc shape's stroke and arrowheads.
// OOXML arc preset: adj1 = start angle, adj2 = end angle (in 60000ths of a degree).
// Default: adj1=16200000 (270°), adj2=0 (0°) — a quarter-circle arc from bottom to right.
func (r *renderer) renderArcBorder(s *AutoShape, x, y, w, h int, bc color.RGBA, pw int) {
	// Get adjustment values (angles in 60000ths of a degree)
	stAng := 16200000 // default start: 270°
	endAng := 0       // default end: 0°
	if s.adjustValues != nil {
		if v, ok := s.adjustValues["adj1"]; ok {
			stAng = v
		}
		if v, ok := s.adjustValues["adj2"]; ok {
			endAng = v
		}
	}

	stRad := float64(stAng) / 60000.0 * math.Pi / 180.0
	endRad := float64(endAng) / 60000.0 * math.Pi / 180.0

	// Ensure we sweep in the positive direction
	if endRad <= stRad {
		endRad += 2 * math.Pi
	}

	rx := float64(w) / 2.0
	ry := float64(h) / 2.0
	cx := float64(x) + rx
	cy := float64(y) + ry

	// Generate arc points
	sweep := endRad - stRad
	steps := maxInt(int(math.Abs(sweep)*(rx+ry)*0.5), 60)
	pts := make([]fpoint, steps+1)
	for i := 0; i <= steps; i++ {
		t := float64(i) / float64(steps)
		a := stRad + sweep*t
		pts[i] = fpoint{cx + rx*math.Cos(a), cy + ry*math.Sin(a)}
	}

	// Draw the arc stroke
	ls := BorderSolid
	if s.border != nil {
		ls = s.border.Style
	}
	if ls == BorderDash || ls == BorderDot {
		r.drawDashedPolylineAA(pts, bc, pw, ls)
	} else {
		for i := 1; i < len(pts); i++ {
			r.drawLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), bc, pw)
		}
	}

	// Draw arrowheads
	intPts := make([][2]int, len(pts))
	for i, p := range pts {
		intPts[i] = [2]int{int(p.x), int(p.y)}
	}
	if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
		r.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, bc, pw, s.headEnd)
	}
	if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
		last := intPts[len(intPts)-1]
		r.drawArrowOnPath(last[0], last[1], intPts, bc, pw, s.tailEnd)
	}
}

func (r *renderer) renderLine(s *LineShape) {
	rotation := s.GetRotation()
	if rotation != 0 {
		// For rotated connectors, compute the path in local coordinates,
		// apply flip and rotation transforms, then draw on the main canvas.
		r.renderLineRotated(s)
		return
	}
	ox := r.emuToPixelX(s.offsetX)
	oy := r.emuToPixelY(s.offsetY)
	r.renderLineAt(s, ox, oy)
}

// renderLineRotated handles connectors with rotation by transforming path points.
func (r *renderer) renderLineRotated(s *LineShape) {
	// Use float64 EMU coordinates throughout to avoid precision loss.
	// When the bounding box is very narrow (e.g. width=10390 EMU -> 1 pixel),
	// computing in pixel space destroys the adjustment value information.
	wEmu := float64(s.width)
	hEmu := float64(s.height)
	oxEmu := float64(s.offsetX)
	oyEmu := float64(s.offsetY)
	rotation := s.GetRotation()

	// Custom geometry path with rotation — convert path to pixel coords,
	// then rotate around the bounding box center.
	if s.customPath != nil && len(s.customPath.Commands) > 0 {
		ox := r.emuToPixelX(s.offsetX)
		oy := r.emuToPixelY(s.offsetY)
		w := r.emuToPixelX(s.width)
		h := r.emuToPixelY(s.height)
		subs := r.customPathToPixelSubpaths(s.customPath, ox, oy, w, h)
		if len(subs) > 0 {
			// Rotate around bounding box center
			cxPx := float64(ox) + float64(w)/2.0
			cyPx := float64(oy) + float64(h)/2.0
			rad := float64(rotation) * math.Pi / 180.0
			cosA := math.Cos(rad)
			sinA := math.Sin(rad)
			rotate := func(pts []fpoint) {
				for i := range pts {
					dx := pts[i].x - cxPx
					dy := pts[i].y - cyPx
					pts[i].x = dx*cosA - dy*sinA + cxPx
					pts[i].y = dx*sinA + dy*cosA + cyPx
				}
			}

			// Rotate every subpath once around the box center, then stroke.
			for _, sub := range subs {
				rotate(sub)
			}
			pw := maxInt(int(float64(s.GetLineWidthEMU())*r.scaleX), 1)
			c := argbToRGBA(s.lineColor)
			ls := s.lineStyle
			for _, sub := range subs {
				if len(sub) < 2 {
					continue
				}
				if ls == BorderDash || ls == BorderDot {
					r.drawDashedPolylineAA(sub, c, pw, ls)
				} else {
					for i := 1; i < len(sub); i++ {
						r.drawLineAA(int(sub[i-1].x), int(sub[i-1].y), int(sub[i].x), int(sub[i].y), c, pw)
					}
				}
			}
			if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
				head := subs[0]
				intPts := make([][2]int, len(head))
				for i, p := range head {
					intPts[i] = [2]int{int(p.x), int(p.y)}
				}
				r.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, c, pw, s.headEnd)
			}
			if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
				tail := subs[len(subs)-1]
				intPts := make([][2]int, len(tail))
				for i, p := range tail {
					intPts[i] = [2]int{int(p.x), int(p.y)}
				}
				last := intPts[len(intPts)-1]
				r.drawArrowOnPath(last[0], last[1], intPts, c, pw, s.tailEnd)
			}
		}
		return
	}

	// Build path in local EMU coordinates (0,0)-(wEmu,hEmu)
	type fpt [2]float64
	var pathPts []fpt

	switch {
	case s.connectorType == "bentConnector3":
		adjPct := 50000.0
		if v, ok := s.adjustValues["adj1"]; ok {
			adjPct = float64(v)
		}
		midX := wEmu * adjPct / 100000.0
		pathPts = []fpt{{0, 0}, {midX, 0}, {midX, hEmu}, {wEmu, hEmu}}

	case s.connectorType == "bentConnector2":
		pathPts = []fpt{{0, 0}, {wEmu, 0}, {wEmu, hEmu}}

	case s.connectorType == "bentConnector4":
		adjPct1 := 50000.0
		adjPct2 := 50000.0
		if v, ok := s.adjustValues["adj1"]; ok {
			adjPct1 = float64(v)
		}
		if v, ok := s.adjustValues["adj2"]; ok {
			adjPct2 = float64(v)
		}
		midX := wEmu * adjPct1 / 100000.0
		midY := hEmu * adjPct2 / 100000.0
		pathPts = []fpt{{0, 0}, {midX, 0}, {midX, midY}, {wEmu, midY}, {wEmu, hEmu}}

	case s.connectorType == "bentConnector5":
		adjPct1 := 50000.0
		adjPct2 := 50000.0
		adjPct3 := 50000.0
		if v, ok := s.adjustValues["adj1"]; ok {
			adjPct1 = float64(v)
		}
		if v, ok := s.adjustValues["adj2"]; ok {
			adjPct2 = float64(v)
		}
		if v, ok := s.adjustValues["adj3"]; ok {
			adjPct3 = float64(v)
		}
		midX1 := wEmu * adjPct1 / 100000.0
		midY := hEmu * adjPct2 / 100000.0
		midX2 := wEmu * adjPct3 / 100000.0
		pathPts = []fpt{{0, 0}, {midX1, 0}, {midX1, midY}, {midX2, midY}, {midX2, hEmu}, {wEmu, hEmu}}

	case strings.HasPrefix(s.connectorType, "curvedConnector"):
		// For curved connectors with rotation, compute endpoints in EMU,
		// rotate, convert to pixels, then delegate to renderCurvedConnector.
		cx := wEmu / 2.0
		cy := hEmu / 2.0
		rad := float64(rotation) * math.Pi / 180.0
		cosA := math.Cos(rad)
		sinA := math.Sin(rad)
		destCX := oxEmu + cx
		destCY := oyEmu + cy

		sx, sy := 0.0, 0.0
		ex, ey := wEmu, hEmu
		if s.flipHorizontal {
			sx, ex = wEmu-sx, wEmu-ex
		}
		if s.flipVertical {
			sy, ey = hEmu-sy, hEmu-ey
		}
		rsx := (sx-cx)*cosA - (sy-cy)*sinA + destCX
		rsy := (sx-cx)*sinA + (sy-cy)*cosA + destCY
		rex := (ex-cx)*cosA - (ey-cy)*sinA + destCX
		rey := (ex-cx)*sinA + (ey-cy)*cosA + destCY

		px1 := int(math.Round(rsx * r.scaleX))
		py1 := int(math.Round(rsy * r.scaleY))
		px2 := int(math.Round(rex * r.scaleX))
		py2 := int(math.Round(rey * r.scaleY))

		pw := maxInt(int(float64(s.GetLineWidthEMU())*r.scaleX), 1)
		c := argbToRGBA(s.lineColor)
		r.renderCurvedConnector(s.connectorType, px1, py1, px2, py2, s.adjustValues, c, pw, s.lineStyle, s.headEnd, s.tailEnd)
		return

	default:
		pathPts = []fpt{{0, 0}, {wEmu, hEmu}}
	}

	// Apply flips in EMU space
	if s.flipHorizontal {
		for i := range pathPts {
			pathPts[i][0] = wEmu - pathPts[i][0]
		}
	}
	if s.flipVertical {
		for i := range pathPts {
			pathPts[i][1] = hEmu - pathPts[i][1]
		}
	}

	// Rotate each point around the center of the bounding box in EMU space
	cx := wEmu / 2.0
	cy := hEmu / 2.0
	rad := float64(rotation) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	destCX := oxEmu + cx
	destCY := oyEmu + cy

	// Transform to slide EMU coordinates, then convert to pixels
	transformed := make([][2]int, len(pathPts))
	for i, pt := range pathPts {
		rx := pt[0] - cx
		ry := pt[1] - cy
		nx := rx*cosA - ry*sinA + destCX
		ny := rx*sinA + ry*cosA + destCY
		transformed[i] = [2]int{
			int(math.Round(nx * r.scaleX)),
			int(math.Round(ny * r.scaleY)),
		}
	}

	pw := maxInt(int(float64(s.GetLineWidthEMU())*r.scaleX), 1)
	c := argbToRGBA(s.lineColor)
	ls := s.lineStyle

	drawSeg := func(ax, ay, bx, by int) {
		if ls == BorderDash || ls == BorderDot {
			r.drawDashedLineAA(ax, ay, bx, by, c, pw, ls)
		} else {
			r.drawLineAA(ax, ay, bx, by, c, pw)
		}
	}

	for i := 0; i+1 < len(transformed); i++ {
		drawSeg(transformed[i][0], transformed[i][1],
			transformed[i+1][0], transformed[i+1][1])
	}

	if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
		r.drawArrowOnPath(transformed[0][0], transformed[0][1], transformed, c, pw, s.headEnd)
	}
	if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
		last := transformed[len(transformed)-1]
		r.drawArrowOnPath(last[0], last[1], transformed, c, pw, s.tailEnd)
	}
}

// renderLineAt draws a line/connector with the bounding box top-left at (ox, oy).
// Flip and adjust values are applied relative to this origin.
func (r *renderer) renderLineAt(s *LineShape, ox, oy int) {
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	// Visual start/end (after flip) — headEnd is at visual start (x1,y1),
	// tailEnd is at visual end (x2,y2). Flip attributes determine which
	// geometric corner maps to the visual start/end.
	gx1 := ox
	gy1 := oy
	gx2 := ox + w
	gy2 := oy + h

	x1, y1, x2, y2 := gx1, gy1, gx2, gy2
	if s.flipHorizontal {
		x1, x2 = x2, x1
	}
	if s.flipVertical {
		y1, y2 = y2, y1
	}
	// lineWidth in EMU, convert to pixels
	pw := maxInt(int(float64(s.GetLineWidthEMU())*r.scaleX), 1)
	c := argbToRGBA(s.lineColor)
	ls := s.lineStyle

	// Custom geometry path (freeform curved arrows, ink annotations, etc.)
	if s.customPath != nil && len(s.customPath.Commands) > 0 {
		subs := r.customPathToPixelSubpaths(s.customPath, ox, oy, w, h)
		if len(subs) > 0 {
			for _, sub := range subs {
				if len(sub) < 2 {
					continue
				}
				if ls == BorderDash || ls == BorderDot {
					r.drawDashedPolylineAA(sub, c, pw, ls)
				} else {
					for i := 1; i < len(sub); i++ {
						r.drawLineAA(int(sub[i-1].x), int(sub[i-1].y), int(sub[i].x), int(sub[i].y), c, pw)
					}
				}
			}
			if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
				head := subs[0]
				intPts := make([][2]int, len(head))
				for i, p := range head {
					intPts[i] = [2]int{int(p.x), int(p.y)}
				}
				r.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, c, pw, s.headEnd)
			}
			if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
				tail := subs[len(subs)-1]
				intPts := make([][2]int, len(tail))
				for i, p := range tail {
					intPts[i] = [2]int{int(p.x), int(p.y)}
				}
				last := intPts[len(intPts)-1]
				r.drawArrowOnPath(last[0], last[1], intPts, c, pw, s.tailEnd)
			}
		}
		return
	}

	// drawSeg draws a line segment respecting the connector's dash style.
	drawSeg := func(ax, ay, bx, by int) {
		if ls == BorderDash || ls == BorderDot {
			r.drawDashedLineAA(ax, ay, bx, by, c, pw, ls)
		} else {
			r.drawLineAA(ax, ay, bx, by, c, pw)
		}
	}

	switch {
	case s.connectorType == "bentConnector3":
		// Elbow connector with 3 segments: horizontal, vertical, horizontal
		adjPct := 50000
		if v, ok := s.adjustValues["adj1"]; ok {
			adjPct = v
		}
		midX := x1 + int(float64(x2-x1)*float64(adjPct)/100000.0)
		drawSeg(x1, y1, midX, y1)
		drawSeg(midX, y1, midX, y2)
		drawSeg(midX, y2, x2, y2)
		pathPts := [][2]int{{x1, y1}, {midX, y1}, {midX, y2}, {x2, y2}}
		if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
			r.drawArrowOnPath(x1, y1, pathPts, c, pw, s.headEnd)
		}
		if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
			r.drawArrowOnPath(x2, y2, pathPts, c, pw, s.tailEnd)
		}

	case s.connectorType == "bentConnector2":
		drawSeg(x1, y1, x2, y1)
		drawSeg(x2, y1, x2, y2)
		pathPts := [][2]int{{x1, y1}, {x2, y1}, {x2, y2}}
		if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
			r.drawArrowOnPath(x1, y1, pathPts, c, pw, s.headEnd)
		}
		if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
			r.drawArrowOnPath(x2, y2, pathPts, c, pw, s.tailEnd)
		}

	case s.connectorType == "bentConnector4":
		adjPct1 := 50000
		adjPct2 := 50000
		if v, ok := s.adjustValues["adj1"]; ok {
			adjPct1 = v
		}
		if v, ok := s.adjustValues["adj2"]; ok {
			adjPct2 = v
		}
		midX := x1 + int(float64(x2-x1)*float64(adjPct1)/100000.0)
		midY := y1 + int(float64(y2-y1)*float64(adjPct2)/100000.0)
		drawSeg(x1, y1, midX, y1)
		drawSeg(midX, y1, midX, midY)
		drawSeg(midX, midY, x2, midY)
		drawSeg(x2, midY, x2, y2)
		pathPts := [][2]int{{x1, y1}, {midX, y1}, {midX, midY}, {x2, midY}, {x2, y2}}
		if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
			r.drawArrowOnPath(x1, y1, pathPts, c, pw, s.headEnd)
		}
		if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
			r.drawArrowOnPath(x2, y2, pathPts, c, pw, s.tailEnd)
		}

	case s.connectorType == "bentConnector5":
		adjPct1 := 50000
		adjPct2 := 50000
		adjPct3 := 50000
		if v, ok := s.adjustValues["adj1"]; ok {
			adjPct1 = v
		}
		if v, ok := s.adjustValues["adj2"]; ok {
			adjPct2 = v
		}
		if v, ok := s.adjustValues["adj3"]; ok {
			adjPct3 = v
		}
		midX1 := x1 + int(float64(x2-x1)*float64(adjPct1)/100000.0)
		midY := y1 + int(float64(y2-y1)*float64(adjPct2)/100000.0)
		midX2 := x1 + int(float64(x2-x1)*float64(adjPct3)/100000.0)
		drawSeg(x1, y1, midX1, y1)
		drawSeg(midX1, y1, midX1, midY)
		drawSeg(midX1, midY, midX2, midY)
		drawSeg(midX2, midY, midX2, y2)
		drawSeg(midX2, y2, x2, y2)
		pathPts := [][2]int{{x1, y1}, {midX1, y1}, {midX1, midY}, {midX2, midY}, {midX2, y2}, {x2, y2}}
		if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
			r.drawArrowOnPath(x1, y1, pathPts, c, pw, s.headEnd)
		}
		if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
			r.drawArrowOnPath(x2, y2, pathPts, c, pw, s.tailEnd)
		}

	case strings.HasPrefix(s.connectorType, "curvedConnector"):
		r.renderCurvedConnector(s.connectorType, x1, y1, x2, y2, s.adjustValues, c, pw, ls, s.headEnd, s.tailEnd)

	default:
		// Straight line connector (line, straightConnector1, etc.)
		drawSeg(x1, y1, x2, y2)
		// headEnd at visual start (x1,y1), tailEnd at visual end (x2,y2).
		// Arrow tip placed at the endpoint, direction from the other end.
		if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
			r.drawArrowHead(x2, y2, x1, y1, c, pw, s.headEnd, false)
		}
		if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
			r.drawArrowHead(x1, y1, x2, y2, c, pw, s.tailEnd, false)
		}
	}
}

// renderCurvedConnector draws a curved connector using cubic Bezier curves.
// OOXML curved connectors (curvedConnector2..5) follow the same waypoint
// logic as bent connectors but replace the right-angle segments with smooth
// S-curves through the waypoints.
func (r *renderer) renderCurvedConnector(connType string, x1, y1, x2, y2 int, adj map[string]int, c color.RGBA, pw int, ls BorderStyle, headEnd, tailEnd *LineEnd) {
	drawBezier := func(bx0, by0, bx1, by1, bx2, by2, bx3, by3 float64) {
		if ls == BorderDash || ls == BorderDot {
			r.drawDashedCubicBezierAA(bx0, by0, bx1, by1, bx2, by2, bx3, by3, c, pw, ls)
		} else {
			r.drawCubicBezierAA(bx0, by0, bx1, by1, bx2, by2, bx3, by3, c, pw)
		}
	}

	// Build waypoints based on connector type (same as bent connectors)
	var waypoints []fpoint
	switch connType {
	case "curvedConnector2":
		waypoints = []fpoint{{float64(x1), float64(y1)}, {float64(x2), float64(y1)}, {float64(x2), float64(y2)}}
	case "curvedConnector3":
		adjPct := 50000
		if v, ok := adj["adj1"]; ok {
			adjPct = v
		}
		midX := float64(x1) + float64(x2-x1)*float64(adjPct)/100000.0
		waypoints = []fpoint{{float64(x1), float64(y1)}, {midX, float64(y1)}, {midX, float64(y2)}, {float64(x2), float64(y2)}}
	case "curvedConnector4":
		adjPct1 := 50000
		adjPct2 := 50000
		if v, ok := adj["adj1"]; ok {
			adjPct1 = v
		}
		if v, ok := adj["adj2"]; ok {
			adjPct2 = v
		}
		midX := float64(x1) + float64(x2-x1)*float64(adjPct1)/100000.0
		midY := float64(y1) + float64(y2-y1)*float64(adjPct2)/100000.0
		waypoints = []fpoint{{float64(x1), float64(y1)}, {midX, float64(y1)}, {midX, midY}, {float64(x2), midY}, {float64(x2), float64(y2)}}
	case "curvedConnector5":
		adjPct1 := 50000
		adjPct2 := 50000
		adjPct3 := 50000
		if v, ok := adj["adj1"]; ok {
			adjPct1 = v
		}
		if v, ok := adj["adj2"]; ok {
			adjPct2 = v
		}
		if v, ok := adj["adj3"]; ok {
			adjPct3 = v
		}
		midX1 := float64(x1) + float64(x2-x1)*float64(adjPct1)/100000.0
		midY := float64(y1) + float64(y2-y1)*float64(adjPct2)/100000.0
		midX2 := float64(x1) + float64(x2-x1)*float64(adjPct3)/100000.0
		waypoints = []fpoint{{float64(x1), float64(y1)}, {midX1, float64(y1)}, {midX1, midY}, {midX2, midY}, {midX2, float64(y2)}, {float64(x2), float64(y2)}}
	default:
		// Unknown curved connector variant, draw as straight
		waypoints = []fpoint{{float64(x1), float64(y1)}, {float64(x2), float64(y2)}}
	}

	if len(waypoints) < 2 {
		return
	}

	// Draw smooth curves through waypoints using cubic Bezier segments.
	// Each pair of consecutive waypoints becomes a Bezier segment where
	// the control points create a smooth S-curve between the two points.
	for i := 0; i < len(waypoints)-1; i++ {
		p0 := waypoints[i]
		p1 := waypoints[i+1]
		// Control points at 1/3 and 2/3 along the segment, but shifted
		// to create the S-curve effect (horizontal→vertical or vertical→horizontal)
		dx := p1.x - p0.x
		dy := p1.y - p0.y
		if math.Abs(dx) > math.Abs(dy) {
			// Primarily horizontal segment: curve vertically at midpoint
			drawBezier(p0.x, p0.y, p0.x+dx/2, p0.y, p0.x+dx/2, p1.y, p1.x, p1.y)
		} else {
			// Primarily vertical segment: curve horizontally at midpoint
			drawBezier(p0.x, p0.y, p0.x, p0.y+dy/2, p1.x, p0.y+dy/2, p1.x, p1.y)
		}
	}

	// Draw arrow heads using the tangent direction at the endpoints
	if headEnd != nil && headEnd.Type != ArrowNone && headEnd.Type != "" {
		// Direction from second waypoint toward first
		p0 := waypoints[0]
		p1 := waypoints[1]
		dx := p1.x - p0.x
		dy := p1.y - p0.y
		// Tangent at start: for our Bezier, the initial tangent points toward the first control point
		var fromX, fromY int
		if math.Abs(dx) > math.Abs(dy) {
			fromX = int(p0.x + dx/2)
			fromY = int(p0.y)
		} else {
			fromX = int(p0.x)
			fromY = int(p0.y + dy/2)
		}
		r.drawArrowHead(fromX, fromY, int(p0.x), int(p0.y), c, pw, headEnd, false)
	}
	if tailEnd != nil && tailEnd.Type != ArrowNone && tailEnd.Type != "" {
		n := len(waypoints)
		pLast := waypoints[n-1]
		pPrev := waypoints[n-2]
		dx := pLast.x - pPrev.x
		dy := pLast.y - pPrev.y
		var fromX, fromY int
		if math.Abs(dx) > math.Abs(dy) {
			fromX = int(pPrev.x + dx/2)
			fromY = int(pLast.y)
		} else {
			fromX = int(pLast.x)
			fromY = int(pPrev.y + dy/2)
		}
		r.drawArrowHead(fromX, fromY, int(pLast.x), int(pLast.y), c, pw, tailEnd, false)
	}
}

// drawArrowOnPath draws an arrow at the visual endpoint (vx,vy) using the
// direction from the visual path. It finds which end of the path is closest to
// the visual point and uses the appropriate segment for direction.
func (r *renderer) drawArrowOnPath(vx, vy int, pathPts [][2]int, c color.RGBA, lineWidth int, le *LineEnd) {
	if len(pathPts) < 2 {
		return
	}
	first := pathPts[0]
	last := pathPts[len(pathPts)-1]
	distFirst := abs(vx-first[0]) + abs(vy-first[1])
	distLast := abs(vx-last[0]) + abs(vy-last[1])

	if distFirst <= distLast {
		// Visual point is at the start of the path.
		// Find first non-zero-length segment for direction.
		for i := 0; i+1 < len(pathPts); i++ {
			dx := pathPts[i+1][0] - pathPts[i][0]
			dy := pathPts[i+1][1] - pathPts[i][1]
			if abs(dx) > 1 || abs(dy) > 1 {
				r.drawArrowHead(pathPts[i+1][0], pathPts[i+1][1], vx, vy, c, lineWidth, le, false)
				return
			}
		}
		r.drawArrowHead(last[0], last[1], vx, vy, c, lineWidth, le, false)
	} else {
		// Visual point is at the end of the path.
		// Find last non-zero-length segment for direction.
		for i := len(pathPts) - 1; i > 0; i-- {
			dx := pathPts[i][0] - pathPts[i-1][0]
			dy := pathPts[i][1] - pathPts[i-1][1]
			if abs(dx) > 1 || abs(dy) > 1 {
				r.drawArrowHead(pathPts[i-1][0], pathPts[i-1][1], vx, vy, c, lineWidth, le, false)
				return
			}
		}
		r.drawArrowHead(first[0], first[1], vx, vy, c, lineWidth, le, false)
	}
}

// drawArrowHead draws an arrow head at one end of a line.
// If atStart is true, the arrow is drawn at (x1,y1) pointing away from (x2,y2).
// If atStart is false, the arrow is drawn at (x2,y2) pointing away from (x1,y1).
func (r *renderer) drawArrowHead(x1, y1, x2, y2 int, c color.RGBA, lineWidth int, le *LineEnd, atStart bool) {
	// Compute arrow size based on line width and arrow size attributes.
	// PowerPoint arrow sizing: the OOXML spec defines arrow length/width in
	// terms of line width multiples. For "med" size on a 2pt line at 96 DPI:
	//   length ≈ 9px, width ≈ 7px
	// We use a formula that matches PowerPoint's rendering closely.
	lw := float64(lineWidth)
	baseLen := lw*3.0 + 4.0
	baseWidth := lw*2.5 + 3.0

	switch le.Length {
	case ArrowSizeSm:
		baseLen *= 0.6
	case ArrowSizeLg:
		baseLen *= 1.6
	}
	switch le.Width {
	case ArrowSizeSm:
		baseWidth *= 0.6
	case ArrowSizeLg:
		baseWidth *= 1.6
	}

	// Minimum arrow size for visibility
	if baseLen < 7 {
		baseLen = 7
	}
	if baseWidth < 5 {
		baseWidth = 5
	}

	// Direction vector
	var dx, dy float64
	if atStart {
		dx = float64(x1 - x2)
		dy = float64(y1 - y2)
	} else {
		dx = float64(x2 - x1)
		dy = float64(y2 - y1)
	}
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 1 {
		return
	}
	dx /= length
	dy /= length

	// Tip point — extend 0.5px past the endpoint so the scanline at the
	// endpoint row hits the very tip of the triangle (the scanline samples
	// at pixel-center y+0.5, so without this offset the tip row is already
	// past the vertex and produces a flat bottom instead of a sharp point).
	var tipX, tipY float64
	if atStart {
		tipX = float64(x1) + dx*0.5
		tipY = float64(y1) + dy*0.5
	} else {
		tipX = float64(x2) + dx*0.5
		tipY = float64(y2) + dy*0.5
	}

	// Base center (behind the tip)
	baseX := tipX - dx*baseLen
	baseY := tipY - dy*baseLen

	// Perpendicular
	perpX := -dy
	perpY := dx

	halfW := baseWidth / 2.0

	switch le.Type {
	case ArrowTriangle:
		// Filled triangle arrow head
		p1 := fpoint{tipX, tipY}
		p2 := fpoint{baseX + perpX*halfW, baseY + perpY*halfW}
		p3 := fpoint{baseX - perpX*halfW, baseY - perpY*halfW}
		pts := []fpoint{p1, p2, p3}
		r.fillPolygon(pts, c)
	case ArrowStealth:
		// Stealth has a notch at the base
		p1 := fpoint{tipX, tipY}
		p2 := fpoint{baseX + perpX*halfW, baseY + perpY*halfW}
		p3 := fpoint{baseX - perpX*halfW, baseY - perpY*halfW}
		notchDepth := baseLen * 0.3
		notchX := baseX + dx*notchDepth
		notchY := baseY + dy*notchDepth
		pts := []fpoint{p1, p2, {notchX, notchY}, p3}
		r.fillPolygon(pts, c)
	case ArrowArrow:
		// Open arrow head — two lines forming a V (not filled)
		p2 := fpoint{baseX + perpX*halfW, baseY + perpY*halfW}
		p3 := fpoint{baseX - perpX*halfW, baseY - perpY*halfW}
		lw := maxInt(lineWidth, 1)
		r.drawLineAA(int(p2.x), int(p2.y), int(tipX), int(tipY), c, lw)
		r.drawLineAA(int(tipX), int(tipY), int(p3.x), int(p3.y), c, lw)
	case ArrowDiamond:
		// Diamond shape
		midX := tipX - dx*baseLen/2
		midY := tipY - dy*baseLen/2
		p1 := fpoint{tipX, tipY}
		p2 := fpoint{midX + perpX*halfW, midY + perpY*halfW}
		p3 := fpoint{baseX, baseY}
		p4 := fpoint{midX - perpX*halfW, midY - perpY*halfW}
		pts := []fpoint{p1, p2, p3, p4}
		r.fillPolygon(pts, c)
	case ArrowOval:
		// Oval/circle at the end
		cx := int(tipX - dx*baseLen/2)
		cy := int(tipY - dy*baseLen/2)
		rad := int(baseLen / 2)
		r.fillEllipseAA(cx-rad, cy-rad, rad*2, rad*2, c)
	}
}

func (r *renderer) renderTable(s *TableShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	if s.numRows == 0 || s.numCols == 0 {
		return
	}

	// Compute column positions using individual widths if available
	colX := make([]int, s.numCols+1)
	colX[0] = x
	if len(s.colWidths) == s.numCols {
		for i, cw := range s.colWidths {
			colX[i+1] = colX[i] + r.emuToPixelX(cw)
		}
	} else {
		cellW := w / s.numCols
		for i := 0; i <= s.numCols; i++ {
			colX[i] = x + i*cellW
		}
	}

	// Compute row positions using individual heights if available
	rowY := make([]int, s.numRows+1)
	rowY[0] = y
	if len(s.rowHeights) == s.numRows {
		for i, rh := range s.rowHeights {
			rowY[i+1] = rowY[i] + r.emuToPixelY(rh)
		}
	} else {
		cellH := h / s.numRows
		for i := 0; i <= s.numRows; i++ {
			rowY[i] = y + i*cellH
		}
	}

	// A <a:tr> height is a *minimum*: PowerPoint grows a row to fit its
	// tallest cell's text plus the cell insets. Honouring the declared height
	// alone makes every later row sit higher than PowerPoint puts it, and the
	// offset compounds down the table.
	for row := 0; row < s.numRows && row < len(s.rows); row++ {
		need := 0
		for ci, cell := range s.rows[row] {
			if cell == nil || cell.hMerge || cell.vMerge || len(cell.paragraphs) == 0 {
				continue
			}
			c0 := ci
			c1 := ci + cell.colSpan
			if c1 > s.numCols {
				c1 = s.numCols
			}
			if c1 <= c0 || c1 >= len(colX) {
				continue
			}
			cw := colX[c1] - colX[c0]
			if cw < 10 {
				cw = 10
			}
			cl, ct, cr, cb := r.cellInsets(cell)
			avail := cw - cl - cr
			if avail < 10 {
				avail = 10
			}
			th := r.measureParagraphsHeight(cell.paragraphs, avail, 1<<20, TextAnchorNone, true)
			if th+ct+cb > need {
				need = th + ct + cb
			}
		}
		if need > 0 && rowY[row+1]-rowY[row] < need {
			grow := need - (rowY[row+1] - rowY[row])
			for k := row + 1; k <= s.numRows; k++ {
				rowY[k] += grow
			}
		}
	}

	for row := 0; row < s.numRows; row++ {
		if row >= len(s.rows) {
			break
		}
		for col := 0; col < len(s.rows[row]); col++ {
			if col >= s.numCols {
				break
			}
			cell := s.rows[row][col]
			// Skip merged continuation cells
			if cell.hMerge || cell.vMerge {
				continue
			}
			cx := colX[col]
			cy := rowY[row]
			// Handle column span
			endCol := col + cell.colSpan
			if endCol > s.numCols {
				endCol = s.numCols
			}
			// Handle row span
			endRow := row + cell.rowSpan
			if endRow > s.numRows {
				endRow = s.numRows
			}
			cellW := colX[endCol] - cx
			cellH := rowY[endRow] - cy
			cellRect := image.Rect(cx, cy, cx+cellW, cy+cellH)
			r.renderFill(cell.fill, cellRect)
			if cell.border != nil {
				r.renderCellBorders(cell.border, cellRect)
			} else {
				r.drawRect(cellRect, color.RGBA{A: 255}, 1)
			}
			cl, ct, cr, cb := r.cellInsets(cell)
			tw := cellW - cl - cr
			th := cellH - ct - cb
			if tw < 1 {
				tw = 1
			}
			if th < 1 {
				th = 1
			}
			// <a:tcPr anchor>: a cell's text sits where the cell says, and
			// an untouched cell defaults to the top. The bottom-anchored
			// header row of the r29 deck makes the difference visible.
			anchor := TextAnchorNone
			switch cell.anchor {
			case "t":
				anchor = TextAnchorTop
			case "ctr":
				anchor = TextAnchorMiddle
			case "b":
				anchor = TextAnchorBottom
			}
			r.drawParagraphs(cell.paragraphs, cx+cl, cy+ct, tw, th, anchor, true)
		}
	}
}

// cellInsets returns a table cell's insets in pixels: left, top, right,
// bottom. A cell that declares none gets the table default of 0.1" left and
// right and 0.05" top and bottom, which at 1600px across a 10" slide is 16px
// and 8px — not the few pixels of slack a table would tolerate, since the
// insets also narrow the width a cell's text may wrap into.
func (r *renderer) cellInsets(cell *TableCell) (l, t, right, b int) {
	ml, mr, mt, mb, ok := cell.GetMargins()
	if !ok {
		ml, mr = DefaultCellMarginLR, DefaultCellMarginLR
		mt, mb = DefaultCellMarginTB, DefaultCellMarginTB
	}
	return r.emuToPixelX(int64(ml)), r.emuToPixelY(int64(mt)),
		r.emuToPixelX(int64(mr)), r.emuToPixelY(int64(mb))
}

func (r *renderer) renderCellBorders(cb *CellBorders, rect image.Rectangle) {
	drawBorder := func(b *Border, x1, y1, x2, y2 int) {
		if b == nil || b.Style == BorderNone {
			return
		}
		pw := maxInt(int(float64(b.Width)*12700.0*r.scaleX), 1)
		r.drawLineThick(x1, y1, x2, y2, argbToRGBA(b.Color), pw)
	}
	drawBorder(cb.Top, rect.Min.X, rect.Min.Y, rect.Max.X, rect.Min.Y)
	drawBorder(cb.Bottom, rect.Min.X, rect.Max.Y-1, rect.Max.X, rect.Max.Y-1)
	drawBorder(cb.Left, rect.Min.X, rect.Min.Y, rect.Min.X, rect.Max.Y)
	drawBorder(cb.Right, rect.Max.X-1, rect.Min.Y, rect.Max.X-1, rect.Max.Y)
}

// --- Fill rendering ---

func (r *renderer) renderFill(fill *Fill, rect image.Rectangle) {
	if fill == nil || fill.Type == FillNone {
		return
	}
	switch fill.Type {
	case FillSolid:
		fc := argbToRGBA(fill.Color)
		fc = r.scaleAlpha(fc)
		r.fillRectBlend(rect, fc)
	case FillGradientLinear:
		r.fillGradientLinear(rect, fill)
	case FillGradientPath:
		r.fillGradientPath(rect, fill)
	}
}

// renderCustomPathFill fills a custom geometry path within the given shape bounds.
func (r *renderer) renderCustomPathFill(cp *CustomGeomPath, fill *Fill, ox, oy, w, h int) {
	if fill == nil || fill.Type == FillNone || cp == nil || len(cp.Commands) == 0 {
		return
	}
	// Convert path coordinates to pixel coordinates
	pts := r.customPathToPixelPoints(cp, ox, oy, w, h)
	if len(pts) < 3 {
		return
	}
	fc := argbToRGBA(fill.Color)
	fc = r.scaleAlpha(fc)
	r.fillPolygon(pts, fc)
}

// customPathToPixelPoints converts a custom geometry path to pixel-space fpoints.
func (r *renderer) customPathToPixelPoints(cp *CustomGeomPath, ox, oy, w, h int) []fpoint {
	var pts []fpoint
	for _, sub := range r.customPathToPixelSubpaths(cp, ox, oy, w, h) {
		pts = append(pts, sub...)
	}
	return pts
}

// customPathToPixelSubpaths converts a custom geometry path to pixel-space
// subpaths, splitting at every moveTo. Filling ignores the split (a moveTo
// mid-path was always a fill hole, not a seam), but stroking must respect it:
// VML ink annotations carry several strokes per shape and a bridging segment
// between them draws a line PowerPoint never does.
func (r *renderer) customPathToPixelSubpaths(cp *CustomGeomPath, ox, oy, w, h int) [][]fpoint {
	if cp == nil || cp.Width <= 0 || cp.Height <= 0 {
		return nil
	}
	scX := float64(w) / float64(cp.Width)
	scY := float64(h) / float64(cp.Height)

	toPixel := func(p PathPoint) fpoint {
		return fpoint{float64(ox) + float64(p.X)*scX, float64(oy) + float64(p.Y)*scY}
	}

	var subs [][]fpoint
	var cur []fpoint
	var lastPt fpoint
	endSub := func() {
		if len(cur) > 0 {
			subs = append(subs, cur)
			cur = nil
		}
	}

	for _, cmd := range cp.Commands {
		switch cmd.Type {
		case "moveTo":
			if len(cmd.Pts) > 0 {
				endSub()
				p := toPixel(cmd.Pts[0])
				cur = []fpoint{p}
				lastPt = p
			}
		case "lnTo":
			if len(cmd.Pts) > 0 {
				p := toPixel(cmd.Pts[0])
				if cur == nil {
					cur = []fpoint{p}
				} else {
					cur = append(cur, p)
				}
				lastPt = p
			}
		case "cubicBezTo":
			// Flatten cubic bezier into line segments for accurate curves
			if len(cmd.Pts) >= 3 {
				cp1 := toPixel(cmd.Pts[0])
				cp2 := toPixel(cmd.Pts[1])
				ep := toPixel(cmd.Pts[2])
				bezPts := r.flattenCubicBezier(lastPt.x, lastPt.y, cp1.x, cp1.y, cp2.x, cp2.y, ep.x, ep.y, 0)
				cur = append(cur, bezPts...)
				cur = append(cur, ep)
				lastPt = ep
			}
		case "quadBezTo":
			// Flatten quadratic bezier by converting to cubic
			if len(cmd.Pts) >= 2 {
				cp1 := toPixel(cmd.Pts[0])
				ep := toPixel(cmd.Pts[1])
				// Convert quadratic to cubic: CP1' = P0 + 2/3*(CP-P0), CP2' = EP + 2/3*(CP-EP)
				c1x := lastPt.x + 2.0/3.0*(cp1.x-lastPt.x)
				c1y := lastPt.y + 2.0/3.0*(cp1.y-lastPt.y)
				c2x := ep.x + 2.0/3.0*(cp1.x-ep.x)
				c2y := ep.y + 2.0/3.0*(cp1.y-ep.y)
				bezPts := r.flattenCubicBezier(lastPt.x, lastPt.y, c1x, c1y, c2x, c2y, ep.x, ep.y, 0)
				cur = append(cur, bezPts...)
				cur = append(cur, ep)
				lastPt = ep
			}
		case "close":
			// close is implicit in fillPolygon
		case "arcTo":
			// OOXML arcTo: wR/hR are ellipse radii in path coords,
			// stAng/swAng are in 60000ths of a degree.
			// The arc is drawn on an ellipse whose center is computed so
			// that the arc starts at lastPt.
			wR := float64(cmd.WR) * scX
			hR := float64(cmd.HR) * scY
			stAngDeg := float64(cmd.StAng) / 60000.0
			swAngDeg := float64(cmd.SwAng) / 60000.0
			stRad := stAngDeg * math.Pi / 180.0
			swRad := swAngDeg * math.Pi / 180.0

			if wR < 0.5 || hR < 0.5 {
				// Degenerate arc — skip
				break
			}

			// Center of the ellipse: lastPt is on the ellipse at stAng
			cx := lastPt.x - wR*math.Cos(stRad)
			cy := lastPt.y - hR*math.Sin(stRad)

			// Number of steps proportional to arc length
			steps := maxInt(int(math.Abs(swRad)*(wR+hR)*0.5), 8)
			angleStep := swRad / float64(steps)
			for i := 1; i <= steps; i++ {
				a := stRad + angleStep*float64(i)
				p := fpoint{cx + wR*math.Cos(a), cy + hR*math.Sin(a)}
				cur = append(cur, p)
				lastPt = p
			}
		}
	}
	endSub()
	return subs
}

// scaleAlpha applies the overlayOpacityScale to semi-transparent colors.
func (r *renderer) scaleAlpha(c color.RGBA) color.RGBA {
	scale := r.overlayOpacityScale
	if scale <= 0 || scale >= 1.0 {
		return c
	}
	if c.A < 255 && c.A > 0 {
		c.A = uint8(float64(c.A) * scale)
	}
	return c
}

// gradientStopsRGBA normalises a Fill's gradient into sorted stop positions
// (0..1) with colours. The N-stop list wins when present; otherwise the
// legacy start/middle/end fields are used.
func gradientStopsRGBA(fill *Fill) ([]float64, [][4]uint8) {
	if fill == nil {
		return nil, nil
	}
	if len(fill.Stops) >= 2 {
		pos := make([]float64, len(fill.Stops))
		cols := make([][4]uint8, len(fill.Stops))
		for i, s := range fill.Stops {
			t := float64(s.Pos) / 100000.0
			if t < 0 {
				t = 0
			} else if t > 1 {
				t = 1
			}
			pos[i] = t
			c := argbToRGBA(s.Color)
			cols[i] = [4]uint8{c.R, c.G, c.B, c.A}
		}
		return pos, cols
	}
	start := argbToRGBA(fill.Color)
	end := argbToRGBA(fill.EndColor)
	if fill.MidPos > 0 && fill.MidPos < 100000 {
		mid := argbToRGBA(fill.MidColor)
		return []float64{0, float64(fill.MidPos) / 100000.0, 1},
			[][4]uint8{{start.R, start.G, start.B, start.A}, {mid.R, mid.G, mid.B, mid.A}, {end.R, end.G, end.B, end.A}}
	}
	return []float64{0, 1},
		[][4]uint8{{start.R, start.G, start.B, start.A}, {end.R, end.G, end.B, end.A}}
}

// gradStopColor evaluates an N-stop piecewise-linear ramp at t (sRGB space,
// the same space the two-stop paths have always interpolated in).
func gradStopColor(pos []float64, cols [][4]uint8, t float64) [4]uint8 {
	if t <= pos[0] {
		return cols[0]
	}
	n := len(pos)
	if t >= pos[n-1] {
		return cols[n-1]
	}
	for i := 1; i < n; i++ {
		if t <= pos[i] {
			span := pos[i] - pos[i-1]
			seg := (t - pos[i-1]) / span
			it := 1 - seg
			a, b := cols[i-1], cols[i]
			return [4]uint8{
				uint8(float64(a[0])*it + float64(b[0])*seg),
				uint8(float64(a[1])*it + float64(b[1])*seg),
				uint8(float64(a[2])*it + float64(b[2])*seg),
				uint8(float64(a[3])*it + float64(b[3])*seg),
			}
		}
	}
	return cols[n-1]
}

func (r *renderer) fillGradientLinear(rect image.Rectangle, fill *Fill) {
	pos, cols := gradientStopsRGBA(fill)
	if pos == nil {
		return
	}
	w := rect.Dx()
	h := rect.Dy()
	if w <= 0 || h <= 0 {
		return
	}
	rad := float64(fill.Rotation) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	cx := float64(w) / 2
	cy := float64(h) / 2
	maxProj := math.Abs(cx*cosA) + math.Abs(cy*sinA)
	if maxProj < 1 {
		maxProj = 1
	}
	invMaxProj := 1.0 / (2 * maxProj)

	// Pre-compute row-independent part
	pix := r.img.Pix
	bounds := r.img.Bounds()
	stride := r.img.Stride

	for py := rect.Min.Y; py < rect.Max.Y; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		dyf := float64(py-rect.Min.Y) - cy
		rowBase := dyf*sinA + maxProj
		off := (py-bounds.Min.Y)*stride + (maxInt(rect.Min.X, bounds.Min.X)-bounds.Min.X)*4
		for px := maxInt(rect.Min.X, bounds.Min.X); px < minInt(rect.Max.X, bounds.Max.X); px++ {
			dxf := float64(px-rect.Min.X) - cx
			t := (dxf*cosA + rowBase) * invMaxProj
			if t < 0 {
				t = 0
			} else if t > 1 {
				t = 1
			}
			outC := gradStopColor(pos, cols, t)
			pix[off] = outC[0]
			pix[off+1] = outC[1]
			pix[off+2] = outC[2]
			pix[off+3] = outC[3]
			off += 4
		}
	}
}

// pathGradGeom carries the normalisation of a path gradient over a box, per
// the r39 pinned semantics (measured against PowerPoint COM exports with
// chromatic variant decks):
//
//   - The focus is the centre of the rectangle <a:fillToRect> carves out of
//     the box.
//   - The tile is the box grown/shrunk by the <a:tileRect> insets (negative
//     insets grow it).
//   - Focus strictly inside the tile: t = |P−F| / D, D the distance from the
//     focus to the FARTHEST TILE CORNER (not the box corner — a tile grown
//     by tileRect moves the t=1 ring outward with it).
//   - Focus on the tile boundary: Euclidean collapses (the whole tile is on
//     one side of the focus) and PowerPoint switches to a max-metric: t is
//     the largest of the per-axis distances from the focus, each normalised
//     by the focus-to-farthest-tile-edge distance on that axis.
//
// t is clamped to [0,1] — path gradients do not tile.
type pathGradGeom struct {
	fx, fy     float64
	dInv       float64 // euclidean regime: 1/D
	farX, farY float64 // max-metric regime: focus-to-far-edge extents
	euclidean  bool
}

func pathGradientGeometry(w, h float64, fill *Fill) pathGradGeom {
	// Focus: centre of the fillToRect rectangle. l/t/r/b insets in
	// 0..100000 box units.
	fl := float64(fill.FillTo[0]) / 100000.0
	ft := float64(fill.FillTo[1]) / 100000.0
	fr := float64(fill.FillTo[2]) / 100000.0
	fb := float64(fill.FillTo[3]) / 100000.0
	fx := w * (1 + fl - fr) / 2
	fy := h * (1 + ft - fb) / 2

	// Tile rectangle in box-local coordinates.
	tl := float64(fill.TileTo[0]) / 100000.0
	tt := float64(fill.TileTo[1]) / 100000.0
	tr := float64(fill.TileTo[2]) / 100000.0
	tb := float64(fill.TileTo[3]) / 100000.0
	tminX := w * tl
	tminY := h * tt
	tmaxX := w * (1 - tr)
	tmaxY := h * (1 - tb)

	// Regime: interior focus (any real margin) is Euclidean; a focus sitting
	// on an edge (margin ~0) is max-metric. Real decks use either a tile
	// centred on the focus or the bare box with a corner focus.
	const eps = 0.5
	margin := fx - tminX
	if tmaxX-fx < margin {
		margin = tmaxX - fx
	}
	if fy-tminY < margin {
		margin = fy - tminY
	}
	if tmaxY-fy < margin {
		margin = tmaxY - fy
	}
	g := pathGradGeom{fx: fx, fy: fy}
	if margin > eps {
		g.euclidean = true
		// Farthest tile corner.
		d := math.Hypot(fx-tminX, fy-tminY)
		if v := math.Hypot(tmaxX-fx, fy-tminY); v > d {
			d = v
		}
		if v := math.Hypot(fx-tminX, tmaxY-fy); v > d {
			d = v
		}
		if v := math.Hypot(tmaxX-fx, tmaxY-fy); v > d {
			d = v
		}
		if d < 1 {
			d = 1
		}
		g.dInv = 1 / d
	} else {
		g.farX = math.Max(fx-tminX, tmaxX-fx)
		g.farY = math.Max(fy-tminY, tmaxY-fy)
		if g.farX < 1 {
			g.farX = 1
		}
		if g.farY < 1 {
			g.farY = 1
		}
	}
	return g
}

// pathGradientT evaluates the gradient parameter at a box-local point.
func pathGradientT(g pathGradGeom, px, py float64) float64 {
	var t float64
	if g.euclidean {
		t = math.Hypot(px-g.fx, py-g.fy) * g.dInv
	} else {
		tx := math.Abs(px-g.fx) / g.farX
		ty := math.Abs(py-g.fy) / g.farY
		t = math.Max(tx, ty)
	}
	if t < 0 {
		t = 0
	} else if t > 1 {
		t = 1
	}
	return t
}

func (r *renderer) fillGradientPath(rect image.Rectangle, fill *Fill) {
	pos, cols := gradientStopsRGBA(fill)
	if pos == nil {
		return
	}
	w := rect.Dx()
	h := rect.Dy()
	if w <= 0 || h <= 0 {
		return
	}
	g := pathGradientGeometry(float64(w), float64(h), fill)

	pix := r.img.Pix
	bounds := r.img.Bounds()
	stride := r.img.Stride

	for py := rect.Min.Y; py < rect.Max.Y; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		pyLocal := float64(py-rect.Min.Y) + 0.5
		off := (py-bounds.Min.Y)*stride + (maxInt(rect.Min.X, bounds.Min.X)-bounds.Min.X)*4
		for px := maxInt(rect.Min.X, bounds.Min.X); px < minInt(rect.Max.X, bounds.Max.X); px++ {
			pxLocal := float64(px-rect.Min.X) + 0.5
			outC := gradStopColor(pos, cols, pathGradientT(g, pxLocal, pyLocal))
			pix[off] = outC[0]
			pix[off+1] = outC[1]
			pix[off+2] = outC[2]
			pix[off+3] = outC[3]
			off += 4
		}
	}
}

func lerpColor(a, b color.RGBA, t float64) color.RGBA {
	it := 1 - t
	return color.RGBA{
		R: uint8(float64(a.R)*it + float64(b.R)*t),
		G: uint8(float64(a.G)*it + float64(b.G)*t),
		B: uint8(float64(a.B)*it + float64(b.B)*t),
		A: uint8(float64(a.A)*it + float64(b.A)*t),
	}
}

// --- Shadow rendering ---

func (r *renderer) renderShadow(shadow *Shadow, rect image.Rectangle) {
	if r.draft || shadow == nil || !shadow.Visible {
		return
	}
	rad := float64(shadow.Direction) * math.Pi / 180.0
	// Distance is in points and scaleX is pixels per EMU: the offset used to
	// be dist*scaleX, which at 1600 px wide rounds to zero for any realistic
	// distance - shape shadows were parsed but never visibly displaced.
	dist := float64(shadow.Distance) * 12700 * r.scaleX
	dx := int(dist * math.Cos(rad))
	dy := int(dist * math.Sin(rad))
	shadowColor := argbToRGBA(shadow.Color)
	shadowColor.A = uint8(float64(shadow.Alpha) * 255 / 100)

	blurPx := int(float64(shadow.BlurRadius)*12700*r.scaleX + 0.5)
	if blurPx <= 0 {
		r.fillRectBlend(rect.Add(image.Pt(dx, dy)), shadowColor)
		return
	}

	// Rasterise the shape's silhouette as an offscreen alpha mask, box-blur
	// it and composite it offset - the same machinery the text shadow uses.
	// PowerPoint's blurRad names a Gaussian whose sigma is about half the
	// radius in pixels; a box blur of radius blurPx/2 lands on that sigma,
	// which is the fade length the COM export shows under a themed shape.
	radius := maxInt(1, blurPx/2)
	pad := radius*3 + 2
	bw := rect.Dx() + 2*pad
	bh := rect.Dy() + 2*pad
	if bw <= 0 || bh <= 0 || rect.Dx() <= 0 || rect.Dy() <= 0 {
		return
	}
	mask := image.NewAlpha(image.Rect(0, 0, bw, bh))
	for y := pad; y < pad+rect.Dy(); y++ {
		row := mask.Pix[y*mask.Stride:]
		for x := pad; x < pad+rect.Dx(); x++ {
			row[x] = 255
		}
	}
	boxBlurAlpha(mask, radius, 3)
	r.compositeShadowMask(shadowColor, mask, rect.Min.X-pad+dx, rect.Min.Y-pad+dy)
}

// renderShadowPolygon is the frame-outline variant of renderShadow: the
// silhouette rasterised into the blur mask is the given polygon rather than
// the bounding rectangle, so a snipped-corner picture frame does not spray
// shadow through its cut corners.
func (r *renderer) renderShadowPolygon(shadow *Shadow, pts []fpoint) {
	if r.draft || shadow == nil || !shadow.Visible || len(pts) < 3 {
		return
	}
	rad := float64(shadow.Direction) * math.Pi / 180.0
	dist := float64(shadow.Distance) * 12700 * r.scaleX
	dx := int(dist * math.Cos(rad))
	dy := int(dist * math.Sin(rad))
	shadowColor := argbToRGBA(shadow.Color)
	shadowColor.A = uint8(float64(shadow.Alpha) * 255 / 100)

	blurPx := int(float64(shadow.BlurRadius)*12700*r.scaleX + 0.5)
	if blurPx <= 0 {
		r.fillPolygon(pts, shadowColor)
		return
	}
	radius := maxInt(1, blurPx/2)
	pad := radius*3 + 2

	// Normalise the polygon into the padded mask's coordinate space.
	minX, minY := pts[0].x, pts[0].y
	maxX, maxY := minX, minY
	for _, p := range pts[1:] {
		minX = math.Min(minX, p.x)
		minY = math.Min(minY, p.y)
		maxX = math.Max(maxX, p.x)
		maxY = math.Max(maxY, p.y)
	}
	bw := int(maxX-minX) + 2*pad + 2
	bh := int(maxY-minY) + 2*pad + 2
	if bw <= 0 || bh <= 0 {
		return
	}
	shifted := make([]fpoint, len(pts))
	for i, p := range pts {
		shifted[i] = fpoint{p.x - minX + float64(pad), p.y - minY + float64(pad)}
	}
	mask := polygonAlphaMask(shifted, bw, bh)
	boxBlurAlpha(mask, radius, 3)
	r.compositeShadowMask(shadowColor, mask, int(minX)-pad+dx, int(minY)-pad+dy)
}

// compositeShadowMask blends a blurred shadow silhouette at the given offset.
func (r *renderer) compositeShadowMask(shadowColor color.RGBA, mask *image.Alpha, ox, oy int) {
	bounds := r.img.Bounds()
	baseA := float64(shadowColor.A) / 255
	bw := mask.Bounds().Dx()
	bh := mask.Bounds().Dy()
	for py := 0; py < bh; py++ {
		imgY := oy + py
		if imgY < bounds.Min.Y || imgY >= bounds.Max.Y {
			continue
		}
		for px := 0; px < bw; px++ {
			a := mask.AlphaAt(px, py).A
			if a == 0 {
				continue
			}
			imgX := ox + px
			if imgX < bounds.Min.X || imgX >= bounds.Max.X {
				continue
			}
			c := shadowColor
			c.A = uint8(float64(a)*baseA + 0.5)
			if c.A > 0 {
				r.blendPixel(imgX, imgY, c)
			}
		}
	}
}

// polygonAlphaMask rasterises a polygon (even-odd scanline, the same rule
// fillPolygon uses) into an alpha mask of the given size. Coordinates are in
// mask space.
func polygonAlphaMask(pts []fpoint, w, h int) *image.Alpha {
	mask := image.NewAlpha(image.Rect(0, 0, w, h))
	if len(pts) < 3 {
		return mask
	}
	n := len(pts)
	intersections := make([]float64, 0, n)
	for y := 0; y < h; y++ {
		fy := float64(y) + 0.5
		intersections = intersections[:0]
		for i := 0; i < n; i++ {
			j := (i + 1) % n
			py1, py2 := pts[i].y, pts[j].y
			if py1 > py2 {
				py1, py2 = py2, py1
			}
			if fy < py1 || fy >= py2 {
				continue
			}
			dy := pts[j].y - pts[i].y
			if dy == 0 {
				continue
			}
			t := (fy - pts[i].y) / dy
			intersections = append(intersections, pts[i].x+t*(pts[j].x-pts[i].x))
		}
		sort.Float64s(intersections)
		row := mask.Pix[y*mask.Stride:]
		for i := 0; i+1 < len(intersections); i += 2 {
			x1 := int(math.Ceil(intersections[i]))
			x2 := int(math.Floor(intersections[i+1]))
			for x := maxInt(x1, 0); x <= minInt(x2, w-1); x++ {
				row[x] = 255
			}
		}
	}
	return mask
}

func (r *renderer) renderShadowRounded(shadow *Shadow, rect image.Rectangle, radius int) {
	if r.draft || shadow == nil || !shadow.Visible {
		return
	}
	rad := float64(shadow.Direction) * math.Pi / 180.0
	// Same points→EMU→pixels conversion as renderShadow.
	dist := float64(shadow.Distance) * 12700 * r.scaleX
	dx := int(dist * math.Cos(rad))
	dy := int(dist * math.Sin(rad))
	shadowColor := argbToRGBA(shadow.Color)
	shadowColor.A = uint8(float64(shadow.Alpha) * 255 / 100)
	shadowRect := rect.Add(image.Pt(dx, dy))

	blur := shadow.BlurRadius
	if blur <= 0 {
		sw := shadowRect.Dx()
		sh := shadowRect.Dy()
		r.fillRoundedRect(shadowRect.Min.X, shadowRect.Min.Y, sw, sh, radius, shadowColor)
		return
	}

	steps := minInt(blur, 10)
	outerRect := shadowRect.Inset(-steps)
	tmpW := outerRect.Dx()
	tmpH := outerRect.Dy()
	if tmpW <= 0 || tmpH <= 0 {
		return
	}
	tmp := image.NewRGBA(image.Rect(0, 0, tmpW, tmpH))
	tmpR := r.subRenderer(tmp)

	for i := steps; i >= 0; i-- {
		t := float64(i) / float64(steps)
		alpha := uint8(float64(shadowColor.A) * (1 - t*t))
		c := color.RGBA{R: shadowColor.R, G: shadowColor.G, B: shadowColor.B, A: alpha}
		expanded := shadowRect.Inset(-i)
		ex := expanded.Min.X - outerRect.Min.X
		ey := expanded.Min.Y - outerRect.Min.Y
		ew := expanded.Dx()
		eh := expanded.Dy()
		er := radius + i
		tmpR.fillRoundedRect(ex, ey, ew, eh, er, c)
	}

	bounds := r.img.Bounds()
	for py := 0; py < tmpH; py++ {
		ddy := outerRect.Min.Y + py
		if ddy < bounds.Min.Y || ddy >= bounds.Max.Y {
			continue
		}
		for px := 0; px < tmpW; px++ {
			ddx := outerRect.Min.X + px
			if ddx < bounds.Min.X || ddx >= bounds.Max.X {
				continue
			}
			sc := tmp.RGBAAt(px, py)
			if sc.A == 0 {
				continue
			}
			r.blendPixel(ddx, ddy, sc)
		}
	}
}

// --- Drawing primitives ---

func (r *renderer) drawRect(rect image.Rectangle, c color.RGBA, width int) {
	for i := 0; i < width; i++ {
		// Top and bottom horizontal lines
		r.fillRectBlend(image.Rect(rect.Min.X, rect.Min.Y+i, rect.Max.X, rect.Min.Y+i+1), c)
		r.fillRectBlend(image.Rect(rect.Min.X, rect.Max.Y-1-i, rect.Max.X, rect.Max.Y-i), c)
		// Left and right vertical lines
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			r.blendPixel(rect.Min.X+i, y, c)
			r.blendPixel(rect.Max.X-1-i, y, c)
		}
	}
}

// picFramePoints returns the picture frame's preset outline in absolute
// coordinates, or nil when the frame is a plain rectangle. Only presets with
// a polygon point function are clipped; the rest draw unclipped, which is
// what every deck measured so far uses (snip2DiagRect on framed screenshots).
func (r *renderer) picFramePoints(prst string, x, y, w, h int) []fpoint {
	switch AutoShapeType(prst) {
	case AutoShapeSnip2DiagRect:
		return r.snip2DiagRectPoints(x, y, w, h, nil)
	}
	return nil
}

// shiftPts translates a polygon by (dx, dy).
func shiftPts(pts []fpoint, dx, dy float64) []fpoint {
	out := make([]fpoint, len(pts))
	for i, p := range pts {
		out[i] = fpoint{p.x + dx, p.y + dy}
	}
	return out
}

// applyDuotone repaints every pixel on the straight line between the two
// duotone colours at the pixel's Rec.709 luma.
//
// COM calibration (grayscale ramp + saturated patches, exported through
// PowerPoint): out = A + (B-A)*t with t = (0.2126R + 0.7152G + 0.0722B)/255,
// computed and interpolated in sRGB gamma space — the sixteen ramp bands land
// on round(A+(B-A)*v/255) band-exact, and the pure primaries map to
// 54/182/18 = the three weights times 255. Alpha passes through untouched.
func applyDuotone(img *image.RGBA, a, b Color) {
	ar, ag, ab := float64(a.GetRed()), float64(a.GetGreen()), float64(a.GetBlue())
	br, bg, bb := float64(b.GetRed()), float64(b.GetGreen()), float64(b.GetBlue())
	bounds := img.Bounds()
	for py := bounds.Min.Y; py < bounds.Max.Y; py++ {
		for px := bounds.Min.X; px < bounds.Max.X; px++ {
			i := img.PixOffset(px, py)
			a8 := img.Pix[i+3]
			if a8 == 0 {
				continue
			}
			// Un-premultiply for the luma, re-premultiply on the way back.
			inv := 255.0 / float64(a8)
			r := float64(img.Pix[i]) * inv
			g := float64(img.Pix[i+1]) * inv
			bl := float64(img.Pix[i+2]) * inv
			t := (0.2126*r + 0.7152*g + 0.0722*bl) / 255.0
			if t > 1 {
				t = 1
			} else if t < 0 {
				t = 0
			}
			img.Pix[i] = uint8((ar+(br-ar)*t)*float64(a8)/255.0 + 0.5)
			img.Pix[i+1] = uint8((ag+(bg-ag)*t)*float64(a8)/255.0 + 0.5)
			img.Pix[i+2] = uint8((ab+(bb-ab)*t)*float64(a8)/255.0 + 0.5)
		}
	}
}

func (r *renderer) drawRectBorder(rect image.Rectangle, c color.RGBA, width int, style BorderStyle) {
	if style == BorderSolid || style == BorderNone {
		r.drawRect(rect, c, width)
		return
	}
	dashLen, gapLen := 6, 4
	if style == BorderDot {
		dashLen, gapLen = 2, 2
	}
	for i := 0; i < width; i++ {
		r.drawDashedHLine(rect.Min.X, rect.Max.X, rect.Min.Y+i, c, dashLen, gapLen)
		r.drawDashedHLine(rect.Min.X, rect.Max.X, rect.Max.Y-1-i, c, dashLen, gapLen)
		r.drawDashedVLine(rect.Min.X+i, rect.Min.Y, rect.Max.Y, c, dashLen, gapLen)
		r.drawDashedVLine(rect.Max.X-1-i, rect.Min.Y, rect.Max.Y, c, dashLen, gapLen)
	}
}

func (r *renderer) drawDashedHLine(x1, x2, y int, c color.RGBA, dashLen, gapLen int) {
	period := dashLen + gapLen
	for x := x1; x < x2; x++ {
		if (x-x1)%period < dashLen {
			r.blendPixel(x, y, c)
		}
	}
}

func (r *renderer) drawDashedVLine(x, y1, y2 int, c color.RGBA, dashLen, gapLen int) {
	period := dashLen + gapLen
	for y := y1; y < y2; y++ {
		if (y-y1)%period < dashLen {
			r.blendPixel(x, y, c)
		}
	}
}

func (r *renderer) drawLineThick(x1, y1, x2, y2 int, c color.RGBA, width int) {
	if width <= 1 {
		r.drawLine(x1, y1, x2, y2, c)
		return
	}
	dx := float64(x2 - x1)
	dy := float64(y2 - y1)
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 0.5 {
		r.blendPixel(x1, y1, c)
		return
	}
	nx := -dy / length
	ny := dx / length
	hw := float64(width) / 2.0
	for i := 0; i < width; i++ {
		offset := -hw + float64(i) + 0.5
		r.drawLine(x1+int(offset*nx), y1+int(offset*ny), x2+int(offset*nx), y2+int(offset*ny), c)
	}
}

func (r *renderer) drawLine(x1, y1, x2, y2 int, c color.RGBA) {
	dx := abs(x2 - x1)
	dy := abs(y2 - y1)
	sx, sy := 1, 1
	if x1 > x2 {
		sx = -1
	}
	if y1 > y2 {
		sy = -1
	}
	err := dx - dy
	for {
		r.blendPixel(x1, y1, c)
		if x1 == x2 && y1 == y2 {
			break
		}
		e2 := 2 * err
		if e2 > -dy {
			err -= dy
			x1 += sx
		}
		if e2 < dx {
			err += dx
			y1 += sy
		}
	}
}

func (r *renderer) drawLineAA(x1, y1, x2, y2 int, c color.RGBA, width int) {
	if width <= 1 {
		r.drawLineWu(float64(x1), float64(y1), float64(x2), float64(y2), c)
		return
	}
	dx := float64(x2 - x1)
	dy := float64(y2 - y1)
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 0.5 {
		r.blendPixel(x1, y1, c)
		return
	}
	nx := -dy / length
	ny := dx / length
	hw := float64(width) / 2.0
	for i := 0; i < width; i++ {
		offset := -hw + float64(i) + 0.5
		ox := offset * nx
		oy := offset * ny
		r.drawLineWu(float64(x1)+ox, float64(y1)+oy, float64(x2)+ox, float64(y2)+oy, c)
	}
}

// drawDashedLineAA draws a dashed or dotted anti-aliased line.
func (r *renderer) drawDashedLineAA(x1, y1, x2, y2 int, c color.RGBA, width int, style BorderStyle) {
	if style == BorderSolid || style == BorderNone {
		r.drawLineAA(x1, y1, x2, y2, c, width)
		return
	}
	dx := float64(x2 - x1)
	dy := float64(y2 - y1)
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 1 {
		r.blendPixel(x1, y1, c)
		return
	}
	dashLen := 12.0
	gapLen := 6.0
	if style == BorderDot {
		dashLen = 3.0
		gapLen = 3.0
	}
	// Scale dash/gap by line width for visual consistency
	if width > 1 {
		dashLen *= float64(width) * 0.4
		gapLen *= float64(width) * 0.4
	}
	ux := dx / length
	uy := dy / length
	pos := 0.0
	drawing := true
	segStart := 0.0
	for pos < length {
		segLen := dashLen
		if !drawing {
			segLen = gapLen
		}
		segEnd := pos + segLen
		if segEnd > length {
			segEnd = length
		}
		if drawing {
			sx := x1 + int(ux*segStart)
			sy := y1 + int(uy*segStart)
			ex := x1 + int(ux*segEnd)
			ey := y1 + int(uy*segEnd)
			r.drawLineAA(sx, sy, ex, ey, c, width)
		}
		pos = segEnd
		segStart = segEnd
		drawing = !drawing
	}
}

// drawDashedPolylineAA draws a dashed/dotted polyline with continuous dash pattern
// across all segments, so the dash state carries over from one segment to the next.
func (r *renderer) drawDashedPolylineAA(pts []fpoint, c color.RGBA, width int, style BorderStyle) {
	if len(pts) < 2 {
		return
	}
	dashLen := 12.0
	gapLen := 6.0
	if style == BorderDot {
		dashLen = 3.0
		gapLen = 3.0
	}
	if width > 1 {
		dashLen *= float64(width) * 0.4
		gapLen *= float64(width) * 0.4
	}
	drawing := true
	remain := dashLen // remaining length in current dash/gap phase

	for i := 1; i < len(pts); i++ {
		sx, sy := pts[i-1].x, pts[i-1].y
		ex, ey := pts[i].x, pts[i].y
		dx := ex - sx
		dy := ey - sy
		segLen := math.Sqrt(dx*dx + dy*dy)
		if segLen < 0.5 {
			continue
		}
		ux := dx / segLen
		uy := dy / segLen
		pos := 0.0
		for pos < segLen {
			step := remain
			if pos+step > segLen {
				step = segLen - pos
			}
			if drawing {
				ax := int(sx + ux*pos)
				ay := int(sy + uy*pos)
				bx := int(sx + ux*(pos+step))
				by := int(sy + uy*(pos+step))
				r.drawLineAA(ax, ay, bx, by, c, width)
			}
			pos += step
			remain -= step
			if remain <= 0 {
				drawing = !drawing
				if drawing {
					remain = dashLen
				} else {
					remain = gapLen
				}
			}
		}
	}
}

// drawCubicBezierAA draws a cubic Bezier curve using adaptive subdivision.
func (r *renderer) drawCubicBezierAA(x0, y0, x1, y1, x2, y2, x3, y3 float64, c color.RGBA, width int) {
	// Flatten the Bezier into line segments
	pts := r.flattenCubicBezier(x0, y0, x1, y1, x2, y2, x3, y3, 0)
	pts = append([]fpoint{{x0, y0}}, pts...)
	pts = append(pts, fpoint{x3, y3})
	for i := 1; i < len(pts); i++ {
		r.drawLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), c, width)
	}
}

// drawDashedCubicBezierAA draws a dashed cubic Bezier curve.
func (r *renderer) drawDashedCubicBezierAA(x0, y0, x1, y1, x2, y2, x3, y3 float64, c color.RGBA, width int, style BorderStyle) {
	if style == BorderSolid || style == BorderNone {
		r.drawCubicBezierAA(x0, y0, x1, y1, x2, y2, x3, y3, c, width)
		return
	}
	pts := r.flattenCubicBezier(x0, y0, x1, y1, x2, y2, x3, y3, 0)
	pts = append([]fpoint{{x0, y0}}, pts...)
	pts = append(pts, fpoint{x3, y3})
	for i := 1; i < len(pts); i++ {
		r.drawDashedLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), c, width, style)
	}
}

// flattenCubicBezier recursively subdivides a cubic Bezier into line segments.
func (r *renderer) flattenCubicBezier(x0, y0, x1, y1, x2, y2, x3, y3 float64, depth int) []fpoint {
	if depth > 8 {
		return nil
	}
	// Check if the curve is flat enough
	dx := x3 - x0
	dy := y3 - y0
	d := math.Sqrt(dx*dx + dy*dy)
	if d < 0.5 {
		return nil
	}
	// Distance of control points from the line (x0,y0)-(x3,y3)
	d1 := math.Abs((x1-x0)*dy-(y1-y0)*dx) / d
	d2 := math.Abs((x2-x0)*dy-(y2-y0)*dx) / d
	if d1+d2 < 1.0 {
		return nil
	}
	// Subdivide at t=0.5
	mx01 := (x0 + x1) / 2
	my01 := (y0 + y1) / 2
	mx12 := (x1 + x2) / 2
	my12 := (y1 + y2) / 2
	mx23 := (x2 + x3) / 2
	my23 := (y2 + y3) / 2
	mx012 := (mx01 + mx12) / 2
	my012 := (my01 + my12) / 2
	mx123 := (mx12 + mx23) / 2
	my123 := (my12 + my23) / 2
	mx0123 := (mx012 + mx123) / 2
	my0123 := (my012 + my123) / 2

	left := r.flattenCubicBezier(x0, y0, mx01, my01, mx012, my012, mx0123, my0123, depth+1)
	right := r.flattenCubicBezier(mx0123, my0123, mx123, my123, mx23, my23, x3, y3, depth+1)
	result := append(left, fpoint{mx0123, my0123})
	result = append(result, right...)
	return result
}

func (r *renderer) drawLineWu(x0, y0, x1, y1 float64, c color.RGBA) {
	steep := math.Abs(y1-y0) > math.Abs(x1-x0)
	if steep {
		x0, y0 = y0, x0
		x1, y1 = y1, x1
	}
	if x0 > x1 {
		x0, x1 = x1, x0
		y0, y1 = y1, y0
	}
	dx := x1 - x0
	dy := y1 - y0
	gradient := 0.0
	if dx != 0 {
		gradient = dy / dx
	}

	// First endpoint
	xend := math.Round(x0)
	yend := y0 + gradient*(xend-x0)
	xgap := 1.0 - fpart(x0+0.5)
	xpxl1 := int(xend)
	ypxl1 := int(math.Floor(yend))
	if steep {
		r.blendPixelF(ypxl1, xpxl1, c, (1-fpart(yend))*xgap)
		r.blendPixelF(ypxl1+1, xpxl1, c, fpart(yend)*xgap)
	} else {
		r.blendPixelF(xpxl1, ypxl1, c, (1-fpart(yend))*xgap)
		r.blendPixelF(xpxl1, ypxl1+1, c, fpart(yend)*xgap)
	}
	intery := yend + gradient

	// Second endpoint
	xend = math.Round(x1)
	yend = y1 + gradient*(xend-x1)
	xgap = fpart(x1 + 0.5)
	xpxl2 := int(xend)
	ypxl2 := int(math.Floor(yend))
	if steep {
		r.blendPixelF(ypxl2, xpxl2, c, (1-fpart(yend))*xgap)
		r.blendPixelF(ypxl2+1, xpxl2, c, fpart(yend)*xgap)
	} else {
		r.blendPixelF(xpxl2, ypxl2, c, (1-fpart(yend))*xgap)
		r.blendPixelF(xpxl2, ypxl2+1, c, fpart(yend)*xgap)
	}

	for x := xpxl1 + 1; x < xpxl2; x++ {
		iy := int(math.Floor(intery))
		f := fpart(intery)
		if steep {
			r.blendPixelF(iy, x, c, 1-f)
			r.blendPixelF(iy+1, x, c, f)
		} else {
			r.blendPixelF(x, iy, c, 1-f)
			r.blendPixelF(x, iy+1, c, f)
		}
		intery += gradient
	}
}

func fpart(x float64) float64 { return x - math.Floor(x) }

// --- Ellipse rendering (anti-aliased) ---

func (r *renderer) fillEllipseAA(cx, cy, w, h int, c color.RGBA) {
	if w <= 0 || h <= 0 {
		return
	}
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry
	invRx2 := 1.0 / (rx * rx)
	invRy2 := 1.0 / (ry * ry)
	aaThreshold := 0.05

	bounds := r.img.Bounds()
	pix := r.img.Pix
	stride := r.img.Stride

	for py := cy; py < cy+h; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		dyNorm := float64(py) + 0.5 - centerY
		dy2 := dyNorm * dyNorm * invRy2
		if dy2 > 1.0 {
			continue
		}
		hExtent := rx * math.Sqrt(1.0-dy2)
		minPx := maxInt(int(centerX-hExtent), cx)
		maxPx := minInt(int(centerX+hExtent+1), cx+w)
		minPx = maxInt(minPx, bounds.Min.X)
		maxPx = minInt(maxPx, bounds.Max.X)

		rowOff := (py-bounds.Min.Y)*stride + (minPx-bounds.Min.X)*4
		for px := minPx; px < maxPx; px++ {
			dxNorm := float64(px) + 0.5 - centerX
			d := dxNorm*dxNorm*invRx2 + dy2
			if d <= 1.0 {
				edge := 1.0 - d
				if edge < aaThreshold {
					r.blendPixelF(px, py, c, edge/aaThreshold)
				} else if c.A == 255 {
					pix[rowOff] = c.R
					pix[rowOff+1] = c.G
					pix[rowOff+2] = c.B
					pix[rowOff+3] = 255
				} else {
					a := uint32(c.A)
					ia := 255 - a
					pix[rowOff] = uint8((uint32(c.R)*a + uint32(pix[rowOff])*ia) / 255)
					pix[rowOff+1] = uint8((uint32(c.G)*a + uint32(pix[rowOff+1])*ia) / 255)
					pix[rowOff+2] = uint8((uint32(c.B)*a + uint32(pix[rowOff+2])*ia) / 255)
					pix[rowOff+3] = uint8(uint32(pix[rowOff+3]) + (255-uint32(pix[rowOff+3]))*a/255)
				}
			}
			rowOff += 4
		}
	}

}

func (r *renderer) drawEllipseAA(cx, cy, w, h int, c color.RGBA, lineWidth int) {
	if w <= 0 || h <= 0 {
		return
	}
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry
	lw := float64(lineWidth)
	minR := math.Min(rx, ry)
	if minR < 1 {
		minR = 1
	}
	halfLW := lw / 2
	threshold := halfLW + 1

	for py := cy - lineWidth - 1; py < cy+h+lineWidth+1; py++ {
		dyNorm := (float64(py) + 0.5 - centerY) / ry
		dy2 := dyNorm * dyNorm
		if dy2 > 1.5 { // quick reject for rows far outside
			continue
		}
		for px := cx - lineWidth - 1; px < cx+w+lineWidth+1; px++ {
			dxNorm := (float64(px) + 0.5 - centerX) / rx
			d := math.Sqrt(dxNorm*dxNorm + dy2)
			distPx := math.Abs(d-1.0) * minR
			if distPx < threshold {
				coverage := 1.0
				if distPx > halfLW {
					coverage = 1.0 - (distPx - halfLW)
				}
				if coverage > 0 {
					r.blendPixelF(px, py, c, coverage)
				}
			}
		}
	}
}

// Legacy compatibility wrappers
func (r *renderer) fillEllipse(cx, cy, w, h int, c color.RGBA) { r.fillEllipseAA(cx, cy, w, h, c) }
func (r *renderer) drawEllipse(cx, cy, w, h int, c color.RGBA) { r.drawEllipseAA(cx, cy, w, h, c, 1) }

// --- Rounded rectangle ---

func (r *renderer) fillRoundedRect(x, y, w, h, radius int, c color.RGBA) {
	if radius <= 0 {
		r.fillRectBlend(image.Rect(x, y, x+w, y+h), c)
		return
	}
	radius = minInt(radius, minInt(w/2, h/2))
	r2 := float64(radius * radius)

	// Fill center rectangle (no corner checks needed)
	r.fillRectBlend(image.Rect(x+radius, y, x+w-radius, y+h), c)
	// Fill left/right strips (excluding corners)
	r.fillRectBlend(image.Rect(x, y+radius, x+radius, y+h-radius), c)
	r.fillRectBlend(image.Rect(x+w-radius, y+radius, x+w, y+h-radius), c)

	// Fill corners with circle test
	corners := [4][2]int{
		{x + radius, y + radius},         // top-left center
		{x + w - radius, y + radius},     // top-right center
		{x + radius, y + h - radius},     // bottom-left center
		{x + w - radius, y + h - radius}, // bottom-right center
	}
	cornerRects := [4]image.Rectangle{
		{Min: image.Pt(x, y), Max: image.Pt(x+radius, y+radius)},
		{Min: image.Pt(x+w-radius, y), Max: image.Pt(x+w, y+radius)},
		{Min: image.Pt(x, y+h-radius), Max: image.Pt(x+radius, y+h)},
		{Min: image.Pt(x+w-radius, y+h-radius), Max: image.Pt(x+w, y+h)},
	}
	for ci := 0; ci < 4; ci++ {
		ccx, ccy := corners[ci][0], corners[ci][1]
		cr := cornerRects[ci]
		for py := cr.Min.Y; py < cr.Max.Y; py++ {
			dy := float64(py - ccy)
			for px := cr.Min.X; px < cr.Max.X; px++ {
				dx := float64(px - ccx)
				if dx*dx+dy*dy <= r2 {
					r.blendPixel(px, py, c)
				}
			}
		}
	}
}

func (r *renderer) drawRoundedRect(x, y, w, h, radius int, c color.RGBA, lineWidth int) {
	r.drawLineThick(x+radius, y, x+w-radius, y, c, lineWidth)
	r.drawLineThick(x+radius, y+h-1, x+w-radius, y+h-1, c, lineWidth)
	r.drawLineThick(x, y+radius, x, y+h-radius, c, lineWidth)
	r.drawLineThick(x+w-1, y+radius, x+w-1, y+h-radius, c, lineWidth)
	r.drawArc(x, y, radius*2, radius*2, c, math.Pi, 1.5*math.Pi, lineWidth)
	r.drawArc(x+w-radius*2, y, radius*2, radius*2, c, 1.5*math.Pi, 2*math.Pi, lineWidth)
	r.drawArc(x, y+h-radius*2, radius*2, radius*2, c, 0.5*math.Pi, math.Pi, lineWidth)
	r.drawArc(x+w-radius*2, y+h-radius*2, radius*2, radius*2, c, 0, 0.5*math.Pi, lineWidth)
}

func (r *renderer) drawArc(cx, cy, w, h int, c color.RGBA, startAngle, endAngle float64, lineWidth int) {
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry
	// Use enough steps for smooth arc
	circumference := math.Pi * (rx + ry) * (endAngle - startAngle) / (2 * math.Pi)
	steps := maxInt(int(circumference*2), 30)
	angleStep := (endAngle - startAngle) / float64(steps)

	var prevPx, prevPy int
	for i := 0; i <= steps; i++ {
		angle := startAngle + angleStep*float64(i)
		px := int(centerX + rx*math.Cos(angle))
		py := int(centerY + ry*math.Sin(angle))
		if i > 0 && (px != prevPx || py != prevPy) {
			r.drawLineThick(prevPx, prevPy, px, py, c, lineWidth)
		}
		prevPx, prevPy = px, py
	}
}

// --- Can (cylinder) preset ---

// canTopRadius returns the vertical semi-axis of the can's top ellipse. The
// preset's adj is a percentage of the height with PowerPoint's default 25000,
// and the top ellipse takes half of that as its semi-axis — 12.5% of the
// height at the default, which is what the COM export measures.
func canTopRadius(h int, adjustValues map[string]int) int {
	adj := 25000
	if v, ok := adjustValues["adj"]; ok {
		adj = v
	}
	if adj < 0 {
		adj = 0
	}
	if adj > 50000 {
		adj = 50000
	}
	return h * adj / 200000
}

// fillCan fills the cylinder silhouette: a body between the two ellipse
// centres plus two full ellipses. The three overlaps are the same colour, so
// the union is the shape PowerPoint fills.
func (r *renderer) fillCan(x, y, w, h int, c color.RGBA, adjustValues map[string]int) {
	ry := canTopRadius(h, adjustValues)
	if 2*ry > h {
		ry = h / 2
	}
	if ry <= 0 {
		r.fillRectFast(image.Rect(x, y, x+w, y+h), c)
		return
	}
	r.fillRectFast(image.Rect(x, y+ry, x+w, y+h-ry), c)
	r.fillEllipseAA(x, y+h-2*ry, w, 2*ry, c)
	r.fillEllipseAA(x, y, w, 2*ry, c)
}

// drawCan strokes the cylinder outline: the top rim as a full ellipse, the
// two side lines, and the bottom cap as the ellipse's lower half.
func (r *renderer) drawCan(x, y, w, h int, c color.RGBA, lineWidth int, adjustValues map[string]int) {
	ry := canTopRadius(h, adjustValues)
	if 2*ry > h {
		ry = h / 2
	}
	if ry <= 0 {
		r.drawRect(image.Rect(x, y, x+w, y+h), c, lineWidth)
		return
	}
	r.drawEllipseAA(x, y, w, 2*ry, c, lineWidth)
	r.drawLineThick(x, y+ry, x, y+h-ry, c, lineWidth)
	r.drawLineThick(x+w-1, y+ry, x+w-1, y+h-ry, c, lineWidth)
	r.drawArc(x, y+h-2*ry, w, 2*ry, c, 0, math.Pi, lineWidth)
}

// --- Polygon shapes ---

type fpoint struct{ x, y float64 }

// fillPolygon fills a polygon using scanline algorithm with sort.Float64s.
func (r *renderer) fillPolygon(pts []fpoint, c color.RGBA) {
	if len(pts) < 3 {
		return
	}
	minY, maxY := pts[0].y, pts[0].y
	for _, p := range pts[1:] {
		if p.y < minY {
			minY = p.y
		}
		if p.y > maxY {
			maxY = p.y
		}
	}

	n := len(pts)
	// Pre-allocate intersection buffer
	intersections := make([]float64, 0, n)

	for y := int(minY); y <= int(maxY); y++ {
		fy := float64(y) + 0.5
		intersections = intersections[:0]
		for i := 0; i < n; i++ {
			j := (i + 1) % n
			py1, py2 := pts[i].y, pts[j].y
			if py1 > py2 {
				py1, py2 = py2, py1
			}
			if fy < py1 || fy >= py2 {
				continue
			}
			dy := pts[j].y - pts[i].y
			if dy == 0 {
				continue
			}
			t := (fy - pts[i].y) / dy
			intersections = append(intersections, pts[i].x+t*(pts[j].x-pts[i].x))
		}
		sort.Float64s(intersections)
		for i := 0; i+1 < len(intersections); i += 2 {
			x1 := int(math.Ceil(intersections[i]))
			x2 := int(math.Floor(intersections[i+1]))
			if x1 <= x2 {
				if c.A == 255 {
					r.fillRectFast(image.Rect(x1, y, x2+1, y+1), c)
				} else {
					r.fillRectBlend(image.Rect(x1, y, x2+1, y+1), c)
				}
			}
		}
	}
}

// fillPolygonGradient fills a polygon with a gradient. box is the shape's
// bounding rectangle — the gradient geometry (focus, tile, linear vector)
// lives in shape-box space, not the polygon's own bbox: a bent-up arrow only
// covers part of its box, and PowerPoint spreads its gradient across the
// whole box. Path gradients take the r39 pinned model; linear gradients keep
// the legacy bbox projection.
func (r *renderer) fillPolygonGradient(pts []fpoint, box image.Rectangle, fill *Fill) {
	if len(pts) < 3 || fill == nil {
		return
	}
	if fill.Type == FillGradientPath {
		r.fillPolygonPathGradient(pts, box, fill)
		return
	}
	startC := argbToRGBA(fill.Color)
	endC := argbToRGBA(fill.EndColor)

	// Compute bounding box
	minX, minY, maxX, maxY := pts[0].x, pts[0].y, pts[0].x, pts[0].y
	for _, p := range pts[1:] {
		if p.x < minX {
			minX = p.x
		}
		if p.y < minY {
			minY = p.y
		}
		if p.x > maxX {
			maxX = p.x
		}
		if p.y > maxY {
			maxY = p.y
		}
	}
	bw := maxX - minX
	bh := maxY - minY
	if bw <= 0 || bh <= 0 {
		return
	}

	rad := float64(fill.Rotation) * math.Pi / 180.0
	cosA := math.Cos(rad)
	sinA := math.Sin(rad)
	cx := bw / 2
	cy := bh / 2
	maxProj := math.Abs(cx*cosA) + math.Abs(cy*sinA)
	if maxProj < 1 {
		maxProj = 1
	}
	invMaxProj := 1.0 / (2 * maxProj)

	n := len(pts)
	intersections := make([]float64, 0, n)
	bounds := r.img.Bounds()
	pix := r.img.Pix
	stride := r.img.Stride

	for y := int(minY); y <= int(maxY); y++ {
		if y < bounds.Min.Y || y >= bounds.Max.Y {
			continue
		}
		fy := float64(y) + 0.5
		intersections = intersections[:0]
		for i := 0; i < n; i++ {
			j := (i + 1) % n
			py1, py2 := pts[i].y, pts[j].y
			if py1 > py2 {
				py1, py2 = py2, py1
			}
			if fy < py1 || fy >= py2 {
				continue
			}
			dy := pts[j].y - pts[i].y
			if dy == 0 {
				continue
			}
			t := (fy - pts[i].y) / dy
			intersections = append(intersections, pts[i].x+t*(pts[j].x-pts[i].x))
		}
		sort.Float64s(intersections)

		dyf := float64(y) - minY - cy
		rowBase := dyf*sinA + maxProj

		for i := 0; i+1 < len(intersections); i += 2 {
			x1 := int(math.Ceil(intersections[i]))
			x2 := int(math.Floor(intersections[i+1]))
			if x1 > x2 {
				continue
			}
			if x1 < bounds.Min.X {
				x1 = bounds.Min.X
			}
			if x2 >= bounds.Max.X {
				x2 = bounds.Max.X - 1
			}
			off := (y-bounds.Min.Y)*stride + (x1-bounds.Min.X)*4
			for px := x1; px <= x2; px++ {
				dxf := float64(px) - minX - cx
				t := (dxf*cosA + rowBase) * invMaxProj
				if t < 0 {
					t = 0
				} else if t > 1 {
					t = 1
				}
				it := 1 - t
				pix[off] = uint8(float64(startC.R)*it + float64(endC.R)*t)
				pix[off+1] = uint8(float64(startC.G)*it + float64(endC.G)*t)
				pix[off+2] = uint8(float64(startC.B)*it + float64(endC.B)*t)
				pix[off+3] = uint8(float64(startC.A)*it + float64(endC.A)*t)
				off += 4
			}
		}
	}
}

func (r *renderer) drawPolygon(pts []fpoint, c color.RGBA, width int) {
	n := len(pts)
	for i := 0; i < n; i++ {
		j := (i + 1) % n
		r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[j].x), int(pts[j].y), c, width)
	}
}

// fillPolygonPathGradient fills a polygon with a path gradient whose geometry
// (focus, tile) lives in shape-box space. Each scanline is clipped to the
// polygon exactly like fillPolygon does; the colour comes from the box-space
// gradient parameter, so the ramp lines up with sibling shapes that share the
// box (an arrow's head and shaft continue one another's gradient).
func (r *renderer) fillPolygonPathGradient(pts []fpoint, box image.Rectangle, fill *Fill) {
	pos, cols := gradientStopsRGBA(fill)
	if pos == nil || box.Dx() <= 0 || box.Dy() <= 0 {
		return
	}
	g := pathGradientGeometry(float64(box.Dx()), float64(box.Dy()), fill)

	minY, maxY := pts[0].y, pts[0].y
	for _, p := range pts[1:] {
		if p.y < minY {
			minY = p.y
		}
		if p.y > maxY {
			maxY = p.y
		}
	}

	n := len(pts)
	intersections := make([]float64, 0, n)
	bounds := r.img.Bounds()
	pix := r.img.Pix
	stride := r.img.Stride

	for y := int(minY); y <= int(maxY); y++ {
		if y < bounds.Min.Y || y >= bounds.Max.Y {
			continue
		}
		fy := float64(y) + 0.5
		intersections = intersections[:0]
		for i := 0; i < n; i++ {
			j := (i + 1) % n
			py1, py2 := pts[i].y, pts[j].y
			if py1 > py2 {
				py1, py2 = py2, py1
			}
			if fy < py1 || fy >= py2 {
				continue
			}
			dy := pts[j].y - pts[i].y
			if dy == 0 {
				continue
			}
			t := (fy - pts[i].y) / dy
			intersections = append(intersections, pts[i].x+t*(pts[j].x-pts[i].x))
		}
		sort.Float64s(intersections)

		pyLocal := float64(y-box.Min.Y) + 0.5
		for i := 0; i+1 < len(intersections); i += 2 {
			x1 := int(math.Ceil(intersections[i]))
			x2 := int(math.Floor(intersections[i+1]))
			if x1 > x2 {
				continue
			}
			if x1 < bounds.Min.X {
				x1 = bounds.Min.X
			}
			if x2 >= bounds.Max.X {
				x2 = bounds.Max.X - 1
			}
			off := (y-bounds.Min.Y)*stride + (x1-bounds.Min.X)*4
			for px := x1; px <= x2; px++ {
				pxLocal := float64(px-box.Min.X) + 0.5
				outC := gradStopColor(pos, cols, pathGradientT(g, pxLocal, pyLocal))
				pix[off] = outC[0]
				pix[off+1] = outC[1]
				pix[off+2] = outC[2]
				pix[off+3] = outC[3]
				off += 4
			}
		}
	}
}

func (r *renderer) fillTriangle(x, y, w, h int, c color.RGBA) {
	r.fillPolygon([]fpoint{
		{float64(x) + float64(w)/2, float64(y)},
		{float64(x + w), float64(y + h)},
		{float64(x), float64(y + h)},
	}, c)
}

func (r *renderer) drawTriangle(x, y, w, h int, c color.RGBA, width int) {
	r.drawPolygon([]fpoint{
		{float64(x) + float64(w)/2, float64(y)},
		{float64(x + w), float64(y + h)},
		{float64(x), float64(y + h)},
	}, c, width)
}

func (r *renderer) fillDiamond(x, y, w, h int, c color.RGBA) {
	cx, cy := float64(x)+float64(w)/2, float64(y)+float64(h)/2
	r.fillPolygon([]fpoint{{cx, float64(y)}, {float64(x + w), cy}, {cx, float64(y + h)}, {float64(x), cy}}, c)
}

func (r *renderer) drawDiamond(x, y, w, h int, c color.RGBA, width int) {
	cx, cy := float64(x)+float64(w)/2, float64(y)+float64(h)/2
	r.drawPolygon([]fpoint{{cx, float64(y)}, {float64(x + w), cy}, {cx, float64(y + h)}, {float64(x), cy}}, c, width)
}

func (r *renderer) fillRegularPolygon(x, y, w, h, sides int, startAngle float64, c color.RGBA) {
	pts := regularPolygonPoints(x, y, w, h, sides, startAngle)
	r.fillPolygon(pts, c)
}

func regularPolygonPoints(x, y, w, h, sides int, startAngle float64) []fpoint {
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2
	rx := float64(w) / 2
	ry := float64(h) / 2
	pts := make([]fpoint, sides)
	for i := 0; i < sides; i++ {
		angle := startAngle + float64(i)*2*math.Pi/float64(sides)
		pts[i] = fpoint{cx + rx*math.Cos(angle), cy + ry*math.Sin(angle)}
	}
	return pts
}

func (r *renderer) fillPentagon(x, y, w, h int, c color.RGBA) {
	r.fillRegularPolygon(x, y, w, h, 5, -math.Pi/2, c)
}

func (r *renderer) fillHexagon(x, y, w, h int, c color.RGBA) {
	r.fillRegularPolygon(x, y, w, h, 6, 0, c)
}

// flowChartPreparationPoints returns the 6 vertices for the OOXML flowChartPreparation shape.
// Unlike a regular hexagon inscribed in an ellipse, this shape fills the entire bounding box:
// left/right points at mid-height, top/bottom edges span full width minus an inset.
func flowChartPreparationPoints(x, y, w, h int) []fpoint {
	inset := float64(w) / 5.0
	fx, fy, fw, fh := float64(x), float64(y), float64(w), float64(h)
	return []fpoint{
		{fx, fy + fh/2},            // left point
		{fx + inset, fy},           // top-left
		{fx + fw - inset, fy},      // top-right
		{fx + fw, fy + fh/2},       // right point
		{fx + fw - inset, fy + fh}, // bottom-right
		{fx + inset, fy + fh},      // bottom-left
	}
}

func (r *renderer) fillFlowChartPreparation(x, y, w, h int, c color.RGBA) {
	pts := flowChartPreparationPoints(x, y, w, h)
	r.fillPolygon(pts, c)
}

func (r *renderer) fillStar(x, y, w, h, points int, c color.RGBA) {
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2
	outerRx, outerRy := float64(w)/2, float64(h)/2
	innerRx, innerRy := outerRx*0.4, outerRy*0.4
	n := points * 2
	pts := make([]fpoint, n)
	for i := 0; i < n; i++ {
		angle := -math.Pi/2 + float64(i)*2*math.Pi/float64(n)
		rx, ry := outerRx, outerRy
		if i%2 == 1 {
			rx, ry = innerRx, innerRy
		}
		pts[i] = fpoint{cx + rx*math.Cos(angle), cy + ry*math.Sin(angle)}
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillArrowRight(x, y, w, h int, c color.RGBA) {
	shaftH := float64(h) * 0.4
	headW := float64(w) * 0.35
	shaftW := float64(w) - headW
	top := float64(y) + (float64(h)-shaftH)/2
	bot := top + shaftH
	r.fillPolygon([]fpoint{
		{float64(x), top}, {float64(x) + shaftW, top}, {float64(x) + shaftW, float64(y)},
		{float64(x + w), float64(y) + float64(h)/2},
		{float64(x) + shaftW, float64(y + h)}, {float64(x) + shaftW, bot}, {float64(x), bot},
	}, c)
}

func (r *renderer) fillArrowLeft(x, y, w, h int, c color.RGBA) {
	shaftH := float64(h) * 0.4
	headW := float64(w) * 0.35
	top := float64(y) + (float64(h)-shaftH)/2
	bot := top + shaftH
	r.fillPolygon([]fpoint{
		{float64(x + w), top}, {float64(x) + headW, top}, {float64(x) + headW, float64(y)},
		{float64(x), float64(y) + float64(h)/2},
		{float64(x) + headW, float64(y + h)}, {float64(x) + headW, bot}, {float64(x + w), bot},
	}, c)
}

func (r *renderer) fillArrowUp(x, y, w, h int, c color.RGBA) {
	shaftW := float64(w) * 0.4
	headH := float64(h) * 0.35
	left := float64(x) + (float64(w)-shaftW)/2
	right := left + shaftW
	r.fillPolygon([]fpoint{
		{float64(x) + float64(w)/2, float64(y)},
		{float64(x + w), float64(y) + headH}, {right, float64(y) + headH},
		{right, float64(y + h)}, {left, float64(y + h)},
		{left, float64(y) + headH}, {float64(x), float64(y) + headH},
	}, c)
}

func (r *renderer) fillArrowDown(x, y, w, h int, c color.RGBA) {
	shaftW := float64(w) * 0.4
	headH := float64(h) * 0.35
	shaftTop := float64(h) - headH
	left := float64(x) + (float64(w)-shaftW)/2
	right := left + shaftW
	r.fillPolygon([]fpoint{
		{left, float64(y)}, {right, float64(y)},
		{right, float64(y) + shaftTop}, {float64(x + w), float64(y) + shaftTop},
		{float64(x) + float64(w)/2, float64(y + h)},
		{float64(x), float64(y) + shaftTop}, {left, float64(y) + shaftTop},
	}, c)
}

func (r *renderer) fillHeart(x, y, w, h int, c color.RGBA) {
	cx := float64(x) + float64(w)/2
	topY := float64(y) + float64(h)*0.3
	halfW := float64(w) / 2
	hScale := float64(h) * 0.7

	for py := y; py < y+h; py++ {
		ny := 1 - (float64(py)-topY)/hScale
		ny2 := ny * ny
		ny3 := ny2 * ny
		for px := x; px < x+w; px++ {
			nx := (float64(px) - cx) / halfW
			nx2 := nx * nx
			val := (nx2 + ny2 - 1)
			val = val * val * val
			val -= nx2 * ny3
			if val <= 0 {
				r.blendPixel(px, py, c)
			}
		}
	}
}

func (r *renderer) fillPlus(x, y, w, h int, c color.RGBA) {
	armW := w / 3
	armH := h / 3
	r.fillRectBlend(image.Rect(x, y+armH, x+w, y+h-armH), c)
	r.fillRectBlend(image.Rect(x+armW, y, x+w-armW, y+h), c)
}

func (r *renderer) fillChevron(x, y, w, h int, c color.RGBA) {
	notch := w / 4
	pts := []fpoint{
		{float64(x), float64(y)},
		{float64(x + w - notch), float64(y)},
		{float64(x + w), float64(y + h/2)},
		{float64(x + w - notch), float64(y + h)},
		{float64(x), float64(y + h)},
		{float64(x + notch), float64(y + h/2)},
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillParallelogram(x, y, w, h int, c color.RGBA) {
	offset := w / 4
	pts := []fpoint{
		{float64(x + offset), float64(y)},
		{float64(x + w), float64(y)},
		{float64(x + w - offset), float64(y + h)},
		{float64(x), float64(y + h)},
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillLeftRightArrow(x, y, w, h int, c color.RGBA) {
	headW := w / 4
	bodyH := h / 3
	pts := []fpoint{
		{float64(x), float64(y + h/2)},
		{float64(x + headW), float64(y)},
		{float64(x + headW), float64(y + bodyH)},
		{float64(x + w - headW), float64(y + bodyH)},
		{float64(x + w - headW), float64(y)},
		{float64(x + w), float64(y + h/2)},
		{float64(x + w - headW), float64(y + h)},
		{float64(x + w - headW), float64(y + h - bodyH)},
		{float64(x + headW), float64(y + h - bodyH)},
		{float64(x + headW), float64(y + h)},
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillRtTriangle(x, y, w, h int, c color.RGBA) {
	pts := []fpoint{
		{float64(x), float64(y + h)},
		{float64(x), float64(y)},
		{float64(x + w), float64(y + h)},
	}
	r.fillPolygon(pts, c)
}

func (r *renderer) fillHomePlate(x, y, w, h int, c color.RGBA) {
	notch := w / 5
	pts := []fpoint{
		{float64(x), float64(y)},
		{float64(x + w - notch), float64(y)},
		{float64(x + w), float64(y + h/2)},
		{float64(x + w - notch), float64(y + h)},
		{float64(x), float64(y + h)},
	}
	r.fillPolygon(pts, c)
}

// fillWedgeRoundRectCallout draws a rounded-rectangle callout shape.
// OOXML wedgeRoundRectCallout has three adjust values:
//
//	adj1: X offset of callout tip from center (1/100000 of width, default -20833)
//	adj2: Y offset of callout tip from center (1/100000 of height, default 62500)
//	adj3: corner radius (1/100000 of min(w,h), default 16667)
func (r *renderer) fillWedgeRoundRectCallout(x, y, w, h int, c color.RGBA, adj map[string]int) {
	adj1v := -20833
	adj2v := 62500
	adj3v := 16667
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
		if v, ok := adj["adj3"]; ok {
			adj3v = v
		}
	}
	fw, fh := float64(w), float64(h)
	ss := math.Min(fw, fh)
	radius := int(ss * float64(adj3v) / 100000.0)
	if radius < 0 {
		radius = 0
	}

	// Draw the rounded rectangle body
	r.fillRoundedRect(x, y, w, h, radius, c)

	// Compute callout tip position (relative to shape top-left)
	tipX := float64(x) + fw/2 + fw*float64(adj1v)/100000.0
	tipY := float64(y) + fh/2 + fh*float64(adj2v)/100000.0

	// Determine wedge base: two points on the nearest edge of the rectangle.
	// The wedge base width is about 1/6 of the edge length.
	fx, fy := float64(x), float64(y)
	cx, cy := fx+fw/2, fy+fh/2

	// Determine which edge the wedge originates from based on tip position
	dx := tipX - cx
	dy := tipY - cy
	var bx1, by1, bx2, by2 float64
	wedgeHalf := fw / 12 // half-width of wedge base

	if math.Abs(dy)*fw >= math.Abs(dx)*fh {
		// Tip is more above/below → wedge on top or bottom edge
		wedgeHalf = fw / 12
		baseCX := tipX
		if baseCX < fx+float64(radius) {
			baseCX = fx + float64(radius)
		}
		if baseCX > fx+fw-float64(radius) {
			baseCX = fx + fw - float64(radius)
		}
		if dy >= 0 {
			// Bottom edge
			bx1, by1 = baseCX-wedgeHalf, fy+fh
			bx2, by2 = baseCX+wedgeHalf, fy+fh
		} else {
			// Top edge
			bx1, by1 = baseCX+wedgeHalf, fy
			bx2, by2 = baseCX-wedgeHalf, fy
		}
	} else {
		// Tip is more left/right → wedge on left or right edge
		wedgeHalf = fh / 12
		baseCY := tipY
		if baseCY < fy+float64(radius) {
			baseCY = fy + float64(radius)
		}
		if baseCY > fy+fh-float64(radius) {
			baseCY = fy + fh - float64(radius)
		}
		if dx >= 0 {
			// Right edge
			bx1, by1 = fx+fw, baseCY-wedgeHalf
			bx2, by2 = fx+fw, baseCY+wedgeHalf
		} else {
			// Left edge
			bx1, by1 = fx, baseCY+wedgeHalf
			bx2, by2 = fx, baseCY-wedgeHalf
		}
	}

	// Draw the wedge triangle
	wedge := []fpoint{
		{bx1, by1},
		{tipX, tipY},
		{bx2, by2},
	}
	r.fillPolygon(wedge, c)
}

// drawWedgeRoundRectCalloutBorder draws the border of a wedgeRoundRectCallout shape.
func (r *renderer) drawWedgeRoundRectCalloutBorder(x, y, w, h int, bc color.RGBA, pw int, adj map[string]int) {
	adj1v := -20833
	adj2v := 62500
	adj3v := 16667
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
		if v, ok := adj["adj3"]; ok {
			adj3v = v
		}
	}
	fw, fh := float64(w), float64(h)
	ss := math.Min(fw, fh)
	radius := int(ss * float64(adj3v) / 100000.0)
	if radius < 0 {
		radius = 0
	}

	tipX := float64(x) + fw/2 + fw*float64(adj1v)/100000.0
	tipY := float64(y) + fh/2 + fh*float64(adj2v)/100000.0

	fx, fy := float64(x), float64(y)
	cx, cy := fx+fw/2, fy+fh/2
	dx := tipX - cx
	dy := tipY - cy

	// Determine wedge base position and which edge it's on
	type wedgeInfo struct {
		bx1, by1, bx2, by2 float64
		edge               int // 0=bottom, 1=top, 2=right, 3=left
	}
	var wi wedgeInfo
	if math.Abs(dy)*fw >= math.Abs(dx)*fh {
		wedgeHalf := fw / 12
		baseCX := tipX
		if baseCX < fx+float64(radius) {
			baseCX = fx + float64(radius)
		}
		if baseCX > fx+fw-float64(radius) {
			baseCX = fx + fw - float64(radius)
		}
		if dy >= 0 {
			wi = wedgeInfo{baseCX - wedgeHalf, fy + fh, baseCX + wedgeHalf, fy + fh, 0}
		} else {
			wi = wedgeInfo{baseCX + wedgeHalf, fy, baseCX - wedgeHalf, fy, 1}
		}
	} else {
		wedgeHalf := fh / 12
		baseCY := tipY
		if baseCY < fy+float64(radius) {
			baseCY = fy + float64(radius)
		}
		if baseCY > fy+fh-float64(radius) {
			baseCY = fy + fh - float64(radius)
		}
		if dx >= 0 {
			wi = wedgeInfo{fx + fw, baseCY - wedgeHalf, fx + fw, baseCY + wedgeHalf, 2}
		} else {
			wi = wedgeInfo{fx, baseCY + wedgeHalf, fx, baseCY - wedgeHalf, 3}
		}
	}

	// Draw rounded rect border with gap for wedge, then draw wedge lines.
	// For simplicity, draw the full rounded rect border then overdraw the wedge.
	r.drawRoundedRect(x, y, w, h, radius, bc, pw)

	// Draw wedge lines from base points to tip
	r.drawLineAA(int(wi.bx1), int(wi.by1), int(tipX), int(tipY), bc, pw)
	r.drawLineAA(int(tipX), int(tipY), int(wi.bx2), int(wi.by2), bc, pw)
}

// snip2SameRectPoints computes the polygon points for a snip2SameRect shape.
// In OOXML snip2SameRect, adj1 controls the bottom-left and bottom-right snip,
// adj2 controls the top-left and top-right snip.
func (r *renderer) snip2SameRectPoints(x, y, w, h int, adj map[string]int) []fpoint {
	adj1v := 16667 // default snip for bottom corners
	adj2v := 0     // default snip for top corners
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
	}
	ss := minInt(w, h)
	snipBot := float64(ss) * float64(adj1v) / 100000.0
	snipTop := float64(ss) * float64(adj2v) / 100000.0
	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)

	return []fpoint{
		{fx + snipTop, fy},           // top-left snip end
		{fx + fw - snipTop, fy},      // top-right snip start
		{fx + fw, fy + snipTop},      // top-right snip end
		{fx + fw, fy + fh - snipBot}, // bottom-right snip start
		{fx + fw - snipBot, fy + fh}, // bottom-right snip end
		{fx + snipBot, fy + fh},      // bottom-left snip start
		{fx, fy + fh - snipBot},      // bottom-left snip end
		{fx, fy + snipTop},           // top-left snip start
	}
}

func (r *renderer) fillSnip2SameRect(x, y, w, h int, c color.RGBA, adj map[string]int) {
	pts := r.snip2SameRectPoints(x, y, w, h, adj)
	r.fillPolygon(pts, c)
}

// bentUpArrowPoints builds the bentUpArrow preset outline (bottom
// horizontal bar, right end bending up into an arrowhead). Measured against
// PowerPoint COM exports at adj defaults and eight variant combinations:
//
//	shaft thickness = min(adj1, 50000)·ss           (adj1 default 25000)
//	head base width = min(adj2·2, 100000)·ss        (adj2 default 25000)
//	head triangle height = min(adj3, 50000)·ss      (adj3 default 25000)
//
// with ss = min(w,h). The head's base right corner sits on the shape's right
// edge and the vertical arm is centred on the head's base midpoint; arm and
// bar share the shaft thickness.
func (r *renderer) bentUpArrowPoints(x, y, w, h int, adj map[string]int) []fpoint {
	adj1v, adj2v, adj3v := 25000, 25000, 25000
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
		if v, ok := adj["adj3"]; ok {
			adj3v = v
		}
	}
	ss := float64(minInt(w, h))
	st := ss * math.Min(float64(adj1v)/100000.0, 0.5)
	tw := ss * math.Min(float64(adj2v)/100000.0*2.0, 1.0)
	th := ss * math.Min(float64(adj3v)/100000.0, 0.5)

	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)
	cx := fx + fw - tw/2 // arm centre: head base midpoint, base right corner on the right edge
	baseY := fy + th     // head base (triangle bottom edge)
	barTop := fy + fh - st
	armL, armR := cx-st/2, cx+st/2

	return []fpoint{
		{fx, fy + fh},      // bar bottom-left
		{armR, fy + fh},    // bar bottom-right (= arm bottom-right)
		{armR, baseY},      // arm right edge up to the head base
		{cx + tw/2, baseY}, // head base right corner
		{cx, fy},           // head apex (top edge)
		{cx - tw/2, baseY}, // head base left corner
		{armL, baseY},      // arm left edge at the head base
		{armL, barTop},     // arm left edge down to the bar top
		{fx, barTop},       // bar top-left
	}
}

// snip2DiagRectPoints builds the snip2DiagRect preset outline: a rectangle
// with the top-right and bottom-left corners cut off at 45 degrees. Each
// snip measures adj·ss/100000 along both edges (default 16667, max 50000).
func (r *renderer) snip2DiagRectPoints(x, y, w, h int, adj map[string]int) []fpoint {
	adj1v, adj2v := 16667, 16667
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
	}
	ss := float64(minInt(w, h))
	snipBR := ss * math.Min(float64(adj1v)/100000.0, 0.5) // bottom-left corner
	snipTR := ss * math.Min(float64(adj2v)/100000.0, 0.5) // top-right corner
	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)

	return []fpoint{
		{fx, fy},               // top-left
		{fx + fw - snipTR, fy}, // top-right snip start
		{fx + fw, fy + snipTR}, // top-right snip end
		{fx + fw, fy + fh},     // bottom-right
		{fx + snipBR, fy + fh}, // bottom-left snip end (on bottom edge)
		{fx, fy + fh - snipBR}, // bottom-left snip start (on left edge)
	}
}

func (r *renderer) fillBentArrow(x, y, w, h int, c color.RGBA, adj map[string]int) {
	// OOXML bentArrow preset geometry.
	// L-shaped arrow: vertical shaft going up, then turns right with arrowhead.
	// adj1 = shaft width as fraction of width / 100000 (default 25000)
	// adj2 = arrowhead extra width / 100000 (default 25000)
	// adj3 = arrowhead length as fraction of width / 100000 (default 25000)
	// adj4 = bend position as fraction of height / 100000 (default 43750)
	adj1v := 25000
	adj2v := 25000
	adj3v := 25000
	adj4v := 43750
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
		if v, ok := adj["adj3"]; ok {
			adj3v = v
		}
		if v, ok := adj["adj4"]; ok {
			adj4v = v
		}
	}

	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)

	shaftW := fw * float64(adj1v) / 100000.0
	headExtra := fw * float64(adj2v) / 100000.0
	headLen := fw * float64(adj3v) / 100000.0
	bendY := fy + fh*float64(adj4v)/100000.0

	tipX := fx + fw
	arrowCenterY := bendY - shaftW/2
	arrowBaseX := tipX - headLen
	arrowTop := arrowCenterY - shaftW/2 - headExtra
	arrowBot := arrowCenterY + shaftW/2 + headExtra

	// Corner radius for rounded corners
	cornerR := shaftW * 0.85
	if cornerR < 1 {
		cornerR = 1
	}

	pts := []fpoint{
		{fx, fy + fh}, // bottom-left
	}

	// Outer corner: rounded arc from vertical outer edge to horizontal top
	// The outer corner is at (fx, bendY - shaftW)
	outerCornerX := fx
	outerCornerY := bendY - shaftW
	outerR := cornerR
	// Clamp outer radius so it doesn't exceed available space
	maxOuterR := math.Min(outerCornerY-(fy), fw*0.3)
	if outerR > maxOuterR && maxOuterR > 0 {
		outerR = maxOuterR
	}
	// Arc from vertical (going up) to horizontal (going right)
	// Arc center at (outerCornerX + outerR, outerCornerY + outerR)
	ocx := outerCornerX + outerR
	ocy := outerCornerY + outerR
	arcSteps := 12
	// Start point: on the vertical edge, approaching the corner from below
	pts = append(pts, fpoint{fx, ocy})
	for i := 0; i <= arcSteps; i++ {
		t := float64(i) / float64(arcSteps)
		angle := math.Pi + t*math.Pi/2.0 // π to 3π/2
		ax := ocx + outerR*math.Cos(angle)
		ay := ocy + outerR*math.Sin(angle)
		pts = append(pts, fpoint{ax, ay})
	}

	pts = append(pts,
		fpoint{arrowBaseX, bendY - shaftW}, // top edge to arrowhead base
		fpoint{arrowBaseX, arrowTop},       // arrowhead top
		fpoint{tipX, arrowCenterY},         // arrowhead tip
		fpoint{arrowBaseX, arrowBot},       // arrowhead bottom
		fpoint{arrowBaseX, bendY},          // bottom of horizontal shaft
	)

	// Inner corner: rounded arc from horizontal bottom to vertical inner edge
	innerX := fx + shaftW
	innerR := cornerR
	// Clamp inner radius
	maxInnerR := math.Min(fh-fh*float64(adj4v)/100000.0, shaftW*0.9)
	if innerR > maxInnerR && maxInnerR > 0 {
		innerR = maxInnerR
	}
	cxArc := innerX + innerR
	cyArc := bendY + innerR
	pts = append(pts, fpoint{cxArc, bendY}) // start of inner arc
	for i := 0; i <= arcSteps; i++ {
		t := float64(i) / float64(arcSteps)
		angle := math.Pi/2.0 + t*math.Pi/2.0 // π/2 to π
		ax := cxArc + innerR*math.Cos(angle)
		ay := cyArc - innerR*math.Sin(angle)
		pts = append(pts, fpoint{ax, ay})
	}

	pts = append(pts, fpoint{innerX, fy + fh}) // bottom of inner vertical edge
	r.fillPolygon(pts, c)
}

func (r *renderer) fillUturnArrow(x, y, w, h int, c color.RGBA, adj map[string]int) {
	// OOXML uturnArrow preset geometry (from presetShapeDefinitions.xml).
	// Two vertical shafts connected by arcs at the TOP.
	// The RIGHT shaft has an arrowhead pointing DOWN.
	// The U-turn opening is at the BOTTOM.
	//
	// adj1 = shaft thickness (fraction of ss / 100000, default 25000)
	// adj2 = arrowhead half-extra width (fraction of ss / 100000, default 25000)
	// adj3 = arrowhead height (fraction of ss / 100000, default 25000)
	// adj4 = bend diameter (fraction of ss / 100000, default 43750)
	// adj5 = total height used (fraction of h / 100000, default 75000)
	adj1v := 25000
	adj2v := 25000
	adj3v := 25000
	adj4v := 43750
	adj5v := 75000
	if adj != nil {
		if v, ok := adj["adj1"]; ok {
			adj1v = v
		}
		if v, ok := adj["adj2"]; ok {
			adj2v = v
		}
		if v, ok := adj["adj3"]; ok {
			adj3v = v
		}
		if v, ok := adj["adj4"]; ok {
			adj4v = v
		}
		if v, ok := adj["adj5"]; ok {
			adj5v = v
		}
	}

	fw := float64(w)
	fh := float64(h)
	fx := float64(x)
	fy := float64(y)
	ss := math.Min(fw, fh)

	// Guide calculations per OOXML spec
	th := ss * float64(adj1v) / 100000.0  // shaft thickness
	aw2 := ss * float64(adj2v) / 100000.0 // arrowhead half-extra-width
	th2 := th / 2.0
	dh2 := aw2 - th2                     // arrowhead extension beyond shaft edge
	y5 := fh * float64(adj5v) / 100000.0 // total height
	ah := ss * float64(adj3v) / 100000.0 // arrowhead height
	y4 := y5 - ah                        // arrowhead base y

	x9 := fw - dh2
	bw := x9 / 2.0
	bs := math.Min(bw, y4)
	maxAdj4 := bs * 100000.0 / ss
	a4 := math.Min(float64(adj4v), maxAdj4)
	if a4 < 0 {
		a4 = 0
	}
	bd := ss * a4 / 100000.0 // bend diameter (arc radius)
	bd3 := bd - th
	bd2 := math.Max(bd3, 0) // inner bend diameter

	x3 := th + bd2
	x8 := fw - aw2
	x6 := x8 - aw2
	x7 := x6 + dh2
	x4 := x9 - bd

	// Clamp values
	if y4 < 0 {
		y4 = 0
	}
	if y5 > fh {
		y5 = fh
	}

	pts := make([]fpoint, 0, 100)
	steps := 40

	// Path per OOXML spec:
	// 1. moveTo (0, h) - bottom-left
	pts = append(pts, fpoint{fx, fy + fh})

	// 2. lineTo (0, bd) - up left shaft
	pts = append(pts, fpoint{fx, fy + bd})

	// 3. arcTo wR=bd, hR=bd, stAng=180°, swAng=90° (left-to-top arc)
	// Arc center is at (bd, bd). Arc goes from 180° to 270° (CCW in screen coords = upward-right)
	if bd > 0 {
		arcCX := fx + bd
		arcCY := fy + bd
		for i := 0; i <= steps; i++ {
			t := float64(i) / float64(steps)
			angle := math.Pi + t*(math.Pi/2) // 180° to 270°
			px := arcCX + bd*math.Cos(angle)
			py := arcCY + bd*math.Sin(angle)
			pts = append(pts, fpoint{px, py})
		}
	}

	// 4. lineTo (x4, 0) - across the top
	pts = append(pts, fpoint{fx + x4, fy})

	// 5. arcTo wR=bd, hR=bd, stAng=270°, swAng=90° (top-to-right arc)
	// Arc center is at (x4, bd). Arc goes from 270° to 360°
	if bd > 0 {
		arcCX := fx + x4
		arcCY := fy + bd
		for i := 0; i <= steps; i++ {
			t := float64(i) / float64(steps)
			angle := 3*math.Pi/2 + t*(math.Pi/2) // 270° to 360°
			px := arcCX + bd*math.Cos(angle)
			py := arcCY + bd*math.Sin(angle)
			pts = append(pts, fpoint{px, py})
		}
	}

	// 6. lineTo (x9, y4) - down right shaft to arrowhead base
	pts = append(pts, fpoint{fx + x9, fy + y4})

	// 7. lineTo (r, y4) - arrowhead right wing
	pts = append(pts, fpoint{fx + fw, fy + y4})

	// 8. lineTo (x8, y5) - arrowhead tip (pointing down)
	pts = append(pts, fpoint{fx + x8, fy + y5})

	// 9. lineTo (x6, y4) - arrowhead left wing
	pts = append(pts, fpoint{fx + x6, fy + y4})

	// 10. lineTo (x7, y4) - inner right shaft at arrowhead base
	pts = append(pts, fpoint{fx + x7, fy + y4})

	// 11. lineTo (x7, x3) - up inner right shaft to inner bend
	pts = append(pts, fpoint{fx + x7, fy + x3})

	// 12. arcTo wR=bd2, hR=bd2, stAng=0°, swAng=-90° (inner arc, right-to-top)
	// Arc center is at (x7-bd2, x3). Arc goes from 0° to -90° (= 270°)
	if bd2 > 0 {
		arcCX := fx + x7 - bd2
		arcCY := fy + x3
		for i := 0; i <= steps; i++ {
			t := float64(i) / float64(steps)
			angle := 0 - t*(math.Pi/2) // 0° to -90°
			px := arcCX + bd2*math.Cos(angle)
			py := arcCY + bd2*math.Sin(angle)
			pts = append(pts, fpoint{px, py})
		}
	}

	// 13. lineTo (x3, th) - across inner top
	pts = append(pts, fpoint{fx + x3, fy + th})

	// 14. arcTo wR=bd2, hR=bd2, stAng=270°, swAng=-90° (inner arc, top-to-left)
	// Arc center is at (x3, th+bd2). Arc goes from 270° to 180°
	if bd2 > 0 {
		arcCX := fx + x3
		arcCY := fy + th + bd2
		for i := 0; i <= steps; i++ {
			t := float64(i) / float64(steps)
			angle := 3*math.Pi/2 - t*(math.Pi/2) // 270° to 180°
			px := arcCX + bd2*math.Cos(angle)
			py := arcCY + bd2*math.Sin(angle)
			pts = append(pts, fpoint{px, py})
		}
	}

	// 15. lineTo (th, h) - down inner left shaft to bottom
	pts = append(pts, fpoint{fx + th, fy + fh})

	// 16. close (implicit - polygon closes back to start)

	r.fillPolygon(pts, c)
}

// --- Text rendering ---

// fontSizePixels converts a Font's point size into pixels at the current render
// scale, applying any normAutofit scaling. 1pt = 12700 EMU; scaleX converts
// EMU to pixels.
func (r *renderer) fontSizePixels(f *Font) float64 {
	sizePt := float64(f.Size)
	if sizePt <= 0 {
		sizePt = 10
	}
	// Apply normAutofit font scale if set
	if r.fontScale > 0 && r.fontScale != 1.0 {
		sizePt *= r.fontScale
	}
	return sizePt * 12700.0 * r.scaleX
}

// fontFaceFor looks a face up in the font cache. measure selects the
// HintingNone face used for text layout.
func (r *renderer) fontFaceFor(name string, sizePixels float64, bold, italic, measure bool) font.Face {
	if r.fontCache == nil || name == "" {
		return nil
	}
	if measure {
		return r.fontCache.GetMeasureFace(name, sizePixels, bold, italic)
	}
	return r.fontCache.GetFace(name, sizePixels, bold, italic)
}

// resolveFace finds a face for f: the requested font first, then the East Asian
// font, then the configured fallback chain, then the built-in one. It reports
// which name matched and how, so substitutions can be reported to the caller.
func (r *renderer) resolveFace(f *Font, sizePixels float64, measure bool) (font.Face, string, FontFallbackKind) {
	if face := r.fontFaceFor(f.Name, sizePixels, f.Bold, f.Italic, measure); face != nil {
		return face, f.Name, FontResolved
	}
	if f.NameEA != "" {
		if face := r.fontFaceFor(f.NameEA, sizePixels, f.Bold, f.Italic, measure); face != nil {
			return face, f.NameEA, FontSubstituted
		}
	}
	for _, name := range r.fontFallback {
		if face := r.fontFaceFor(name, sizePixels, f.Bold, f.Italic, measure); face != nil {
			return face, name, FontSubstituted
		}
	}
	return nil, "", FontMissing
}

// textClass says which kind of face a character has to be drawn with. A run is
// split wherever the class changes, because one face cannot serve all three.
type textClass uint8

const (
	// textClassLatin covers everything the declared text face owns: Latin,
	// Greek, Cyrillic, digits, spaces, ordinary punctuation.
	textClassLatin textClass = iota
	// textClassCJK covers Han, Kana, Hangul, CJK punctuation and the enclosed
	// forms — the characters a CJK face has to supply.
	textClassCJK
	// textClassSymbol covers pictographs: emoji, dingbats, the miscellaneous
	// symbol blocks. No text face carries glyphs for these.
	textClassSymbol
	numTextClasses = 3
)

// classOrder lists the classes in the order a fallback is considered.
var classOrder = [numTextClasses]textClass{textClassLatin, textClassCJK, textClassSymbol}

// resolveClassFace finds a face that can actually draw the characters of the
// given class in sample, for f. It returns the face and the name of the font
// that matched, or nil and "" when nothing installed covers the text.
//
// "The name resolves" is not a strong enough test, and using it is what makes
// tofu possible. A document whose generator copied one font-family list into
// both <a:latin> and <a:ea> declares a Latin face as its East Asian font; that
// face is installed, so the request looks satisfied, and every Chinese
// character then renders as the font's .notdef box while the font diagnostics
// report a perfect match. Each candidate is therefore checked against the
// characters the run actually contains, so a Latin face is skipped in favour of
// the East Asian fallback chain.
//
// The same argument applies to a symbol face, which is why the class is a
// parameter: an emoji is no more drawable by the declared text face than a
// Chinese character is by a Latin-only one, and it used to be drawn with the
// declared face for exactly the same reason — nothing asked the question.
func (r *renderer) resolveClassFace(f *Font, sizePixels float64, measure bool, sample string, class textClass) (font.Face, string) {
	if r.fontCache == nil {
		return nil, ""
	}

	candidates := r.classCandidates(f, class)

	// A font covering every character of the sample in this class is the
	// answer, and the first such candidate wins.
	bestName, bestCovered := "", 0
	for _, name := range candidates {
		covered, total := r.countCoveredClass(name, f.Bold, f.Italic, sample, class)
		if total == 0 {
			continue
		}
		if covered == total {
			if face := r.fontFaceFor(name, sizePixels, f.Bold, f.Italic, measure); face != nil {
				return face, name
			}
			continue
		}
		if covered > bestCovered {
			bestName, bestCovered = name, covered
		}
	}

	// No font covers the whole sample. Surrendering here would return nil, and
	// splitRunByClass would then draw the entire run — including the characters
	// that ARE drawable — with the Latin face. So take the font that covers the
	// most, and let only the genuinely missing characters show a .notdef box.
	if bestName != "" {
		if face := r.fontFaceFor(bestName, sizePixels, f.Bold, f.Italic, measure); face != nil {
			return face, bestName
		}
	}
	return nil, ""
}

// classCandidates returns the font names to try for a class, in preference
// order: the font the document declared for this kind of text first, then that
// class's fallback chain.
//
// Declaration first is what keeps a run's typeface stable. A document that
// declares a face able to draw its symbols — or its Chinese — is honoured, and
// the coverage check is what rejects a declaration that cannot: a text face that
// merely looks resolved never wins a class it has no glyphs for.
func (r *renderer) classCandidates(f *Font, class textClass) []string {
	declared := make([]string, 0, 2)
	switch class {
	case textClassCJK:
		// East Asian text is declared by <a:ea> alone.
		declared = append(declared, f.NameEA)
	case textClassSymbol:
		// A symbol declared with <a:sym> names the font that carries the run's
		// private-use glyphs, so it outranks everything else. Otherwise either
		// of the run's two font names may be the one that carries the glyphs.
		declared = append(declared, f.NameSym, f.NameEA, f.Name)
	}

	var chain []string
	switch class {
	case textClassCJK:
		chain = r.cjkFallback
	case textClassSymbol:
		chain = r.symbolFallback
	}
	if len(declared) == 0 && len(chain) == 0 {
		return nil
	}

	out := make([]string, 0, len(declared)+len(chain))
	seen := make(map[string]struct{}, len(declared)+len(chain))
	for _, list := range [][]string{declared, chain} {
		for _, name := range list {
			if name == "" {
				continue
			}
			key := lowerFontName(name)
			if _, dup := seen[key]; dup {
				continue
			}
			seen[key] = struct{}{}
			out = append(out, name)
		}
	}
	return out
}

// countCoveredClass reports how many of the sample's characters of the given
// class the named font can draw, and how many such characters the sample
// contains. A total of zero means the sample holds none of that class, so the
// question does not apply.
func (r *renderer) countCoveredClass(name string, bold, italic bool, sample string, class textClass) (covered, total int) {
	if r.fontCache == nil || name == "" {
		return 0, 0
	}
	for _, ch := range sample {
		if classOf(ch) != class {
			continue
		}
		total++
		if r.fontCache.CoversRune(name, bold, italic, ch) {
			covered++
		}
	}
	return covered, total
}

// resolveCJKFace finds a face for the East Asian characters of sample.
func (r *renderer) resolveCJKFace(f *Font, sizePixels float64, measure bool, sample string) (font.Face, string) {
	return r.resolveClassFace(f, sizePixels, measure, sample, textClassCJK)
}

// resolveSymbolFace finds a face for the pictographs of sample — the emoji and
// dingbats a text face has no glyphs for. Without it those characters are drawn
// by whatever face the name resolves to, which is a .notdef box on the slide.
func (r *renderer) resolveSymbolFace(f *Font, sizePixels float64, measure bool, sample string) (font.Face, string) {
	return r.resolveClassFace(f, sizePixels, measure, sample, textClassSymbol)
}

// countCoveredCJK reports how many of the sample's CJK characters the named
// font can draw, and how many such characters the sample contains.
func (r *renderer) countCoveredCJK(name string, bold, italic bool, sample string) (covered, total int) {
	return r.countCoveredClass(name, bold, italic, sample, textClassCJK)
}

// coversCJK reports whether the named font has a glyph for every East Asian
// character in sample.
//
// A sample containing no East Asian characters cannot be answered by this
// question, so it reports false: vouching for a font that was never tested is
// how a Latin face ends up drawing East Asian text. Callers only reach here for
// text that contains CJK.
func (r *renderer) coversCJK(name string, bold, italic bool, sample string) bool {
	covered, total := r.countCoveredCJK(name, bold, italic, sample)
	return total > 0 && covered == total
}

// noteFontFallback reports a font request that was not satisfied exactly. Fully
// resolved requests are skipped because they are not interesting.
func (r *renderer) noteFontFallback(f *Font, used string, kind FontFallbackKind) {
	if kind == FontResolved {
		return
	}
	if f.Name == "" && kind == FontSubstituted {
		// No font was requested, so a default face was picked. Not a
		// substitution worth reporting.
		return
	}
	r.recordFontUsage(f.Name, used, kind, f.Bold, f.Italic)
}

// recordFontUsage files one font resolution outcome with the diagnostics sink
// and the fallback callback. Resolved requests are skipped because a caller
// cannot act on them.
func (r *renderer) recordFontUsage(requested, used string, kind FontFallbackKind, bold, italic bool) {
	if kind == FontResolved || (r.fontDiag == nil && r.onFontFallback == nil) {
		return
	}
	u := FontUsage{
		Requested: requested,
		Used:      used,
		Kind:      kind,
		Bold:      bold,
		Italic:    italic,
	}
	r.fontDiag.record(u)
	if r.onFontFallback != nil {
		r.onFontFallback(u)
	}
}

// noteCJKFont reports the outcome of choosing an East Asian font for f.
//
// A declared East Asian font that had to be replaced is the case worth
// surfacing, and the one that used to pass unmentioned: the name resolves, so
// nothing looks wrong, while the text is drawn as missing-glyph boxes. Silent
// tofu is exactly the condition FontDiagnostics exists to expose.
func (r *renderer) noteCJKFont(f *Font, used string) {
	if r.fontDiag == nil && r.onFontFallback == nil {
		return
	}
	if f.NameEA != "" && f.NameEA != used {
		kind := FontSubstituted
		if used == "" {
			kind = FontMissing
		}
		r.recordFontUsage(f.NameEA, used, kind, f.Bold, f.Italic)
		return
	}
	if used == "" {
		// No East Asian font was declared and none installed can draw the text,
		// so every East Asian character becomes a missing-glyph box.
		r.recordFontUsage(f.Name, "", FontMissing, f.Bold, f.Italic)
	}
}

// noteSymbolFont reports the outcome of choosing a face for the pictographs in
// a run, so an emoji that had to be drawn from a symbol face — or could not be
// drawn at all — is visible in the diagnostics instead of only on the slide.
//
// It is reported against the run's text font, because that is the face the
// document asked for and the one the reader can do something about: the cure is
// to install a symbol font, or to write the character in a font that has it.
func (r *renderer) noteSymbolFont(f *Font, used string) {
	if r.fontDiag == nil && r.onFontFallback == nil {
		return
	}
	if used == "" {
		// Nothing installed can draw the pictographs, so they become
		// missing-glyph boxes.
		r.recordFontUsage(f.Name, "", FontMissing, f.Bold, f.Italic)
		return
	}
	if f.Name != "" && !strings.EqualFold(f.Name, used) && !strings.EqualFold(f.NameEA, used) {
		r.recordFontUsage(f.Name, used, FontSubstituted, f.Bold, f.Italic)
	}
}

// getFace returns a font.Face for the given Font. When no installed font can be
// found it falls back to basicfont.Face7x13, which renders non-ASCII glyphs as
// blank boxes; the miss is reported through FontDiagnostics / OnFontFallback.
func (r *renderer) getFace(f *Font) font.Face {
	if r.fontCache == nil {
		return basicfont.Face7x13
	}
	face, used, kind := r.resolveFace(f, r.fontSizePixels(f), false)
	if face == nil {
		r.noteFontFallback(f, "", FontMissing)
		return basicfont.Face7x13
	}
	r.noteFontFallback(f, used, kind)
	return face
}

// getCJKFace returns a font face able to draw the East Asian text in sample,
// plus the name of the font it came from, or nil and "" when no installed font
// covers that text.
func (r *renderer) getCJKFace(f *Font, sample string) (font.Face, string) {
	if r.fontCache == nil {
		return nil, ""
	}
	return r.resolveCJKFace(f, r.fontSizePixels(f), false, sample)
}

// getMeasureFace returns a font.Face with HintingNone for text measurement.
// PowerPoint uses unhinted glyph metrics for text layout, so using HintingNone
// produces glyph advances that match PowerPoint's wrapping positions. It
// returns nil when no installed font matches, leaving the caller to fall back.
func (r *renderer) getMeasureFace(f *Font) font.Face {
	if r.fontCache == nil {
		return nil
	}
	face, used, kind := r.resolveFace(f, r.fontSizePixels(f), true)
	if face == nil {
		r.noteFontFallback(f, "", FontMissing)
		return nil
	}
	r.noteFontFallback(f, used, kind)
	return face
}

// getCJKMeasureFace returns a HintingNone font.Face for CJK text measurement,
// plus the name of the font it came from. It is coverage-aware in exactly the
// same way as getCJKFace, so measurement and drawing always agree on which font
// a run uses — otherwise line wrapping would be computed against a different
// face's advances than the one the glyphs are drawn with.
func (r *renderer) getCJKMeasureFace(f *Font, sample string) (font.Face, string) {
	if r.fontCache == nil {
		return nil, ""
	}
	return r.resolveCJKFace(f, r.fontSizePixels(f), true, sample)
}

// getSymbolFace returns a font face able to draw the pictographs in sample —
// the emoji and dingbats — plus the name of the font it came from, or nil and
// "" when no installed font covers them.
func (r *renderer) getSymbolFace(f *Font, sample string) (font.Face, string) {
	if r.fontCache == nil {
		return nil, ""
	}
	return r.resolveSymbolFace(f, r.fontSizePixels(f), false, sample)
}

// getSymbolMeasureFace is getSymbolFace with the HintingNone face used for
// layout, so measurement and drawing agree on which font a run uses.
func (r *renderer) getSymbolMeasureFace(f *Font, sample string) (font.Face, string) {
	if r.fontCache == nil {
		return nil, ""
	}
	return r.resolveSymbolFace(f, r.fontSizePixels(f), true, sample)
}

// containsCJK returns true if the string contains any CJK characters.
func containsCJK(s string) bool {
	for _, r := range s {
		if isCJK(r) {
			return true
		}
	}
	return false
}

// containsSymbol returns true if the string contains any pictographs, i.e. any
// character that has to be drawn from a symbol face.
func containsSymbol(s string) bool {
	for _, r := range s {
		if isSymbolRune(r) {
			return true
		}
	}
	return false
}

// buildParaTextRuns builds textRun slices for a paragraph's elements,
// using HintingNone measure faces for width calculation and HintingFull
// render faces for drawing. This is the single place where render/measure
// face pairs are created, ensuring consistent layout across measurement
// and drawing code paths.
func (r *renderer) buildParaTextRuns(elements []ParagraphElement) []textRun {
	var runs []textRun
	for _, elem := range elements {
		switch e := elem.(type) {
		case *TextRun:
			// A field run is evaluated here, at draw time: "slidenum" becomes
			// the number of the slide on the canvas, whatever the file's cached
			// <a:t> claims. Substituting before the empty check keeps a field
			// whose cache is empty (and whose run is therefore the paragraph's
			// only one) from vanishing.
			text := e.text
			if e.fieldType == "slidenum" {
				text = strconv.Itoa(r.slideNumber)
			}
			if text == "" {
				continue
			}
			f := e.font
			if f == nil {
				f = NewFont()
			}
			// A baseline-shifted run draws at two thirds of its declared
			// size: PowerPoint shrinks the glyphs (the COM export's
			// subscript measures ~0.68x the surrounding run's advance
			// width) while the file keeps the declared size. Faces resolve
			// from the scaled copy; the run itself keeps the original font,
			// so the baseline shift below uses the real size.
			faceFont := f
			if f.Superscript || f.Subscript {
				scaled := *f
				scaled.Size = int(float64(f.Size)*(2.0/3.0) + 0.5)
				faceFont = &scaled
			}
			if (containsCJK(text) || containsSymbol(text)) && r.fontCache != nil {
				sizePt := float64(faceFont.Size)
				if sizePt <= 0 {
					sizePt = 10
				}
				if r.fontScale > 0 && r.fontScale != 1.0 {
					sizePt *= r.fontScale
				}
				scaledPt := sizePt * 12700.0 * r.scaleX
				var faces, measures [numTextClasses]font.Face
				var wins [numTextClasses]textWin
				faces[textClassLatin] = r.fontCache.GetFace(faceFont.Name, scaledPt, faceFont.Bold, faceFont.Italic)
				if faces[textClassLatin] == nil {
					faces[textClassLatin] = r.getFace(faceFont)
				}
				measures[textClassLatin] = r.getMeasureFace(faceFont)
				if a, d, ok := r.fontCache.WinVerticalMetrics(faceFont.Name, scaledPt, faceFont.Bold, faceFont.Italic); ok {
					wins[textClassLatin] = textWin{asc: int(a + 0.5), desc: int(d + 0.5), ascF: a, descF: d}
				}
				// Face selection is driven by the characters in this run, not by
				// the declared font names: a name that resolves can still lack
				// the glyphs, and then the text draws as empty boxes.
				if containsCJK(text) {
					var used string
					faces[textClassCJK], used = r.getCJKFace(faceFont, text)
					measures[textClassCJK], _ = r.getCJKMeasureFace(faceFont, text)
					r.noteCJKFont(f, used)
					if a, d, ok := r.fontCache.WinVerticalMetrics(used, scaledPt, faceFont.Bold, faceFont.Italic); ok {
						wins[textClassCJK] = textWin{asc: int(a + 0.5), desc: int(d + 0.5), ascF: a, descF: d}
					}
				}
				if containsSymbol(text) {
					var used string
					faces[textClassSymbol], used = r.getSymbolFace(faceFont, text)
					measures[textClassSymbol], _ = r.getSymbolMeasureFace(faceFont, text)
					r.noteSymbolFont(f, used)
					if a, d, ok := r.fontCache.WinVerticalMetrics(used, scaledPt, faceFont.Bold, faceFont.Italic); ok {
						wins[textClassSymbol] = textWin{asc: int(a + 0.5), desc: int(d + 0.5), ascF: a, descF: d}
					}
				}
				runs = append(runs, r.splitRunByClass(text, f, faces, measures, wins)...)
			} else {
				face := r.getFace(faceFont)
				mf := r.getMeasureFace(faceFont)
				tr := textRun{
					text:        text,
					font:        f,
					face:        face,
					measureFace: mf,
					width:       measureStringWithKern(face, text).Ceil(),
				}
				if r.fontCache != nil {
					if a, d, ok := r.fontCache.WinVerticalMetrics(faceFont.Name, r.fontSizePixels(faceFont), faceFont.Bold, faceFont.Italic); ok {
						tr.winAsc = int(a + 0.5)
						tr.winDesc = int(d + 0.5)
						tr.winAscF = a
						tr.winDescF = d
					}
				}
				runs = append(runs, tr)
			}
		case *BreakElement:
			runs = append(runs, textRun{text: "\n"})
		}
	}
	return runs
}

// splitRunByClass splits a text run wherever the character class changes, so
// each segment is drawn with a face that has glyphs for it. This is what keeps
// CJK text off a Latin face, and pictographs off both.
//
// faces and measures are indexed by textClass and are supplied by the caller,
// which resolves one face per class for the whole run — so the sample every
// face was chosen against is the same one the segments are cut from.
func (r *renderer) splitRunByClass(text string, f *Font, faces, measures [numTextClasses]font.Face, wins [numTextClasses]textWin) []textRun {
	// A class with no face of its own — no installed font covers its
	// characters — is drawn with the Latin face, which the caller always
	// supplies. Normalising here is also what keeps the split conservative:
	// below, a class whose face IS the Latin face is not a boundary, so text
	// that used to be drawn in one piece still is.
	if faces[textClassLatin] == nil {
		for _, c := range classOrder {
			if faces[c] != nil {
				faces[textClassLatin], measures[textClassLatin] = faces[c], measures[c]
				break
			}
		}
	}
	if faces[textClassLatin] == nil {
		faces[textClassLatin], measures[textClassLatin] = basicfont.Face7x13, basicfont.Face7x13
	}
	for _, c := range classOrder {
		if faces[c] == nil {
			faces[c], measures[c] = faces[textClassLatin], measures[textClassLatin]
		}
	}

	runs := make([]textRun, 0, 3)
	var buf strings.Builder
	cur := textClassLatin
	first := true

	flush := func(class textClass) {
		if buf.Len() == 0 {
			return
		}
		seg := buf.String()
		face := faces[class]
		runs = append(runs, textRun{
			text:        seg,
			font:        f,
			face:        face,
			measureFace: measures[class],
			width:       measureStringWithKern(face, seg).Ceil(),
			winAsc:      wins[class].asc,
			winDesc:     wins[class].desc,
			winAscF:     wins[class].ascF,
			winDescF:    wins[class].descF,
		})
	}

	for _, ch := range text {
		next := classOf(ch)
		if faces[next] == faces[textClassLatin] {
			// This class has no face but the text face, so drawing the
			// character where it is drawn now is the whole of the change.
			next = textClassLatin
		}
		if isEmojiJoiner(ch) {
			// A joiner has no glyph of its own: it belongs to whichever
			// segment it is modifying, or a sequence would be cut in half.
			next = cur
		}
		if !first && next != cur {
			flush(cur)
			buf.Reset()
		}
		buf.WriteRune(ch)
		cur = next
		first = false
	}
	flush(cur)
	return runs
}

// textRun holds a measured run of text with its formatting.
type textRun struct {
	text        string
	font        *Font
	face        font.Face // render face (HintingFull) for drawing
	measureFace font.Face // measure face (HintingNone) for layout; nil falls back to face
	width       int
	// isBullet marks the prefix run buildBulletRun creates. A bullet takes
	// its metrics from whatever face it draws with — often Arial, where the
	// master's buFont points — and that face's ascent can exceed the text
	// face's. PowerPoint sizes a line by its text, not its bullet, so line
	// metrics ignore bullet runs and the glyph rides the text baseline.
	isBullet bool
	// winAsc/winDesc are the GDI vertical metrics (OS/2 usWinAscent and
	// usWinDescent) of the face this run draws with, scaled to the face's
	// pixel size. PowerPoint places the baseline winAsc below the line top
	// and advances lines by winAsc+winDesc, but Go's font.Face.Metrics
	// reports the hhea values instead — for Calibri a fifth of an em apart
	// (hhea ascent 0.75em vs win ascent 0.952em), which put every baseline
	// 13px too high on 28pt text. Zero means "unknown": the run then falls
	// back to face.Metrics in buildTextLine.
	winAsc  int
	winDesc int
	// The same pair before rounding. normAutofit's two levers both act on
	// the ascent, and PowerPoint applies them to the unrounded metric —
	// rounding between the two is a one-pixel-per-line loss (round 28).
	winAscF  float64
	winDescF float64
}

// textWin carries the win-metric pair for one text class through
// splitRunByClass, which stamps them onto each segment run it flushes.
type textWin struct {
	asc  int
	desc int
	// Unrounded pair, carried so the autofit levers can be applied to it
	// (round 28).
	ascF  float64
	descF float64
}

// mface returns the face to use for measurement. If a dedicated measure face
// is set it is preferred; otherwise the render face is used.
func (tr *textRun) mface() font.Face {
	if tr.measureFace != nil {
		return tr.measureFace
	}
	return tr.face
}

// textLine holds a line of text runs with total metrics.
type textLine struct {
	runs       []textRun
	width      int
	ascent     int
	descent    int
	lineHeight int
	// Unrounded win ascent/descent for the line: the levers must be applied
	// to these, not to the rounded ascent (round 28). Left at zero for CJK
	// lines, whose baseline placement the golden image still arbitrates.
	ascentF  float64
	descendF float64
	hasCJK   bool
}

// buildTextLine measures a slice of textRuns and returns a textLine.
func (r *renderer) buildTextLine(runs []textRun) textLine {
	var tl textLine
	tl.runs = runs
	maxHeight := 0 // track font's recommended line-to-line height (includes line gap)
	hasCJK := false
	hasWin := false
	// A bullet run must not contribute to the line's metrics when real text
	// shares the line: the bullet's face (the master's buFont, often Arial)
	// can be taller than the text face and would push every glyph down. When
	// the line holds nothing else, the bullet is all there is and keeps its
	// own metrics.
	textRuns := runs
	hasTextRun := false
	for _, run := range runs {
		if !run.isBullet {
			hasTextRun = true
			break
		}
	}
	if hasTextRun {
		textRuns = make([]textRun, 0, len(runs))
		for _, run := range runs {
			if !run.isBullet {
				textRuns = append(textRuns, run)
			} else {
				tl.width += run.width
			}
		}
	}
	for _, run := range textRuns {
		tl.width += run.width
		if !hasCJK && containsCJK(run.text) {
			hasCJK = true
		}
		if run.winAsc > 0 || run.winDesc > 0 {
			// GDI metrics recorded at face-creation time — PowerPoint's
			// baseline placement. These win over the hhea-derived
			// face.Metrics, which sit a line-gap (and for fonts like
			// Calibri a fifth of an em) away from them. The float form is
			// kept beside the integer one: the baseline is the metric
			// after *both* autofit levers, and rounding it before the
			// second lever is applied costs a pixel per line that only
			// autofit pages pay (slide35: 62.63 → 63 → 56.7 → 57 where
			// PowerPoint draws 56.37 → 56).
			hasWin = true
			af, df := run.winAscF, run.winDescF
			if af <= 0 {
				af = float64(run.winAsc)
			}
			if df <= 0 {
				df = float64(run.winDesc)
			}
			if af > tl.ascentF {
				tl.ascentF = af
				tl.ascent = run.winAsc
			}
			if df > tl.descendF {
				tl.descendF = df
				tl.descent = run.winDesc
			}
			if h := run.winAsc + run.winDesc; h > maxHeight {
				maxHeight = h
			}
			continue
		}
		if run.face == nil {
			continue
		}
		// Use measure face (HintingNone) for line metrics when available.
		// HintingFull rounds ascent/descent to pixel boundaries, inflating
		// line heights. HintingNone gives ideal metrics matching PowerPoint.
		metricFace := run.mface()
		metrics := metricFace.Metrics()
		asc := metrics.Ascent.Ceil()
		desc := metrics.Descent.Ceil()
		if asc > tl.ascent {
			tl.ascent = asc
		}
		if desc > tl.descent {
			tl.descent = desc
		}
		if h := metrics.Height.Ceil(); h > maxHeight {
			maxHeight = h
		}
	}
	// Use the font's recommended height (ascent + descent + line gap) so that
	// default single spacing matches PowerPoint's behaviour. When the font
	// reports no line gap, fall back to ascent + descent.
	tl.lineHeight = maxHeight
	if tl.lineHeight < tl.ascent+tl.descent {
		tl.lineHeight = tl.ascent + tl.descent
	}
	// CJK fonts report a larger metrics.Height (line gap) than what
	// PowerPoint's DirectWrite renderer uses. Even with HintingNone, the
	// OS/2 table values are slightly larger. Cap line height to
	// ascent+descent (no extra line gap) for CJK lines.
	if hasCJK {
		adSum := tl.ascent + tl.descent
		if tl.lineHeight > adSum {
			tl.lineHeight = adSum
		}
	}
	// But the line *advance* is not the font's metrics: PowerPoint spaces
	// lines at 1.2 × the largest font size on the line (its "single"
	// spacing) whatever face serves the glyphs. COM single-factor variants
	// of slide14 of the comparison deck pinned this: swapping the master
	// body font between Calibri (win-metric height 1.221em), Arial
	// (1.117em) and Segoe UI (1.332em) moved the rendered line pitch by
	// exactly nothing, while 36pt/40pt master sizes moved it to 1.2×size
	// rounded per line (96/107 px). The win metrics still place the
	// baseline — ascent below the line top (round 22) — they just do not
	// size the line box.
	//
	// Those variants were Latin, so the rule is applied to Latin lines
	// only. CJK keeps the metrics-derived height it was tuned against:
	// the cjk_wrap golden measures 3.2% of pixels against the 1.2×
	// advance and passes against ascent+descent, and nothing in this
	// round's evidence speaks to which of the two PowerPoint really uses
	// there. Reserving the question beats silently moving a golden.
	if hasWin && !hasCJK {
		adv := 0
		for _, run := range textRuns {
			if run.font == nil {
				continue
			}
			if h := int(1.2*r.fontSizePixels(run.font) + 0.5); h > adv {
				adv = h
			}
		}
		if adv > 0 {
			tl.lineHeight = adv
		}
	}
	if tl.lineHeight < 1 {
		tl.lineHeight = 14
	}
	tl.hasCJK = hasCJK
	return tl
}

// applyLnSpcReduction shrinks a line advance by normAutofit's
// lnSpcReduction percentage. PowerPoint applies it to every line of the text
// body that carries the attribute — it is the second lever (besides
// fontScale) that presses overflowing text back into its shape. The
// experiment that pinned the semantics showed the whole line box scales:
// the ascent rides the same factor, which lifts the first glyphs toward the
// line top rather than only pulling consecutive lines together.
func (r *renderer) applyLnSpcReduction(lh int) int {
	if r.lnSpcReduction <= 0 {
		return lh
	}
	return int(float64(lh)*(1.0-r.lnSpcReduction) + 0.5)
}

// baselineOffset is how far below the line top the baseline sits.
//
// normAutofit's fontScale and lnSpcReduction both act on the win ascent, and
// PowerPoint applies them to the unrounded metric. Rounding between the two
// levers — which is what reading the already-rounded ascent back does — costs
// a pixel of baseline per line, and only autofit pages pay it: slide35's 32pt
// body measures 62.63px after fontScale, rounds to 63, then to 57 after the
// reduction, where the export shows 56.37 → 56 (round 28).
//
// CJK lines keep the rounded path: the win metrics that place their baseline
// are the ones the golden image was tuned against, and this round's evidence
// (a Latin deck) says nothing about them.
func (r *renderer) baselineOffset(l textLine) int {
	if l.ascentF > 0 && !l.hasCJK {
		f := l.ascentF
		if r.lnSpcReduction > 0 {
			f *= 1.0 - r.lnSpcReduction
		}
		return int(f + 0.5)
	}
	return r.applyLnSpcReduction(l.ascent)
}

// paraSpaceBefore resolves a paragraph's space-before to hundredths of a
// point. A declared spcPts is already stored; a declared spcPct is resolved
// against the paragraph's own font size here, at layout time, because the
// run sizes are not known when the pPr closes. The percentage is of the font
// size itself, not of the 1.2× line: the COM experiment on slide35 measured
// the inherited 20% as 20% × 32pt and not 20% × 1.2 × 32pt.
//
// The first paragraph of a text block never gets any of it. The COM variants
// of slide1 proved this in both directions: deleting every declared spcBef
// and enlarging them fivefold both rendered bit-identical to the original —
// PowerPoint ignores a first paragraph's space-before whether it was declared
// on the paragraph or inherited from the master.
//
// The declared and the inherited percentage resolve differently, each pinned
// by its own variant experiment: a declared spcPct is a percentage of the
// font size itself (slide35's 20% measured 14px on 32pt runs), while the
// master-inherited one rides the whole line — 1.2 × the size (slide34's
// tripled 20% moved every gap by 2 × 17.5px on the same 32pt runs).
func (r *renderer) paraSpaceBefore(para *Paragraph, firstPara bool) int {
	if firstPara {
		return 0
	}
	if para.spaceBefore > 0 {
		return para.spaceBefore
	}
	if para.spaceBeforePct > 0 {
		return int(r.paragraphFontSize(para)*float64(para.spaceBeforePct)/100000.0*100 + 0.5)
	}
	if para.inheritedSpaceBeforePct > 0 {
		// The inherited percentage resolves against 1.2 × the run's own
		// size (round 26), but it does NOT ride the normAutofit levers
		// (fontScale, lnSpcReduction) the way the line advance does.
		// Round 27 tried both and let the deck decide: with the levers
		// multiplied in, slide35 — the only autofit page whose body
		// inherits a 20% space-before — scored 6.45% against the
		// PowerPoint export; without them 5.29%, and the pages that
		// carry no autofit were untouched. The two-factor fit round 26
		// rested on (20% × 1.2 × 32 × 0.925 × 0.9 = 14.2px) was a
		// coincidence: the unscaled 20% × 1.2 × 28 = 14.9px sits in the
		// same pixel, and only one of the two survives a second page.
		v := r.paragraphFontSize(para) * 1.2 * float64(para.inheritedSpaceBeforePct) / 100000.0 * 100
		return int(v + 0.5)
	}
	return para.spaceBefore
}

// paragraphFontSize is the size a paragraph's spacing percentages resolve
// against: the first sized run; for a runless paragraph the endParaRPr size
// (slide34's empty paragraph resolves its 20% against the 18pt endParaRPr —
// doubling that sz doubled the gap component); a fallback otherwise.
func (r *renderer) paragraphFontSize(para *Paragraph) float64 {
	for _, elem := range para.elements {
		if tr, ok := elem.(*TextRun); ok && tr.font != nil && tr.font.Size > 0 {
			return float64(tr.font.Size)
		}
	}
	if para.endParaRPrSize > 0 {
		return float64(para.endParaRPrSize) / 100.0
	}
	return 18.0
}

// emptyParagraphLineHeight sizes a runless paragraph's line box. The old
// hard-coded 14px is why slide34's inter-group gap measured 66px against the
// export's 99px: PowerPoint draws the empty line at the endParaRPr size —
// its 18pt endParaRPr gives a 48px line, and vC (sz 1800→3600) grew the gap
// by exactly that line's height. The height comes from the same win metrics
// the real lines use, through whatever face the default resolution picks.
func (r *renderer) emptyParagraphLineHeight(para *Paragraph) int {
	if para.endParaRPrSize <= 0 {
		return 14
	}
	sizePt := float64(para.endParaRPrSize) / 100.0
	f := NewFont()
	f.Size = int(sizePt + 0.5)
	// Line advance is 1.2 × the size regardless of the font's vertical
	// metrics (see buildTextLine); the empty line is no exception — the
	// COM variant that doubled an empty paragraph's endParaRPr size grew
	// the following gap by exactly the 1.2× line plus its spacing.
	return int(1.2*r.fontSizePixels(f) + 0.5)
}

// measureParagraphsHeight estimates the total pixel height needed to render
// the given paragraphs within the specified width, replicating the same line
// building and spacing logic used by drawParagraphs.
func (r *renderer) measureParagraphsHeight(paragraphs []*Paragraph, w, h int, anchor TextAnchorType, wordWrap bool) int {
	if len(paragraphs) == 0 {
		return 0
	}
	type lineInfo struct {
		lineHeight  int
		spaceBefore int
		spaceAfter  int
		lineSpacing int
	}
	var allLines []lineInfo
	ordinals := bulletOrdinals(paragraphs)

	for pi, para := range paragraphs {
		marginLeft := 0
		marginRight := 0
		indent := 0
		if para.alignment != nil {
			marginLeft = r.emuToPixelX(para.alignment.MarginLeft)
			marginRight = r.emuToPixelX(para.alignment.MarginRight)
			indent = r.emuToPixelX(para.alignment.Indent)
		}
		var paraRuns []textRun
		if bRun := r.bulletRunFor(para, ordinals[pi], indent); bRun.text != "" {
			paraRuns = append(paraRuns, bRun)
		}
		paraRuns = append(paraRuns, r.buildParaTextRuns(para.elements)...)
		baseW := w - marginLeft - marginRight
		firstLineW := baseW - indent
		if firstLineW < 10 {
			firstLineW = w
		}
		if baseW < 10 {
			baseW = w
		}
		if !wordWrap {
			firstLineW = 999999
			baseW = 999999
		}
		lines := r.wrapRunLine(paraRuns, baseW)
		if indent != 0 && len(lines) > 0 && wordWrap {
			lines = r.wrapRunLineWithIndent(paraRuns, firstLineW, baseW)
		}
		if len(lines) == 0 {
			lines = []textLine{{lineHeight: r.emptyParagraphLineHeight(para)}}
		}
		for i, line := range lines {
			li := lineInfo{
				lineHeight:  line.lineHeight,
				lineSpacing: para.lineSpacing,
			}
			if i == 0 {
				li.spaceBefore = r.hundredthPtToPixelY(r.paraSpaceBefore(para, pi == 0))
			}
			if i == len(lines)-1 {
				li.spaceAfter = r.hundredthPtToPixelY(para.spaceAfter)
			}
			allLines = append(allLines, li)
		}
	}

	totalH := 0
	for _, li := range allLines {
		// spaceBefore applies to every paragraph except the block's first —
		// same as the draw path.
		totalH += li.spaceBefore
		lh := li.lineHeight
		if li.lineSpacing < 0 {
			lh = int(float64(lh) * float64(-li.lineSpacing) / 100000.0)
		} else if li.lineSpacing > 0 {
			lh = r.hundredthPtToPixelY(li.lineSpacing)
		}
		totalH += r.applyLnSpcReduction(lh)
		totalH += li.spaceAfter
	}
	return totalH
}

// measureMaxLineWidth returns the maximum line width across all paragraphs
// after word-wrapping. This is used to detect horizontal text overflow.
func (r *renderer) measureMaxLineWidth(paragraphs []*Paragraph, w int, wordWrap bool) int {
	if len(paragraphs) == 0 {
		return 0
	}
	maxW := 0
	ordinals := bulletOrdinals(paragraphs)
	for pi, para := range paragraphs {
		marginLeft := 0
		marginRight := 0
		indent := 0
		if para.alignment != nil {
			marginLeft = r.emuToPixelX(para.alignment.MarginLeft)
			marginRight = r.emuToPixelX(para.alignment.MarginRight)
			indent = r.emuToPixelX(para.alignment.Indent)
		}
		var paraRuns []textRun
		if bRun := r.bulletRunFor(para, ordinals[pi], indent); bRun.text != "" {
			paraRuns = append(paraRuns, bRun)
		}
		paraRuns = append(paraRuns, r.buildParaTextRuns(para.elements)...)
		baseW := w - marginLeft - marginRight
		firstLineW := baseW - indent
		if firstLineW < 10 {
			firstLineW = w
		}
		if baseW < 10 {
			baseW = w
		}
		if !wordWrap {
			firstLineW = 999999
			baseW = 999999
		}
		lines := r.wrapRunLine(paraRuns, baseW)
		if indent != 0 && len(lines) > 0 && wordWrap {
			lines = r.wrapRunLineWithIndent(paraRuns, firstLineW, baseW)
		}
		for _, line := range lines {
			if line.width > maxW {
				maxW = line.width
			}
		}
	}
	return maxW
}

// drawParagraphs renders paragraphs within the given bounding box.
func (r *renderer) drawParagraphs(paragraphs []*Paragraph, x, y, w, h int, anchor TextAnchorType, wordWrap bool) {
	if len(paragraphs) == 0 {
		return
	}

	// Build all lines from all paragraphs, tracking per-paragraph spacing
	type lineInfo struct {
		line        textLine
		spaceBefore int
		spaceAfter  int
		lineSpacing int // 0 means default (single)
		hAlign      HorizontalAlignment
		paraIdx     int  // index into paragraphs slice
		isFirst     bool // first line of paragraph
		isLast      bool // last line of paragraph
	}
	var allLines []lineInfo
	ordinals := bulletOrdinals(paragraphs)

	for pi, para := range paragraphs {
		align := HorizontalLeft
		marginLeft := 0
		marginRight := 0
		indent := 0
		if para.alignment != nil {
			align = para.alignment.Horizontal
			marginLeft = r.emuToPixelX(para.alignment.MarginLeft)
			marginRight = r.emuToPixelX(para.alignment.MarginRight)
			indent = r.emuToPixelX(para.alignment.Indent)
		}

		// Build runs for this paragraph
		var paraRuns []textRun

		// Bullet run
		if bRun := r.bulletRunFor(para, ordinals[pi], indent); bRun.text != "" {
			paraRuns = append(paraRuns, bRun)
		}

		paraRuns = append(paraRuns, r.buildParaTextRuns(para.elements)...)

		// Wrap runs into lines.
		// In PowerPoint, indent only affects the first line of a paragraph.
		// Continuation lines use the full width minus margins only.
		baseW := w - marginLeft - marginRight
		firstLineW := baseW - indent
		if firstLineW < 10 {
			firstLineW = w
		}
		if baseW < 10 {
			baseW = w
		}
		if !wordWrap {
			firstLineW = 999999
			baseW = 999999
		}
		// First, wrap using the continuation-line width (wider), then check
		// if the first line exceeds the first-line width and re-wrap if needed.
		lines := r.wrapRunLine(paraRuns, baseW)
		if indent != 0 && len(lines) > 0 && wordWrap {
			// Re-wrap with first-line width to handle indent correctly
			lines = r.wrapRunLineWithIndent(paraRuns, firstLineW, baseW)
		}
		if len(lines) == 0 {
			// Empty paragraph still takes space
			lines = []textLine{{lineHeight: r.emptyParagraphLineHeight(para)}}
		}

		for i, line := range lines {
			li := lineInfo{
				line:        line,
				lineSpacing: para.lineSpacing,
				hAlign:      align,
				paraIdx:     pi,
				isFirst:     i == 0,
				isLast:      i == len(lines)-1,
			}
			if i == 0 {
				// spaceBefore is in hundredths of a point from spcPts
				li.spaceBefore = r.hundredthPtToPixelY(r.paraSpaceBefore(para, pi == 0))
			}
			if i == len(lines)-1 {
				li.spaceAfter = r.hundredthPtToPixelY(para.spaceAfter)
			}
			allLines = append(allLines, li)
		}
	}

	// Calculate total height
	totalH := 0
	for _, li := range allLines {
		// spaceBefore only ever rides the first line of a paragraph, and
		// PowerPoint applies it before the first paragraph too — no skip.
		totalH += li.spaceBefore
		lh := li.line.lineHeight
		if li.lineSpacing < 0 {
			// spcPct: negative value, percentage * 1000 (e.g. -150000 = 150%)
			lh = int(float64(lh) * float64(-li.lineSpacing) / 100000.0)
		} else if li.lineSpacing > 0 {
			// spcPts: hundredths of a point (e.g. 1200 = 12pt)
			lh = r.hundredthPtToPixelY(li.lineSpacing)
		}
		totalH += r.applyLnSpcReduction(lh)
		totalH += li.spaceAfter
	}

	// Vertical anchor offset
	startY := y
	switch anchor {
	case TextAnchorMiddle:
		startY = y + (h-totalH)/2
	case TextAnchorBottom:
		startY = y + h - totalH
		// Do NOT clamp startY for bottom anchor. PowerPoint allows
		// bottom-anchored text to overflow upward above the shape
		// boundary, which is the expected behaviour for shapes like
		// timeline annotation boxes.
	}

	curY := startY
	for _, li := range allLines {
		// Same rule as the height measure above: every paragraph carries its
		// spaceBefore except the block's first, which PowerPoint ignores.
		curY += li.spaceBefore

		lh := li.line.lineHeight
		if li.lineSpacing < 0 {
			lh = int(float64(lh) * float64(-li.lineSpacing) / 100000.0)
		} else if li.lineSpacing > 0 {
			lh = r.hundredthPtToPixelY(li.lineSpacing)
		}
		lh = r.applyLnSpcReduction(lh)

		// Horizontal alignment
		lineX := x
		para := paragraphs[li.paraIdx]
		if para.alignment != nil {
			lineX += r.emuToPixelX(para.alignment.MarginLeft)
			if li.isFirst {
				lineX += r.emuToPixelX(para.alignment.Indent)
			}
		}

		switch li.hAlign {
		case HorizontalCenter:
			lineX = x + (w-li.line.width)/2
		case HorizontalRight:
			lineX = x + w - li.line.width
			if para.alignment != nil {
				lineX -= r.emuToPixelX(para.alignment.MarginRight)
			}
		}

		// The baseline sits ascent-scaled below the line top: lnSpcReduction
		// shrinks the ascent along with the advance (see applyLnSpcReduction),
		// which is what lifts the first line's glyphs toward the inset.
		baseline := curY + r.baselineOffset(li.line)

		// Draw each run
		drawX := lineX
		for _, run := range li.line.runs {
			if run.text == "\n" || run.text == "" {
				continue
			}
			if run.face == nil {
				continue
			}
			fc := color.RGBA{A: 255}
			if run.font != nil {
				fc = argbToRGBA(run.font.Color)
			}

			runBaseline := baseline
			if run.font != nil {
				// The shift is what the file says: baseline is a shift in
				// thousandths of a percent of the run's own size, and
				// PowerPoint writes +30000 for a superscript and -25000 for
				// a subscript.
				if run.font.Superscript {
					runBaseline -= int(r.fontSizePixels(run.font)*0.30 + 0.5)
				} else if run.font.Subscript {
					runBaseline += int(r.fontSizePixels(run.font)*0.25 + 0.5)
				}
			}

			// Tab characters: PowerPoint advances to the next default tab
			// stop — one inch from the paragraph's text origin — instead of
			// drawing the control character, which the face has no glyph for
			// and which otherwise comes out as a .notdef box (slide35's
			// sub-bullets open with a tab, as do the "Memory:\t224 MB" info
			// tables). Runs with tabs take a dedicated path; shadows on
			// tabbed runs are not supported here.
			if strings.Contains(run.text, "\t") {
				startX := drawX
				tabPx := r.emuToPixelX(914400)
				if tabPx < 1 {
					tabPx = 1
				}
				segs := strings.Split(run.text, "\t")
				for si, seg := range segs {
					if seg != "" {
						sd := &font.Drawer{
							Dst:  r.img,
							Src:  image.NewUniform(drawSrc(fc)),
							Face: run.face,
							Dot:  fixed.P(drawX, runBaseline),
						}
						sd.DrawString(seg)
						if run.font != nil && run.font.Bold {
							sd.Dot = fixed.P(drawX+1, runBaseline)
							sd.DrawString(seg)
						}
						if run.font != nil && run.font.Underline != UnderlineNone {
							uy := runBaseline + 2
							w := measureStringWithKern(run.face, seg).Ceil()
							r.drawUnderline(drawX, drawX+w, uy, fc, run.font.Underline)
						}
						drawX += measureStringWithKern(run.face, seg).Ceil()
					}
					if si < len(segs)-1 {
						// The tab jumps to the next one-inch stop, measured
						// from the text area's left edge — not from lineX,
						// which already carries the paragraph's margin and
						// indent (a level-1 sub-bullet's first tab must land
						// on its marL, one stop out, not two).
						off := drawX - x
						drawX = x + (off/tabPx+1)*tabPx
					}
				}
				if run.font != nil && run.font.Strikethrough {
					sy := runBaseline - li.line.ascent/3
					r.drawLine(startX, sy, drawX, sy, fc)
				}
				continue
			}

			// Run-level text shadow: the glyph pass is repeated offset along
			// the shadow direction and blended with the shadow colour before
			// the real text goes down on top of it. Same geometry as the
			// shape shadow: distance in points scaled to pixels, direction in
			// degrees with y growing downward.
			if run.font != nil && run.font.Shadow != nil && run.font.Shadow.Visible && !r.draft {
				sh := run.font.Shadow
				rad := float64(sh.Direction) * math.Pi / 180.0
				// Distance is in points; scaleX is pixels per EMU, so the
				// points go to EMU first (12700 per point).
				dist := float64(sh.Distance) * 12700 * r.scaleX
				dx := int(dist * math.Cos(rad))
				dy := int(dist * math.Sin(rad))
				blurPx := int(float64(sh.BlurRadius)*12700*r.scaleX + 0.5)
				r.drawTextShadow(run.text, run.face, drawX+dx, runBaseline+dy, sh, blurPx)
			}

			d := &font.Drawer{
				Dst:  r.img,
				Src:  image.NewUniform(drawSrc(fc)),
				Face: run.face,
				Dot:  fixed.P(drawX, runBaseline),
			}
			d.DrawString(run.text)

			// Synthetic bold: if bold was requested but the font face is the
			// regular weight (no bold variant found), re-draw with a 1px
			// horizontal offset to embolden the glyphs.
			if run.font != nil && run.font.Bold {
				d2 := &font.Drawer{
					Dst:  r.img,
					Src:  image.NewUniform(drawSrc(fc)),
					Face: run.face,
					Dot:  fixed.P(drawX+1, runBaseline),
				}
				d2.DrawString(run.text)
			}

			// Underline
			if run.font != nil && run.font.Underline != UnderlineNone {
				uy := runBaseline + 2
				r.drawUnderline(drawX, drawX+run.width, uy, fc, run.font.Underline)
			}

			// Strikethrough
			if run.font != nil && run.font.Strikethrough {
				sy := runBaseline - li.line.ascent/3
				r.drawLine(drawX, sy, drawX+run.width, sy, fc)
			}

			drawX += run.width
		}

		curY += lh
		curY += li.spaceAfter
	}
}

// drawTextShadow draws the shadow copy of a text run: the glyphs rasterised at
// the already-offset baseline (x, y), coloured and blended at the shadow's
// alpha. With a blur radius the glyphs go into an offscreen alpha mask that is
// box-blurred (three passes, which tracks PowerPoint's gaussian closely enough
// at deck render sizes) and composited through the mask, so the shadow is soft
// the way the golden export shows it instead of a hard second glyph.
func (r *renderer) drawTextShadow(text string, face font.Face, x, y int, sh *Shadow, blurPx int) {
	if text == "" || face == nil {
		return
	}
	scol := argbToRGBA(sh.Color)
	baseA := float64(sh.Alpha) / 100
	if blurPx < 1 {
		c := scol
		c.A = uint8(baseA*255 + 0.5)
		d := &font.Drawer{
			Dst:  r.img,
			Src:  image.NewUniform(drawSrc(c)),
			Face: face,
			Dot:  fixed.P(x, y),
		}
		d.DrawString(text)
		return
	}

	// Offscreen coverage mask, padded wide enough for the blur to spread.
	metrics := face.Metrics()
	asc := metrics.Ascent.Ceil()
	desc := metrics.Descent.Ceil()
	adv := measureStringWithKern(face, text).Ceil()
	pad := blurPx*3 + 2
	bx := x - pad
	by := y - asc - pad
	bw := adv + 2*pad
	bh := asc + desc + 2*pad
	if bw <= 0 || bh <= 0 || bx >= r.img.Bounds().Max.X || by >= r.img.Bounds().Max.Y {
		return
	}
	mask := image.NewAlpha(image.Rect(0, 0, bw, bh))
	d := &font.Drawer{
		Dst:  mask,
		Src:  image.NewUniform(color.Alpha{A: 255}),
		Face: face,
		Dot:  fixed.P(pad, pad+asc),
	}
	d.DrawString(text)
	boxBlurAlpha(mask, blurPx, 3)

	bounds := r.img.Bounds()
	for py := 0; py < bh; py++ {
		imgY := by + py
		if imgY < bounds.Min.Y || imgY >= bounds.Max.Y {
			continue
		}
		for px := 0; px < bw; px++ {
			a := mask.AlphaAt(px, py).A
			if a == 0 {
				continue
			}
			imgX := bx + px
			if imgX < bounds.Min.X || imgX >= bounds.Max.X {
				continue
			}
			c := scol
			c.A = uint8(float64(a)*baseA + 0.5)
			if c.A > 0 {
				r.blendPixel(imgX, imgY, c)
			}
		}
	}
}

// boxBlurAlpha box-blurs an alpha mask in place. Three passes of a box blur
// approximate a gaussian; radius is in pixels.
func boxBlurAlpha(m *image.Alpha, radius, passes int) {
	if radius < 1 || passes < 1 {
		return
	}
	b := m.Bounds()
	w, h := b.Dx(), b.Dy()
	if w == 0 || h == 0 {
		return
	}
	src := make([]uint8, w*h)
	dst := make([]uint8, w*h)
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			src[y*w+x] = m.AlphaAt(b.Min.X+x, b.Min.Y+y).A
		}
	}
	div := 2*radius + 1
	clamp := func(v, lo, hi int) int {
		if v < lo {
			return lo
		}
		if v > hi {
			return hi
		}
		return v
	}
	for p := 0; p < passes; p++ {
		// Horizontal, with a running sum over the window.
		for y := 0; y < h; y++ {
			row := y * w
			sum := 0
			for x := -radius; x <= radius; x++ {
				sum += int(src[row+clamp(x, 0, w-1)])
			}
			for x := 0; x < w; x++ {
				dst[row+x] = uint8(sum / div)
				sum += int(src[row+clamp(x+radius+1, 0, w-1)]) - int(src[row+clamp(x-radius, 0, w-1)])
			}
		}
		// Vertical, same running sum over the horizontal result.
		for x := 0; x < w; x++ {
			sum := 0
			for y := -radius; y <= radius; y++ {
				sum += int(dst[clamp(y, 0, h-1)*w+x])
			}
			for y := 0; y < h; y++ {
				src[y*w+x] = uint8(sum / div)
				sum += int(dst[clamp(y+radius+1, 0, h-1)*w+x]) - int(dst[clamp(y-radius, 0, h-1)*w+x])
			}
		}
	}
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			m.SetAlpha(b.Min.X+x, b.Min.Y+y, color.Alpha{A: src[y*w+x]})
		}
	}
}

// drawUnderline draws an underline of the given style.
func (r *renderer) drawUnderline(x1, x2, y int, c color.RGBA, style UnderlineType) {
	switch style {
	case UnderlineSingle:
		r.drawLine(x1, y, x2, y, c)
	case UnderlineDouble:
		r.drawLine(x1, y-1, x2, y-1, c)
		r.drawLine(x1, y+1, x2, y+1, c)
	case UnderlineHeavy:
		r.drawLine(x1, y-1, x2, y-1, c)
		r.drawLine(x1, y, x2, y, c)
		r.drawLine(x1, y+1, x2, y+1, c)
	case UnderlineDash:
		r.drawDashedHLine(x1, x2, y, c, 6, 3)
	case UnderlineWavy:
		for px := x1; px < x2; px++ {
			wy := y + int(math.Sin(float64(px-x1)*0.5)*2)
			r.blendPixel(px, wy, c)
		}
	default:
		r.drawLine(x1, y, x2, y, c)
	}
}

// bulletOrdinals returns, for each paragraph, the number its auto-numbered
// bullet displays; non-numbered paragraphs get 0.
//
// PowerPoint keeps a single element for a numbered bullet (<a:buAutoNum>) and
// numbers a list implicitly: startAt carries the number of the list's *first*
// item and every following paragraph continues the count. The model stores one
// Bullet per paragraph, so the running count has to be reconstructed here.
// Without it buildBulletRun read the same StartAt for every paragraph and a
// numbered list rendered as "1. 1. 1.".
//
// The count is kept per outline level: a numbered list at level 0 is not
// broken by the indented bullet-less or character-bulleted paragraphs under
// its items (slide35's "1. … 2. …" interleaved with unnumbered sub-points),
// only by another numbered paragraph at the same level with a different
// format. Two other breaks remain: a paragraph with no bullet or a character
// bullet ends its own level's list, and — see below — numbered paragraphs
// restart their own level only.
//
// A sequence is broken by a paragraph with no bullet, a character bullet, or a
// different number format — a new list then starts from its own startAt. Both
// the measuring passes and the drawing pass call this, so they cannot disagree
// about how wide the bullet is.
func bulletOrdinals(paragraphs []*Paragraph) []int {
	ordinals := make([]int, len(paragraphs))
	type seq struct {
		format  string
		running int
	}
	state := map[int]*seq{}
	for i, para := range paragraphs {
		level := 0
		if para.alignment != nil && para.alignment.Level > 0 {
			level = para.alignment.Level
		}
		b := para.bullet
		if b == nil || (b.Type != BulletTypeNumeric && b.Type != BulletTypeAutoNum) {
			// The break ends only this level's list; the numbered list one
			// level out continues over it.
			delete(state, level)
			continue
		}
		num := b.NumFormat
		if num == "" {
			num = defaultNumFormat
		}
		s := state[level]
		if s == nil || num != s.format {
			running := b.StartAt
			if running < defaultBulletStart {
				running = defaultBulletStart
			}
			s = &seq{format: num, running: running}
			state[level] = s
		} else {
			s.running++
		}
		ordinals[i] = s.running
	}
	return ordinals
}

// bulletRunFor builds the bullet prefix run for a paragraph and pads it so the
// text lands where PowerPoint puts it. PowerPoint draws the bullet at
// marL+indent but flows the text from marL — the bullet is followed by
// something tab-like, not by its own advance. With a hanging indent
// (indent < 0) that landing is -indent away from the bullet's pen position, so
// a bullet narrower than the hang is padded to it; a wider one keeps its
// advance (PowerPoint then tabs to the next stop, which we approximate).
func (r *renderer) bulletRunFor(para *Paragraph, ordinal, indent int) textRun {
	if para.bullet == nil || para.bullet.Type == BulletTypeNone {
		return textRun{}
	}
	bRun := r.buildBulletRun(para.bullet, para, ordinal)
	if bRun.text == "" {
		return textRun{}
	}
	if indent < 0 && bRun.width < -indent {
		bRun.width = -indent
	}
	return bRun
}

// buildBulletRun creates a textRun for a bullet prefix. ordinal is the number
// an auto-numbered bullet shows, as returned by bulletOrdinals.
func (r *renderer) buildBulletRun(b *Bullet, para *Paragraph, ordinal int) textRun {
	if b == nil || b.Type == BulletTypeNone {
		return textRun{}
	}

	// Determine bullet font
	bulletFont := NewFont()
	bulletFont.Size = 10
	// Try to get size from first text run
	for _, elem := range para.elements {
		if tr, ok := elem.(*TextRun); ok && tr.font != nil {
			bulletFont.Size = tr.font.Size
			bulletFont.Color = tr.font.Color
			break
		}
	}
	// <a:buSzPct> sizes the bullet as a percentage of the text, and the reader
	// and the writer both carry Bullet.Size — only the renderer never looked at
	// it, so a bullet set to 200% drew at 100% in the preview while the file
	// said otherwise.
	if b.Size > 0 && b.Size != 100 {
		bulletFont.Size = bulletFont.Size * b.Size / 100
		if bulletFont.Size < 1 {
			bulletFont.Size = 1
		}
	}
	if b.Color != nil {
		bulletFont.Color = *b.Color
	}
	if b.Font != "" {
		bulletFont.Name = b.Font
	}

	var text string
	switch b.Type {
	case BulletTypeChar:
		// The writer substitutes <a:buChar char="•"/> for an empty character;
		// the renderer drew a bare space, so an empty bullet showed nothing
		// where the file had a bullet.
		char := b.Style
		if char == "" {
			char = defaultBulletChar
		}
		text = char + " "
	case BulletTypeNumeric, BulletTypeAutoNum:
		num := ordinal
		if num < defaultBulletStart {
			num = b.StartAt
		}
		if num < defaultBulletStart {
			num = defaultBulletStart
		}
		format := b.NumFormat
		if format == "" {
			format = defaultNumFormat
		}
		text = formatBulletNumber(num, format) + " "
	}

	// Handle symbol font characters (Wingdings, Symbol, etc.).
	// These fonts use a special encoding where characters map to the
	// Unicode Private Use Area (U+F000 + byte value) in TrueType.
	// First try rendering with the actual symbol font via PUA mapping;
	// if the font is not available, fall back to Unicode equivalents.
	if b.Type == BulletTypeChar && isSymbolFont(bulletFont.Name) {
		// Try PUA mapping with the actual symbol font first
		puaText := symbolToPUA(b.Style)
		puaFont := *bulletFont // copy
		face := r.getFace(&puaFont)
		if face != nil && r.fontCache != nil && r.fontCache.GetFace(bulletFont.Name, 12, false, false) != nil {
			// Use only the symbol glyph without trailing space — the space
			// character in symbol fonts often renders as .notdef (black box).
			// A gap is added via width padding below instead.
			text = puaText
		} else {
			// Font not available — fall back to Unicode equivalent
			mapped := mapSymbolChar(bulletFont.Name, b.Style)
			text = mapped + " "
			// Use the paragraph's text font instead of the symbol font
			bulletFont.Name = ""
			for _, elem := range para.elements {
				if tr, ok := elem.(*TextRun); ok && tr.font != nil {
					bulletFont.Name = tr.font.Name
					bulletFont.NameEA = tr.font.NameEA
					break
				}
			}
			if bulletFont.Name == "" {
				bulletFont.Name = "Calibri"
			}
		}
	}

	face := r.getFace(bulletFont)
	w := font.MeasureString(face, text).Ceil()
	// For symbol fonts rendered via PUA (no trailing space in text),
	// add a small gap so the bullet doesn't touch the text.
	if b.Type == BulletTypeChar && isSymbolFont(bulletFont.Name) {
		gap := int(bulletFont.Size / 3)
		if gap < 2 {
			gap = 2
		}
		w += gap
	}
	br := textRun{
		text:     text,
		font:     bulletFont,
		face:     face,
		width:    w,
		isBullet: true,
	}
	if r.fontCache != nil {
		if a, d, ok := r.fontCache.WinVerticalMetrics(bulletFont.Name, r.fontSizePixels(bulletFont), bulletFont.Bold, bulletFont.Italic); ok {
			br.winAsc = int(a + 0.5)
			br.winDesc = int(d + 0.5)
			br.winAscF = a
			br.winDescF = d
		}
	}
	return br
}

// isSymbolFont returns true if the font name is a symbol/dingbats font
// whose characters need mapping to Unicode equivalents.
func isSymbolFont(name string) bool {
	n := strings.ToLower(name)
	return n == "wingdings" || n == "wingdings 2" || n == "wingdings 3" ||
		n == "symbol" || n == "webdings"
}

// symbolToPUA maps a symbol font character to the Unicode Private Use Area.
// Symbol fonts like Wingdings store glyphs at U+F000 + original byte value
// in their TrueType cmap table.
func symbolToPUA(ch string) string {
	if len(ch) == 0 {
		return ch
	}
	r := []rune(ch)[0]
	if r < 0x100 {
		return string(rune(0xF000 + r))
	}
	return ch
}

// mapSymbolChar maps a character from a symbol font to a Unicode equivalent.
// Symbol fonts like Wingdings encode characters at code points that don't
// correspond to their visual appearance in Unicode.
func mapSymbolChar(fontName, ch string) string {
	if len(ch) == 0 {
		return "•"
	}
	r := []rune(ch)[0]
	n := strings.ToLower(fontName)

	if n == "wingdings" {
		// Wingdings character map (code point → Unicode equivalent)
		switch r {
		case 0xD8: // bowtie (two triangles forming a butterfly/wing shape)
			return "\u22C8"
		case 0xA8: // filled circle
			return "●"
		case 0x6C: // bullet
			return "●"
		case 0x6E: // filled square
			return "■"
		case 0x71: // open circle
			return "○"
		case 0x75, 0xA7: // diamond
			return "◆"
		case 0x76: // open diamond
			return "◇"
		case 0x77: // filled triangle right
			return "▶"
		case 0xFC: // check mark
			return "✓"
		case 0xFB: // cross mark
			return "✗"
		case 0xE0: // right arrow
			return "→"
		case 0xDF: // left arrow
			return "←"
		case 0xE1: // up arrow
			return "↑"
		case 0xE2: // down arrow
			return "↓"
		case 0xF0: // right pointing triangle
			return "►"
		case 0x9F: // star
			return "★"
		case 0xAB: // dash
			return "–"
		default:
			return "•" // fallback to standard bullet
		}
	}

	if n == "wingdings 2" {
		return "•"
	}

	if n == "wingdings 3" {
		switch r {
		case 0x75: // triangle right
			return "▶"
		case 0x76: // triangle left
			return "◀"
		default:
			return "•"
		}
	}

	if n == "symbol" {
		switch r {
		case 0xB7: // middle dot
			return "·"
		case 0xD8: // empty set
			return "∅"
		default:
			return string(r) // Symbol font mostly maps to Unicode directly
		}
	}

	return "•" // fallback
}

// formatBulletNumber formats a number according to the bullet format.
func formatBulletNumber(num int, format string) string {
	switch format {
	case NumFormatRomanUcPeriod:
		return toRoman(num) + "."
	case NumFormatRomanLcPeriod:
		return strings.ToLower(toRoman(num)) + "."
	case NumFormatAlphaUcPeriod:
		if num >= 1 && num <= 26 {
			return string(rune('A'+num-1)) + "."
		}
		return fmt.Sprintf("%d.", num)
	case NumFormatAlphaLcPeriod:
		if num >= 1 && num <= 26 {
			return string(rune('a'+num-1)) + "."
		}
		return fmt.Sprintf("%d.", num)
	case NumFormatAlphaLcParen:
		if num >= 1 && num <= 26 {
			return string(rune('a'+num-1)) + ")"
		}
		return fmt.Sprintf("%d)", num)
	case NumFormatArabicParen:
		return fmt.Sprintf("%d)", num)
	default: // arabicPeriod
		return fmt.Sprintf("%d.", num)
	}
}

// toRoman converts an integer to a Roman numeral string.
func toRoman(num int) string {
	if num <= 0 || num > 3999 {
		return fmt.Sprintf("%d", num)
	}
	vals := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	syms := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
	var buf strings.Builder
	for i, v := range vals {
		for num >= v {
			buf.WriteString(syms[i])
			num -= v
		}
	}
	return buf.String()
}

// isCJK reports whether the rune is a CJK character that can be broken
// at any position (no spaces between characters).
//
// The predicate answers a question about the *character*, not about the fonts
// installed on this machine: "is this a CJK-typographic character, which should
// be set full-width by, and take its glyph from, a CJK-capable face". The
// enclosed symbol blocks below are included because they are used inside CJK
// text and are drawn from CJK faces — a deck that writes ①②③ is writing CJK
// typography even though the code points live in a symbol block.
//
// Classifying them here also hands them to the coverage-based font selection in
// coversCJK. That is the point: every face-selection decision downstream is
// gated on isCJK, so a character outside this set never gets a coverage check
// at all. That is how ①②③ reached a Latin face and rendered as .notdef boxes
// while the font diagnostics reported a perfect match.
//
// Pictographs are in the same position and are classified by isSymbolRune, not
// here: they are drawn from a symbol face, not from a CJK one, and merging the
// two would send an emoji to a face that has no glyph for it either.
func isCJK(r rune) bool {
	return unicode.Is(unicode.Han, r) ||
		unicode.Is(unicode.Hangul, r) ||
		unicode.Is(unicode.Hiragana, r) ||
		unicode.Is(unicode.Katakana, r) ||
		(r >= 0x3000 && r <= 0x303F) || // CJK Symbols and Punctuation
		(r >= 0x3200 && r <= 0x32FF) || // Enclosed CJK Letters and Months (㈠ ㉑)
		(r >= 0x3300 && r <= 0x33FF) || // CJK Compatibility (㎡ ㍿)
		(r >= 0x2460 && r <= 0x24FF) || // Enclosed Alphanumerics (① ② ③)
		(r >= 0xFF00 && r <= 0xFFEF) // Fullwidth Forms
}

// isSymbolRune reports whether r is a pictograph — an emoji, a dingbat, one of
// the miscellaneous symbols — which has to be drawn from a symbol face.
//
// These are the characters a deck uses as bullet markers and icons, and no text
// face carries them: the face a document declares for body text is a text face,
// and the CJK chain the fallback supplies is made of text faces too. So an emoji
// classified as ordinary Latin text reaches the declared face, finds no glyph,
// and is drawn as that face's .notdef box — the hollow rectangle that reads as a
// white square on the slide. Classifying it here is what routes it to the
// symbol chain, where the coverage check can find a face that has it.
//
// The enclosed forms (① ② ③, ㈠, ㎡) are deliberately absent: those are CJK
// typography drawn from CJK faces, and isCJK owns them.
func isSymbolRune(r rune) bool {
	switch {
	case r >= 0x1F000 && r <= 0x1FAFF:
		return true // Mahjong … Symbols and Pictographs Extended-A (🐱 💙 🏆 🧶 🐾)
	case r >= 0x2600 && r <= 0x27BF:
		return true // Miscellaneous Symbols (⚖ ☀) and Dingbats (✂ ➔)
	case r >= 0x2B00 && r <= 0x2BFF:
		return true // Miscellaneous Symbols and Arrows (⬛ ⬅ ⭐)
	case r >= 0x2300 && r <= 0x23FF:
		return true // Miscellaneous Technical (⌚ ⏰ ⏳), drawn as emoji
	case r >= 0xF000 && r <= 0xF0FF:
		// Symbol-font private use: the TrueType cmap of Symbol, Wingdings and
		// Webdings maps U+F000+byte, and a run declares that font with
		// <a:sym>. PowerPoint writes these code points verbatim into <a:t>.
		return true
	}
	return false
}

// isEmojiJoiner reports whether r is a zero-width character that carries no
// glyph of its own and only has meaning inside a sequence: the zero-width
// joiner, the variation selectors, and the enclosing keycap.
//
// The splitter keeps such a character in the segment it is modifying. Cut out
// on its own it would be drawn by whichever face its class resolved to, and a
// sequence like 👩👩👧 would be broken where the joiner sits.
func isEmojiJoiner(r rune) bool {
	switch r {
	case 0x200D, // zero-width joiner
		0xFE0E, 0xFE0F, // variation selectors 15/16
		0x20E3: // combining enclosing keycap
		return true
	}
	return false
}

// classOf classifies a character by the face that has to draw it.
func classOf(r rune) textClass {
	switch {
	case isCJK(r):
		return textClassCJK
	case isSymbolRune(r):
		return textClassSymbol
	default:
		return textClassLatin
	}
}

// isCJKClosingPunct returns true for CJK closing punctuation that must not
// start a new line (禁则処理 — line-start prohibited characters).
func isCJKClosingPunct(r rune) bool {
	switch r {
	case '）', '】', '》', '」', '』', '〉', '〕', '｝', '］',
		'。', '，', '、', '；', '：', '！', '？', '…',
		')', ']', '}', '>', '.', ',', ';', ':', '!', '?':
		return true
	}
	return false
}

// isCJKOpeningPunct returns true for CJK opening punctuation that must not
// end a line (line-end prohibited characters).
func isCJKOpeningPunct(r rune) bool {
	switch r {
	case '（', '【', '《', '「', '『', '〈', '〔', '｛', '［',
		'(', '[', '{', '<',
		'\u201C', '\u2018': // " and '
		return true
	}
	return false
}

// isClosingPunctRun returns true if the run text consists entirely of
// closing punctuation characters that should not start a new line.
func isClosingPunctRun(text string) bool {
	if text == "" {
		return false
	}
	for _, r := range text {
		if !isCJKClosingPunct(r) {
			return false
		}
	}
	return true
}

// splitCJKAware splits text into wrappable segments.
// CJK characters become individual segments; Latin words stay grouped.
// Spaces are preserved as separate segments to avoid inflating word widths.
func splitCJKAware(text string) []string {
	if text == "" {
		return nil
	}
	// Fast path: pure ASCII text (no CJK possible)
	ascii := true
	for i := 0; i < len(text); i++ {
		if text[i] >= 0x80 {
			ascii = false
			break
		}
	}
	if ascii {
		return splitASCIIWords(text)
	}
	// Slow path: handle CJK characters
	runes := []rune(text)
	segments := make([]string, 0, len(runes)/2+1)
	start := 0
	for i, r := range runes {
		if isCJK(r) {
			if i > start {
				segments = append(segments, string(runes[start:i]))
			}
			segments = append(segments, string(r))
			start = i + 1
		} else if r == ' ' || r == '\t' {
			if i > start {
				segments = append(segments, string(runes[start:i]))
			}
			segments = append(segments, string(r))
			start = i + 1
		}
	}
	if start < len(runes) {
		segments = append(segments, string(runes[start:]))
	}
	// Apply kinsoku (禁則処理): merge closing punctuation into the preceding
	// segment so it cannot start a new line.
	// Also merge opening punctuation into the following segment so it cannot
	// end a line.
	if len(segments) > 1 {
		merged := make([]string, 0, len(segments))
		for i, seg := range segments {
			rs := []rune(seg)
			if i > 0 && len(rs) == 1 && isCJKClosingPunct(rs[0]) && len(merged) > 0 {
				merged[len(merged)-1] += seg
			} else {
				merged = append(merged, seg)
			}
		}
		// Second pass: merge opening punctuation with the following segment
		if len(merged) > 1 {
			merged2 := make([]string, 0, len(merged))
			for i := 0; i < len(merged); i++ {
				rs := []rune(merged[i])
				if len(rs) == 1 && isCJKOpeningPunct(rs[0]) && i+1 < len(merged) {
					// Merge opening punct with next segment
					merged[i+1] = merged[i] + merged[i+1]
				} else {
					merged2 = append(merged2, merged[i])
				}
			}
			segments = merged2
		} else {
			segments = merged
		}
	}
	return segments
}

// splitASCIIWords splits ASCII text into words and spaces as separate segments.
func splitASCIIWords(text string) []string {
	segments := make([]string, 0, 8)
	start := 0
	for i := 0; i < len(text); i++ {
		if text[i] == ' ' || text[i] == '\t' {
			if i > start {
				segments = append(segments, text[start:i])
			}
			segments = append(segments, text[i:i+1])
			start = i + 1
		}
	}
	if start < len(text) {
		segments = append(segments, text[start:])
	}
	return segments
}

// measureStringWithKern measures the advance width of a string using the face's
// GlyphAdvance and Kern methods. Unlike font.MeasureString, this accounts for
// kerning pairs, producing measurements closer to what PowerPoint's DirectWrite
// renderer computes.
func measureStringWithKern(face font.Face, s string) fixed.Int26_6 {
	var advance fixed.Int26_6
	prevR := rune(-1)
	for _, r := range s {
		if prevR >= 0 {
			advance += face.Kern(prevR, r)
		}
		a, ok := face.GlyphAdvance(r)
		if ok {
			advance += a
		}
		prevR = r
	}
	return advance
}

// wrapRunLine wraps text runs into multiple lines that fit within maxWidth.
func (r *renderer) wrapRunLine(runs []textRun, maxWidth int) []textLine {
	if len(runs) == 0 {
		return nil
	}
	if maxWidth <= 0 {
		maxWidth = 1
	}

	// Reserve 1px safety margin. HintingFull glyph bitmaps can be slightly
	// wider than the advance width reported by GlyphAdvance, causing the
	// last character on a line to overshoot the text box edge by a fraction
	// of a pixel. Subtracting 1px from the wrap limit prevents this.
	maxW26_6 := fixed.I(maxWidth - 1)
	if maxW26_6 < fixed.I(1) {
		maxW26_6 = fixed.I(1)
	}

	var lines []textLine
	var currentRuns []textRun
	var currentWidth fixed.Int26_6 // fixed-point accumulation avoids Ceil rounding buildup

	for _, run := range runs {
		if run.text == "\n" {
			lines = append(lines, r.buildTextLine(currentRuns))
			currentRuns = nil
			currentWidth = 0
			continue
		}
		if run.face == nil {
			continue
		}

		mf := run.mface()
		// Use the larger of measure-face and render-face widths for wrapping.
		// Measure face (HintingNone) matches PowerPoint's layout but can be
		// narrower than the render face (HintingFull). Using the max prevents
		// fitting more characters than the render face can actually display.
		runMW := measureStringWithKern(mf, run.text)
		runRW := measureStringWithKern(run.face, run.text)
		runW := runMW
		if runRW > runW {
			runW = runRW
		}

		// If the run fits, add it whole
		if currentWidth+runW <= maxW26_6 {
			currentRuns = append(currentRuns, run)
			currentWidth += runW
			continue
		}

		// Closing punctuation (e.g. ）】》) must not start a new line
		// (kinsoku / 禁則処理). Keep it on the current line even if it
		// slightly overflows.
		if isClosingPunctRun(run.text) {
			currentRuns = append(currentRuns, run)
			currentWidth += runW
			continue
		}

		// Run doesn't fit — try to split into wrappable segments (CJK-aware)
		segments := splitCJKAware(run.text)

		if len(segments) <= 1 {
			// Single segment doesn't fit, force it on new line
			if len(currentRuns) > 0 {
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
			}
			currentRuns = append(currentRuns, run)
			currentWidth = runW
			continue
		}

		// Split by segments
		var partial strings.Builder
		for _, seg := range segments {
			test := partial.String() + seg
			twM := measureStringWithKern(mf, test)
			twR := measureStringWithKern(run.face, test)
			tw := twM
			if twR > tw {
				tw = twR
			}
			if currentWidth+tw > maxW26_6 && (len(currentRuns) > 0 || partial.Len() > 0) {
				if partial.Len() > 0 {
					pText := partial.String()
					currentRuns = append(currentRuns, textRun{
						text:        pText,
						font:        run.font,
						face:        run.face,
						measureFace: run.measureFace,
						width:       measureStringWithKern(run.face, pText).Ceil(),
						isBullet:    run.isBullet,
						winAsc:      run.winAsc,
						winDesc:     run.winDesc,
					})
				}
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
				partial.Reset()
				// The break consumes the whitespace it happened at: PowerPoint
				// never renders, at the start of a continuation line, the
				// space the previous line broke on — and real decks carry
				// double spaces ("Project  declared") that made the wrapped
				// line start visibly indented.
				if strings.TrimSpace(seg) != "" {
					partial.WriteString(seg)
				}
			} else if partial.Len() > 0 || len(currentRuns) > 0 || strings.TrimSpace(seg) != "" {
				// A fresh line never opens with whitespace: the line break
				// consumes what it broke on, and the second of the deck's
				// double spaces would otherwise surface as an indented
				// continuation line.
				partial.WriteString(seg)
			}
		}
		if partial.Len() > 0 {
			pText := partial.String()
			pwM := measureStringWithKern(mf, pText)
			pwR := measureStringWithKern(run.face, pText)
			pw := pwM
			if pwR > pw {
				pw = pwR
			}
			wr := textRun{
				text:        pText,
				font:        run.font,
				face:        run.face,
				measureFace: run.measureFace,
				width:       measureStringWithKern(run.face, pText).Ceil(),
				isBullet:    run.isBullet,
				winAsc:      run.winAsc,
				winDesc:     run.winDesc,
			}
			currentRuns = append(currentRuns, wr)
			currentWidth += pw
		}
	}

	if len(currentRuns) > 0 {
		lines = append(lines, r.buildTextLine(currentRuns))
	}

	return lines
}

// wrapRunLineWithIndent wraps text runs using different widths for the first
// line (which includes the paragraph indent) and continuation lines.
func (r *renderer) wrapRunLineWithIndent(runs []textRun, firstLineWidth, contLineWidth int) []textLine {
	if len(runs) == 0 {
		return nil
	}
	if firstLineWidth <= 0 {
		firstLineWidth = 1
	}
	if contLineWidth <= 0 {
		contLineWidth = 1
	}

	lineIdx := 0
	getMaxW := func() fixed.Int26_6 {
		w := contLineWidth
		if lineIdx == 0 {
			w = firstLineWidth
		}
		// 1px safety margin, same as wrapRunLine.
		if w > 1 {
			w--
		}
		mw := fixed.I(w)
		return mw
	}

	var lines []textLine
	var currentRuns []textRun
	var currentWidth fixed.Int26_6

	for _, run := range runs {
		if run.text == "\n" {
			lines = append(lines, r.buildTextLine(currentRuns))
			currentRuns = nil
			currentWidth = 0
			lineIdx++
			continue
		}
		if run.face == nil {
			continue
		}

		mf := run.mface()
		maxW := getMaxW()
		// Use the larger of measure-face and render-face widths for wrapping,
		// same logic as wrapRunLine.
		runMW := measureStringWithKern(mf, run.text)
		runRW := measureStringWithKern(run.face, run.text)
		runW := runMW
		if runRW > runW {
			runW = runRW
		}

		if currentWidth+runW <= maxW {
			currentRuns = append(currentRuns, run)
			currentWidth += runW
			continue
		}

		if isClosingPunctRun(run.text) {
			currentRuns = append(currentRuns, run)
			currentWidth += runW
			continue
		}

		segments := splitCJKAware(run.text)
		if len(segments) <= 1 {
			if len(currentRuns) > 0 {
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
				lineIdx++
			}
			currentRuns = append(currentRuns, run)
			currentWidth = runW
			continue
		}

		var partial strings.Builder
		for _, seg := range segments {
			test := partial.String() + seg
			twM := measureStringWithKern(mf, test)
			twR := measureStringWithKern(run.face, test)
			tw := twM
			if twR > tw {
				tw = twR
			}
			maxW = getMaxW()
			if currentWidth+tw > maxW && (len(currentRuns) > 0 || partial.Len() > 0) {
				if partial.Len() > 0 {
					pText := partial.String()
					currentRuns = append(currentRuns, textRun{
						text:        pText,
						font:        run.font,
						face:        run.face,
						measureFace: run.measureFace,
						width:       measureStringWithKern(run.face, pText).Ceil(),
						isBullet:    run.isBullet,
						winAsc:      run.winAsc,
						winDesc:     run.winDesc,
					})
				}
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
				lineIdx++
				partial.Reset()
				// A break consumes the whitespace it happened at — see
				// wrapRunLine.
				if strings.TrimSpace(seg) != "" {
					partial.WriteString(seg)
				}
			} else if partial.Len() > 0 || len(currentRuns) > 0 || strings.TrimSpace(seg) != "" {
				// A fresh line never opens with whitespace: the line break
				// consumes what it broke on, and the second of the deck's
				// double spaces would otherwise surface as an indented
				// continuation line.
				partial.WriteString(seg)
			}
		}
		if partial.Len() > 0 {
			pText := partial.String()
			pwM := measureStringWithKern(mf, pText)
			pwR := measureStringWithKern(run.face, pText)
			pw := pwM
			if pwR > pw {
				pw = pwR
			}
			currentRuns = append(currentRuns, textRun{
				text:        pText,
				font:        run.font,
				face:        run.face,
				measureFace: run.measureFace,
				width:       measureStringWithKern(run.face, pText).Ceil(),
				winAsc:      run.winAsc,
				winDesc:     run.winDesc,
			})
			currentWidth += pw
		}
	}

	if len(currentRuns) > 0 {
		lines = append(lines, r.buildTextLine(currentRuns))
	}

	return lines
}

// drawStringCentered draws a string centered in the given rectangle.
func (r *renderer) drawStringCentered(text string, face font.Face, c color.RGBA, rect image.Rectangle) {
	if text == "" || face == nil {
		return
	}
	tw := font.MeasureString(face, text).Ceil()
	metrics := face.Metrics()
	th := (metrics.Ascent + metrics.Descent).Ceil()
	cx := rect.Min.X + (rect.Dx()-tw)/2
	cy := rect.Min.Y + (rect.Dy()-th)/2 + metrics.Ascent.Ceil()
	d := &font.Drawer{
		Dst:  r.img,
		Src:  image.NewUniform(drawSrc(c)),
		Face: face,
		Dot:  fixed.P(cx, cy),
	}
	d.DrawString(text)
}

// --- Chart rendering ---

// defaultChartPalette is the default color palette for chart series.
var defaultChartPalette = []color.RGBA{
	{R: 79, G: 129, B: 189, A: 255},
	{R: 192, G: 80, B: 77, A: 255},
	{R: 155, G: 187, B: 89, A: 255},
	{R: 128, G: 100, B: 162, A: 255},
	{R: 75, G: 172, B: 198, A: 255},
	{R: 247, G: 150, B: 70, A: 255},
	{R: 119, G: 44, B: 42, A: 255},
	{R: 77, G: 93, B: 58, A: 255},
}

// chartColors returns the default color palette for chart series.
func chartColors() []color.RGBA {
	return defaultChartPalette
}

// getSeriesColor returns the color for a series, using its FillColor if set, otherwise a palette color.
func getSeriesColor(s *ChartSeries, idx int, palette []color.RGBA) color.RGBA {
	if s.FillColor.ARGB != "" && s.FillColor.ARGB != "00000000" {
		return argbToRGBA(s.FillColor)
	}
	return palette[idx%len(palette)]
}

// seriesPointColor resolves a bar's colour: a <c:dPt> override for that
// category index wins, then the series fill, then the palette.
func seriesPointColor(ser *ChartSeries, si, ci int, palette []color.RGBA) color.RGBA {
	if c, ok := ser.PointColors[ci]; ok && c.ARGB != "" && c.ARGB != "00000000" {
		return argbToRGBA(c)
	}
	return getSeriesColor(ser, si, palette)
}

// --- Image scaling ---

// applyLumAdjust applies a picture's <a:lum bright contrast> to already-scaled
// pixels, in place.
//
// Both numbers are in 1/1000 of a percent and both were measured off
// PowerPoint's own output (r29 COM probe: flat-colour pictures, one adjustment
// each, read back pixel by pixel):
//
//   - bright is a plain additive offset on every sRGB channel, bright/100000:
//     bright 25000 takes black to #3F3F3F and #4F81BD to #8EC0FC, bright
//     -50000 takes white to #7F7F7F — no gamma, no clamp before the add.
//   - contrast is an affine about mid-grey, #7F7F7F, whose scale is 1+k below
//     zero but 1/(1-k) above it: contrast 50000 doubles the distance from
//     mid-grey (#4F81BD -> #1E82FA), contrast -50000 halves it (#4F81BD ->
//     #67809E), and -100000 collapses the whole image onto #7F7F7F. The two
//     branches are what the export shows; a single 1+k or 1/(1-k) on both
//     sides misses one half badly.
//
// Channels clip at 0 and 255, which is how a 50% contrast boost leaves black
// and white untouched.
func applyLumAdjust(img *image.RGBA, bright, contrast int) {
	offset := float64(bright) / 100000.0 * 255.0
	scale := 1.0
	if contrast > 0 {
		scale = 1.0 / (1.0 - float64(contrast)/100000.0)
	} else if contrast < 0 {
		scale = 1.0 + float64(contrast)/100000.0
	}
	var lut [256]uint8
	for i := 0; i < 256; i++ {
		v := (float64(i)-127.5)*scale + 127.5 + offset
		if v < 0 {
			v = 0
		}
		if v > 255 {
			v = 255
		}
		lut[i] = uint8(v)
	}
	b := img.Bounds()
	for y := b.Min.Y; y < b.Max.Y; y++ {
		row := img.Pix[img.PixOffset(b.Min.X, y):img.PixOffset(b.Max.X, y)]
		for i := 0; i+3 < len(row); i += 4 {
			row[i] = lut[row[i]]
			row[i+1] = lut[row[i+1]]
			row[i+2] = lut[row[i+2]]
		}
	}
}

// scaleImageBilinear scales an image to the target width and height using bilinear interpolation.
func scaleImageBilinear(src image.Image, dstW, dstH int) *image.RGBA {
	if dstW <= 0 || dstH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, 1, 1))
	}
	bounds := src.Bounds()
	srcW := bounds.Dx()
	srcH := bounds.Dy()
	if srcW <= 0 || srcH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, dstW, dstH))
	}

	dst := image.NewRGBA(image.Rect(0, 0, dstW, dstH))

	xRatio := float64(srcW) / float64(dstW)
	yRatio := float64(srcH) / float64(dstH)

	// Fast path for *image.RGBA source
	if srcRGBA, ok := src.(*image.RGBA); ok {
		for dy := 0; dy < dstH; dy++ {
			sy := float64(dy) * yRatio
			sy0 := int(sy)
			sy1 := sy0 + 1
			if sy1 >= srcH {
				sy1 = srcH - 1
			}
			fy := sy - float64(sy0)
			ify := 1 - fy
			srcOff0 := (sy0+bounds.Min.Y-srcRGBA.Rect.Min.Y)*srcRGBA.Stride + (bounds.Min.X-srcRGBA.Rect.Min.X)*4
			srcOff1 := (sy1+bounds.Min.Y-srcRGBA.Rect.Min.Y)*srcRGBA.Stride + (bounds.Min.X-srcRGBA.Rect.Min.X)*4
			dstOff := dy * dst.Stride

			for dx := 0; dx < dstW; dx++ {
				sx := float64(dx) * xRatio
				sx0 := int(sx)
				sx1 := sx0 + 1
				if sx1 >= srcW {
					sx1 = srcW - 1
				}
				fx := sx - float64(sx0)
				ifx := 1 - fx

				o00 := srcOff0 + sx0*4
				o10 := srcOff0 + sx1*4
				o01 := srcOff1 + sx0*4
				o11 := srcOff1 + sx1*4
				sp := srcRGBA.Pix

				for ch := 0; ch < 4; ch++ {
					top := float64(sp[o00+ch])*ifx + float64(sp[o10+ch])*fx
					bot := float64(sp[o01+ch])*ifx + float64(sp[o11+ch])*fx
					dst.Pix[dstOff+ch] = uint8(top*ify + bot*fy)
				}
				dstOff += 4
			}
		}
		return dst
	}

	// Generic path for other image types
	for dy := 0; dy < dstH; dy++ {
		sy := float64(dy) * yRatio
		sy0 := int(sy)
		sy1 := sy0 + 1
		if sy1 >= srcH {
			sy1 = srcH - 1
		}
		fy := sy - float64(sy0)

		for dx := 0; dx < dstW; dx++ {
			sx := float64(dx) * xRatio
			sx0 := int(sx)
			sx1 := sx0 + 1
			if sx1 >= srcW {
				sx1 = srcW - 1
			}
			fx := sx - float64(sx0)

			r00, g00, b00, a00 := src.At(bounds.Min.X+sx0, bounds.Min.Y+sy0).RGBA()
			r10, g10, b10, a10 := src.At(bounds.Min.X+sx1, bounds.Min.Y+sy0).RGBA()
			r01, g01, b01, a01 := src.At(bounds.Min.X+sx0, bounds.Min.Y+sy1).RGBA()
			r11, g11, b11, a11 := src.At(bounds.Min.X+sx1, bounds.Min.Y+sy1).RGBA()

			lerp := func(v00, v10, v01, v11 uint32) uint8 {
				top := float64(v00)*(1-fx) + float64(v10)*fx
				bot := float64(v01)*(1-fx) + float64(v11)*fx
				v := (top*(1-fy) + bot*fy) / 257.0
				if v > 255 {
					v = 255
				}
				return uint8(v + 0.5)
			}

			off := dy*dst.Stride + dx*4
			dst.Pix[off+0] = lerp(r00, r10, r01, r11)
			dst.Pix[off+1] = lerp(g00, g10, g01, g11)
			dst.Pix[off+2] = lerp(b00, b10, b01, b11)
			dst.Pix[off+3] = lerp(a00, a10, a01, a11)
		}
	}
	return dst
}

// scaleImage scales an image using nearest-neighbor (fast fallback).
func scaleImage(src image.Image, dstW, dstH int) *image.RGBA {
	return scaleImageNearest(src, dstW, dstH)
}

// scaleImageNearest resizes with nearest-neighbour sampling: one source pixel
// per destination pixel, no interpolation. It is the cheap alternative to
// scaleImageBilinear and is used by draft renders.
func scaleImageNearest(src image.Image, dstW, dstH int) *image.RGBA {
	if dstW <= 0 || dstH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, 1, 1))
	}
	bounds := src.Bounds()
	srcW, srcH := bounds.Dx(), bounds.Dy()
	if srcW <= 0 || srcH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, dstW, dstH))
	}

	// Work from an *image.RGBA addressed from (0,0) so the sampling loop can copy
	// whole pixels out of Pix. Calling image.Image.At per pixel boxes a colour
	// and allocates on every call, which made this slower than the bilinear
	// filter it is meant to replace on JPEG (YCbCr) and paletted sources.
	s := normalizeToRGBA(src)

	dst := image.NewRGBA(image.Rect(0, 0, dstW, dstH))
	for dy := 0; dy < dstH; dy++ {
		sy := dy * srcH / dstH
		if sy >= srcH {
			sy = srcH - 1
		}
		srcRow := sy * s.Stride
		dstRow := dy * dst.Stride
		for dx := 0; dx < dstW; dx++ {
			sx := dx * srcW / dstW
			if sx >= srcW {
				sx = srcW - 1
			}
			so := srcRow + sx*4
			do := dstRow + dx*4
			dst.Pix[do] = s.Pix[so]
			dst.Pix[do+1] = s.Pix[so+1]
			dst.Pix[do+2] = s.Pix[so+2]
			dst.Pix[do+3] = s.Pix[so+3]
		}
	}
	return dst
}

// normalizeToRGBA returns src as an *image.RGBA whose pixels are addressed from
// (0,0). The input is returned unchanged when it already has that shape, so the
// common case costs nothing; other formats are converted once with image/draw,
// which is far cheaper than converting per sampled pixel.
func normalizeToRGBA(src image.Image) *image.RGBA {
	if s, ok := src.(*image.RGBA); ok && s.Rect.Min.X == 0 && s.Rect.Min.Y == 0 {
		return s
	}
	b := src.Bounds()
	out := image.NewRGBA(image.Rect(0, 0, b.Dx(), b.Dy()))
	draw.Draw(out, out.Bounds(), src, b.Min, draw.Src)
	return out
}

// scaleForRender resizes a decoded image for compositing, choosing the filter
// that the current quality mode allows. crop, when non-nil, is the source
// sub-rectangle in source pixel coordinates (a:srcRect); the resampler samples
// it directly so fractional crop edges land where PowerPoint puts them.
func (r *renderer) scaleForRender(src image.Image, dstW, dstH int, crop *[4]float64) *image.RGBA {
	if r.draft {
		return scaleImageNearest(applyIntCrop(src, crop), dstW, dstH)
	}
	return scaleImageMitchell(src, dstW, dstH, crop)
}

// applyIntCrop crops src to the whole-pixel approximation of the float crop
// rectangle. Used by the draft path, where the cheap nearest-neighbour
// resampler cannot sample sub-rectangles itself.
func applyIntCrop(src image.Image, crop *[4]float64) image.Image {
	if crop == nil {
		return src
	}
	b := src.Bounds()
	x0 := b.Min.X + int((*crop)[0])
	y0 := b.Min.Y + int((*crop)[1])
	x1 := b.Min.X + int((*crop)[2])
	y1 := b.Min.Y + int((*crop)[3])
	if x1 <= x0 || y1 <= y0 {
		return src
	}
	out := image.NewRGBA(image.Rect(0, 0, x1-x0, y1-y0))
	draw.Draw(out, out.Bounds(), src, image.Pt(x0, y0), draw.Src)
	return out
}

// scaleImageMitchell resamples src so that the source sub-rectangle crop (full
// image when nil) maps onto the dstW×dstH output.
//
// The filter is Mitchell-Netravali with B=C=1/3 — the kernel behind GDI+
// InterpolationModeHighQualityBicubic, which is what PowerPoint uses to scale
// pictures. Sampling is corner-mapped: destination pixel dx covers [dx, dx+1)
// of the destination rectangle and samples source coordinate
// cropLeft + dx*cropWidth/dstW, with no half-pixel offset. Both choices were
// pinned against PowerPoint COM goldens (r38): across the four picture slides
// of the comparison deck, corner mapping beat centre mapping on every one, and
// the Mitchell kernel beat bicubic a=-0.5, bilinear and Lanczos3. Sampling
// happens on the premultiplied sRGB values; edges are clamped.
func scaleImageMitchell(src image.Image, dstW, dstH int, crop *[4]float64) *image.RGBA {
	if dstW <= 0 || dstH <= 0 {
		return image.NewRGBA(image.Rect(0, 0, 1, 1))
	}
	rgba := normalizeToRGBA(src)
	b := rgba.Bounds()
	srcW, srcH := b.Dx(), b.Dy()
	full := [4]float64{0, 0, float64(srcW), float64(srcH)}
	if crop != nil {
		full = *crop
	}
	dst := image.NewRGBA(image.Rect(0, 0, dstW, dstH))
	if srcW <= 0 || srcH <= 0 || full[2]-full[0] <= 0 || full[3]-full[1] <= 0 {
		return dst
	}
	// Identity fast path: no crop and a 1:1 size match needs no filtering.
	if crop == nil && dstW == srcW && dstH == srcH {
		draw.Draw(dst, dst.Bounds(), rgba, b.Min, draw.Src)
		return dst
	}

	const (
		mB = 1.0 / 3.0
		mC = 1.0 / 3.0
	)
	kernel := func(t float64) float64 {
		t = math.Abs(t)
		t2 := t * t
		if t < 1 {
			return ((12-9*mB-6*mC)*t*t2 + (-18+12*mB+6*mC)*t2 + (6 - 2*mB)) / 6
		}
		if t < 2 {
			return ((-mB-6*mC)*t*t2 + (6*mB+30*mC)*t2 + (-12*mB-48*mC)*t + (8*mB + 24*mC)) / 6
		}
		return 0
	}

	cw := full[2] - full[0]
	ch := full[3] - full[1]

	// Per-destination-column source taps and normalised weights. Weights stay
	// float64 through the accumulation: float32 sums land a hair under the
	// .5 rounding boundary on symmetric blends and flip pixels by one.
	type axisTaps struct {
		idx [4]int32
		w   [4]float64
	}
	cols := make([]axisTaps, dstW)
	for dx := range cols {
		sx := full[0] + (float64(dx)+0.0)*cw/float64(dstW)
		f := math.Floor(sx)
		wsum := 0.0
		for i := 0; i < 4; i++ {
			idx := int(f) - 1 + i
			if idx < 0 {
				idx = 0
			}
			if idx >= srcW {
				idx = srcW - 1
			}
			cols[dx].idx[i] = int32(idx)
			cols[dx].w[i] = kernel(sx - f - float64(i-1))
			wsum += cols[dx].w[i]
		}
		if wsum > 0 {
			for i := 0; i < 4; i++ {
				cols[dx].w[i] /= wsum
			}
		}
	}

	srcPix := rgba.Pix
	stride := rgba.Stride
	for dy := 0; dy < dstH; dy++ {
		sy := full[1] + float64(dy)*ch/float64(dstH)
		f := math.Floor(sy)
		var rows axisTaps
		wsum := 0.0
		for j := 0; j < 4; j++ {
			idx := int(f) - 1 + j
			if idx < 0 {
				idx = 0
			}
			if idx >= srcH {
				idx = srcH - 1
			}
			rows.idx[j] = int32(idx)
			rows.w[j] = kernel(sy - f - float64(j-1))
			wsum += rows.w[j]
		}
		if wsum > 0 {
			for j := 0; j < 4; j++ {
				rows.w[j] /= wsum
			}
		}
		dstOff := dy * dst.Stride
		for dx := 0; dx < dstW; dx++ {
			col := &cols[dx]
			var r, g, bl, a float64
			for j := 0; j < 4; j++ {
				wy := rows.w[j]
				if wy == 0 {
					continue
				}
				rowOff := int(rows.idx[j]) * stride
				for i := 0; i < 4; i++ {
					w := wy * col.w[i]
					if w == 0 {
						continue
					}
					o := rowOff + int(col.idx[i])*4
					r += w * float64(srcPix[o])
					g += w * float64(srcPix[o+1])
					bl += w * float64(srcPix[o+2])
					a += w * float64(srcPix[o+3])
				}
			}
			dst.Pix[dstOff] = clampByte(r)
			dst.Pix[dstOff+1] = clampByte(g)
			dst.Pix[dstOff+2] = clampByte(bl)
			dst.Pix[dstOff+3] = clampByte(a)
			dstOff += 4
		}
	}
	return dst
}

// clampByte rounds a filter accumulator to a byte, half-up like the Python
// reference model the sampling semantics were pinned with.
func clampByte(v float64) uint8 {
	if v <= 0 {
		return 0
	}
	if v >= 255 {
		return 255
	}
	return uint8(v + 0.5)
}

// --- Utility functions ---

func abs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// decodeMetafileBitmap attempts to extract a renderable image from WMF/EMF
// metafile data. It first scans for embedded PNG or JPEG data, then falls
// back to parsing WMF DIB (Device Independent Bitmap) records or EMF records.
func decodeMetafileBitmap(data []byte, fc *FontCache) image.Image {
	if len(data) < 10 {
		return nil
	}

	// Try to find embedded PNG (89 50 4E 47) or JPEG (FF D8 FF) inside the data
	if img := findEmbeddedImage(data); img != nil {
		return img
	}

	// WMF: magic 01 00 09 00
	if len(data) > 4 && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x09 && data[3] == 0x00 {
		return decodeWMFDIB(data, fc)
	}

	// Placeable WMF: magic D7 CD C6 9A (22-byte header before standard WMF)
	if len(data) > 26 && data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A {
		return decodeWMFDIB(data[22:], fc)
	}

	// EMF: first DWORD is record type 1 (EMR_HEADER), magic 01 00 00 00
	if len(data) > 8 && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00 {
		return decodeEMFBitmap(data)
	}

	return nil
}

// findEmbeddedImage scans binary data for embedded PNG or JPEG signatures
// and attempts to decode the first one found.
func findEmbeddedImage(data []byte) image.Image {
	for i := 0; i < len(data)-8; i++ {
		// PNG signature: 89 50 4E 47 0D 0A 1A 0A
		if data[i] == 0x89 && data[i+1] == 0x50 && data[i+2] == 0x4E && data[i+3] == 0x47 &&
			data[i+4] == 0x0D && data[i+5] == 0x0A && data[i+6] == 0x1A && data[i+7] == 0x0A {
			if img, _, err := image.Decode(bytes.NewReader(data[i:])); err == nil {
				return img
			}
		}
		// JPEG signature: FF D8 FF
		if data[i] == 0xFF && data[i+1] == 0xD8 && data[i+2] == 0xFF {
			if img, _, err := image.Decode(bytes.NewReader(data[i:])); err == nil {
				return img
			}
		}
	}
	return nil
}

// decodeWMFDIB extracts a DIB bitmap from a WMF file by scanning for
// StretchDIBits (0x0B41) or SetDIBitsToDevice (0x0D33) records that
// contain a BITMAPINFOHEADER.
func decodeWMFDIB(data []byte, fc *FontCache) image.Image {
	if len(data) < 18 {
		return nil
	}

	// Parse WMF header to get window extent
	winW := 102 // default
	winH := 84

	// Collect all drawing operations from WMF records
	type dibRecord struct {
		destX, destY, destW, destH int
		rasterOp                   uint32
		img                        image.Image
		bitCount                   uint16
	}
	type textRecord struct {
		x, y    int
		text    string
		centerH bool // TA_CENTER
	}

	var dibs []dibRecord
	var texts []textRecord
	textAlignCenter := false

	pos := 18
	for pos+6 < len(data) {
		recSize := uint32(data[pos]) | uint32(data[pos+1])<<8 | uint32(data[pos+2])<<16 | uint32(data[pos+3])<<24
		recFunc := uint16(data[pos+4]) | uint16(data[pos+5])<<8
		recBytes := int(recSize) * 2
		if recBytes < 6 || pos+recBytes > len(data) {
			break
		}

		switch recFunc {
		case 0x020C: // SetWindowExt
			if recBytes >= 10 {
				winH = int(int16(uint16(data[pos+6]) | uint16(data[pos+7])<<8))
				winW = int(int16(uint16(data[pos+8]) | uint16(data[pos+9])<<8))
			}

		case 0x0B41, 0x0D33: // StretchDIBits, SetDIBitsToDevice
			if recBytes >= 26 {
				p := pos + 6
				rop := uint32(data[p]) | uint32(data[p+1])<<8 | uint32(data[p+2])<<16 | uint32(data[p+3])<<24
				srcH := int(int16(uint16(data[p+4]) | uint16(data[p+5])<<8))
				srcW := int(int16(uint16(data[p+6]) | uint16(data[p+7])<<8))
				_ = srcH
				_ = srcW
				dstH := int(int16(uint16(data[p+12]) | uint16(data[p+13])<<8))
				dstW := int(int16(uint16(data[p+14]) | uint16(data[p+15])<<8))
				dstY := int(int16(uint16(data[p+16]) | uint16(data[p+17])<<8))
				dstX := int(int16(uint16(data[p+18]) | uint16(data[p+19])<<8))

				// Find BITMAPINFOHEADER
				for j := pos + 6; j+40 <= pos+recBytes; j++ {
					biSz := uint32(data[j]) | uint32(data[j+1])<<8 | uint32(data[j+2])<<16 | uint32(data[j+3])<<24
					if biSz != 40 {
						continue
					}
					biPlanes := uint16(data[j+12]) | uint16(data[j+13])<<8
					if biPlanes != 1 {
						continue
					}
					biBitCount := uint16(data[j+14]) | uint16(data[j+15])<<8
					if biBitCount != 1 && biBitCount != 4 && biBitCount != 8 && biBitCount != 24 && biBitCount != 32 {
						continue
					}
					biW := int32(uint32(data[j+4]) | uint32(data[j+5])<<8 | uint32(data[j+6])<<16 | uint32(data[j+7])<<24)
					biH := int32(uint32(data[j+8]) | uint32(data[j+9])<<8 | uint32(data[j+10])<<16 | uint32(data[j+11])<<24)
					if biW <= 0 || biW > 4096 {
						continue
					}
					absH := biH
					if absH < 0 {
						absH = -absH
					}
					if absH <= 0 || absH > 4096 {
						continue
					}
					if img := parseDIB(data[j:pos+recBytes], recBytes-(j-pos)); img != nil {
						dibs = append(dibs, dibRecord{dstX, dstY, dstW, dstH, rop, img, biBitCount})
					}
					break
				}
			}

		case 0x012E: // SetTextAlign
			if recBytes >= 8 {
				align := uint16(data[pos+6]) | uint16(data[pos+7])<<8
				textAlignCenter = (align & 0x06) == 0x06 // TA_CENTER
			}

		case 0x0A32: // ExtTextOut
			if recBytes >= 14 {
				p := pos + 6
				ty := int(int16(uint16(data[p]) | uint16(data[p+1])<<8))
				tx := int(int16(uint16(data[p+2]) | uint16(data[p+3])<<8))
				count := int(int16(uint16(data[p+4]) | uint16(data[p+5])<<8))
				opts := uint16(data[p+6]) | uint16(data[p+7])<<8
				strOff := 8
				if opts&0x0006 != 0 {
					strOff = 16
				}
				if p+strOff+count <= pos+recBytes && count > 0 {
					raw := data[p+strOff : p+strOff+count]
					text := decodeGBKToUTF8(raw)
					texts = append(texts, textRecord{tx, ty, text, textAlignCenter})
				}
			}
		}

		pos += recBytes
	}

	if len(dibs) == 0 && len(texts) == 0 {
		return nil
	}

	// Render at a higher resolution for quality (up to 4x the WMF logical
	// units). The window extent is read straight out of the file, so it decides
	// how much memory this allocates: a metafile claiming a 3000x3000 logical
	// extent asks for a 12000x12000 canvas, which image.NewRGBA tries to
	// allocate and dies on — a crash the recover boundary in safety.go cannot
	// catch, because it is the allocator failing, not a panic. parseDIB already
	// refuses any dimension over 4096 and the EMF path clamps its canvas to
	// 2000 px, so apply the same ceiling here, giving up the quality multiplier
	// first so that ordinary metafiles keep their 4x.
	const maxMetafileDim = 2000
	scale := 4.0
	if ext := float64(maxInt(winW, winH)); ext > 0 && ext*scale > maxMetafileDim {
		scale = float64(maxMetafileDim) / ext
	}
	imgW := int(float64(winW) * scale)
	imgH := int(float64(winH) * scale)
	if imgW <= 0 || imgH <= 0 {
		imgW = 408
		imgH = 336
	}

	canvas := image.NewRGBA(image.Rect(0, 0, imgW, imgH))
	// Fill with white background
	draw.Draw(canvas, canvas.Bounds(), &image.Uniform{color.White}, image.Point{}, draw.Src)

	// Draw DIBs with mask compositing
	var maskImg image.Image
	for _, d := range dibs {
		dx := int(float64(d.destX) * scale)
		dy := int(float64(d.destY) * scale)
		dw := int(float64(d.destW) * scale)
		dh := int(float64(d.destH) * scale)
		scaled := scaleImageBilinear(d.img, dw, dh)

		if d.rasterOp == 0x008800C6 { // SRCAND - this is the mask
			maskImg = scaled
		} else if d.rasterOp == 0x00660046 && maskImg != nil { // SRCINVERT with mask
			// Apply mask: where mask is black, use the color image; where white, keep background
			for py := 0; py < dh && py < imgH-dy; py++ {
				for px := 0; px < dw && px < imgW-dx; px++ {
					mr, _, _, _ := maskImg.At(px, py).RGBA()
					if mr < 0x8000 { // mask is dark = draw pixel
						canvas.Set(dx+px, dy+py, scaled.At(px, py))
					}
				}
			}
			maskImg = nil
		} else {
			// Simple draw
			draw.Draw(canvas, image.Rect(dx, dy, dx+dw, dy+dh), scaled, image.Point{}, draw.Over)
		}
	}

	// Draw text
	for _, t := range texts {
		tx := int(float64(t.x) * scale)
		ty := int(float64(t.y) * scale)
		drawWMFText(canvas, tx, ty, t.text, scale, t.centerH, fc)
	}

	return canvas
}

// decodeGBKToUTF8 converts GBK/GB2312 encoded bytes to a UTF-8 string.
func decodeGBKToUTF8(data []byte) string {
	decoded, err := simplifiedchinese.GBK.NewDecoder().Bytes(data)
	if err != nil {
		return string(data)
	}
	return string(decoded)
}

// drawWMFText draws text onto the canvas at the given position.
func drawWMFText(canvas *image.RGBA, x, y int, text string, scale float64, centerH bool, fc *FontCache) {
	col := color.Black
	// Try to use a proper font that supports Chinese characters
	var face font.Face
	if fc != nil {
		// Try common Chinese fonts at a size proportional to the scale
		fontSize := 10 * scale
		for _, name := range []string{"microsoft yahei", "微软雅黑", "simsun", "宋体", "simhei", "黑体"} {
			if f := fc.GetFace(name, fontSize, false, false); f != nil {
				face = f
				break
			}
		}
	}
	if face == nil {
		face = basicfont.Face7x13
	}
	d := &font.Drawer{
		Dst:  canvas,
		Src:  image.NewUniform(col),
		Face: face,
		Dot:  fixed.P(x, y+face.Metrics().Ascent.Ceil()),
	}
	if centerH {
		// Measure text width and offset x to center
		textWidth := d.MeasureString(text)
		d.Dot.X = fixed.I(x) - textWidth/2
	}
	d.DrawString(text)
}

// parseDIB parses a BITMAPINFOHEADER + pixel data into an image.
func parseDIB(data []byte, maxLen int) image.Image {
	if len(data) < 40 {
		return nil
	}
	biWidth := int(int32(uint32(data[4]) | uint32(data[5])<<8 | uint32(data[6])<<16 | uint32(data[7])<<24))
	biHeight := int(int32(uint32(data[8]) | uint32(data[9])<<8 | uint32(data[10])<<16 | uint32(data[11])<<24))
	biBitCount := int(uint16(data[14]) | uint16(data[15])<<8)

	if biWidth <= 0 || biWidth > 4096 {
		return nil
	}
	absHeight := biHeight
	bottomUp := true
	if biHeight < 0 {
		absHeight = -biHeight
		bottomUp = false
	}
	if absHeight <= 0 || absHeight > 4096 {
		return nil
	}

	// Calculate palette size
	paletteEntries := 0
	if biBitCount <= 8 {
		paletteEntries = 1 << biBitCount
	}
	paletteSize := paletteEntries * 4 // RGBQUAD = 4 bytes each
	pixelOffset := 40 + paletteSize

	if pixelOffset >= len(data) {
		return nil
	}

	// Read palette
	palette := make([]color.RGBA, paletteEntries)
	for i := 0; i < paletteEntries && 40+i*4+3 < len(data); i++ {
		off := 40 + i*4
		palette[i] = color.RGBA{R: data[off+2], G: data[off+1], B: data[off], A: 255}
	}

	img := image.NewRGBA(image.Rect(0, 0, biWidth, absHeight))
	pixData := data[pixelOffset:]

	// Row stride (padded to 4-byte boundary)
	bitsPerRow := biWidth * biBitCount
	stride := ((bitsPerRow + 31) / 32) * 4

	for row := 0; row < absHeight; row++ {
		srcRow := row
		dstRow := row
		if bottomUp {
			dstRow = absHeight - 1 - row
		}
		_ = srcRow
		rowStart := row * stride
		if rowStart >= len(pixData) {
			break
		}

		for col := 0; col < biWidth; col++ {
			var c color.RGBA
			switch biBitCount {
			case 1:
				byteIdx := rowStart + col/8
				if byteIdx >= len(pixData) {
					continue
				}
				bit := (pixData[byteIdx] >> (7 - uint(col%8))) & 1
				if int(bit) < len(palette) {
					c = palette[bit]
				}
			case 4:
				byteIdx := rowStart + col/2
				if byteIdx >= len(pixData) {
					continue
				}
				var nibble byte
				if col%2 == 0 {
					nibble = (pixData[byteIdx] >> 4) & 0x0F
				} else {
					nibble = pixData[byteIdx] & 0x0F
				}
				if int(nibble) < len(palette) {
					c = palette[nibble]
				}
			case 8:
				byteIdx := rowStart + col
				if byteIdx >= len(pixData) {
					continue
				}
				idx := pixData[byteIdx]
				if int(idx) < len(palette) {
					c = palette[idx]
				}
			case 24:
				byteIdx := rowStart + col*3
				if byteIdx+2 >= len(pixData) {
					continue
				}
				c = color.RGBA{R: pixData[byteIdx+2], G: pixData[byteIdx+1], B: pixData[byteIdx], A: 255}
			case 32:
				byteIdx := rowStart + col*4
				if byteIdx+3 >= len(pixData) {
					continue
				}
				c = color.RGBA{R: pixData[byteIdx+2], G: pixData[byteIdx+1], B: pixData[byteIdx], A: 255}
			default:
				continue
			}
			img.SetRGBA(col, dstRow, c)
		}
	}

	return img
}

// decodeEMFBitmap extracts a bitmap from an EMF (Enhanced Metafile) by
// scanning for EMR_STRETCHDIBITS (0x51) or EMR_BITBLT (0x4C) records
// that contain a BITMAPINFOHEADER.
func decodeEMFBitmap(data []byte) image.Image {
	if len(data) < 88 {
		return nil
	}
	// EMF header: first record is EMR_HEADER (type=1)
	// Each EMR record: DWORD type, DWORD size
	pos := 0
	var bestImg image.Image
	for pos+8 <= len(data) {
		recType := uint32(data[pos]) | uint32(data[pos+1])<<8 | uint32(data[pos+2])<<16 | uint32(data[pos+3])<<24
		recSize := uint32(data[pos+4]) | uint32(data[pos+5])<<8 | uint32(data[pos+6])<<16 | uint32(data[pos+7])<<24

		if recSize < 8 || pos+int(recSize) > len(data) {
			break
		}

		// EMR_STRETCHDIBITS = 0x51, EMR_BITBLT = 0x4C, EMR_SETDIBITSTODEVICE = 0x50
		if recType == 0x51 || recType == 0x4C || recType == 0x50 {
			recData := data[pos : pos+int(recSize)]
			// Scan for BITMAPINFOHEADER (biSize=40) with validation
			for j := 8; j+40 <= len(recData); j++ {
				biSz := uint32(recData[j]) | uint32(recData[j+1])<<8 | uint32(recData[j+2])<<16 | uint32(recData[j+3])<<24
				if biSz != 40 {
					continue
				}
				// Validate: biPlanes must be 1
				biPlanes := uint16(recData[j+12]) | uint16(recData[j+13])<<8
				if biPlanes != 1 {
					continue
				}
				// Validate: biBitCount must be valid
				biBitCount := uint16(recData[j+14]) | uint16(recData[j+15])<<8
				if biBitCount != 1 && biBitCount != 4 && biBitCount != 8 && biBitCount != 24 && biBitCount != 32 {
					continue
				}
				if img := parseDIB(recData[j:], len(recData)-j); img != nil {
					if bestImg == nil || biBitCount > 1 {
						bestImg = img
					}
				}
				break
			}
		}

		// EMR_EOF = 0x0E
		if recType == 0x0E {
			break
		}

		pos += int(recSize)
	}
	if bestImg != nil {
		return bestImg
	}
	// Fallback: try vector rendering for EMFs without embedded bitmaps
	return renderEMFVector(data)
}
