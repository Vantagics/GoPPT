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
	FontCache *FontCache
	// OverlayOpacityScale scales the opacity of semi-transparent shape fills.
	// Value between 0.0 and 1.0. Default 0 means use 1.0 (no change).
	// Set to e.g. 0.5 to halve the opacity of overlays, making dark backgrounds brighter.
	OverlayOpacityScale float64
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
func (p *Presentation) SlideToImage(slideIndex int, opts *RenderOptions) (image.Image, error) {
	if slideIndex < 0 || slideIndex >= len(p.slides) {
		return nil, fmt.Errorf("slide index %d out of range (0-%d)", slideIndex, len(p.slides)-1)
	}
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	if opts.Width <= 0 {
		opts.Width = 960
	}

	slide := p.slides[slideIndex]
	layout := p.layout

	slideW := float64(layout.CX)
	slideH := float64(layout.CY)
	imgW := opts.Width
	imgH := int(float64(imgW) * slideH / slideW)

	scaleX := float64(imgW) / slideW
	scaleY := float64(imgH) / slideH

	img := image.NewRGBA(image.Rect(0, 0, imgW, imgH))

	fc := opts.FontCache
	if fc == nil {
		fc = NewFontCache(opts.FontDirs...)
	}
	dpi := opts.DPI
	if dpi <= 0 {
		dpi = 96
	}

	r := &renderer{
		img:                 img,
		scaleX:              scaleX,
		scaleY:              scaleY,
		fontCache:           fc,
		dpi:                 dpi,
		overlayOpacityScale: opts.OverlayOpacityScale,
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
			r.fillGradientLinear(img.Bounds(), slide.background)
			drawn = true
		case FillGradientPath:
			r.fillGradientPath(img.Bounds(), slide.background)
			drawn = true
		}
	}
	if !drawn {
		r.fillRectFast(img.Bounds(), bgColor)
	}

	// Render shapes in their original XML order (z-order).
	// Shapes that appear earlier in the spTree are behind shapes that appear later,
	// matching PowerPoint's rendering behavior.
	for _, shape := range slide.shapes {
		r.renderShape(shape)
	}

	return img, nil
}

// SlidesToImages renders all slides to images.
func (p *Presentation) SlidesToImages(opts *RenderOptions) ([]image.Image, error) {
	if opts == nil {
		opts = DefaultRenderOptions()
	}
	if opts.FontCache == nil {
		opts.FontCache = NewFontCache(opts.FontDirs...)
	}
	images := make([]image.Image, len(p.slides))
	for i := range p.slides {
		img, err := p.SlideToImage(i, opts)
		if err != nil {
			return nil, fmt.Errorf("slide %d: %w", i, err)
		}
		images[i] = img
	}
	return images, nil
}

// SaveSlideAsImage renders a slide and saves it to a file.
func (p *Presentation) SaveSlideAsImage(slideIndex int, path string, opts *RenderOptions) error {
	img, err := p.SlideToImage(slideIndex, opts)
	if err != nil {
		return err
	}
	return saveImage(img, path, opts)
}

// SaveSlidesAsImages renders all slides and saves them to files.
// The pattern should contain %d for the slide number (1-based), e.g. "slide_%d.png".
func (p *Presentation) SaveSlidesAsImages(pattern string, opts *RenderOptions) error {
	for i := range p.slides {
		path := fmt.Sprintf(pattern, i+1)
		if err := p.SaveSlideAsImage(i, path, opts); err != nil {
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
}

func (r *renderer) renderShape(shape Shape) {
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
	}
}

func (r *renderer) emuToPixelX(emu int64) int { return int(math.Round(float64(emu) * r.scaleX)) }
func (r *renderer) emuToPixelY(emu int64) int { return int(math.Round(float64(emu) * r.scaleY)) }

// hundredthPtToPixelY converts hundredths of a point (from spcPts) to pixels.
// spcPts values are in 1/100 of a point, e.g. 1200 = 12pt.
// 1 point = 12700 EMU, so 1/100 point = 127 EMU.
func (r *renderer) hundredthPtToPixelY(val int) int {
	emu := float64(val) * 127.0
	return int(emu * r.scaleY)
}

func argbToRGBA(c Color) color.RGBA {
	return color.RGBA{R: c.GetRed(), G: c.GetGreen(), B: c.GetBlue(), A: c.GetAlpha()}
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
	draw.Draw(r.img, rect, &image.Uniform{c}, image.Point{}, draw.Over)
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
	tmpR := &renderer{img: tmp, scaleX: r.scaleX, scaleY: r.scaleY, fontCache: r.fontCache, dpi: r.dpi, fontScale: r.fontScale}
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
	rad := -float64(rotation) * math.Pi / 180.0
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
			// Step 1: un-flip (flip is self-inverse, applied in rotated space)
			fx, fy := rx, ry
			if flipH {
				fx = -fx
			}
			if flipV {
				fy = -fy
			}
			// Step 2: un-rotate (inverse rotation)
			sx := fx*cosA + fy*sinA + cx
			sy := -fx*sinA + fy*cosA + cy
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
	defer func() { r.fontScale = prevFontScale }()

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
	// Also apply auto-shrink for AutoFitNone when text still overflows after
	// inset reduction — Go's CJK font metrics often produce larger line heights
	// than PowerPoint, causing text to overflow shapes that fit perfectly in
	// the original authoring environment.
	shouldAutoShrink := false
	if s.autoFit == AutoFitNormal && (s.fontScale == 0 || s.fontScale == 100000) {
		shouldAutoShrink = true
	} else if s.autoFit == AutoFitNone && (s.fontScale == 0 || s.fontScale == 100000) {
		textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
		if textH > h && h > 0 {
			// Text exceeds the full shape height — font metrics are too large
			shouldAutoShrink = true
		}
	}
	if shouldAutoShrink {
		textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, wordWrap)
		if textH > th && th > 0 {
			// Binary search for the right scale factor
			lo, hi := 0.1, 1.0
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

		if !skipText {
			if vertRotation != 0 {
				// For vertical text, draw into a rotated buffer with swapped dimensions.
				vtw, vth := drawTH, tw // text area: width=drawTH, height=tw (before rotation)
				if vtw > 0 && vth > 0 {
					tmp := image.NewRGBA(image.Rect(0, 0, vtw, vth))
					tmpR := &renderer{img: tmp, scaleX: tr.scaleX, scaleY: tr.scaleY, fontCache: tr.fontCache, dpi: tr.dpi, fontScale: tr.fontScale}
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
			if vertRotation != 0 {
				vtw, vth := drawTH, tw
				if vtw > 0 && vth > 0 {
					tmp := image.NewRGBA(image.Rect(0, 0, vtw, vth))
					tmpR := &renderer{img: tmp, scaleX: tr.scaleX, scaleY: tr.scaleY, fontCache: tr.fontCache, dpi: tr.dpi, fontScale: tr.fontScale}
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
	if err != nil {
		r.drawRect(image.Rect(x, y, x+w, y+h), color.RGBA{R: 200, G: 200, B: 200, A: 255}, 1)
		return
	}

	rotation := s.GetRotation()
	flipH := s.GetFlipHorizontal()
	flipV := s.GetFlipVertical()

	drawImg := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		scaledImg := scaleImageBilinear(srcImg, w, h)
		draw.Draw(tr.img, image.Rect(ox, oy, ox+w, oy+h), scaledImg, image.Point{}, draw.Over)
	}

	if rotation != 0 || flipH || flipV {
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawImg)
	} else {
		drawImg(r)
	}
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
	defer func() { r.fontScale = prevFontScale }()

	// Vertical text direction
	vertRotation := 0
	if s.textDirection == "vert" || s.textDirection == "eaVert" || s.textDirection == "wordArtVert" {
		vertRotation = 270
	} else if s.textDirection == "vert270" {
		vertRotation = 90
	}

	drawContent := func(tr *renderer) {
		ox, oy := x, y
		if tr != r {
			ox, oy = 0, 0
		}
		rect := image.Rect(ox, oy, ox+w, oy+h)
		if s.shadow != nil && s.shadow.Visible {
			switch s.shapeType {
			case AutoShapeRoundedRect:
				sRadius := minInt(w, h) * 16667 / 100000
				if s.adjustValues != nil {
					if adj, ok := s.adjustValues["adj"]; ok {
						sRadius = minInt(w, h) * adj / 200000
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
				textH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, true)
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
			if (s.fontScale == 0 || s.fontScale == 100000) {
				atextH := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, true)
				if atextH > h && h > 0 && atextH > th && th > 0 {
					lo, hi := 0.1, 1.0
					for i := 0; i < 10; i++ {
						mid := (lo + hi) / 2
						r.fontScale = mid
						mh := r.measureParagraphsHeight(s.paragraphs, tw, th, s.textAnchor, true)
						if mh > th {
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
					tmpR := &renderer{img: tmp, scaleX: tr.scaleX, scaleY: tr.scaleY, fontCache: tr.fontCache, dpi: tr.dpi, fontScale: tr.fontScale}
					tmpR.drawParagraphs(s.paragraphs, 0, 0, vtw, vth, s.textAnchor, true)
					rotateAndComposite(tr.img, tmp, tx, ty, tw, th, vertRotation)
				}
			} else {
				tr.drawParagraphs(s.paragraphs, tx, ty, tw, th, s.textAnchor, true)
			}
		} else if s.text != "" {
			tr.drawStringCentered(s.text, tr.getFace(NewFont()), color.RGBA{A: 255}, rect)
		}
	}

	// For uturnArrow with 90/270 rotation, swap geometry dimensions.
	needsGeomSwap := s.shapeType == AutoShapeUturnArrow &&
		(rotation == 90 || rotation == 270)

	// For rtTriangle with 90/270 rotation, OOXML ext gives the rotated
	// bounding box size. Draw the mirror-image triangle in the buffer so
	// that after rotation the filled area covers the correct half.
	needsRtTriSwap := s.shapeType == AutoShapeRtTriangle &&
		(rotation == 90 || rotation == 270)

	if needsRtTriSwap {
		drawSwapped := func(tr *renderer) {
			if s.fill != nil && s.fill.Type != FillNone {
				fc := argbToRGBA(s.fill.Color)
				fc = tr.scaleAlpha(fc)
				pts := []fpoint{
					{0, 0},
					{float64(w), 0},
					{float64(w), float64(h)},
				}
				tr.fillPolygon(pts, fc)
			}
		}
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawSwapped)
	} else if needsGeomSwap {
		drawSwapped := func(tr *renderer) {
			if s.fill != nil && s.fill.Type != FillNone {
				fc := argbToRGBA(s.fill.Color)
				fc = tr.scaleAlpha(fc)
				tr.fillUturnArrowTransposed(0, 0, w, h, fc, s.adjustValues)
			}
		}
		r.renderRotated(x, y, w, h, rotation, flipH, flipV, drawSwapped)
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
		r.fillHexagon(x, y, w, h, fc)
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
	case AutoShapeSnip2SameRect:
		r.fillSnip2SameRect(x, y, w, h, fc, s.adjustValues)
	case AutoShapeUturnArrow:
		r.fillUturnArrow(x, y, w, h, fc, s.adjustValues)
	case AutoShapeBentArrow:
		r.fillBentArrow(x, y, w, h, fc, s.adjustValues)
	case AutoShapeArc:
		// Arc preset geometry has no fill by default (it's just a stroke).
		// Skip fill for arc shapes.
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
		pts := regularPolygonPoints(x, y, w, h, 6, 0)
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
	case AutoShapeArc:
		r.renderArcBorder(s, x, y, w, h, bc, pw)
	default:
		r.drawRectBorder(image.Rect(x, y, x+w, y+h), bc, pw, s.border.Style)
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
		pts := r.customPathToPixelPoints(s.customPath, ox, oy, w, h)
		if len(pts) >= 2 {
			// Rotate around bounding box center
			cxPx := float64(ox) + float64(w)/2.0
			cyPx := float64(oy) + float64(h)/2.0
			rad := float64(rotation) * math.Pi / 180.0
			cosA := math.Cos(rad)
			sinA := math.Sin(rad)
			for i := range pts {
				dx := pts[i].x - cxPx
				dy := pts[i].y - cyPx
				pts[i].x = dx*cosA - dy*sinA + cxPx
				pts[i].y = dx*sinA + dy*cosA + cyPx
			}

			pw := maxInt(int(float64(s.GetLineWidthEMU())*r.scaleX), 1)
			c := argbToRGBA(s.lineColor)
			ls := s.lineStyle
			if ls == BorderDash || ls == BorderDot {
				r.drawDashedPolylineAA(pts, c, pw, ls)
			} else {
				for i := 1; i < len(pts); i++ {
					r.drawLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), c, pw)
				}
			}
			intPts := make([][2]int, len(pts))
			for i, p := range pts {
				intPts[i] = [2]int{int(p.x), int(p.y)}
			}
			if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
				r.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, c, pw, s.headEnd)
			}
			if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
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

	// Custom geometry path (freeform curved arrows, etc.)
	if s.customPath != nil && len(s.customPath.Commands) > 0 {
		pts := r.customPathToPixelPoints(s.customPath, ox, oy, w, h)
		if len(pts) >= 2 {
			if ls == BorderDash || ls == BorderDot {
				r.drawDashedPolylineAA(pts, c, pw, ls)
			} else {
				for i := 1; i < len(pts); i++ {
					r.drawLineAA(int(pts[i-1].x), int(pts[i-1].y), int(pts[i].x), int(pts[i].y), c, pw)
				}
			}
			intPts := make([][2]int, len(pts))
			for i, p := range pts {
				intPts[i] = [2]int{int(p.x), int(p.y)}
			}
			if s.headEnd != nil && s.headEnd.Type != ArrowNone && s.headEnd.Type != "" {
				r.drawArrowOnPath(intPts[0][0], intPts[0][1], intPts, c, pw, s.headEnd)
			}
			if s.tailEnd != nil && s.tailEnd.Type != ArrowNone && s.tailEnd.Type != "" {
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

	pad := 3

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
			r.drawParagraphs(cell.paragraphs, cx+pad, cy+pad, cellW-2*pad, cellH-2*pad, TextAnchorNone, true)
		}
	}
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
	if cp.Width <= 0 || cp.Height <= 0 {
		return nil
	}
	scX := float64(w) / float64(cp.Width)
	scY := float64(h) / float64(cp.Height)

	toPixel := func(p PathPoint) fpoint {
		return fpoint{float64(ox) + float64(p.X)*scX, float64(oy) + float64(p.Y)*scY}
	}

	var pts []fpoint
	var lastPt fpoint
	for _, cmd := range cp.Commands {
		switch cmd.Type {
		case "moveTo", "lnTo":
			if len(cmd.Pts) > 0 {
				p := toPixel(cmd.Pts[0])
				pts = append(pts, p)
				lastPt = p
			}
		case "cubicBezTo":
			// Flatten cubic bezier into line segments for accurate curves
			if len(cmd.Pts) >= 3 {
				cp1 := toPixel(cmd.Pts[0])
				cp2 := toPixel(cmd.Pts[1])
				ep := toPixel(cmd.Pts[2])
				bezPts := r.flattenCubicBezier(lastPt.x, lastPt.y, cp1.x, cp1.y, cp2.x, cp2.y, ep.x, ep.y, 0)
				pts = append(pts, bezPts...)
				pts = append(pts, ep)
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
				pts = append(pts, bezPts...)
				pts = append(pts, ep)
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
				pts = append(pts, p)
				lastPt = p
			}
		}
	}
	return pts
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

func (r *renderer) fillGradientLinear(rect image.Rectangle, fill *Fill) {
	startC := argbToRGBA(fill.Color)
	endC := argbToRGBA(fill.EndColor)
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
			it := 1 - t
			pix[off] = uint8(float64(startC.R)*it + float64(endC.R)*t)
			pix[off+1] = uint8(float64(startC.G)*it + float64(endC.G)*t)
			pix[off+2] = uint8(float64(startC.B)*it + float64(endC.B)*t)
			pix[off+3] = uint8(float64(startC.A)*it + float64(endC.A)*t)
			off += 4
		}
	}
}

func (r *renderer) fillGradientPath(rect image.Rectangle, fill *Fill) {
	startC := argbToRGBA(fill.Color)
	endC := argbToRGBA(fill.EndColor)
	w := rect.Dx()
	h := rect.Dy()
	if w <= 0 || h <= 0 {
		return
	}
	cx := float64(w) / 2
	cy := float64(h) / 2
	maxDist := math.Sqrt(cx*cx + cy*cy)
	if maxDist < 1 {
		maxDist = 1
	}
	invMaxDist := 1.0 / maxDist

	pix := r.img.Pix
	bounds := r.img.Bounds()
	stride := r.img.Stride

	for py := rect.Min.Y; py < rect.Max.Y; py++ {
		if py < bounds.Min.Y || py >= bounds.Max.Y {
			continue
		}
		dyf := float64(py-rect.Min.Y) - cy
		dy2 := dyf * dyf
		off := (py-bounds.Min.Y)*stride + (maxInt(rect.Min.X, bounds.Min.X)-bounds.Min.X)*4
		for px := maxInt(rect.Min.X, bounds.Min.X); px < minInt(rect.Max.X, bounds.Max.X); px++ {
			dxf := float64(px-rect.Min.X) - cx
			t := math.Sqrt(dxf*dxf+dy2) * invMaxDist
			if t > 1 {
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
	if shadow == nil || !shadow.Visible {
		return
	}
	rad := float64(shadow.Direction) * math.Pi / 180.0
	dist := float64(shadow.Distance) * r.scaleX
	dx := int(dist * math.Cos(rad))
	dy := int(dist * math.Sin(rad))
	shadowColor := argbToRGBA(shadow.Color)
	shadowColor.A = uint8(float64(shadow.Alpha) * 255 / 100)
	shadowRect := rect.Add(image.Pt(dx, dy))

	blur := shadow.BlurRadius
	if blur <= 0 {
		r.fillRectBlend(shadowRect, shadowColor)
		return
	}

	// Box-blur approximation: render shadow at full alpha, then apply a simple
	// multi-pass box expansion with decreasing alpha from outside in.
	// We draw from outermost ring inward so inner pixels get the strongest alpha.
	steps := minInt(blur, 10)
	for i := steps; i >= 0; i-- {
		t := float64(i) / float64(steps)
		alpha := uint8(float64(shadowColor.A) * (1 - t*t)) // quadratic falloff
		c := color.RGBA{R: shadowColor.R, G: shadowColor.G, B: shadowColor.B, A: alpha}
		expanded := shadowRect.Inset(-i)
		// Only draw the ring (not the interior) for outer layers
		if i > 0 {
			inner := shadowRect.Inset(-(i - 1))
			// Top strip
			r.fillRectBlend(image.Rect(expanded.Min.X, expanded.Min.Y, expanded.Max.X, inner.Min.Y), c)
			// Bottom strip
			r.fillRectBlend(image.Rect(expanded.Min.X, inner.Max.Y, expanded.Max.X, expanded.Max.Y), c)
			// Left strip
			r.fillRectBlend(image.Rect(expanded.Min.X, inner.Min.Y, inner.Min.X, inner.Max.Y), c)
			// Right strip
			r.fillRectBlend(image.Rect(inner.Max.X, inner.Min.Y, expanded.Max.X, inner.Max.Y), c)
		} else {
			r.fillRectBlend(expanded, c)
		}
	}
}


func (r *renderer) renderShadowRounded(shadow *Shadow, rect image.Rectangle, radius int) {
	if shadow == nil || !shadow.Visible {
		return
	}
	rad := float64(shadow.Direction) * math.Pi / 180.0
	dist := float64(shadow.Distance) * r.scaleX
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
	tmpR := &renderer{img: tmp, scaleX: r.scaleX, scaleY: r.scaleY}

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

func (r *renderer) fillPolygonGradient(pts []fpoint, fill *Fill) {
	if len(pts) < 3 || fill == nil {
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
	// OOXML uturnArrow preset geometry.
	// Two vertical shafts connected by a semicircular arc at the BOTTOM.
	// The LEFT shaft has an arrowhead pointing UP.
	//
	// adj1 = shaft width (fraction of w / 100000)
	// adj2 = arrowhead extra width beyond shaft (fraction of w / 100000)
	// adj3 = arrowhead height (fraction of h / 100000)
	// adj4 = horizontal span of U-turn (fraction of w / 100000) — distance
	//        between outer edges of the two shafts
	// adj5 = total height used (fraction of h / 100000)
	adj1v := 25000
	adj2v := 25000
	adj3v := 25000
	adj4v := 43750
	adj5v := 100000
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

	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)

	shaftW := fw * float64(adj1v) / 100000.0
	headExtra := fw * float64(adj2v) / 100000.0
	headH := fh * float64(adj3v) / 100000.0
	uWidth := fw * float64(adj4v) / 100000.0
	totalH := fh * float64(adj5v) / 100000.0

	// Two shafts side by side, connected by U-turn arc at bottom.
	// Left shaft: x=0 to x=shaftW
	// Right shaft: x=(uWidth-shaftW) to x=uWidth
	leftOuter := fx
	leftInner := fx + shaftW
	rightOuter := fx + uWidth
	rightInner := rightOuter - shaftW
	if rightInner < leftInner {
		rightInner = leftInner
	}

	// Arc at BOTTOM connecting the two shafts
	outerRx := uWidth / 2
	gap := rightInner - leftInner
	if gap < 0 {
		gap = 0
	}
	innerRx := gap / 2
	arcCX := (leftOuter + rightOuter) / 2

	// Arc Ry: semicircular — use outerRx as Ry for a circular arc,
	// but cap to available height after arrowhead.
	availH := totalH - headH
	outerRy := outerRx
	if outerRy > availH*0.5 {
		outerRy = availH * 0.5
	}
	if outerRy < 1 {
		outerRy = 1
	}
	innerRy := outerRy * innerRx / outerRx
	if outerRx == 0 {
		innerRy = 0
	}

	shaftTop := fy
	arcCY := fy + totalH - outerRy

	// Arrowhead on LEFT shaft, pointing UP
	arrowCenterX := (leftOuter + leftInner) / 2
	halfHead := shaftW/2 + headExtra
	arrowLeft := arrowCenterX - halfHead
	arrowRight := arrowCenterX + halfHead
	if arrowLeft < fx {
		arrowLeft = fx
	}
	if arrowRight > fx+fw {
		arrowRight = fx + fw
	}
	arrowTipY := shaftTop
	arrowBaseY := shaftTop + headH

	pts := make([]fpoint, 0, 80)
	steps := 40

	// Start: right shaft outer, from top going down to arc
	pts = append(pts, fpoint{rightOuter, shaftTop})
	pts = append(pts, fpoint{rightOuter, arcCY})

	// Outer arc (right to left, curving DOWN)
	for i := 0; i <= steps; i++ {
		angle := math.Pi * float64(i) / float64(steps)
		px := arcCX + outerRx*math.Cos(angle)
		py := arcCY + outerRy*math.Sin(angle)
		pts = append(pts, fpoint{px, py})
	}

	// Left shaft outer, going up to arrowhead base
	pts = append(pts, fpoint{leftOuter, arcCY})
	pts = append(pts, fpoint{leftOuter, arrowBaseY})

	// Arrowhead left wing
	pts = append(pts, fpoint{arrowLeft, arrowBaseY})

	// Arrow tip (pointing up)
	pts = append(pts, fpoint{arrowCenterX, arrowTipY})

	// Arrowhead right wing
	pts = append(pts, fpoint{arrowRight, arrowBaseY})

	// Left shaft inner, going down to arc
	pts = append(pts, fpoint{leftInner, arrowBaseY})
	pts = append(pts, fpoint{leftInner, arcCY})

	// Inner arc (left to right, curving DOWN)
	for i := steps; i >= 0; i-- {
		angle := math.Pi * float64(i) / float64(steps)
		px := arcCX + innerRx*math.Cos(angle)
		py := arcCY + innerRy*math.Sin(angle)
		pts = append(pts, fpoint{px, py})
	}

	// Right shaft inner, going up to top
	pts = append(pts, fpoint{rightInner, arcCY})
	pts = append(pts, fpoint{rightInner, shaftTop})

	r.fillPolygon(pts, c)
}

// fillUturnArrowTransposed draws a U-turn arrow geometry transposed in the
// w×h buffer. The adj fractions that normally use w now use h (visual width)
// and those that use h now use w (visual height). The shafts run horizontally
// (along X) and the U-turn arc is at the right side (high X).
// This is used for 90°/270° rotations where the geometry needs to fill the
// full buffer width to span the full visual height after rotation.
func (r *renderer) fillUturnArrowTransposed(x, y, w, h int, c color.RGBA, adj map[string]int) {
	adj1v := 25000
	adj2v := 25000
	adj3v := 25000
	adj4v := 43750
	adj5v := 100000
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

	fx, fy := float64(x), float64(y)
	fw, fh := float64(w), float64(h)

	// Transposed geometry: shafts run along X, arc connects them vertically.
	// adj1/adj2/adj4 control Y-direction dimensions → use fh (short axis in buffer,
	// becomes visual width after 270° rotation).
	// adj3/adj5 control X-direction dimensions → use fw (long axis in buffer,
	// becomes visual height after rotation, must span all boxes).
	shaftW := fh * float64(adj1v) / 100000.0    // shaft thickness (Y direction)
	headExtra := fh * float64(adj2v) / 100000.0 // extra arrowhead width beyond shaft
	headH := fw * float64(adj3v) / 100000.0     // arrowhead length (X direction)
	uWidth := fh * float64(adj4v) / 100000.0    // U-turn span between shafts (Y direction)
	totalH := fw * float64(adj5v) / 100000.0    // total shaft length (X direction)

	// Two shafts side by side in Y, running along X.
	// U-turn arc at LEFT (low X), arrowhead at RIGHT (high X).
	// After flipV + 270° CW rotation: left→top (U-turn at visual top),
	// right→bottom (arrowhead at visual bottom). But original PPT shows
	// arrowhead pointing UP and U-turn at bottom, so we need:
	// arc at RIGHT (high X) → maps to visual bottom after rotation
	// arrowhead at LEFT (low X) → maps to visual top after rotation
	topOuter := fy
	topInner := fy + shaftW
	botOuter := fy + uWidth
	botInner := botOuter - shaftW
	if botInner < topInner {
		botInner = topInner
	}

	// Arc at RIGHT connecting the two shafts
	outerRy := uWidth / 2
	gap := botInner - topInner
	if gap < 0 {
		gap = 0
	}
	innerRy := gap / 2
	arcCY := (topOuter + botOuter) / 2

	availW := totalH - headH
	outerRx := outerRy // circular arc
	if outerRx > availW*0.5 {
		outerRx = availW * 0.5
	}
	if outerRx < 1 {
		outerRx = 1
	}
	innerRx := outerRx * innerRy / outerRy
	if outerRy == 0 {
		innerRx = 0
	}

	shaftLeft := fx                // left edge (arrowhead end)
	arcCX := fx + totalH - outerRx // arc center near right edge

	// Arrowhead on TOP shaft, pointing LEFT (→ visual top after rotation)
	arrowCenterY := (topOuter + topInner) / 2
	halfHead := shaftW/2 + headExtra
	arrowTop := arrowCenterY - halfHead
	arrowBot := arrowCenterY + halfHead
	if arrowTop < fy {
		arrowTop = fy
	}
	if arrowBot > fy+fh {
		arrowBot = fy + fh
	}
	arrowTipX := shaftLeft
	arrowBaseX := shaftLeft + headH

	pts := make([]fpoint, 0, 80)
	steps := 40

	// Layout: arrowhead at LEFT (low X), U-turn arc at RIGHT (high X).
	// Top shaft has the arrowhead; bottom shaft is plain.
	// Outer path goes clockwise:
	//   bottom-shaft left edge → right along bottom outer → arc curves RIGHT →
	//   left along top outer → arrowhead → back along top inner →
	//   arc inner curves RIGHT (reverse) → right along bottom inner → close

	// 1. Bottom shaft outer: left edge to arc
	pts = append(pts, fpoint{shaftLeft, botOuter})
	pts = append(pts, fpoint{arcCX, botOuter})

	// 2. Outer arc: from bottom (botOuter) to top (topOuter), curving RIGHT
	//    At angle -π/2 (bottom): cy + outerRy = botOuter ✓
	//    At angle +π/2 (top):    cy - outerRy = topOuter ✓
	for i := 0; i <= steps; i++ {
		angle := -math.Pi/2 + math.Pi*float64(i)/float64(steps)
		px := arcCX + outerRx*math.Cos(angle)
		py := arcCY - outerRy*math.Sin(angle)
		pts = append(pts, fpoint{px, py})
	}

	// 3. Top shaft outer: from arc to arrowhead base
	pts = append(pts, fpoint{arcCX, topOuter})
	pts = append(pts, fpoint{arrowBaseX, topOuter})

	// 4. Arrowhead pointing LEFT
	pts = append(pts, fpoint{arrowBaseX, arrowTop})
	pts = append(pts, fpoint{arrowTipX, arrowCenterY})
	pts = append(pts, fpoint{arrowBaseX, arrowBot})

	// 5. Top shaft inner: from arrowhead back to arc
	pts = append(pts, fpoint{arrowBaseX, topInner})
	pts = append(pts, fpoint{arcCX, topInner})

	// 6. Inner arc: from top (topInner) to bottom (botInner), curving RIGHT (reverse)
	for i := steps; i >= 0; i-- {
		angle := -math.Pi/2 + math.Pi*float64(i)/float64(steps)
		px := arcCX + innerRx*math.Cos(angle)
		py := arcCY - innerRy*math.Sin(angle)
		pts = append(pts, fpoint{px, py})
	}

	// 7. Bottom shaft inner: from arc back to left edge
	pts = append(pts, fpoint{arcCX, botInner})
	pts = append(pts, fpoint{shaftLeft, botInner})

	r.fillPolygon(pts, c)
}

// --- Text rendering ---

// getFace returns a font.Face for the given Font, falling back to basicfont.Face7x13.
func (r *renderer) getFace(f *Font) font.Face {
	if r.fontCache == nil {
		return basicfont.Face7x13
	}
	sizePt := float64(f.Size)
	if sizePt <= 0 {
		sizePt = 10
	}
	// Apply normAutofit font scale if set
	if r.fontScale > 0 && r.fontScale != 1.0 {
		sizePt *= r.fontScale
	}
	// Convert point size to pixels using the rendering scale.
	// 1pt = 12700 EMU; scaleX converts EMU to pixels.
	sizePixels := sizePt * 12700.0 * r.scaleX

	face := r.fontCache.GetFace(f.Name, sizePixels, f.Bold, f.Italic)
	if face != nil {
		return face
	}
	// Try East Asian font name if specified
	if f.NameEA != "" {
		face = r.fontCache.GetFace(f.NameEA, sizePixels, f.Bold, f.Italic)
		if face != nil {
			return face
		}
	}
	// CJK fallback names
	for _, fallback := range []string{
		"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
		"Yu Gothic", "Meiryo", "MS Gothic",
		"Malgun Gothic", "Gulim",
		"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
		"Arial", "Helvetica", "DejaVu Sans",
	} {
		face = r.fontCache.GetFace(fallback, sizePixels, f.Bold, f.Italic)
		if face != nil {
			return face
		}
	}
	return basicfont.Face7x13
}

// getCJKFace returns a font face suitable for CJK characters.
// It tries NameEA first, then common CJK fonts.
func (r *renderer) getCJKFace(f *Font) font.Face {
	if r.fontCache == nil {
		return nil
	}
	sizePt := float64(f.Size)
	if sizePt <= 0 {
		sizePt = 10
	}
	// Apply normAutofit font scale if set
	if r.fontScale > 0 && r.fontScale != 1.0 {
		sizePt *= r.fontScale
	}
	sizePixels := sizePt * 12700.0 * r.scaleX

	// Try East Asian font name first
	if f.NameEA != "" {
		face := r.fontCache.GetFace(f.NameEA, sizePixels, f.Bold, f.Italic)
		if face != nil {
			return face
		}
	}
	// CJK fallback
	for _, name := range []string{
		"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
		"Yu Gothic", "Meiryo", "MS Gothic",
		"Malgun Gothic", "Gulim",
		"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
	} {
		face := r.fontCache.GetFace(name, sizePixels, f.Bold, f.Italic)
		if face != nil {
			return face
		}
	}
	return nil
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

// splitRunByCJK splits a text run into sub-runs where CJK and non-CJK
// segments use different font faces. This ensures CJK characters are
// rendered with a CJK-capable font even when the primary font is Latin-only.
func (r *renderer) splitRunByCJK(text string, f *Font, latinFace, cjkFace font.Face) []textRun {
	if cjkFace == nil || latinFace == nil {
		// Can't split, return single run
		face := latinFace
		if face == nil {
			face = cjkFace
		}
		if face == nil {
			face = basicfont.Face7x13
		}
		return []textRun{{
			text:  text,
			font:  f,
			face:  face,
			width: measureStringWithKern(face, text).Ceil(),
		}}
	}

	var runs []textRun
	var buf strings.Builder
	wasCJK := false
	first := true

	for _, ch := range text {
		nowCJK := isCJK(ch)
		if !first && nowCJK != wasCJK {
			// Flush buffer
			seg := buf.String()
			face := latinFace
			if wasCJK {
				face = cjkFace
			}
			runs = append(runs, textRun{
				text:  seg,
				font:  f,
				face:  face,
				width: measureStringWithKern(face, seg).Ceil(),
			})
			buf.Reset()
		}
		buf.WriteRune(ch)
		wasCJK = nowCJK
		first = false
	}
	if buf.Len() > 0 {
		seg := buf.String()
		face := latinFace
		if wasCJK {
			face = cjkFace
		}
		runs = append(runs, textRun{
			text:  seg,
			font:  f,
			face:  face,
			width: measureStringWithKern(face, seg).Ceil(),
		})
	}
	return runs
}

// textRun holds a measured run of text with its formatting.
type textRun struct {
	text  string
	font  *Font
	face  font.Face
	width int
}

// textLine holds a line of text runs with total metrics.
type textLine struct {
	runs       []textRun
	width      int
	ascent     int
	descent    int
	lineHeight int
}

// buildTextLine measures a slice of textRuns and returns a textLine.
func (r *renderer) buildTextLine(runs []textRun) textLine {
	var tl textLine
	tl.runs = runs
	maxHeight := 0 // track font's recommended line-to-line height (includes line gap)
	for _, run := range runs {
		tl.width += run.width
		if run.face == nil {
			continue
		}
		metrics := run.face.Metrics()
		asc := metrics.Ascent.Ceil()
		desc := metrics.Descent.Ceil()
		if asc > tl.ascent {
			tl.ascent = asc
		}
		if desc > tl.descent {
			tl.descent = desc
		}
		// metrics.Height is the recommended line-to-line spacing which includes
		// the font's internal line gap (leading). PowerPoint's default single
		// spacing uses this full height, not just ascent+descent.
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
	if tl.lineHeight < 1 {
		tl.lineHeight = 14
	}
	return tl
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

	for _, para := range paragraphs {
		marginLeft := 0
		marginRight := 0
		indent := 0
		if para.alignment != nil {
			marginLeft = r.emuToPixelX(para.alignment.MarginLeft)
			marginRight = r.emuToPixelX(para.alignment.MarginRight)
			indent = r.emuToPixelX(para.alignment.Indent)
		}
		var paraRuns []textRun
		if para.bullet != nil && para.bullet.Type != BulletTypeNone {
			bRun := r.buildBulletRun(para.bullet, para)
			if bRun.text != "" {
				paraRuns = append(paraRuns, bRun)
			}
		}
		for _, elem := range para.elements {
			switch e := elem.(type) {
			case *TextRun:
				if e.text == "" {
					continue
				}
				f := e.font
				if f == nil {
					f = NewFont()
				}
				if containsCJK(e.text) && r.fontCache != nil {
					sizePt := float64(f.Size)
					if sizePt <= 0 {
						sizePt = 10
					}
					if r.fontScale > 0 && r.fontScale != 1.0 {
						sizePt *= r.fontScale
					}
					scaledPt := sizePt * 12700.0 * r.scaleX
					latinFace := r.fontCache.GetFace(f.Name, scaledPt, f.Bold, f.Italic)
					if latinFace == nil {
						latinFace = r.getFace(f)
					}
					cjkFace := r.getCJKFace(f)
					subRuns := r.splitRunByCJK(e.text, f, latinFace, cjkFace)
					paraRuns = append(paraRuns, subRuns...)
				} else {
					face := r.getFace(f)
					paraRuns = append(paraRuns, textRun{
						text:  e.text,
						font:  f,
						face:  face,
						width: measureStringWithKern(face, e.text).Ceil(),
					})
				}
			case *BreakElement:
				paraRuns = append(paraRuns, textRun{text: "\n"})
			}
		}
		availW := w - marginLeft - marginRight - indent
		if availW < 10 {
			availW = w
		}
		if !wordWrap {
			availW = 999999
		}
		lines := r.wrapRunLine(paraRuns, availW)
		if len(lines) == 0 {
			lines = []textLine{{lineHeight: 14}}
		}
		for i, line := range lines {
			li := lineInfo{
				lineHeight:  line.lineHeight,
				lineSpacing: para.lineSpacing,
			}
			if i == 0 {
				li.spaceBefore = r.hundredthPtToPixelY(para.spaceBefore)
			}
			if i == len(lines)-1 {
				li.spaceAfter = r.hundredthPtToPixelY(para.spaceAfter)
			}
			allLines = append(allLines, li)
		}
	}

	totalH := 0
	for i, li := range allLines {
		if i > 0 {
			totalH += li.spaceBefore
		}
		lh := li.lineHeight
		if li.lineSpacing < 0 {
			lh = int(float64(lh) * float64(-li.lineSpacing) / 100000.0)
		} else if li.lineSpacing > 0 {
			lh = r.hundredthPtToPixelY(li.lineSpacing)
		}
		totalH += lh
		totalH += li.spaceAfter
	}
	return totalH
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
		if para.bullet != nil && para.bullet.Type != BulletTypeNone {
			bRun := r.buildBulletRun(para.bullet, para)
			if bRun.text != "" {
				paraRuns = append(paraRuns, bRun)
			}
		}

		for _, elem := range para.elements {
			switch e := elem.(type) {
			case *TextRun:
				if e.text == "" {
					continue
				}
				f := e.font
				if f == nil {
					f = NewFont()
				}
				// If text contains CJK characters, split into CJK/Latin segments
				// so each segment uses an appropriate font face
				if containsCJK(e.text) && r.fontCache != nil {
					sizePt := float64(f.Size)
					if sizePt <= 0 {
						sizePt = 10
					}
					if r.fontScale > 0 && r.fontScale != 1.0 {
						sizePt *= r.fontScale
					}
					scaledPt := sizePt * 12700.0 * r.scaleX
					latinFace := r.fontCache.GetFace(f.Name, scaledPt, f.Bold, f.Italic)
					if latinFace == nil {
						latinFace = r.getFace(f)
					}
					cjkFace := r.getCJKFace(f)
					subRuns := r.splitRunByCJK(e.text, f, latinFace, cjkFace)
					paraRuns = append(paraRuns, subRuns...)
				} else {
					face := r.getFace(f)
					paraRuns = append(paraRuns, textRun{
						text:  e.text,
						font:  f,
						face:  face,
						width: measureStringWithKern(face, e.text).Ceil(),
					})
				}
			case *BreakElement:
				// Force a new line
				paraRuns = append(paraRuns, textRun{text: "\n"})
			}
		}

		// Wrap runs into lines
		availW := w - marginLeft - marginRight - indent
		if availW < 10 {
			availW = w
		}
		if !wordWrap {
			availW = 999999
		}
		lines := r.wrapRunLine(paraRuns, availW)
		if len(lines) == 0 {
			// Empty paragraph still takes space
			lines = []textLine{{lineHeight: 14}}
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
				li.spaceBefore = r.hundredthPtToPixelY(para.spaceBefore)
			}
			if i == len(lines)-1 {
				li.spaceAfter = r.hundredthPtToPixelY(para.spaceAfter)
			}
			allLines = append(allLines, li)
		}
	}

	// Calculate total height
	totalH := 0
	for i, li := range allLines {
		if i > 0 {
			totalH += li.spaceBefore
		}
		lh := li.line.lineHeight
		if li.lineSpacing < 0 {
			// spcPct: negative value, percentage * 1000 (e.g. -150000 = 150%)
			lh = int(float64(lh) * float64(-li.lineSpacing) / 100000.0)
		} else if li.lineSpacing > 0 {
			// spcPts: hundredths of a point (e.g. 1200 = 12pt)
			lh = r.hundredthPtToPixelY(li.lineSpacing)
		}
		totalH += lh
		totalH += li.spaceAfter
	}

	// Vertical anchor offset
	startY := y
	switch anchor {
	case TextAnchorMiddle:
		startY = y + (h-totalH)/2
		// When text overflows the available area, clamp to top so that
		// text only overflows at the bottom (matching PowerPoint behaviour).
		if startY < y {
			startY = y
		}
	case TextAnchorBottom:
		startY = y + h - totalH
		if startY < y {
			startY = y
		}
	}

	curY := startY
	for i, li := range allLines {
		if i > 0 {
			curY += li.spaceBefore
		}

		lh := li.line.lineHeight
		if li.lineSpacing < 0 {
			lh = int(float64(lh) * float64(-li.lineSpacing) / 100000.0)
		} else if li.lineSpacing > 0 {
			lh = r.hundredthPtToPixelY(li.lineSpacing)
		}

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

		baseline := curY + li.line.ascent

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
				if run.font.Superscript {
					runBaseline -= li.line.ascent / 3
				} else if run.font.Subscript {
					runBaseline += li.line.descent / 2
				}
			}

			d := &font.Drawer{
				Dst:  r.img,
				Src:  image.NewUniform(fc),
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
					Src:  image.NewUniform(fc),
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

// buildBulletRun creates a textRun for a bullet prefix.
func (r *renderer) buildBulletRun(b *Bullet, para *Paragraph) textRun {
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
	if b.Color != nil {
		bulletFont.Color = *b.Color
	}
	if b.Font != "" {
		bulletFont.Name = b.Font
	}

	var text string
	switch b.Type {
	case BulletTypeChar:
		text = b.Style + " "
	case BulletTypeNumeric, BulletTypeAutoNum:
		num := b.StartAt
		if num < 1 {
			num = 1
		}
		text = formatBulletNumber(num, b.NumFormat) + " "
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
	return textRun{
		text:  text,
		font:  bulletFont,
		face:  face,
		width: w,
	}
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
func isCJK(r rune) bool {
	return unicode.Is(unicode.Han, r) ||
		unicode.Is(unicode.Hangul, r) ||
		unicode.Is(unicode.Hiragana, r) ||
		unicode.Is(unicode.Katakana, r) ||
		(r >= 0x3000 && r <= 0x303F) || // CJK Symbols and Punctuation
		(r >= 0xFF00 && r <= 0xFFEF) // Fullwidth Forms
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
		'(', '[', '{', '<':
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
		segments = merged
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

	maxW26_6 := fixed.I(maxWidth)
	// Add a small tolerance (~1%) to account for differences between Go's
	// text measurement and PowerPoint's DirectWrite renderer. Go's opentype
	// package doesn't apply the same GPOS/GSUB shaping as DirectWrite,
	// causing Latin text segments to measure slightly wider.
	maxW26_6 += maxW26_6 / 100

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

		runW := measureStringWithKern(run.face, run.text)

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
			tw := measureStringWithKern(run.face, test)
			if currentWidth+tw > maxW26_6 && (len(currentRuns) > 0 || partial.Len() > 0) {
				if partial.Len() > 0 {
					pText := partial.String()
					currentRuns = append(currentRuns, textRun{
						text:  pText,
						font:  run.font,
						face:  run.face,
						width: measureStringWithKern(run.face, pText).Ceil(),
					})
				}
				lines = append(lines, r.buildTextLine(currentRuns))
				currentRuns = nil
				currentWidth = 0
				partial.Reset()
				partial.WriteString(seg)
			} else {
				partial.WriteString(seg)
			}
		}
		if partial.Len() > 0 {
			pText := partial.String()
			pw := measureStringWithKern(run.face, pText)
			wr := textRun{
				text:  pText,
				font:  run.font,
				face:  run.face,
				width: pw.Ceil(),
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
		Src:  image.NewUniform(c),
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

func (r *renderer) renderChart(s *ChartShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	// Background
	r.fillRectFast(image.Rect(x, y, x+w, y+h), color.RGBA{R: 255, G: 255, B: 255, A: 255})
	r.drawRect(image.Rect(x, y, x+w, y+h), color.RGBA{R: 200, G: 200, B: 200, A: 255}, 1)

	// Title
	titleH := 0
	if s.title != nil && s.title.Visible && s.title.Text != "" {
		face := r.getFace(s.title.Font)
		fc := argbToRGBA(s.title.Font.Color)
		titleH = face.Metrics().Height.Ceil() + 4
		r.drawStringCentered(s.title.Text, face, fc, image.Rect(x, y, x+w, y+titleH))
	}

	// Legend height
	legendH := 0
	if s.legend != nil && s.legend.Visible {
		legendH = 20
	}

	// Plot area
	plotX := x + 40
	plotY := y + titleH + 5
	plotW := w - 50
	plotH := h - titleH - legendH - 15
	if plotW < 10 {
		plotW = 10
	}
	if plotH < 10 {
		plotH = 10
	}

	ct := s.plotArea.GetType()
	if ct == nil {
		return
	}

	switch c := ct.(type) {
	case *BarChart:
		r.renderBarChart(c, plotX, plotY, plotW, plotH)
	case *Bar3DChart:
		r.renderBarChart(&c.BarChart, plotX, plotY, plotW, plotH)
	case *LineChart:
		r.renderLineChart(c, plotX, plotY, plotW, plotH)
	case *PieChart:
		r.renderPieChart(c.Series, plotX, plotY, plotW, plotH)
	case *Pie3DChart:
		r.renderPieChart(c.Series, plotX, plotY, plotW, plotH)
	case *DoughnutChart:
		r.renderDoughnutChart(c, plotX, plotY, plotW, plotH)
	case *AreaChart:
		r.renderAreaChart(c, plotX, plotY, plotW, plotH)
	case *ScatterChart:
		r.renderScatterChart(c, plotX, plotY, plotW, plotH)
	case *RadarChart:
		r.renderRadarChart(c, plotX, plotY, plotW, plotH)
	}

	// Legend
	if s.legend != nil && s.legend.Visible {
		r.renderChartLegend(s, x, y+h-legendH, w, legendH)
	}
}

func (r *renderer) renderBarChart(c *BarChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// Collect all categories and find value range
	cats := c.Series[0].Categories
	minVal := 0.0
	maxVal := 0.0
	first := true
	for _, s := range c.Series {
		for _, cat := range s.Categories {
			v := s.Values[cat]
			if first {
				minVal = v
				maxVal = v
				first = false
			} else {
				if v < minVal {
					minVal = v
				}
				if v > maxVal {
					maxVal = v
				}
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	// Draw axes
	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	nCats := len(cats)
	nSeries := len(c.Series)
	if nCats == 0 {
		return
	}
	catW := pw / nCats
	barW := catW / (nSeries + 1)
	if barW < 1 {
		barW = 1
	}

	for ci, cat := range cats {
		for si, s := range c.Series {
			v := s.Values[cat]
			barH := int(float64(ph) * (v - minVal) / valRange)
			bx := px + ci*catW + (si+1)*barW - barW/2
			by := py + ph - barH
			sc := getSeriesColor(s, si, palette)
			r.fillRectBlend(image.Rect(bx, by, bx+barW-1, py+ph), sc)
		}
	}
}

func (r *renderer) renderLineChart(c *LineChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// Find value range
	minVal := math.MaxFloat64
	maxVal := -math.MaxFloat64
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v < minVal {
				minVal = v
			}
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	// Draw axes
	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}
		prevX, prevY := 0, 0
		for i, cat := range cats {
			v := s.Values[cat]
			ptX := px
			if nPts > 1 {
				ptX = px + i*pw/(nPts-1)
			}
			ptY := py + ph - int(float64(ph)*(v-minVal)/valRange)
			if i > 0 {
				r.drawLineAA(prevX, prevY, ptX, ptY, sc, 2)
			}
			// Draw marker
			r.fillEllipseAA(ptX-2, ptY-2, 5, 5, sc)
			prevX, prevY = ptX, ptY
		}
	}
}

func (r *renderer) renderPieChart(series []*ChartSeries, px, py, pw, ph int) {
	if len(series) == 0 || len(series[0].Categories) == 0 {
		return
	}
	palette := chartColors()
	s := series[0]

	// Sum values
	total := 0.0
	for _, cat := range s.Categories {
		v := s.Values[cat]
		if v > 0 {
			total += v
		}
	}
	if total == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2
	radius := minInt(pw, ph) / 2
	if radius < 5 {
		return
	}

	startAngle := -math.Pi / 2
	for i, cat := range s.Categories {
		v := s.Values[cat]
		if v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		endAngle := startAngle + sweep
		sc := palette[i%len(palette)]
		r.fillPieSlice(cx, cy, radius, startAngle, endAngle, sc)
		startAngle = endAngle
	}
}

// fillPieSlice fills a pie slice using scanline approach with row-level x-range.
func (r *renderer) fillPieSlice(cx, cy, radius int, startAngle, endAngle float64, c color.RGBA) {
	r2 := radius * radius
	for dy := -radius; dy <= radius; dy++ {
		dy2 := dy * dy
		if dy2 > r2 {
			continue
		}
		// Compute max dx for this row
		maxDx := int(math.Sqrt(float64(r2 - dy2)))
		for dx := -maxDx; dx <= maxDx; dx++ {
			angle := math.Atan2(float64(dy), float64(dx))
			if angleInSweep(angle, startAngle, endAngle) {
				r.blendPixel(cx+dx, cy+dy, c)
			}
		}
	}
}

// angleInSweep checks if angle is within the sweep from start to end (going clockwise).
func angleInSweep(angle, start, end float64) bool {
	// Normalize to [0, 2*pi)
	norm := func(a float64) float64 {
		for a < 0 {
			a += 2 * math.Pi
		}
		for a >= 2*math.Pi {
			a -= 2 * math.Pi
		}
		return a
	}
	a := norm(angle)
	s := norm(start)
	e := norm(end)
	if s <= e {
		return a >= s && a <= e
	}
	return a >= s || a <= e
}

func (r *renderer) renderDoughnutChart(c *DoughnutChart, px, py, pw, ph int) {
	if len(c.Series) == 0 || len(c.Series[0].Categories) == 0 {
		return
	}
	palette := chartColors()
	s := c.Series[0]

	total := 0.0
	for _, cat := range s.Categories {
		v := s.Values[cat]
		if v > 0 {
			total += v
		}
	}
	if total == 0 {
		return
	}

	cx := px + pw/2
	cy := py + ph/2
	outerR := minInt(pw, ph) / 2
	innerR := outerR * c.HoleSize / 100
	if outerR < 5 {
		return
	}

	startAngle := -math.Pi / 2
	for i, cat := range s.Categories {
		v := s.Values[cat]
		if v <= 0 {
			continue
		}
		sweep := 2 * math.Pi * v / total
		endAngle := startAngle + sweep
		sc := palette[i%len(palette)]
		r.fillDoughnutSlice(cx, cy, innerR, outerR, startAngle, endAngle, sc)
		startAngle = endAngle
	}
}

// fillDoughnutSlice fills a doughnut slice.
func (r *renderer) fillDoughnutSlice(cx, cy, innerR, outerR int, startAngle, endAngle float64, c color.RGBA) {
	or2 := outerR * outerR
	ir2 := innerR * innerR
	for dy := -outerR; dy <= outerR; dy++ {
		dy2 := dy * dy
		if dy2 > or2 {
			continue
		}
		maxDx := int(math.Sqrt(float64(or2 - dy2)))
		for dx := -maxDx; dx <= maxDx; dx++ {
			d2 := dx*dx + dy2
			if d2 < ir2 {
				continue
			}
			angle := math.Atan2(float64(dy), float64(dx))
			if angleInSweep(angle, startAngle, endAngle) {
				r.blendPixel(cx+dx, cy+dy, c)
			}
		}
	}
}

func (r *renderer) renderAreaChart(c *AreaChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	minVal := math.MaxFloat64
	maxVal := -math.MaxFloat64
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v < minVal {
				minVal = v
			}
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	// Axes
	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		// Semi-transparent fill
		fillC := color.RGBA{R: sc.R, G: sc.G, B: sc.B, A: 128}
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}

		pts := make([]fpoint, 0, nPts+2)
		for i, cat := range cats {
			v := s.Values[cat]
			ptX := float64(px)
			if nPts > 1 {
				ptX = float64(px) + float64(i)*float64(pw)/float64(nPts-1)
			}
			ptY := float64(py+ph) - float64(ph)*(v-minVal)/valRange
			pts = append(pts, fpoint{ptX, ptY})
		}
		// Close polygon along baseline
		pts = append(pts, fpoint{pts[len(pts)-1].x, float64(py + ph)})
		pts = append(pts, fpoint{pts[0].x, float64(py + ph)})
		r.fillPolygon(pts, fillC)

		// Draw line on top
		for i := 0; i < nPts-1; i++ {
			r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[i+1].x), int(pts[i+1].y), sc, 2)
		}
	}
}

func (r *renderer) renderScatterChart(c *ScatterChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// For scatter, categories are X values (parsed as indices), values are Y
	minVal := math.MaxFloat64
	maxVal := -math.MaxFloat64
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v < minVal {
				minVal = v
			}
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if minVal > 0 {
		minVal = 0
	}
	if maxVal <= minVal {
		maxVal = minVal + 1
	}
	valRange := maxVal - minVal

	axisColor := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	r.drawLine(px, py+ph, px+pw, py+ph, axisColor)
	r.drawLine(px, py, px, py+ph, axisColor)

	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}
		for i, cat := range cats {
			v := s.Values[cat]
			ptX := px + (i * pw / maxInt(nPts-1, 1))
			ptY := py + ph - int(float64(ph)*(v-minVal)/valRange)
			r.fillEllipseAA(ptX-3, ptY-3, 7, 7, sc)
		}
	}
}

func (r *renderer) renderRadarChart(c *RadarChart, px, py, pw, ph int) {
	if len(c.Series) == 0 {
		return
	}
	palette := chartColors()

	// Find max value
	maxVal := 0.0
	for _, s := range c.Series {
		for _, v := range s.Values {
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if maxVal == 0 {
		maxVal = 1
	}

	cx := px + pw/2
	cy := py + ph/2
	radius := minInt(pw, ph) / 2

	// Draw radar grid
	gridColor := color.RGBA{R: 200, G: 200, B: 200, A: 255}
	nCats := len(c.Series[0].Categories)
	if nCats == 0 {
		return
	}
	for i := 0; i < nCats; i++ {
		angle := 2*math.Pi*float64(i)/float64(nCats) - math.Pi/2
		ex := cx + int(float64(radius)*math.Cos(angle))
		ey := cy + int(float64(radius)*math.Sin(angle))
		r.drawLine(cx, cy, ex, ey, gridColor)
	}

	// Draw series
	for si, s := range c.Series {
		sc := getSeriesColor(s, si, palette)
		cats := s.Categories
		nPts := len(cats)
		if nPts == 0 {
			continue
		}
		pts := make([]fpoint, nPts)
		for i, cat := range cats {
			v := s.Values[cat]
			angle := 2*math.Pi*float64(i)/float64(nPts) - math.Pi/2
			dist := float64(radius) * v / maxVal
			pts[i] = fpoint{
				x: float64(cx) + dist*math.Cos(angle),
				y: float64(cy) + dist*math.Sin(angle),
			}
		}
		// Draw polygon
		for i := 0; i < nPts; i++ {
			j := (i + 1) % nPts
			r.drawLineAA(int(pts[i].x), int(pts[i].y), int(pts[j].x), int(pts[j].y), sc, 2)
		}
		// Fill with semi-transparent
		fillC := color.RGBA{R: sc.R, G: sc.G, B: sc.B, A: 64}
		r.fillPolygon(pts, fillC)
	}
}

func (r *renderer) renderChartLegend(s *ChartShape, lx, ly, lw, lh int) {
	ct := s.plotArea.GetType()
	if ct == nil {
		return
	}
	palette := chartColors()
	face := r.getFace(s.legend.Font)

	var names []string
	var colors []color.RGBA

	switch c := ct.(type) {
	case *BarChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *Bar3DChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *LineChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *PieChart:
		if len(c.Series) > 0 {
			for i, cat := range c.Series[0].Categories {
				names = append(names, cat)
				colors = append(colors, palette[i%len(palette)])
			}
		}
	case *Pie3DChart:
		if len(c.Series) > 0 {
			for i, cat := range c.Series[0].Categories {
				names = append(names, cat)
				colors = append(colors, palette[i%len(palette)])
			}
		}
	case *DoughnutChart:
		if len(c.Series) > 0 {
			for i, cat := range c.Series[0].Categories {
				names = append(names, cat)
				colors = append(colors, palette[i%len(palette)])
			}
		}
	case *AreaChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *ScatterChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	case *RadarChart:
		for i, ser := range c.Series {
			names = append(names, ser.Title)
			colors = append(colors, getSeriesColor(ser, i, palette))
		}
	}

	if len(names) == 0 {
		return
	}

	// Draw legend entries horizontally centered
	entryW := lw / len(names)
	for i, name := range names {
		ex := lx + i*entryW
		// Color box
		boxSize := 10
		bx := ex + 4
		by := ly + (lh-boxSize)/2
		r.fillRectFast(image.Rect(bx, by, bx+boxSize, by+boxSize), colors[i])
		// Text
		d := &font.Drawer{
			Dst:  r.img,
			Src:  image.NewUniform(color.RGBA{A: 255}),
			Face: face,
			Dot:  fixed.P(bx+boxSize+4, ly+lh/2+4),
		}
		d.DrawString(name)
	}
}

// --- Image scaling ---

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
	return scaleImageBilinear(src, dstW, dstH)
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

	// Render at a higher resolution for quality (4x the WMF logical units)
	scale := 4
	imgW := winW * scale
	imgH := winH * scale
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
		dx := d.destX * scale
		dy := d.destY * scale
		dw := d.destW * scale
		dh := d.destH * scale
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
		tx := t.x * scale
		ty := t.y * scale
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
func drawWMFText(canvas *image.RGBA, x, y int, text string, scale int, centerH bool, fc *FontCache) {
	col := color.Black
	// Try to use a proper font that supports Chinese characters
	var face font.Face
	if fc != nil {
		// Try common Chinese fonts at a size proportional to the scale
		fontSize := float64(10 * scale)
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
