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
	"strings"

	"golang.org/x/image/font"
	"golang.org/x/image/font/basicfont"
	"golang.org/x/image/math/fixed"
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

	// Calculate image dimensions from slide aspect ratio
	slideW := float64(layout.CX)
	slideH := float64(layout.CY)
	imgW := opts.Width
	imgH := int(float64(imgW) * slideH / slideW)

	scaleX := float64(imgW) / slideW
	scaleY := float64(imgH) / slideH

	img := image.NewRGBA(image.Rect(0, 0, imgW, imgH))

	// Fill background
	bgColor := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	if opts.BackgroundColor != nil {
		bgColor = *opts.BackgroundColor
	} else if slide.background != nil && slide.background.Type == FillSolid {
		bgColor = argbToRGBA(slide.background.Color)
	}
	draw.Draw(img, img.Bounds(), &image.Uniform{bgColor}, image.Point{}, draw.Src)

	r := &renderer{
		img:       img,
		scaleX:    scaleX,
		scaleY:    scaleY,
		fontCache: opts.FontCache,
		dpi:       opts.DPI,
	}
	if r.fontCache == nil {
		r.fontCache = NewFontCache(opts.FontDirs...)
	}
	if r.dpi <= 0 {
		r.dpi = 96
	}

	// Render shapes in order
	for _, shape := range slide.shapes {
		r.renderShape(shape)
	}

	return img, nil
}


// SlidesToImages renders all slides to images.
func (p *Presentation) SlidesToImages(opts *RenderOptions) ([]image.Image, error) {
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
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return fmt.Errorf("create directory: %w", err)
		}
	}

	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create file: %w", err)
	}
	defer f.Close()

	switch opts.Format {
	case ImageFormatJPEG:
		quality := opts.JPEGQuality
		if quality <= 0 || quality > 100 {
			quality = 90
		}
		return jpeg.Encode(f, img, &jpeg.Options{Quality: quality})
	default:
		return png.Encode(f, img)
	}
}

// --- renderer ---

type renderer struct {
	img       *image.RGBA
	scaleX    float64
	scaleY    float64
	fontCache *FontCache
	dpi       float64
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
	case *GroupShape:
		for _, gs := range s.shapes {
			r.renderShape(gs)
		}
	}
}

func (r *renderer) emuToPixelX(emu int64) int {
	return int(float64(emu) * r.scaleX)
}

func (r *renderer) emuToPixelY(emu int64) int {
	return int(float64(emu) * r.scaleY)
}

func argbToRGBA(c Color) color.RGBA {
	return color.RGBA{
		R: c.GetRed(),
		G: c.GetGreen(),
		B: c.GetBlue(),
		A: c.GetAlpha(),
	}
}


// --- Shape rendering ---

func (r *renderer) renderRichText(s *RichTextShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	rect := image.Rect(x, y, x+w, y+h)

	// Fill background if set
	if s.fill != nil && s.fill.Type == FillSolid {
		fillColor := argbToRGBA(s.fill.Color)
		draw.Draw(r.img, rect, &image.Uniform{fillColor}, image.Point{}, draw.Over)
	}

	// Draw border if set
	if s.border != nil && s.border.Style != BorderNone {
		borderColor := argbToRGBA(s.border.Color)
		bw := s.border.Width
		if bw <= 0 {
			bw = 1
		}
		pw := int(float64(bw) * r.scaleX)
		if pw < 1 {
			pw = 1
		}
		r.drawRect(rect, borderColor, pw)
	}

	// Draw text
	r.drawParagraphs(s.paragraphs, x, y, w, h)
}

func (r *renderer) renderDrawing(s *DrawingShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	if len(s.data) == 0 {
		return
	}

	// Decode image data
	reader := bytes.NewReader(s.data)
	srcImg, _, err := image.Decode(reader)
	if err != nil {
		// If we can't decode, just draw a placeholder rectangle
		rect := image.Rect(x, y, x+w, y+h)
		r.drawRect(rect, color.RGBA{R: 200, G: 200, B: 200, A: 255}, 1)
		return
	}

	// Scale and draw the image
	dstRect := image.Rect(x, y, x+w, y+h)
	scaledImg := scaleImage(srcImg, w, h)
	draw.Draw(r.img, dstRect, scaledImg, image.Point{}, draw.Over)
}

func (r *renderer) renderAutoShape(s *AutoShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)
	rect := image.Rect(x, y, x+w, y+h)

	// Fill
	if s.fill != nil && s.fill.Type == FillSolid {
		fillColor := argbToRGBA(s.fill.Color)
		switch s.shapeType {
		case AutoShapeEllipse:
			r.fillEllipse(x, y, w, h, fillColor)
		default:
			draw.Draw(r.img, rect, &image.Uniform{fillColor}, image.Point{}, draw.Over)
		}
	}

	// Border
	if s.border != nil && s.border.Style != BorderNone {
		borderColor := argbToRGBA(s.border.Color)
		bw := s.border.Width
		if bw <= 0 {
			bw = 1
		}
		pw := int(float64(bw) * r.scaleX)
		if pw < 1 {
			pw = 1
		}
		switch s.shapeType {
		case AutoShapeEllipse:
			r.drawEllipse(x, y, w, h, borderColor)
		default:
			r.drawRect(rect, borderColor, pw)
		}
	}

	// Text inside auto shape
	if s.text != "" {
		f := NewFont()
		face := r.getFace(f)
		textColor := color.RGBA{R: 0, G: 0, B: 0, A: 255}
		r.drawStringCentered(s.text, face, textColor, rect)
	}
}

func (r *renderer) renderLine(s *LineShape) {
	x1 := r.emuToPixelX(s.offsetX)
	y1 := r.emuToPixelY(s.offsetY)
	x2 := r.emuToPixelX(s.offsetX + s.width)
	y2 := r.emuToPixelY(s.offsetY + s.height)
	lineColor := argbToRGBA(s.lineColor)
	r.drawLine(x1, y1, x2, y2, lineColor)
}

func (r *renderer) renderTable(s *TableShape) {
	x := r.emuToPixelX(s.offsetX)
	y := r.emuToPixelY(s.offsetY)
	w := r.emuToPixelX(s.width)
	h := r.emuToPixelY(s.height)

	if s.numRows == 0 || s.numCols == 0 {
		return
	}

	cellW := w / s.numCols
	cellH := h / s.numRows

	for row := 0; row < s.numRows; row++ {
		for col := 0; col < s.numCols; col++ {
			cx := x + col*cellW
			cy := y + row*cellH
			cellRect := image.Rect(cx, cy, cx+cellW, cy+cellH)

			cell := s.rows[row][col]

			// Cell fill
			if cell.fill != nil && cell.fill.Type == FillSolid {
				fillColor := argbToRGBA(cell.fill.Color)
				draw.Draw(r.img, cellRect, &image.Uniform{fillColor}, image.Point{}, draw.Over)
			}

			// Cell border
			r.drawRect(cellRect, color.RGBA{R: 0, G: 0, B: 0, A: 255}, 1)

			// Cell text
			r.drawParagraphs(cell.paragraphs, cx+2, cy+2, cellW-4, cellH-4)
		}
	}
}


// --- Drawing primitives ---

func (r *renderer) drawRect(rect image.Rectangle, c color.RGBA, width int) {
	for i := 0; i < width; i++ {
		// Top
		for x := rect.Min.X; x < rect.Max.X; x++ {
			r.setPixel(x, rect.Min.Y+i, c)
		}
		// Bottom
		for x := rect.Min.X; x < rect.Max.X; x++ {
			r.setPixel(x, rect.Max.Y-1-i, c)
		}
		// Left
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			r.setPixel(rect.Min.X+i, y, c)
		}
		// Right
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			r.setPixel(rect.Max.X-1-i, y, c)
		}
	}
}

func (r *renderer) drawLine(x1, y1, x2, y2 int, c color.RGBA) {
	// Bresenham's line algorithm
	dx := abs(x2 - x1)
	dy := abs(y2 - y1)
	sx := 1
	if x1 > x2 {
		sx = -1
	}
	sy := 1
	if y1 > y2 {
		sy = -1
	}
	err := dx - dy

	for {
		r.setPixel(x1, y1, c)
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

func (r *renderer) fillEllipse(cx, cy, w, h int, c color.RGBA) {
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry

	for py := cy; py < cy+h; py++ {
		for px := cx; px < cx+w; px++ {
			dx := (float64(px) + 0.5 - centerX) / rx
			dy := (float64(py) + 0.5 - centerY) / ry
			if dx*dx+dy*dy <= 1.0 {
				r.setPixel(px, py, c)
			}
		}
	}
}

func (r *renderer) drawEllipse(cx, cy, w, h int, c color.RGBA) {
	rx := float64(w) / 2
	ry := float64(h) / 2
	centerX := float64(cx) + rx
	centerY := float64(cy) + ry

	steps := int(math.Max(float64(w), float64(h)) * 4)
	if steps < 100 {
		steps = 100
	}
	for i := 0; i < steps; i++ {
		angle := 2 * math.Pi * float64(i) / float64(steps)
		px := int(centerX + rx*math.Cos(angle))
		py := int(centerY + ry*math.Sin(angle))
		r.setPixel(px, py, c)
	}
}

func (r *renderer) setPixel(x, y int, c color.RGBA) {
	bounds := r.img.Bounds()
	if x >= bounds.Min.X && x < bounds.Max.X && y >= bounds.Min.Y && y < bounds.Max.Y {
		r.img.SetRGBA(x, y, c)
	}
}

func abs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}


// --- Text rendering ---

// getFace returns a TrueType font.Face for the given Font, falling back to basicfont.
func (r *renderer) getFace(f *Font) font.Face {
	if f == nil {
		f = NewFont()
	}
	sizePt := float64(f.Size)
	if sizePt <= 0 {
		sizePt = 10
	}
	// Scale font size to match the rendered image resolution.
	// sizePt is the design-time point size. We scale by scaleY (EMU->pixel ratio)
	// and convert pt->px via DPI/72.
	scaledPt := sizePt * r.scaleY * r.dpi / 72.0

	name := f.Name
	if name == "" {
		name = "Calibri"
	}

	face := r.fontCache.GetFace(name, scaledPt, f.Bold, f.Italic)
	if face != nil {
		return face
	}

	// Try common fallbacks
	for _, fallback := range []string{"arial", "helvetica", "dejavu sans", "liberation sans", "noto sans"} {
		face = r.fontCache.GetFace(fallback, scaledPt, f.Bold, f.Italic)
		if face != nil {
			return face
		}
	}

	return basicfont.Face7x13
}

// textRun holds rendering info for a single text run.
type textRun struct {
	text  string
	face  font.Face
	color color.RGBA
}

// textLine holds a wrapped line of text runs.
type textLine struct {
	runs      []textRun
	width     int
	height    int
	alignment HorizontalAlignment
}

func buildTextLine(runs []textRun, align HorizontalAlignment) textLine {
	totalW := 0
	maxH := 0
	for _, r := range runs {
		totalW += font.MeasureString(r.face, r.text).Ceil()
		h := r.face.Metrics().Height.Ceil()
		if h > maxH {
			maxH = h
		}
	}
	if maxH <= 0 {
		maxH = 14
	}
	return textLine{runs: runs, width: totalW, height: maxH, alignment: align}
}

func (r *renderer) drawParagraphs(paragraphs []*Paragraph, x, y, w, h int) {
	var allLines []textLine

	for _, para := range paragraphs {
		align := HorizontalLeft
		if para.alignment != nil {
			align = para.alignment.Horizontal
		}

		var runs []textRun
		for _, elem := range para.elements {
			switch e := elem.(type) {
			case *TextRun:
				face := r.getFace(e.font)
				tc := color.RGBA{R: 0, G: 0, B: 0, A: 255}
				if e.font != nil {
					tc = argbToRGBA(e.font.Color)
				}
				runs = append(runs, textRun{text: e.text, face: face, color: tc})
			case *BreakElement:
				if len(runs) > 0 {
					allLines = append(allLines, buildTextLine(runs, align))
					runs = nil
				} else {
					allLines = append(allLines, textLine{height: 14, alignment: align})
				}
			}
		}
		if len(runs) > 0 {
			allLines = append(allLines, buildTextLine(runs, align))
		} else if len(para.elements) == 0 {
			allLines = append(allLines, textLine{height: 14, alignment: align})
		}
	}

	// Word-wrap lines that exceed width
	var wrappedLines []textLine
	for _, line := range allLines {
		if line.width <= w || w <= 0 || len(line.runs) == 0 {
			wrappedLines = append(wrappedLines, line)
			continue
		}
		wrappedLines = append(wrappedLines, wrapRunLine(line, w)...)
	}

	// Draw
	curY := y
	for _, line := range wrappedLines {
		curY += line.height
		if curY > y+h {
			break
		}

		drawX := x
		switch line.alignment {
		case HorizontalCenter:
			drawX = x + (w-line.width)/2
		case HorizontalRight:
			drawX = x + w - line.width
		}

		for _, run := range line.runs {
			d := &font.Drawer{
				Dst:  r.img,
				Src:  &image.Uniform{run.color},
				Face: run.face,
				Dot:  fixed.P(drawX, curY),
			}
			d.DrawString(run.text)
			drawX += font.MeasureString(run.face, run.text).Ceil()
		}
	}
}

// wrapRunLine wraps a textLine into multiple lines that fit within maxWidth.
func wrapRunLine(line textLine, maxWidth int) []textLine {
	// Flatten all runs into words with their formatting
	type styledWord struct {
		word  string
		face  font.Face
		color color.RGBA
	}

	var words []styledWord
	for _, run := range line.runs {
		for i, w := range strings.Fields(run.text) {
			if i > 0 {
				w = " " + w
			}
			words = append(words, styledWord{word: w, face: run.face, color: run.color})
		}
	}

	if len(words) == 0 {
		return []textLine{line}
	}

	var result []textLine
	var curRuns []textRun
	curWidth := 0

	for _, sw := range words {
		ww := font.MeasureString(sw.face, sw.word).Ceil()
		if curWidth+ww > maxWidth && curWidth > 0 {
			result = append(result, buildTextLine(curRuns, line.alignment))
			curRuns = nil
			curWidth = 0
			// Trim leading space on wrapped word
			sw.word = strings.TrimLeft(sw.word, " ")
			ww = font.MeasureString(sw.face, sw.word).Ceil()
		}
		curRuns = append(curRuns, textRun{text: sw.word, face: sw.face, color: sw.color})
		curWidth += ww
	}
	if len(curRuns) > 0 {
		result = append(result, buildTextLine(curRuns, line.alignment))
	}
	return result
}

func (r *renderer) drawStringCentered(text string, face font.Face, c color.RGBA, rect image.Rectangle) {
	textW := font.MeasureString(face, text).Ceil()
	lineH := face.Metrics().Height.Ceil()
	x := rect.Min.X + (rect.Dx()-textW)/2
	y := rect.Min.Y + (rect.Dy()+lineH)/2

	d := &font.Drawer{
		Dst:  r.img,
		Src:  &image.Uniform{c},
		Face: face,
		Dot:  fixed.P(x, y),
	}
	d.DrawString(text)
}

// scaleImage scales an image to the target width and height using nearest-neighbor.
func scaleImage(src image.Image, dstW, dstH int) image.Image {
	if dstW <= 0 || dstH <= 0 {
		return src
	}
	srcBounds := src.Bounds()
	srcW := srcBounds.Dx()
	srcH := srcBounds.Dy()
	if srcW == 0 || srcH == 0 {
		return src
	}

	dst := image.NewRGBA(image.Rect(0, 0, dstW, dstH))
	for y := 0; y < dstH; y++ {
		for x := 0; x < dstW; x++ {
			srcX := srcBounds.Min.X + x*srcW/dstW
			srcY := srcBounds.Min.Y + y*srcH/dstH
			dst.Set(x, y, src.At(srcX, srcY))
		}
	}
	return dst
}
