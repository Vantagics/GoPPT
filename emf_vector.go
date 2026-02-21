package gopresentation

import (
	"image"
	"image/color"
	"math"
	"sort"
)

// renderEMFVector renders an EMF containing vector drawing commands to an image.
func renderEMFVector(data []byte) image.Image {
	if len(data) < 88 {
		return nil
	}
	u32 := func(d []byte) uint32 {
		return uint32(d[0]) | uint32(d[1])<<8 | uint32(d[2])<<16 | uint32(d[3])<<24
	}
	i32 := func(d []byte) int32 { return int32(u32(d)) }
	i16 := func(d []byte) int16 { return int16(uint16(d[0]) | uint16(d[1])<<8) }

	boundsL := int(i32(data[8:12]))
	boundsT := int(i32(data[12:16]))
	boundsR := int(i32(data[16:20]))
	boundsB := int(i32(data[20:24]))

	// First pass: read transforms and identify "clip" fills to skip.
	// Pattern: FILLPATH → CLOSEFIGURE → ABORTPATH means the fill is a clip op.
	winOrgX, winOrgY := 0, 0
	winExtX, winExtY := 1, 1
	vpOrgX, vpOrgY := 0, 0
	vpExtX, vpExtY := 1, 1
	skipFills := map[int]bool{} // file offsets of FILLPATH records to skip

	pos := 0
	for pos+8 <= len(data) {
		rt := u32(data[pos : pos+4])
		rs := u32(data[pos+4 : pos+8])
		if rs < 8 || pos+int(rs) > len(data) {
			break
		}
		rec := data[pos : pos+int(rs)]
		switch rt {
		case 0x09:
			if len(rec) >= 16 {
				winExtX = int(i32(rec[8:12]))
				winExtY = int(i32(rec[12:16]))
			}
		case 0x0A:
			if len(rec) >= 16 {
				winOrgX = int(i32(rec[8:12]))
				winOrgY = int(i32(rec[12:16]))
			}
		case 0x0B:
			if len(rec) >= 16 {
				vpExtX = int(i32(rec[8:12]))
				vpExtY = int(i32(rec[12:16]))
			}
		case 0x0C:
			if len(rec) >= 16 {
				vpOrgX = int(i32(rec[8:12]))
				vpOrgY = int(i32(rec[12:16]))
			}
		case 0x3D: // FILLPATH - check if followed by CLOSEFIGURE + ABORTPATH
			fillPos := pos
			np := pos + int(rs)
			if np+8 <= len(data) {
				nt := u32(data[np : np+4])
				ns := u32(data[np+4 : np+8])
				if nt == 0x3C && ns >= 8 { // CLOSEFIGURE
					np2 := np + int(ns)
					if np2+8 <= len(data) {
						nt2 := u32(data[np2 : np2+4])
						if nt2 == 0x40 { // ABORTPATH
							skipFills[fillPos] = true
						}
					}
				}
			}
		}
		if rt == 0x0E {
			break
		}
		pos += int(rs)
	}

	devW := boundsR - boundsL
	devH := boundsB - boundsT
	if devW <= 0 || devH <= 0 {
		return nil
	}

	scale := 1.0
	target := 300.0
	if float64(devW) < target || float64(devH) < target {
		sx := target / float64(devW)
		sy := target / float64(devH)
		if sx < sy {
			scale = sx
		} else {
			scale = sy
		}
	}
	imgW := int(float64(devW)*scale) + 2
	imgH := int(float64(devH)*scale) + 2
	if imgW > 2000 {
		imgW = 2000
	}
	if imgH > 2000 {
		imgH = 2000
	}

	img := image.NewRGBA(image.Rect(0, 0, imgW, imgH))
	offX := float64(-boundsL)*scale + 1
	offY := float64(-boundsT)*scale + 1

	toImg := func(lx, ly int) (float64, float64) {
		var dx, dy float64
		if winExtX != 0 {
			dx = float64(lx-winOrgX) * float64(vpExtX) / float64(winExtX)
		}
		if winExtY != 0 {
			dy = float64(ly-winOrgY) * float64(vpExtY) / float64(winExtY)
		}
		dx += float64(vpOrgX)
		dy += float64(vpOrgY)
		return dx*scale + offX, dy*scale + offY
	}

	type emfBrush struct {
		style   uint32
		r, g, b uint8
	}
	type emfPen struct {
		style   uint32
		r, g, b uint8
	}
	brushes := map[uint32]emfBrush{}
	pens := map[uint32]emfPen{}
	var curBrush emfBrush
	var curPen emfPen
	nullBrush := false
	nullPen := true
	type pp struct{ x, y float64 }
	var path []pp
	var lastPath []pp
	var curX, curY float64
	hasDrawing := false

	setPixel := func(x, y int, c color.RGBA) {
		if x >= 0 && x < imgW && y >= 0 && y < imgH {
			off := y*img.Stride + x*4
			img.Pix[off] = c.R
			img.Pix[off+1] = c.G
			img.Pix[off+2] = c.B
			img.Pix[off+3] = c.A
		}
	}

	emfFill := func(pts []pp, c color.RGBA) {
		if len(pts) < 3 || c.A == 0 {
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
		xs := make([]float64, 0, n)
		for y := int(minY); y <= int(maxY); y++ {
			if y < 0 || y >= imgH {
				continue
			}
			fy := float64(y) + 0.5
			xs = xs[:0]
			for i := 0; i < n; i++ {
				j := (i + 1) % n
				y1, y2 := pts[i].y, pts[j].y
				if y1 > y2 {
					y1, y2 = y2, y1
				}
				if fy < y1 || fy >= y2 {
					continue
				}
				dy := pts[j].y - pts[i].y
				if dy == 0 {
					continue
				}
				t := (fy - pts[i].y) / dy
				xs = append(xs, pts[i].x+t*(pts[j].x-pts[i].x))
			}
			sort.Float64s(xs)
			for i := 0; i+1 < len(xs); i += 2 {
				x1 := int(math.Ceil(xs[i]))
				x2 := int(math.Floor(xs[i+1]))
				if x1 < 0 {
					x1 = 0
				}
				if x2 >= imgW {
					x2 = imgW - 1
				}
				for px := x1; px <= x2; px++ {
					setPixel(px, y, c)
				}
			}
		}
	}

	emfLine := func(x0, y0, x1, y1 int, c color.RGBA) {
		dx, dy := x1-x0, y1-y0
		if dx < 0 {
			dx = -dx
		}
		if dy < 0 {
			dy = -dy
		}
		sx, sy := 1, 1
		if x0 > x1 {
			sx = -1
		}
		if y0 > y1 {
			sy = -1
		}
		e := dx - dy
		for {
			setPixel(x0, y0, c)
			if x0 == x1 && y0 == y1 {
				break
			}
			e2 := 2 * e
			if e2 > -dy {
				e -= dy
				x0 += sx
			}
			if e2 < dx {
				e += dx
				y0 += sy
			}
		}
	}

	emfStroke := func(pts []pp, c color.RGBA) {
		for i := 0; i+1 < len(pts); i++ {
			emfLine(int(pts[i].x), int(pts[i].y), int(pts[i+1].x), int(pts[i+1].y), c)
		}
		if len(pts) >= 3 {
			emfLine(int(pts[len(pts)-1].x), int(pts[len(pts)-1].y), int(pts[0].x), int(pts[0].y), c)
		}
	}

	flatBez := func(p0, p1, p2, p3 pp) []pp {
		r := make([]pp, 16)
		for i := 1; i <= 16; i++ {
			t := float64(i) / 16.0
			it := 1 - t
			r[i-1] = pp{
				it*it*it*p0.x + 3*it*it*t*p1.x + 3*it*t*t*p2.x + t*t*t*p3.x,
				it*it*it*p0.y + 3*it*it*t*p1.y + 3*it*t*t*p2.y + t*t*t*p3.y,
			}
		}
		return r
	}

	addBez := func(rec []byte, useCur bool) {
		if len(rec) < 28 {
			return
		}
		cnt := u32(rec[24:28])
		if cnt < 3 || 28+cnt*4 > uint32(len(rec)) {
			return
		}
		pts := make([]pp, cnt)
		for i := uint32(0); i < cnt; i++ {
			pts[i].x, pts[i].y = toImg(int(i16(rec[28+i*4:])), int(i16(rec[28+i*4+2:])))
		}
		if useCur || cnt%3 == 0 {
			for i := 0; i+2 < len(pts); i += 3 {
				flat := flatBez(pp{curX, curY}, pts[i], pts[i+1], pts[i+2])
				path = append(path, flat...)
				curX, curY = flat[len(flat)-1].x, flat[len(flat)-1].y
			}
		} else if len(pts) >= 4 {
			curX, curY = pts[0].x, pts[0].y
			path = append(path, pts[0])
			for i := 1; i+2 < len(pts); i += 3 {
				flat := flatBez(pp{curX, curY}, pts[i], pts[i+1], pts[i+2])
				path = append(path, flat...)
				curX, curY = flat[len(flat)-1].x, flat[len(flat)-1].y
			}
		}
	}

	penColor := func() color.RGBA {
		if nullPen {
			return color.RGBA{}
		}
		return color.RGBA{curPen.r, curPen.g, curPen.b, 255}
	}
	brushColor := func() color.RGBA {
		if nullBrush {
			return color.RGBA{}
		}
		return color.RGBA{curBrush.r, curBrush.g, curBrush.b, 255}
	}

	doFill := func(skip bool) {
		if !skip && len(path) >= 3 {
			bc := brushColor()
			if bc.A > 0 {
				emfFill(path, bc)
				hasDrawing = true
			}
		}
		lastPath = make([]pp, len(path))
		copy(lastPath, path)
		path = path[:0]
	}

	doStrokeFill := func() {
		pts := path
		if len(pts) < 3 {
			pts = lastPath
		}
		if len(pts) >= 3 {
			bc := brushColor()
			if bc.A > 0 {
				emfFill(pts, bc)
				hasDrawing = true
			}
			pc := penColor()
			if pc.A > 0 {
				emfStroke(pts, pc)
				hasDrawing = true
			}
		}
		path = path[:0]
		lastPath = nil
	}

	readPts16 := func(rec []byte) []pp {
		if len(rec) < 28 {
			return nil
		}
		cnt := u32(rec[24:28])
		if 28+cnt*4 > uint32(len(rec)) {
			return nil
		}
		pts := make([]pp, cnt)
		for i := uint32(0); i < cnt; i++ {
			pts[i].x, pts[i].y = toImg(int(i16(rec[28+i*4:])), int(i16(rec[28+i*4+2:])))
		}
		return pts
	}

	// Second pass: render
	pos = 0
	for pos+8 <= len(data) {
		rt := u32(data[pos : pos+4])
		rs := u32(data[pos+4 : pos+8])
		if rs < 8 || pos+int(rs) > len(data) {
			break
		}
		rec := data[pos : pos+int(rs)]
		switch rt {
		case 0x26: // CREATEPEN
			if len(rec) >= 28 {
				ih := u32(rec[8:12])
				pens[ih] = emfPen{u32(rec[12:16]), rec[24], rec[25], rec[26]}
			}
		case 0x27: // CREATEBRUSHINDIRECT
			if len(rec) >= 20 {
				ih := u32(rec[8:12])
				brushes[ih] = emfBrush{u32(rec[12:16]), rec[16], rec[17], rec[18]}
			}
		case 0x28: // DELETEOBJECT
			if len(rec) >= 12 {
				ih := u32(rec[8:12])
				delete(brushes, ih)
				delete(pens, ih)
			}
		case 0x25: // SELECTOBJECT
			if len(rec) >= 12 {
				ih := u32(rec[8:12])
				if ih >= 0x80000000 {
					switch ih {
					case 0x80000000:
						curBrush = emfBrush{0, 255, 255, 255}
						nullBrush = false
					case 0x80000004:
						curBrush = emfBrush{0, 0, 0, 0}
						nullBrush = false
					case 0x80000005:
						nullBrush = true
					case 0x80000006:
						curPen = emfPen{0, 255, 255, 255}
						nullPen = false
					case 0x80000007:
						curPen = emfPen{0, 0, 0, 0}
						nullPen = false
					case 0x80000008:
						nullPen = true
					}
				} else {
					if b, ok := brushes[ih]; ok {
						curBrush = b
						nullBrush = b.style == 1
					}
					if p, ok := pens[ih]; ok {
						curPen = p
						nullPen = p.style == 5
					}
				}
			}
		case 0x1B: // MOVETOEX
			if len(rec) >= 16 {
				curX, curY = toImg(int(i32(rec[8:12])), int(i32(rec[12:16])))
				path = append(path, pp{curX, curY})
			}
		case 0x36: // LINETO
			if len(rec) >= 16 {
				x, y := toImg(int(i32(rec[8:12])), int(i32(rec[12:16])))
				path = append(path, pp{x, y})
				curX, curY = x, y
			}
		case 0x3A: // BEGINPATH
			path = path[:0]
		case 0x3B: // ENDPATH
		case 0x3C: // CLOSEFIGURE
		case 0x3D: // FILLPATH
			doFill(skipFills[pos])
		case 0x3E: // STROKEANDFILLPATH
			doStrokeFill()
		case 0x3F: // STROKEPATH
			if len(path) >= 2 {
				pc := penColor()
				if pc.A > 0 {
					emfStroke(path, pc)
					hasDrawing = true
				}
			}
			path = path[:0]
		case 0x40: // ABORTPATH
			path = path[:0]
		case 0x43: // SELECTCLIPPATH
			path = path[:0]
		case 0x59: // POLYGON16
			pts := readPts16(rec)
			if len(pts) > 0 {
				path = append(path, pts...)
				curX, curY = pts[len(pts)-1].x, pts[len(pts)-1].y
			}
		case 0x58: // POLYBEZIER16
			addBez(rec, false)
		case 0x5B: // POLYBEZIERTO16
			addBez(rec, true)
		case 0x5A: // POLYLINE16
			pts := readPts16(rec)
			for _, p := range pts {
				path = append(path, p)
				curX, curY = p.x, p.y
			}
		case 0x5F: // POLYDRAW16
			if len(rec) >= 28 {
				cnt := u32(rec[24:28])
				if cnt > 0 && 28+cnt*4+cnt <= uint32(len(rec)) {
					ptsOff := 28
					typesOff := 28 + int(cnt)*4
					for i := uint32(0); i < cnt; i++ {
						ix, iy := toImg(int(i16(rec[ptsOff+int(i)*4:])), int(i16(rec[ptsOff+int(i)*4+2:])))
						typ := rec[typesOff+int(i)]
						switch typ & 0x06 {
						case 0x02:
							path = append(path, pp{ix, iy})
							curX, curY = ix, iy
						case 0x04:
							path = append(path, pp{ix, iy})
							curX, curY = ix, iy
						default:
							curX, curY = ix, iy
							path = append(path, pp{ix, iy})
						}
					}
				}
			}
		case 0x2A: // ELLIPSE
			if len(rec) >= 24 {
				l, t, r, b := int(i32(rec[8:12])), int(i32(rec[12:16])), int(i32(rec[16:20])), int(i32(rec[20:24]))
				cx, cy := float64(l+r)/2, float64(t+b)/2
				rx, ry := float64(r-l)/2, float64(b-t)/2
				pts := make([]pp, 32)
				for i := 0; i < 32; i++ {
					a := 2 * math.Pi * float64(i) / 32
					ix, iy := toImg(int(cx+rx*math.Cos(a)), int(cy+ry*math.Sin(a)))
					pts[i] = pp{ix, iy}
				}
				bc := brushColor()
				if bc.A > 0 {
					emfFill(pts, bc)
					hasDrawing = true
				}
			}
		case 0x2B: // RECTANGLE
			if len(rec) >= 24 {
				x0, y0 := toImg(int(i32(rec[8:12])), int(i32(rec[12:16])))
				x1, y1 := toImg(int(i32(rec[16:20])), int(i32(rec[20:24])))
				pts := []pp{{x0, y0}, {x1, y0}, {x1, y1}, {x0, y1}}
				bc := brushColor()
				if bc.A > 0 {
					emfFill(pts, bc)
					hasDrawing = true
				}
			}
		}
		if rt == 0x0E {
			break
		}
		pos += int(rs)
	}
	if !hasDrawing {
		return nil
	}
	return img
}
