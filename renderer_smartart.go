package gopresentation

import (
	"fmt"
	"image"
	"image/color"
	"math"
	"sort"
	"strings"

	"golang.org/x/image/font"
	"golang.org/x/image/math/fixed"
)

// smartArtWrap breaks a label into lines no wider than maxW measured at face.
func smartArtWrap(text string, face font.Face, maxW int) []string {
	text = strings.TrimSpace(text)
	if text == "" {
		return nil
	}
	words := strings.Fields(text)
	if len(words) == 0 {
		return nil
	}
	var lines []string
	line := words[0]
	for _, w := range words[1:] {
		cand := line + " " + w
		if font.MeasureString(face, cand).Ceil() > maxW && line != "" {
			lines = append(lines, line)
			line = w
		} else {
			line = cand
		}
	}
	lines = append(lines, line)
	return lines
}

// smartArtFitFace shrinks the label's face until the wrapped lines fit the
// box height; the corpus's Office exports big labels on big boxes — the
// "Max size (65 pt)" box — at whatever size the autofit allows.
func (r *renderer) smartArtFitFace(text string, maxW, maxH int) (font.Face, []string) {
	for size := 36; size >= 8; size -= 3 {
		f := NewFont()
		f.Size = size
		face := r.getFace(f)
		lines := smartArtWrap(text, face, maxW)
		if len(lines) == 0 {
			return face, nil
		}
		m := face.Metrics()
		th := (m.Ascent + m.Descent).Ceil()
		if len(lines)*th <= maxH {
			return face, lines
		}
	}
	f := NewFont()
	f.Size = 8
	return r.getFace(f), smartArtWrap(text, r.getFace(f), maxW)
}

// smartArtDrawLines paints wrapped label lines, each centred in its slot.
func (r *renderer) smartArtDrawLines(lines []string, face font.Face, c color.RGBA, rect image.Rectangle) {
	if len(lines) == 0 || rect.Dy() < 4 {
		return
	}
	m := face.Metrics()
	th := (m.Ascent + m.Descent).Ceil()
	total := th * len(lines)
	y := rect.Min.Y + (rect.Dy()-total)/2 + m.Ascent.Ceil()
	for _, line := range lines {
		tw := font.MeasureString(face, line).Ceil()
		cx := rect.Min.X + (rect.Dx()-tw)/2
		d := &font.Drawer{Dst: r.img, Src: image.NewUniform(drawSrc(c)), Face: face, Dot: fixed.P(cx, y)}
		d.DrawString(line)
		y += th
	}
}

// The v1 SmartArt layout engine: it lays a SmartArtModel's tree out in one of
// four geometry families and paints each node with a theme-accent fill and
// centred white text. Geometry follows the family the model declared:
//
//	lin    — children side by side across the frame, full height
//	snake  — nodes in a grid, two per row while more than two siblings
//	cycle  — nodes on a ring, as ellipses
//	hier   — levels of the tree stacked top-down, parents centred over
//	         their children's span, with a connector dropped down
//
// It is not the layout definition's constraint solver, and the reader says
// so honestly; it replaces a stark stand-in box with a diagram that reads
// like the family the deck asked for.

// renderSmartArt lays the model out inside the stand-in's frame.
func (r *renderer) renderSmartArt(model *SmartArtModel, x, y, w, h int) {
	if model == nil || model.Root == nil || len(model.Root.Children) == 0 {
		return
	}
	nodes := flattenSmartArtTop(model.Root.Children)
	face := r.getFace(NewFont())
	if model.Process {
		if model.ProcessTwoBox {
			r.smartArtTwoBox(nodes, model.Colorful, face, x, y, w, h)
		} else {
			r.smartArtProcessArrow(nodes, model.Colorful, face, x, y, w, h)
		}
		return
	}
	if model.BulletCols {
		r.smartArtBulletCols(nodes, model.Colorful, model.AccentARGB, face, x, y, w, h)
		return
	}
	if model.VertBoxes {
		r.smartArtVertBoxes(nodes, model.Colorful, model.AccentARGB, face, x, y, w, h)
		return
	}
	switch model.Algo {
	case "cycle":
		r.smartArtCycle(nodes, model.Colorful, model.Process, face, x, y, w, h)
	case "hier":
		r.smartArtHier(nodes, model.Colorful, model.Process, face, x, y, w, h)
	case "snake":
		r.smartArtSnake(nodes, model.Colorful, model.Process, model.RTLSnake, face, x, y, w, h)
	default:
		r.smartArtLin(nodes, model.Colorful, model.Process, model.Vert, model.CardList, face, x, y, w, h)
	}
}

// flattenSmartArtTop unwraps a single empty root wrapper the tree may carry.
func flattenSmartArtTop(nodes []*SmartArtNode) []*SmartArtNode {
	if len(nodes) == 1 && nodes[0].Text == "" && len(nodes[0].Children) > 0 {
		return nodes[0].Children
	}
	return nodes
}

// smartArtSlot is the palette slot for node index i. A colourful quick
// style cycles the accents per node (starting one slot in, where the
// fill-color-list deck's red/green/purple headers live); a uniform quick
// style pins every node to the first accent.
func smartArtSlot(i int, colorful bool) color.RGBA {
	palette := chartColors()
	if colorful {
		return palette[(i+1)%len(palette)]
	}
	return palette[0]
}

// smartArtBody is the tinted light fill under a node's colour band.
func smartArtBody(slot color.RGBA) color.RGBA {
	// two-tenths of the accent over eight of white: the pale body the
	// colourful quick style paints under the header band.
	return color.RGBA{
		R: uint8((int(slot.R)*2 + 255*8) / 10),
		G: uint8((int(slot.G)*2 + 255*8) / 10),
		B: uint8((int(slot.B)*2 + 255*8) / 10),
		A: 255,
	}
}

// smartArtNode paints one node box: the fill (a coloured header band over a
// pale body when the box is tall enough) and the white label centred in it.
func (r *renderer) smartArtNode(n *SmartArtNode, i int, colorful, twoToneOnly bool, face font.Face, rect image.Rectangle) {
	if rect.Dx() < 4 || rect.Dy() < 4 {
		return
	}
	white := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	slot := smartArtSlot(i, colorful)
	twoTone := colorful || twoToneOnly
	if twoTone && rect.Dy() > 24 {
		bandH := rect.Dy() * 22 / 100
		if bandH > 46*rect.Dy()/100 {
			bandH = rect.Dy() * 46 / 100
		}
		head := image.Rect(rect.Min.X, rect.Min.Y, rect.Max.X, rect.Min.Y+bandH)
		body := image.Rect(rect.Min.X, rect.Min.Y+bandH, rect.Max.X, rect.Max.Y)
		r.fillRectFast(head, slot)
		r.fillRectFast(body, smartArtBody(slot))
		face, lines := r.smartArtFitFace(n.Text, head.Dx(), head.Dy())
		r.smartArtDrawLines(lines, face, white, head)
		return
	}
	r.fillRectFast(rect, slot)
	face, lines := r.smartArtFitFace(n.Text, rect.Dx()-8, rect.Dy()-8)
	inner := image.Rect(rect.Min.X+3, rect.Min.Y+2, rect.Max.X-3, rect.Max.Y-2)
	r.smartArtDrawLines(lines, face, white, inner)
}

// smartArtLeaves flattens a tree to its leaf nodes — the list and process
// families draw one box per content item, not per group node.
func smartArtLeaves(nodes []*SmartArtNode) []*SmartArtNode {
	var out []*SmartArtNode
	var visit func(n *SmartArtNode)
	visit = func(n *SmartArtNode) {
		if len(n.Children) == 0 {
			out = append(out, n)
			return
		}
		for _, c := range n.Children {
			visit(c)
		}
	}
	for _, n := range nodes {
		visit(n)
	}
	return out
}

// smartArtLin lays the children across the frame, each full height.
func (r *renderer) smartArtLin(nodes []*SmartArtNode, colorful, twoTone, vert, cardRow bool, face font.Face, x, y, w, h int) {
	// loCatId "list" with a hierarchy: the vert tree draws parent headers
	// over child rows — the table-list deck's Parent/Child 1/Child 2 look.
	// This check runs BEFORE leaf flattening so the hierarchy survives.
	if vert {
		hasKids := false
		for _, node := range nodes {
			if len(node.Children) > 0 {
				hasKids = true
				break
			}
		}
		if hasKids {
			r.smartArtVertTree(nodes, colorful, twoTone, face, x, y, w, h)
			return
		}
	}
	nodes = smartArtLeaves(nodes)
	n := len(nodes)
	if n == 0 || w < 1 {
		return
	}
	if vert {
		// loCatId "list" without hierarchy: nodes flow top-to-bottom, full
		// width each, with a vertical gap (the linear-rule-vert stack). A
		// single node fills its frame — the LO "default" family's lone card
		// spans ~93% of it, where a stacked list keeps 15% breathing room.
		ch := h / n
		insetPct := 15
		if n == 1 {
			insetPct = 4
		}
		for i, node := range nodes {
			rect := image.Rect(x, y+i*ch, x+w, y+(i+1)*ch)
			if i == n-1 {
				rect.Max.Y = y + h
			}
			inset := ch * insetPct / 100
			rect.Min.Y += inset
			rect.Max.Y -= inset
			r.smartArtNode(node, i, colorful, twoTone, face, rect)
		}
		return
	}
	cw := w / n
	// The colourful list lays nodes with a gutter between them — the
	// fill-color-list export keeps ~4% of each column as breathing room —
	// and the header band reads at a fifth of the column, not a third.
	gap := cw * 4 / 100
	// The hList cards float inside the frame: ~10% gutters and ~8% vertical
	// margins top and bottom (measured off the fill-color-list export).
	marginY := 0
	if cardRow {
		gap = cw * 10 / 100
		marginY = h * 8 / 100
	}
	for i, node := range nodes {
		rect := image.Rect(x+i*cw, y+marginY, x+(i+1)*cw, y+h-marginY)
		if i == n-1 {
			rect.Max.X = x + w
		}
		if gap > 0 {
			if i > 0 {
				rect.Min.X += gap / 2
			}
			if i < n-1 {
				rect.Max.X -= gap / 2
			}
		}
		r.smartArtNode(node, i, colorful, twoTone, face, rect)
	}
}

// smartArtSnake lays the children in a grid: one entry per row while there
// are two or fewer (the font-size deck's stacked boxes), otherwise two per
// row (the block-diagram wrap the rightoleft deck's export shows).
func (r *renderer) smartArtSnake(nodes []*SmartArtNode, colorful, twoTone, rtl bool, face font.Face, x, y, w, h int) {
	nodes = smartArtLeaves(nodes)
	n := len(nodes)
	if n == 0 || w < 1 || h < 1 {
		return
	}
	cols := 2
	if n <= 2 {
		cols = 1
	}
	// The RTL grid centres at ~72% of the frame width, margins and all —
	// the rightoleft export keeps ~14% of the frame free on each side.
	if rtl {
		inset := int(float64(w) * 0.14)
		x += inset
		w -= 2 * inset
	}
	rows := (n + cols - 1) / cols
	cw := w / cols
	ch := h / rows
	gx := cw * 6 / 100
	gy := ch * 12 / 100
	for i, node := range nodes {
		row := i / cols
		col := i % cols
		if rtl {
			// The grDir="tR" branch flows right to left: mirror the column.
			col = cols - 1 - col
		}
		rect := image.Rect(x+col*cw, y+row*ch, x+(col+1)*cw, y+(row+1)*ch)
		if col < cols-1 {
			rect.Max.X -= gx
		}
		if row < rows-1 {
			rect.Max.Y -= gy
		}
		if i > 0 && col > 0 {
			rect.Min.X += gx / 2
		}
		if i > 0 && row > 0 {
			rect.Min.Y += gy / 2
		}
		// centre the odd-count last row (the pre-mirror column index tells
		// whether the row started at the flow edge)
		if i == n-1 && i%cols == 0 && cols > 1 {
			centre := x + w/2
			off := centre - (rect.Min.X+rect.Max.X)/2
			rect = rect.Add(image.Pt(off, 0))
		}
		r.smartArtNode(node, i, colorful, twoTone, face, rect)
	}
}

// smartArtCycle places the children on a ring, drawn as ellipses with their
// white label centred.
func (r *renderer) smartArtCycle(nodes []*SmartArtNode, colorful, twoTone bool, face font.Face, x, y, w, h int) {
	n := len(nodes)
	if n == 0 || w < 1 || h < 1 {
		return
	}
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2
	radius := float64(minInt(w, h)) / 2 * 0.70
	// Node size: the Office export's circles read at ~28% of the frame's
	// short side with clear gaps to their neighbours; the chord between
	// adjacent ring positions still caps the diameter so dense rings (the
	// autoTxRot deck's 12 nodes) never overlap.
	d := 2 * radius * math.Sin(math.Pi/float64(n)) * 0.78
	if d > float64(minInt(w, h))*0.28 {
		d = float64(minInt(w, h)) * 0.28
	}
	// The label scales with its circle (the export's caps read ~28% of it).
	lf := NewFont()
	lf.Size = maxInt(10, minInt(64, int(d*0.28)))
	labelFace := r.getFace(lf)
	arrowSlot := smartArtBodyMix(smartArtSlot(0, colorful), 0.45)
	for i, node := range nodes {
		angle := -math.Pi/2 + 2*math.Pi*float64(i)/float64(n)
		ncx := cx + radius*math.Cos(angle)
		ncy := cy + radius*math.Sin(angle)
		ix := int(math.Round(ncx - d/2))
		iy := int(math.Round(ncy - d/2))
		nodeRect := image.Rect(ix, iy, ix+int(d), iy+int(d))
		r.fillEllipseAA(nodeRect.Min.X, nodeRect.Min.Y, nodeRect.Dx(), nodeRect.Dy(), smartArtSlot(i, colorful))
		r.drawStringCentered(node.Text, labelFace, color.RGBA{R: 255, G: 255, B: 255, A: 255}, nodeRect)
	}
	// A small chevron rides the ring between each pair of neighbours,
	// pointing along the flow.
	for i := 0; i < n; i++ {
		mid := -math.Pi/2 + 2*math.Pi*(float64(i)+0.5)/float64(n)
		acx := cx + radius*math.Cos(mid)
		acy := cy + radius*math.Sin(mid)
		// tangential direction (clockwise flow)
		tx, ty := -math.Sin(mid), math.Cos(mid)
		sz := d * 0.30
		hx, hy := tx*sz/2, ty*sz/2
		nx, ny := -ty, tx
		pts := []fpoint{
			{acx + hx, acy + hy},
			{acx - hx + nx*sz*0.45, acy - hy + ny*sz*0.45},
			{acx - hx - nx*sz*0.45, acy - hy - ny*sz*0.45},
		}
		r.fillPolygon(pts, arrowSlot)
	}
}

func smartArtCountLeaves(n *SmartArtNode) int {
	if len(n.Children) == 0 {
		return 1
	}
	c := 0
	for _, ch := range n.Children {
		c += smartArtCountLeaves(ch)
	}
	return c
}

// smartArtHier draws the classic tidy org tree, the shape loTypeId
// orgChart1's Office export shows: every node gets the SAME box, leaves sit
// in equal columns side by side, a parent is centred over the span of its
// subtree's columns, and elbow connectors (down-stub, bus, down-stubs) link
// the rows. Connector lines paint first so boxes cover them.
func (r *renderer) smartArtHier(nodes []*SmartArtNode, colorful, twoTone bool, face font.Face, x, y, w, h int) {
	if len(nodes) == 0 || w < 1 || h < 1 {
		return
	}
	// The depth count adds a level for every assistant band on the deepest
	// chain: the org chart's assistant row sits between the parent and the
	// children, so the children shift one row down.
	var maxDepth func(list []*SmartArtNode) int
	maxDepth = func(list []*SmartArtNode) int {
		md := 0
		for _, node := range list {
			d := maxDepth(node.Children)
			for _, c := range node.Children {
				if c.Asst {
					d++
					break
				}
			}
			if d > md {
				md = d
			}
		}
		return md + 1
	}
	deep := maxDepth(nodes)
	colW := float64(w) / float64(totalHierLeaves(nodes))
	if colW < 1 {
		return
	}
	rowH := float64(h) / float64(deep)
	bw := int(colW * 0.92)
	if bw < 24 {
		bw = 24
	}
	// Org-chart cards keep a near-fixed width regardless of the leaf count:
	// a two-leaf parent's columns would otherwise stretch its card across
	// half the frame (the Office export's cards read at ~26% of it).
	if bw > int(float64(w)*0.30) {
		bw = int(float64(w) * 0.30)
	}
	bh := int(rowH * 0.72)
	if bh > bw*9/10 {
		bh = bw * 9 / 10
	}

	type entry struct {
		node  *SmartArtNode
		depth int
		cx    float64
		cy    float64
		leafL float64
		leafR float64
	}
	var boxes []entry
	var edges [][2]float64
	colCursor := 0.0

	var layout func(n *SmartArtNode, depth int)
	layout = func(n *SmartArtNode, depth int) {
		if len(n.Children) == 0 {
			colX := colCursor
			colCursor += colW
			// colX is frame-relative; the box centres land in slide space,
			// so the frame's own x joins here — a frame that doesn't sit at
			// the slide origin (org-chart's graphicFrame starts a third of
			// the way in) laid its whole tree off to the left without it.
			cx := float64(x) + colX + colW/2
			cy := float64(y) + (float64(depth)+0.5)*rowH
			boxes = append(boxes, entry{n, depth, cx, cy, float64(x) + colX, float64(x) + colX + colW})
			return
		}
		// Assistants hang off the trunk in their own band between this
		// parent and the children; only the regular children take leaf
		// columns, and they shift a row down when a band exists.
		var kids, assts []*SmartArtNode
		for _, c := range n.Children {
			if c.Asst {
				assts = append(assts, c)
			} else {
				kids = append(kids, c)
			}
		}
		kidDepth := depth + 1
		if len(assts) > 0 {
			kidDepth = depth + 2
		}
		var kidCx []float64
		var kidCy []float64
		firstL, lastR := 0.0, 0.0
		for i, c := range kids {
			layout(c, kidDepth)
			ch := boxes[len(boxes)-1]
			kidCx = append(kidCx, ch.cx)
			kidCy = append(kidCy, float64(y)+(float64(kidDepth)+0.5)*rowH)
			if i == 0 {
				firstL = ch.leafL
			}
			lastR = ch.leafR
		}
		cx := (firstL + lastR) / 2
		cy := float64(y) + (float64(depth)+0.5)*rowH
		boxes = append(boxes, entry{n, depth, cx, cy, firstL, lastR})
		// Assistant boxes: one band row between the parent and the children
		// row, over the first child's column, with a horizontal tie to the
		// trunk (multiple assistants stack within the band).
		for j, a := range assts {
			acy := float64(y) + (float64(depth)+1.5)*rowH + float64(j)*rowH*0.4
			acx := cx
			if len(kidCx) > 0 {
				acx = kidCx[0]
			}
			boxes = append(boxes, entry{a, kidDepth, acx, acy, firstL, firstL + colW})
			if len(assts) == 1 {
				edges = append(edges, [2]float64{acx + float64(bw)/2, acy}, [2]float64{cx, acy})
			}
		}
		// elbow pieces: parent stub down, bus across, child stubs down
		busY := float64(y) + (float64(depth)+1.0)*rowH - rowH*0.18
		parentBot := float64(y) + (float64(depth)+0.5)*rowH + float64(bh)/2
		edges = append(edges, [2]float64{cx, parentBot}, [2]float64{cx, busY})
		for i := range kidCx {
			edges = append(edges, [2]float64{cx, busY}, [2]float64{kidCx[i], busY})
			edges = append(edges, [2]float64{kidCx[i], busY}, [2]float64{kidCx[i], kidCy[i]})
		}
	}
	for _, node := range nodes {
		layout(node, 0)
	}

	grey := color.RGBA{R: 130, G: 130, B: 130, A: 255}
	for i := 0; i+1 < len(edges); i += 2 {
		r.drawLineAA(int(edges[i][0]), int(edges[i][1]), int(edges[i+1][0]), int(edges[i+1][1]), grey, 1)
	}

	sort.SliceStable(boxes, func(i, j int) bool { return boxes[i].depth < boxes[j].depth })
	for i := range boxes {
		b := boxes[i]
		cy := b.cy
		rect := image.Rect(int(b.cx)-bw/2, int(cy)-bh/2, int(b.cx)+bw/2, int(cy)+bh/2)
		if rect.Min.X < 0 {
			rect = rect.Add(image.Pt(-rect.Min.X, 0))
		}
		r.fillRectFast(rect, smartArtSlot(b.depth, colorful))
		// Fit the label like the engine's other families: wrap to the box
		// width and shrink until the lines fit the box height — the org
		// chart's two-paragraph manager card and narrow employee boxes
		// clipped their text at a fixed size.
		tf, lines := r.smartArtFitFace(b.node.Text, rect.Dx()-8, rect.Dy()-6)
		r.smartArtDrawLines(lines, tf, color.RGBA{R: 255, G: 255, B: 255, A: 255}, rect)
	}
}

// totalHierLeaves counts the leaves across the top-level subtrees.
// Assistant nodes hang off the trunk instead of taking a leaf column, so
// they don't count.
func totalHierLeaves(nodes []*SmartArtNode) int {
	n := 0
	for _, node := range nodes {
		n += hierLeafCount(node)
	}
	if n == 0 {
		n = 1
	}
	return n
}

func hierLeafCount(n *SmartArtNode) int {
	if len(n.Children) == 0 {
		if n.Asst {
			return 0
		}
		return 1
	}
	sum := 0
	for _, c := range n.Children {
		sum += hierLeafCount(c)
	}
	return sum
}

// smartArtVertTree renders a parent-children hierarchy vertically: each
// parent draws as a full-width header band, its children as a horizontal
// row below (the table-list deck's Parent/Child 1/Child 2 layout).
func (r *renderer) smartArtVertTree(nodes []*SmartArtNode, colorful, twoTone bool, face font.Face, x, y, w, h int) {
	// count total rows: one per top-level node + one per child row
	rows := 0
	for _, node := range nodes {
		rows++
		if len(node.Children) > 0 {
			rows++
		}
	}
	rowH := float64(h) / float64(rows)
	curY := float64(y)
	for i, node := range nodes {
		// parent header: full width
		hdrH := int(rowH)
		hdrRect := image.Rect(x, int(curY), x+w, int(curY)+hdrH)
		r.fillRectFast(hdrRect, smartArtSlot(i, colorful))
		face2, lines := r.smartArtFitFace(node.Text, hdrRect.Dx()-16, hdrRect.Dy()-8)
		inner := image.Rect(hdrRect.Min.X+8, hdrRect.Min.Y+4, hdrRect.Max.X-8, hdrRect.Max.Y-4)
		r.smartArtDrawLines(lines, face2, color.RGBA{R: 255, G: 255, B: 255, A: 255}, inner)
		curY += float64(hdrH)
		// children row: side by side below the parent
		if len(node.Children) > 0 {
			kids := node.Children
			kw := w / len(kids)
			kidH := int(rowH)
			for j, kid := range kids {
				kidRect := image.Rect(x+j*kw, int(curY), x+(j+1)*kw, int(curY)+kidH)
				if j == len(kids)-1 {
					kidRect.Max.X = x + w
				}
				r.fillRectFast(kidRect, smartArtSlot(i, colorful))
				face3, klines := r.smartArtFitFace(kid.Text, kidRect.Dx()-16, kidRect.Dy()-8)
				kinner := image.Rect(kidRect.Min.X+8, kidRect.Min.Y+4, kidRect.Max.X-8, kidRect.Max.Y-4)
				r.smartArtDrawLines(klines, face3, color.RGBA{R: 255, G: 255, B: 255, A: 255}, kinner)
			}
			curY += float64(kidH)
		}
	}
}

// smartArtBulletCols draws the pList family (vertical bullet list): a tinted
// container band across the top carrying one empty rounded box per column,
// then the columns themselves as tall rounded cards below, each holding its
// node label over a bulleted child list. Geometry and colours are measured
// off the autofit-sync Office export: the band spans the full frame at 45%
// of its height, the boxes and cards share the column x-spans (29% of the
// frame wide with 3% gutters and 3% side insets), the cards run from the
// band's bottom edge down to the frame bottom, and the whole family wears
// one accent — the container near-white, the boxes a step darker, the cards
// the raw accent.
func (r *renderer) smartArtBulletCols(nodes []*SmartArtNode, colorful bool, accentARGB string, face font.Face, x, y, w, h int) {
	n := len(nodes)
	if n == 0 || w < 40 || h < 40 {
		return
	}
	white := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	slot := smartArtSlot(0, colorful)
	if accentARGB != "" {
		// A uniform quick style wears the deck theme's own accent1.
		slot = argbToRGBA(Color{ARGB: accentARGB})
	}
	// The band's tint runs ~75% of the way to white; the empty boxes inside
	// sit a step darker at ~68% (measured off the autofit-sync export).
	band := smartArtBodyMix(slot, 0.75)
	boxFill := smartArtBodyMix(slot, 0.68)
	bandH := h * 45 / 100
	r.fillRoundedRect(x, y, w, bandH, h*4/100, band)
	// Column spans: 3% side insets, 3% gutters.
	inset := w * 3 / 100
	gutter := w * 3 / 100
	colW := (w - 2*inset - (n-1)*gutter) / n
	if colW < 8 {
		return
	}
	boxInset := h * 63 / 1000
	boxTop := y + boxInset
	boxH := bandH - 2*boxInset
	radius := h * 15 / 1000
	for i := range nodes {
		bx := x + inset + i*(colW+gutter)
		r.fillRoundedRect(bx, boxTop, colW, boxH, radius, boxFill)
	}
	// The cards hang right below the band and run to the frame bottom.
	cardY := y + bandH + 2
	cardH := y + h - cardY
	textX := w * 25 / 1000 // the label/bullet inset inside a card, in px
	for i, node := range nodes {
		bx := x + inset + i*(colW+gutter)
		r.fillRoundedRect(bx, cardY, colW, cardH, radius, slot)
		r.smartArtBulletText(node, face, white, bx+textX, cardY+textX, colW-2*textX, cardH-2*textX)
	}
}

// smartArtBulletText paints one card's content: the node label left-aligned
// near the top, then one bullet line per child, wrapping each to the card
// width and shrinking until every child fits (the autofit the deck is named
// for).
func (r *renderer) smartArtBulletText(node *SmartArtNode, face font.Face, c color.RGBA, tx, ty, tw, th int) {
	if tw < 20 || th < 20 {
		return
	}
	children := node.Children
	label := node.Text
	// Collect the bullet lines once per size: label first, then "• child".
	lines := func(f font.Face, lw int) []string {
		out := smartArtWrap(label, f, lw)
		for _, kid := range children {
			out = append(out, smartArtWrap("• "+kid.Text, f, lw)...)
		}
		return out
	}
	face, all := func() (font.Face, []string) {
		for size := 14; size >= 8; size -= 2 {
			f := NewFont()
			f.Size = size
			fc := r.getFace(f)
			if ls := lines(fc, tw); len(ls) > 0 {
				m := fc.Metrics()
				if len(ls)*int((m.Ascent+m.Descent).Ceil()) <= th {
					return fc, ls
				}
			}
		}
		f := NewFont()
		f.Size = 8
		return r.getFace(f), lines(r.getFace(f), tw)
	}()
	m := face.Metrics()
	thh := int((m.Ascent + m.Descent).Ceil())
	y := ty
	for _, line := range all {
		if y+thh > ty+th {
			break
		}
		d := &font.Drawer{Dst: r.img, Src: image.NewUniform(drawSrc(c)), Face: face, Dot: fixed.P(tx, y+m.Ascent.Ceil())}
		d.DrawString(line)
		y += thh
	}
}

// smartArtVertBoxes draws the list1 family (vertical box list): each row is
// a full-width thin-outlined white box with a solid rounded card overlapping
// its upper-left, the node label sitting small at the card's top-left.
// Geometry measured off the linear-rule-vert Office export (1600px wide, one
// 776px frame of three rows): the outlined box spans the whole frame at 63%
// of a row's pitch, starting 35% in; the card runs from the row top to 73%
// of the pitch at 69% of the frame width, offset 5.4% from the left; the
// card corners round at ~1.4% of the frame width.
func (r *renderer) smartArtVertBoxes(nodes []*SmartArtNode, colorful bool, accentARGB string, face font.Face, x, y, w, h int) {
	nodes = smartArtLeaves(nodes)
	n := len(nodes)
	if n == 0 || w < 40 || h < 40 {
		return
	}
	slot := smartArtSlot(0, colorful)
	if accentARGB != "" {
		slot = argbToRGBA(Color{ARGB: accentARGB})
	}
	white := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	pitch := float64(h) / float64(n)
	outlineW := maxInt(int(float64(w)*0.003), 2)
	cardX := x + int(float64(w)*0.054)
	cardW := int(float64(w) * 0.694)
	radius := int(float64(w) * 0.014)
	// The label rides the card's top-left, sized like the export's ~11pt
	// italic on a 188px-tall card.
	lf := NewFont()
	lf.Size = 11
	lf.SetItalic(true)
	labelFace := r.getFace(lf)
	m := labelFace.Metrics()
	lh := int((m.Ascent + m.Descent).Ceil())
	for i, node := range nodes {
		rowY := float64(y) + float64(i)*pitch
		// The outlined box sits behind, lower than the card.
		obY := int(rowY + pitch*0.351)
		obH := int(pitch * 0.633)
		if obY+obH > y+h {
			obH = y + h - obY
		}
		r.fillRectFast(image.Rect(x, obY, x+w, obY+obH), white)
		r.drawRectBorder(image.Rect(x, obY, x+w, obY+obH), argbToRGBA(Color{ARGB: hexARGB(slot)}), outlineW, BorderSolid)
		// The solid card overlaps the box's top edge.
		cY := int(rowY)
		cH := int(pitch * 0.726)
		r.fillRoundedRect(cardX, cY, cardW, cH, radius, slot)
		// Label at the card's top-left.
		tx := cardX + int(float64(w)*0.025)
		ty := cY + int(float64(cH)*0.12)
		d := &font.Drawer{Dst: r.img, Src: image.NewUniform(drawSrc(white)), Face: labelFace, Dot: fixed.P(tx, ty+lh)}
		d.DrawString(node.Text)
	}
}

// hexARGB renders an RGBA back into the ARGB hex string the colour helpers
// take.
func hexARGB(c color.RGBA) string {
	return fmt.Sprintf("%02X%02X%02X%02X", c.A, c.R, c.G, c.B)
}

// smartArtBodyMix blends an accent toward white by the given weight — the
// family's container tints.
func smartArtBodyMix(slot color.RGBA, weight float64) color.RGBA {
	mix := func(v uint8) uint8 {
		return uint8(float64(v)*(1-weight) + 255*weight)
	}
	return color.RGBA{R: mix(slot.R), G: mix(slot.G), B: mix(slot.B), A: 255}
}

// smartArtProcessArrow draws the hProcess3 family's single arrow shape:
// a rect body with a triangular tip pointing right, all children's texts
// evenly spaced inside in white. The arrow fills with one accent — the
// office export for the linear-rule deck shows this arrow, not separate
// boxes per item.
func (r *renderer) smartArtProcessArrow(nodes []*SmartArtNode, colorful bool, face font.Face, x, y, w, h int) {
	if len(nodes) == 0 || w < 1 || h < 1 {
		return
	}
	slot := smartArtSlot(0, colorful)
	// The hProcess3 composite packs the arrow into the upper-middle of its
	// frame and notches the profile: the body is a low rect, the head's
	// back edge rises above the body top, and the tip lands ~72% down the
	// arrow's height (measured off the linear-rule export: the arrow spans
	// 19%..63% of the frame, body top at 18% of the arrow, head back at
	// 75% of the width, tip at the right edge, body's bottom-right corner
	// pulled in to 92.5%).
	ax, ay, aw, ah := float64(x), float64(y)+float64(h)*0.19, float64(w), float64(h)*0.44
	pts := [][2]float64{
		{ax, ay + ah*0.18},
		{ax + aw*0.754, ay + ah*0.18},
		{ax + aw*0.754, ay},
		{ax + aw, ay + ah*0.72},
		{ax + aw*0.925, ay + ah},
		{ax, ay + ah},
	}
	var fpts []fpoint
	for _, p := range pts {
		fpts = append(fpts, fpoint{p[0], p[1]})
	}
	r.fillPolygon(fpts, slot)

	// texts evenly spaced across the body (the head excluded), centred in
	// the body band
	white := color.RGBA{R: 255, G: 255, B: 255, A: 255}
	n := len(nodes)
	bodyW := aw * 0.754
	slotW := bodyW / float64(n)
	bodyTop := ay + ah*0.18
	bodyH := ah * 0.82
	for i, node := range nodes {
		cx := ax + slotW*(float64(i)+0.5)
		nodeRect := image.Rect(int(cx-slotW/2), int(bodyTop), int(cx+slotW/2), int(bodyTop+bodyH))
		face2, lines := r.smartArtFitFace(node.Text, nodeRect.Dx()-8, nodeRect.Dy()-6)
		inner := image.Rect(nodeRect.Min.X+4, nodeRect.Min.Y+2, nodeRect.Max.X-4, nodeRect.Max.Y-2)
		r.smartArtDrawLines(lines, face2, white, inner)
	}
}

// smartArtTwoBox draws the accent-process family's layout: each node
// renders as a blue header box (top-left) with a pale body box (larger,
// below-right, overlapping) — the office export's header/body pair with
// a thin border on the body box. The texts go in their respective boxes.
func (r *renderer) smartArtTwoBox(nodes []*SmartArtNode, colorful bool, face font.Face, x, y, w, h int) {
	n := len(nodes)
	if n == 0 || w < 1 || h < 1 {
		return
	}
	cw := w / n
	white := color.RGBA{R: 255, G: 255, B: 255, A: 255}

	for i, node := range nodes {
		nx := x + i*cw
		slot := smartArtSlot(i, colorful)
		body := smartArtBody(slot)

		// header box: blue, covers the top-left area (smaller)
		hdrRect := image.Rect(nx, y, nx+cw*7/8, y+h*45/100)
		r.fillRectFast(hdrRect, slot)
		face2, lines := r.smartArtFitFace(node.Text, hdrRect.Dx()-12, hdrRect.Dy()-8)
		inner := image.Rect(hdrRect.Min.X+6, hdrRect.Min.Y+3, hdrRect.Max.X-6, hdrRect.Max.Y-3)
		r.smartArtDrawLines(lines, face2, white, inner)

		// body box: pale with a thin border, larger, below-right (overlapping)
		bodyRect := image.Rect(nx+cw/8, y+h*30/100, nx+cw-4, y+h*88/100)
		r.fillRectFast(bodyRect, body)
		r.drawRectBorder(bodyRect, color.RGBA{R: 120, G: 120, B: 120, A: 255}, 1, BorderSolid)

		// children as bullet text lines inside the body box
		kidY := bodyRect.Min.Y + 8
		kidFace := NewFont()
		kidFace.Size = 11
		kidFace2 := r.getFace(kidFace)
		for _, kid := range node.Children {
			kidRect := image.Rect(bodyRect.Min.X+10, kidY, bodyRect.Max.X-10, kidY+16)
			r.drawStringCentered("• "+kid.Text, kidFace2, color.RGBA{R: 50, G: 50, B: 50, A: 255}, kidRect)
			kidY += 16
		}
	}
}
