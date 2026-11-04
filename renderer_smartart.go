package gopresentation

import (
	"image"
	"image/color"
	"math"
	"sort"

	"golang.org/x/image/font"
)

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
	switch model.Algo {
	case "cycle":
		r.smartArtCycle(nodes, model.Colorful, face, x, y, w, h)
	case "hier":
		r.smartArtHier(nodes, model.Colorful, face, x, y, w, h)
	case "snake":
		r.smartArtSnake(nodes, model.Colorful, face, x, y, w, h)
	default:
		r.smartArtLin(nodes, model.Colorful, face, x, y, w, h)
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
func (r *renderer) smartArtNode(n *SmartArtNode, i int, colorful bool, face font.Face, rect image.Rectangle) {
	if rect.Dx() < 4 || rect.Dy() < 4 {
		return
	}
	slot := smartArtSlot(i, colorful)
	if colorful && rect.Dy() > 24 {
		bandH := rect.Dy() * 30 / 100
		if bandH > 46*rect.Dy()/100 {
			bandH = rect.Dy() * 46 / 100
		}
		head := image.Rect(rect.Min.X, rect.Min.Y, rect.Max.X, rect.Min.Y+bandH)
		body := image.Rect(rect.Min.X, rect.Min.Y+bandH, rect.Max.X, rect.Max.Y)
		r.fillRectFast(head, slot)
		r.fillRectFast(body, smartArtBody(slot))
		r.drawStringCentered(n.Text, face, color.RGBA{R: 255, G: 255, B: 255, A: 255}, head)
		return
	}
	r.fillRectFast(rect, slot)
	r.drawStringCentered(n.Text, face, color.RGBA{R: 255, G: 255, B: 255, A: 255}, rect)
}

// smartArtLin lays the children across the frame, each full height.
func (r *renderer) smartArtLin(nodes []*SmartArtNode, colorful bool, face font.Face, x, y, w, h int) {
	n := len(nodes)
	if n == 0 || w < 1 {
		return
	}
	cw := w / n
	for i, node := range nodes {
		rect := image.Rect(x+i*cw, y, x+(i+1)*cw, y+h)
		if i == n-1 {
			rect.Max.X = x + w
		}
		r.smartArtNode(node, i, colorful, face, rect)
	}
}

// smartArtSnake lays the children in a grid: one entry per row while there
// are two or fewer (the font-size deck's stacked boxes), otherwise two per
// row (the block-diagram wrap the rightoleft deck's export shows).
func (r *renderer) smartArtSnake(nodes []*SmartArtNode, colorful bool, face font.Face, x, y, w, h int) {
	n := len(nodes)
	if n == 0 || w < 1 || h < 1 {
		return
	}
	cols := 2
	if n <= 2 {
		cols = 1
	}
	rows := (n + cols - 1) / cols
	cw := w / cols
	ch := h / rows
	for i, node := range nodes {
		row := i / cols
		col := i % cols
		rect := image.Rect(x+col*cw, y+row*ch, x+(col+1)*cw, y+(row+1)*ch)
		if col == cols-1 {
			rect.Max.X = x + w
		}
		if row == rows-1 {
			rect.Max.Y = y + h
		}
		r.smartArtNode(node, i, colorful, face, rect)
	}
}

// smartArtCycle places the children on a ring, drawn as ellipses with their
// white label centred.
func (r *renderer) smartArtCycle(nodes []*SmartArtNode, colorful bool, face font.Face, x, y, w, h int) {
	n := len(nodes)
	if n == 0 || w < 1 || h < 1 {
		return
	}
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2
	radius := float64(minInt(w, h)) / 2 * 0.55
	d := float64(minInt(w, h)) / 4.4
	for i, node := range nodes {
		angle := -math.Pi/2 + 2*math.Pi*float64(i)/float64(n)
		ncx := cx + radius*math.Cos(angle)
		ncy := cy + radius*math.Sin(angle)
		ix := int(math.Round(ncx - d/2))
		iy := int(math.Round(ncy - d/2))
		nodeRect := image.Rect(ix, iy, ix+int(d), iy+int(d))
		r.fillEllipseAA(nodeRect.Min.X, nodeRect.Min.Y, nodeRect.Dx(), nodeRect.Dy(), smartArtSlot(i, colorful))
		r.drawStringCentered(node.Text, face, color.RGBA{R: 255, G: 255, B: 255, A: 255}, nodeRect)
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
func (r *renderer) smartArtHier(nodes []*SmartArtNode, colorful bool, face font.Face, x, y, w, h int) {
	if len(nodes) == 0 || w < 1 || h < 1 {
		return
	}
	var maxDepth func(list []*SmartArtNode) int
	maxDepth = func(list []*SmartArtNode) int {
		md := 0
		for _, node := range list {
			if d := maxDepth(node.Children); d > md {
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
			cy := float64(y) + (float64(depth)+0.5)*rowH
			boxes = append(boxes, entry{n, depth, colX + colW/2, cy, colX, colX + colW})
			return
		}
		var kidCx []float64
		var kidCy []float64
		firstL, lastR := 0.0, 0.0
		for i, c := range n.Children {
			layout(c, depth+1)
			ch := boxes[len(boxes)-1]
			kidCx = append(kidCx, ch.cx)
			kidCy = append(kidCy, float64(y)+(float64(depth+1)+0.5)*rowH)
			if i == 0 {
				firstL = ch.leafL
			}
			lastR = ch.leafR
		}
		cx := (firstL + lastR) / 2
		cy := float64(y) + (float64(depth)+0.5)*rowH
		boxes = append(boxes, entry{n, depth, cx, cy, firstL, lastR})
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
		cy := float64(y) + (float64(b.depth)+0.5)*rowH
		rect := image.Rect(int(b.cx)-bw/2, int(cy)-bh/2, int(b.cx)+bw/2, int(cy)+bh/2)
		if rect.Min.X < 0 {
			rect = rect.Add(image.Pt(-rect.Min.X, 0))
		}
		r.fillRectFast(rect, smartArtSlot(b.depth, colorful))
		tf := NewFont()
		tf.Size = maxInt(11, minInt(38, bh*27/100))
		r.drawStringCentered(b.node.Text, r.getFace(tf), color.RGBA{R: 255, G: 255, B: 255, A: 255}, rect)
	}
}

// totalHierLeaves counts the leaves across the top-level subtrees.
func totalHierLeaves(nodes []*SmartArtNode) int {
	n := 0
	for _, node := range nodes {
		n += smartArtCountLeaves(node)
	}
	if n == 0 {
		n = 1
	}
	return n
}
