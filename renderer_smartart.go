package gopresentation

import (
	"image"
	"image/color"
	"math"

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

// smartArtHier draws the tree top-down: a parent sits centred over the span
// its subtree's leaves occupy, one connector drops from parent to row, and
// children lay left-to-right within that span.
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
	rowH := float64(h) / float64(deep)
	total := 0
	for _, node := range nodes {
		total += smartArtCountLeaves(node)
	}
	if total == 0 {
		return
	}

	// drawAt lays one subtree inside [spanX, spanX+spanW], at its row.
	var drawAt func(node *SmartArtNode, depth int, spanX, spanW float64, colorBase int) int
	drawAt = func(node *SmartArtNode, depth int, spanX, spanW float64, colorBase int) int {
		yTop := int(float64(y) + float64(depth)*rowH)
		yBot := int(float64(y) + (float64(depth)+1)*rowH)
		side := spanW
		if side > float64(w)/4 {
			side = float64(w) / 4
		}
		if side < 16 {
			side = 16
		}
		// box centered over its children's span (over its whole span for a leaf)
		cx := spanX + spanW/2
		rect := image.Rect(int(cx-side/2), yTop+4, int(cx+side/2), yBot-6)
		if len(node.Children) == 0 {
			rect = image.Rect(int(spanX)+2, yTop+4, int(spanX+spanW)-2, yBot-6)
		}
		if rect.Dx() > 4 && rect.Dy() > 4 {
			r.smartArtNode(node, colorBase, colorful, face, rect)
		}
		colorBase++
		if len(node.Children) > 0 {
			midX := int(cx)
			childTop := yTop + rect.Dy() + 4
			r.drawLineAA(midX, yTop+6+rect.Dy(), midX, childTop+int(rowH)/6, color.RGBA{R: 130, G: 130, B: 130, A: 255}, 1)
			kidX := spanX
			for _, child := range node.Children {
				kidLeaves := float64(smartArtCountLeaves(child))
				kidSpan := spanW * kidLeaves / float64(smartArtCountLeaves(node))
				drawAt(child, depth+1, kidX, kidSpan, colorBase)
				kidX += kidSpan
			}
		}
		return colorBase
	}
	colorBase := 0
	spanX := float64(x)
	for _, node := range nodes {
		spanW := float64(w) * float64(smartArtCountLeaves(node)) / float64(total)
		colorBase = drawAt(node, 0, spanX, spanW, colorBase)
		spanX += spanW
	}
}
