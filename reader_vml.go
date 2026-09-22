package gopresentation

import (
	"archive/zip"
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// relTypeVMLDrawing is the relationship type of a legacy VML drawing part.
const relTypeVMLDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"

// PowerPoint stores freehand ink annotations (marks made with the pen tool)
// in legacy VML drawing parts: ppt/drawings/vmlDrawingN.vml, related from the
// slide's .rels but referenced by no element of the slide XML itself. The
// shapes are positioned absolutely on the slide and drawn on top of it.
//
// Each <v:shape> carries the geometry twice: a `path` attribute in VML path
// syntax (the vector centerline any renderer can stroke) and an `o:ink`
// attribute holding the binary ISF blob PowerPoint actually renders from.
// Experiments against COM exports (r32) show PowerPoint ignores the VML
// strokecolor and strokeweight for annotation ink — both live in the ISF —
// but the ISF values mirror the declared attributes in real files, so the
// declared ones are what this library renders with.
//
// The ink is modelled as a stroked freeform connector (LineShape with a
// custom path): the reader parses it, the renderer strokes it, and the writer
// already serialises custom paths as <a:custGeom>, so the marks survive a
// read/write round trip as ordinary DrawingML.
func (r *PPTXReader) readSlideInk(zr *zip.Reader, slide *Slide, rels []xmlRelForRead, slidePath string) {
	for _, rel := range rels {
		if rel.Type != relTypeVMLDrawing {
			continue
		}
		target := rel.Target
		if !strings.HasPrefix(target, "ppt/") {
			dir := strings.TrimSuffix(slidePath, "/"+lastPathComponent(slidePath))
			target = resolveRelativePath(dir, target)
		}
		data, err := readFileFromZip(zr, target)
		if err != nil {
			continue
		}
		for _, shape := range parseVMLDrawing(data) {
			slide.shapes = append(slide.shapes, shape)
		}
	}
}

// parseVMLDrawing extracts every stroked <v:shape> from a VML drawing part,
// in document order. Shapes carrying fills are not representable as a stroked
// connector and come back as an UnsupportedShape placeholder instead, so the
// gap stays visible rather than silently empty.
func parseVMLDrawing(data []byte) []Shape {
	var out []Shape
	var cur *vmlShape
	groupDepth := 0

	flush := func() {
		if cur == nil {
			return
		}
		if s := cur.shape(); s != nil {
			out = append(out, s)
		}
		cur = nil
	}

	decoder := xml.NewDecoder(strings.NewReader(string(data)))
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "group":
				groupDepth++
			case "shape":
				if groupDepth == 0 {
					flush()
					cur = parseVMLShapeAttrs(t)
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "group":
				groupDepth--
			case "shape":
				flush()
			}
		}
	}
	flush()
	return out
}

// vmlShape holds one <v:shape>'s rendering-relevant attributes.
type vmlShape struct {
	id           string
	leftPt       float64
	topPt        float64
	widthPt      float64
	heightPt     float64
	hasBox       bool
	coordOriginX int64
	coordOriginY int64
	coordSizeX   int64
	coordSizeY   int64
	path         string
	filled       bool
	stroked      bool
	strokeColor  string
	strokeWeight string
}

// parseVMLShapeAttrs reads the attributes of one <v:shape> start tag.
func parseVMLShapeAttrs(start xml.StartElement) *vmlShape {
	v := &vmlShape{filled: true, stroked: true, strokeColor: "#000000"}
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			v.id = attr.Value
		case "style":
			v.parseStyle(attr.Value)
		case "coordorigin":
			if x, y, ok := parseVMLPoint(attr.Value); ok {
				v.coordOriginX, v.coordOriginY = x, y
			}
		case "coordsize":
			if x, y, ok := parseVMLPoint(attr.Value); ok {
				v.coordSizeX, v.coordSizeY = x, y
			}
		case "path":
			v.path = attr.Value
		case "filled":
			v.filled = attr.Value != "f"
		case "stroked":
			v.stroked = attr.Value != "f"
		case "strokecolor":
			v.strokeColor = attr.Value
		case "strokeweight":
			v.strokeWeight = attr.Value
		}
	}
	return v
}

// parseStyle reads left/top/width/height out of a VML style attribute. The
// value spans lines in real files ("style='position:absolute;left:151pt;\n
// top:226.5pt;...'"), so the text is normalised first.
func (v *vmlShape) parseStyle(style string) {
	style = strings.Join(strings.Fields(style), " ")
	for _, key := range []string{"left", "top", "width", "height"} {
		i := strings.Index(style, key+":")
		if i < 0 {
			continue
		}
		rest := style[i+len(key)+1:]
		if j := strings.IndexAny(rest, ";}"); j >= 0 {
			rest = rest[:j]
		}
		rest = strings.TrimSpace(rest)
		val, unit := parseVMLLength(rest)
		if unit == "px" {
			val *= 0.75
		}
		switch key {
		case "left":
			v.leftPt = val
		case "top":
			v.topPt = val
		case "width":
			v.widthPt = val
		case "height":
			v.heightPt = val
		}
	}
	v.hasBox = v.widthPt > 0 && v.heightPt > 0
}

// parseVMLLength splits "151pt" into 151 and "pt". A bare number is points —
// PowerPoint always writes the unit; lengths without one are interpreted in
// the style's default coordinate, which for positioned shapes is points too.
func parseVMLLength(s string) (float64, string) {
	s = strings.TrimSpace(s)
	unit := "pt"
	for _, u := range []string{"pt", "px", "mm", "cm", "in"} {
		if strings.HasSuffix(s, u) {
			unit = u
			s = strings.TrimSuffix(s, u)
			break
		}
	}
	val, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil {
		return 0, unit
	}
	return val, unit
}

// parseVMLPoint reads "x,y" (VML coordorigin/coordsize).
func parseVMLPoint(s string) (int64, int64, bool) {
	parts := strings.SplitN(s, ",", 2)
	if len(parts) != 2 {
		return 0, 0, false
	}
	x, errX := strconv.ParseInt(strings.TrimSpace(parts[0]), 10, 64)
	y, errY := strconv.ParseInt(strings.TrimSpace(parts[1]), 10, 64)
	return x, y, errX == nil && errY == nil
}

// shape converts the parsed attributes into a model shape: a stroked freeform
// connector for stroke-only ink, an unsupported placeholder when a fill is
// involved, or nil when nothing would be drawn at all.
func (v *vmlShape) shape() Shape {
	if v.path == "" || (!v.stroked && !v.filled) {
		return nil
	}
	if v.filled {
		return NewUnsupportedShape("VML ink shape with a fill is not representable as a stroked connector")
	}
	if !v.hasBox {
		// No style box: the path coordinate space is used directly, in the
		// units the file chose. This does not occur in PowerPoint output.
		return nil
	}
	cmds := parseVMLPath(v.path, v.coordOriginX, v.coordOriginY)
	if len(cmds) == 0 {
		return nil
	}

	weightPt := 1.0
	if v.strokeWeight != "" {
		val, unit := parseVMLLength(v.strokeWeight)
		if unit == "px" {
			val *= 0.75
		}
		if val > 0 {
			weightPt = val
		}
	}

	line := NewLineShape()
	line.SetName("Ink " + v.id)
	line.SetPosition(int64(math.Round(v.leftPt*12700)), int64(math.Round(v.topPt*12700)))
	line.SetSize(int64(math.Round(v.widthPt*12700)), int64(math.Round(v.heightPt*12700)))
	line.lineWidthEMU = int(math.Round(weightPt * 12700))
	line.lineColor = NewColor(v.strokeColor)
	line.customPath = &CustomGeomPath{
		Width:    v.coordSizeX,
		Height:   v.coordSizeY,
		Commands: cmds,
	}
	return line
}

// parseVMLPath compiles a VML `path` attribute into absolute custom-geometry
// commands. Two file quirks are load-bearing:
//
//   - An empty field between separators means zero ("v6,,12,-1,18,-1" is a
//     six-number curve with dy1=0). Splitting on the separators as a run
//     instead of per-comma silently drops those zeros and every stroke
//     distorts.
//   - `v` is the *relative* cubic bezier: all three points offset from the
//     current point (SVG `c` semantics, not chained deltas). Rasterising both
//     readings against the PowerPoint export picks SVG semantics by a wide
//     margin (r32 probe: precision 0.945 vs 0.609).
//
// coordorigin is subtracted so the path lives in 0..coordsize, the space the
// renderer maps onto the shape's frame.
func parseVMLPath(path string, originX, originY int64) []PathCommand {
	type token struct {
		cmd byte
		num float64
	}
	var toks []token
	i, n := 0, len(path)
	for i < n {
		ch := path[i]
		switch {
		case ch >= 'a' && ch <= 'z' || ch >= 'A' && ch <= 'Z':
			toks = append(toks, token{cmd: ch})
			i++
		default:
			j := i
			for j < n {
				c := path[j]
				if c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' {
					break
				}
				j++
			}
			run := path[i:j]
			// Per-comma split: an empty field between two commas is a
			// legitimate zero, not noise.
			fields := strings.Split(run, ",")
			for _, f := range fields {
				f = strings.TrimSpace(f)
				if f == "" {
					toks = append(toks, token{num: 0})
					continue
				}
				val, err := strconv.ParseFloat(f, 64)
				if err != nil {
					// Space-separated fields land here one string at a time.
					for _, sf := range strings.Fields(f) {
						if v, err := strconv.ParseFloat(sf, 64); err == nil {
							toks = append(toks, token{num: v})
						}
					}
					continue
				}
				toks = append(toks, token{num: val})
			}
			i = j
		}
	}

	var out []PathCommand
	var curX, curY float64
	at := func(k int) float64 { return toks[k].num }

	// emit rounds an absolute path-space point into the coordorigin-normalised
	// integer space the model carries.
	emit := func(x, y float64) PathPoint {
		return PathPoint{X: int64(math.Round(x)) - originX, Y: int64(math.Round(y)) - originY}
	}
	newCmd := func(typ string, pts ...PathPoint) PathCommand {
		return PathCommand{Type: typ, Pts: pts}
	}

	i = 0
	for i < len(toks) {
		t := toks[i]
		if t.cmd == 0 {
			i++ // stray number outside any command
			continue
		}
		i++
		switch t.cmd {
		case 'm': // absolute moveto; extra pairs repeat
			for i+1 < len(toks) && toks[i].cmd == 0 && toks[i+1].cmd == 0 {
				curX, curY = at(i), at(i+1)
				out = append(out, newCmd("moveTo", emit(curX, curY)))
				i += 2
			}
		case 't': // relative moveto
			for i+1 < len(toks) && toks[i].cmd == 0 && toks[i+1].cmd == 0 {
				curX += at(i)
				curY += at(i + 1)
				out = append(out, newCmd("moveTo", emit(curX, curY)))
				i += 2
			}
		case 'l': // absolute lineto; pairs repeat
			for i+1 < len(toks) && toks[i].cmd == 0 && toks[i+1].cmd == 0 {
				curX, curY = at(i), at(i+1)
				out = append(out, newCmd("lnTo", emit(curX, curY)))
				i += 2
			}
		case 'r': // relative lineto
			for i+1 < len(toks) && toks[i].cmd == 0 && toks[i+1].cmd == 0 {
				curX += at(i)
				curY += at(i + 1)
				out = append(out, newCmd("lnTo", emit(curX, curY)))
				i += 2
			}
		case 'c': // absolute cubic; sextets repeat
			for i+5 < len(toks) && toks[i].cmd == 0 && toks[i+1].cmd == 0 &&
				toks[i+2].cmd == 0 && toks[i+3].cmd == 0 && toks[i+4].cmd == 0 && toks[i+5].cmd == 0 {
				out = append(out, newCmd("cubicBezTo",
					emit(at(i), at(i+1)), emit(at(i+2), at(i+3)), emit(at(i+4), at(i+5))))
				curX, curY = at(i+4), at(i+5)
				i += 6
			}
		case 'v': // relative cubic: all three points from the current point
			for i+5 < len(toks) && toks[i].cmd == 0 && toks[i+1].cmd == 0 &&
				toks[i+2].cmd == 0 && toks[i+3].cmd == 0 && toks[i+4].cmd == 0 && toks[i+5].cmd == 0 {
				out = append(out, newCmd("cubicBezTo",
					emit(curX+at(i), curY+at(i+1)),
					emit(curX+at(i+2), curY+at(i+3)),
					emit(curX+at(i+4), curY+at(i+5))))
				curX += at(i + 4)
				curY += at(i + 5)
				i += 6
			}
		case 'x': // close
			out = append(out, newCmd("close"))
		case 'e': // end of subpath: the next moveTo starts the next one
			// no model equivalent; stroking splits at moveTo
		default:
			// Unknown command (qb, ae, ...): skip its numbers by stopping the
			// command loop; the next letter token resumes parsing.
		}
	}
	return out
}
