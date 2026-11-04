package gopresentation

import (
	"bytes"
	"encoding/xml"
	"errors"
	"sort"
	"strconv"
	"strings"
)

// A SmartArt data model, read from a diagram's data part when the package
// carries no cached drawing: the tree of content points with their texts and
// the layout family the deck named (the root point's loTypeId, checked
// against the layout part's own <dgm:alg>).
//
// Implementing the full layout engine — constraint solving over the layout
// definition's rule tree — is out of scope. What v1 carries is the geometry
// family (lin, snake, cycle, hier) plus per-node fills cycled over the theme
// accents, which turns the "Unsupported: SmartArt diagram" gap into a
// recognisable diagram for the LibreOffice-authored decks whose packages
// never cache a drawing.
type SmartArtModel struct {
	// Algo is the layout family: "lin", "snake", "cycle" or "hier".
	Algo string
	// Root's children are the content nodes hanging off the doc point.
	Root *SmartArtNode
	// Colorful marks the deck's quick style as per-node accent cycling
	// (csCatId "colorful"): each node takes the next palette slot, and list
	// diagrams paint a coloured header band over a pale body. Without it the
	// quick style is uniform — every node wears one accent — which is what
	// the LibreOffice-authored sample decks' Office exports show.
	Colorful bool
}

// SmartArtNode is one content point of the model.
type SmartArtNode struct {
	Text     string
	Children []*SmartArtNode
}

// smartartPoint is one <dgm:pt> of the data part as parsed.
type smartartPoint struct {
	id   string
	typ  string // "", "node", "doc", "asst", and the transit/pres kinds
	text string
}

var errSmartArtNoModel = errors.New("smartart: no usable data model")

// parseSmartArtModel reads a data part into a SmartArtModel. An error means
// no usable tree was found and the caller keeps its stand-in.
func parseSmartArtModel(data []byte) (*SmartArtModel, error) {
	pts, loType, colorful, err := parseSmartArtPoints(data)
	if err != nil {
		return nil, err
	}
	cxns, err := parseSmartArtConnections(data)
	if err != nil {
		return nil, err
	}
	model, err := buildSmartArtModel(pts, cxns, loType)
	if err != nil {
		return nil, err
	}
	model.Colorful = colorful
	return model, nil
}

// parseSmartArtPoints collects the points and the doc point's loTypeId.
func parseSmartArtPoints(data []byte) ([]smartartPoint, string, bool, error) {
	var (
		pts           []smartartPoint
		cur           smartartPoint
		inPt          bool
		inT           bool
		textParts     []string
		loType        string
		modelColorful bool
	)
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pt":
				inPt = true
				cur = smartartPoint{}
				textParts = nil
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "modelId":
						cur.id = attr.Value
					case "type":
						cur.typ = attr.Value
					}
				}
			case "prSet":
				if inPt && cur.typ == "doc" {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "loTypeId":
							if i := strings.LastIndex(attr.Value, "layout/"); i >= 0 {
								loType = attr.Value[i+len("layout/"):]
							}
						case "csCatId":
							// The colourful quick style cycles the accents per
							// node; every other family paints them uniformly.
							if strings.HasPrefix(attr.Value, "colorful") {
								modelColorful = true
							}
						}
					}
				}
			case "t":
				if inPt {
					inT = true
					textParts = nil
				}
			}
		case xml.CharData:
			if inT {
				textParts = append(textParts, string(t))
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "pt":
				if inPt {
					cur.text = strings.Join(textParts, "")
					pts = append(pts, cur)
					inPt = false
				}
			case "t":
				inT = false
			}
		}
	}
	if len(pts) == 0 {
		return nil, "", false, errSmartArtNoModel
	}
	return pts, loType, modelColorful, nil
}

// parseSmartArtConnections collects the parent-child links: an omitted type
// means "parOf" per ECMA-376; presentation and transition links carry
// explicit types and are skipped.
func parseSmartArtConnections(data []byte) ([][3]string, error) {
	var cxns [][3]string
	decoder := xml.NewDecoder(bytes.NewReader(data))
	attrs := map[string]string{}
	inCxn := false
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Local == "cxn" {
				inCxn = true
				attrs = map[string]string{}
				for _, attr := range t.Attr {
					attrs[attr.Name.Local] = attr.Value
				}
			}
		case xml.EndElement:
			if t.Name.Local == "cxn" && inCxn {
				typ := attrs["type"]
				if typ == "" || typ == "parOf" {
					cxns = append(cxns, [3]string{attrs["srcId"], attrs["destId"], attrs["srcOrd"]})
				}
				inCxn = false
			}
		}
	}
	return cxns, nil
}

// buildSmartArtModel assembles the tree from points and links.
func buildSmartArtModel(pts []smartartPoint, cxns [][3]string, loType string) (*SmartArtModel, error) {
	byID := make(map[string]int, len(pts))
	for i, p := range pts {
		byID[p.id] = i
	}
	var docID string
	for _, p := range pts {
		if p.typ == "doc" {
			docID = p.id
			break
		}
	}
	if docID == "" {
		return nil, errSmartArtNoModel
	}
	var childLists = map[string][]struct {
		id  string
		ord int
	}{}
	for _, cx := range cxns {
		if _, ok := byID[cx[0]]; !ok {
			continue
		}
		ord, _ := strconv.Atoi(cx[2])
		childLists[cx[0]] = append(childLists[cx[0]], struct {
			id  string
			ord int
		}{cx[1], ord})
	}
	for id := range childLists {
		list := childLists[id]
		sort.SliceStable(list, func(i, j int) bool { return list[i].ord < list[j].ord })
		childLists[id] = list
	}

	var build func(id string) *SmartArtNode
	build = func(id string) *SmartArtNode {
		idx, ok := byID[id]
		if !ok {
			return nil
		}
		p := pts[idx]
		if p.typ == "parTrans" || p.typ == "sibTrans" || p.typ == "pres" {
			return nil
		}
		n := &SmartArtNode{Text: p.text}
		for _, e := range childLists[id] {
			if child := build(e.id); child != nil {
				n.Children = append(n.Children, child)
			}
		}
		return n
	}
	var roots []*SmartArtNode
	for _, e := range childLists[docID] {
		if n := build(e.id); n != nil {
			roots = append(roots, n)
		}
	}
	// A group node that carries neither text nor kept children renders as an
	// empty box; drop it so spacer points LibreOffice emits stay invisible.
	var prune func(n *SmartArtNode) bool
	prune = func(n *SmartArtNode) bool {
		kept := n.Children[:0]
		for _, c := range n.Children {
			if prune(c) {
				kept = append(kept, c)
			}
		}
		n.Children = kept
		return n.Text != "" || len(n.Children) > 0
	}
	filtered := roots[:0]
	for _, r := range roots {
		if prune(r) {
			filtered = append(filtered, r)
		}
	}
	if len(filtered) == 0 {
		return nil, errSmartArtNoModel
	}
	return &SmartArtModel{Algo: smartArtFamily(loType), Root: &SmartArtNode{Children: filtered}}, nil
}

// smartArtFamily maps the deck's declared layout family to one of the four
// geometry families the v1 engine lays out.
func smartArtFamily(loType string) string {
	switch {
	case strings.Contains(loType, "cycle"),
		strings.Contains(loType, "arrow"),
		strings.Contains(loType, "radial"),
		strings.Contains(loType, "circle"):
		return "cycle"
	case strings.Contains(loType, "orgChart"),
		strings.Contains(loType, "hier"),
		strings.Contains(loType, "tree"):
		return "hier"
	case strings.Contains(loType, "hList"),
		strings.Contains(loType, "list"),
		strings.Contains(loType, "lin"):
		return "lin"
	default:
		return "snake"
	}
}
