package gopresentation

import (
	"bytes"
	"encoding/xml"
	"path"
	"strings"
	"testing"
)

// Integrity of a package the writer produced.
//
// Reader-side tests can only catch content that came back wrong. These start
// from what was written: the shapes are still there, and every relationship the
// package refers to actually resolves. That second check covers a whole class of
// bug at once — a shape emitted with an rId nothing defines, or pointing at a
// chart part that was never written — which is the kind of damage that makes
// PowerPoint announce a file needs repair.

// presentationWithGroupedContent puts one of every relationship-consuming shape
// kind inside a group, including a group nested in a group.
func presentationWithGroupedContent(t *testing.T) *Presentation {
	t.Helper()
	p := New()

	outer := p.GetActiveSlide().CreateGroupShape()
	outer.BaseShape.SetOffsetX(100000).SetOffsetY(100000)
	outer.BaseShape.SetWidth(8000000).SetHeight(6000000)

	picture := NewDrawingShape()
	picture.SetImageData(twoTonePNG(t), "image/png")
	picture.BaseShape.SetOffsetX(200000).SetOffsetY(200000)
	picture.BaseShape.SetWidth(1000000).SetHeight(1000000)
	outer.AddShape(picture)

	chart := NewChartShape()
	chart.BaseShape.SetOffsetX(2000000).SetOffsetY(200000)
	chart.BaseShape.SetWidth(3000000).SetHeight(2000000)
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("S", []string{"A", "B"}, []float64{1, 2}))
	chart.GetPlotArea().SetType(bar)
	outer.AddShape(chart)

	linked := NewRichTextShape()
	linked.BaseShape.SetOffsetX(200000).SetOffsetY(3000000)
	linked.BaseShape.SetWidth(4000000).SetHeight(1000000)
	run := linked.CreateTextRun("external link")
	run.SetHyperlink(NewHyperlink("https://example.com/"))
	outer.AddShape(linked)

	inner := NewGroupShape()
	inner.BaseShape.SetOffsetX(4000000).SetOffsetY(3000000)
	inner.BaseShape.SetWidth(3000000).SetHeight(2000000)
	leaf := NewRichTextShape()
	leaf.BaseShape.SetOffsetX(4000000).SetOffsetY(3000000)
	leaf.BaseShape.SetWidth(3000000).SetHeight(1000000)
	leaf.CreateTextRun("NESTED-LEAF")
	inner.AddShape(leaf)
	outer.AddShape(inner)

	return p
}

// TestGroupedContentSurvivesRoundTrip covers what the group writer used to drop.
// Every shape kind has to be written from inside a group, not only from the top
// level of the slide: a missing case does not error, it just loses the child.
func TestGroupedContentSurvivesRoundTrip(t *testing.T) {
	pres := roundTrip(t, presentationWithGroupedContent(t))
	shapes := pres.GetAllSlides()[0].GetShapes()

	var pictures, charts, groups int
	for _, s := range flattenShapes(shapes) {
		switch s.(type) {
		case *DrawingShape:
			pictures++
		case *ChartShape:
			charts++
		case *GroupShape:
			groups++
		}
	}
	if pictures != 1 {
		t.Errorf("pictures after round trip = %d, want 1", pictures)
	}
	if charts != 1 {
		t.Errorf("charts after round trip = %d, want 1 (a chart inside a group used to be dropped)", charts)
	}
	if groups != 2 {
		t.Errorf("groups after round trip = %d, want 2 (a group inside a group used to be dropped)", groups)
	}
	if !treeContainsText(shapes, "NESTED-LEAF") {
		t.Error("the leaf text of the nested group was lost")
	}

	// The rebuilt package must still render.
	opts := DefaultRenderOptions()
	opts.Width = 320
	if _, err := pres.SlidesToImages(opts); err != nil {
		t.Errorf("render after round trip: %v", err)
	}
}

// TestWrittenPackageHasNoDanglingRelationships checks the package as a whole:
// every r:id, r:embed and r:link in every part must be defined by that part's
// relationship file, and every internal target must name a part that exists.
func TestWrittenPackageHasNoDanglingRelationships(t *testing.T) {
	pkg := writeToBytes(t, presentationWithGroupedContent(t))
	parts := zipParts(t, pkg)

	checked := 0
	for name, content := range parts {
		if !strings.HasSuffix(name, ".xml") || strings.HasSuffix(name, ".rels") {
			continue
		}
		rels, relsErr := readPackageRelationships(parts, relsPathFor(name))
		if relsErr != nil {
			t.Errorf("parse relationship file for %s: %v", name, relsErr)
			continue
		}
		for _, id := range relationshipRefs(t, name, content) {
			entry, ok := rels[id]
			if !ok {
				t.Errorf("%s references %s, which its relationship file does not define", name, id)
				continue
			}
			checked++
			if entry.external {
				continue
			}
			target := resolveRelTarget(relsPathFor(name), entry.target)
			if _, ok := parts[target]; !ok {
				t.Errorf("%s: relationship %s points at %s, which is not in the package", name, id, target)
			}
		}
	}

	// The fixture has a picture, a chart and an external hyperlink, so a run that
	// resolved nothing would mean this test stopped exercising the writer.
	if checked < 3 {
		t.Fatalf("only %d relationship references were checked; the fixture is not exercising the writer", checked)
	}
}

// relationshipRefs returns the relationship ids a part refers to. The attributes
// live in the officeDocument relationships namespace, which is what the r: prefix
// maps to in every part the writer emits.
func relationshipRefs(t *testing.T, name string, content []byte) []string {
	t.Helper()
	var refs []string
	dec := xml.NewDecoder(bytes.NewReader(content))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		for _, attr := range start.Attr {
			if attr.Name.Space != nsOfficeDocRels {
				continue
			}
			switch attr.Name.Local {
			case "id", "embed", "link":
				refs = append(refs, attr.Value)
			}
		}
	}
	return refs
}

// relEntry is one relationship as defined in a .rels part.
type relEntry struct {
	target   string
	external bool
}

// relsPathFor returns the relationship part that describes a given part.
func relsPathFor(part string) string {
	dir, base := path.Split(part)
	return dir + "_rels/" + base + ".rels"
}

// resolveRelTarget turns a relationship target into the package path it names.
// Targets are relative to the directory of the part that owns the relationship,
// which is the directory above the _rels folder.
func resolveRelTarget(relsPart string, target string) string {
	base := path.Dir(path.Dir(relsPart))
	return path.Join(base, target)
}

// readPackageRelationships parses one .rels part into an id -> relationship map.
func readPackageRelationships(parts map[string][]byte, relsPart string) (map[string]relEntry, error) {
	data, ok := parts[relsPart]
	if !ok {
		return nil, nil
	}
	var doc struct {
		Relationships []struct {
			ID         string `xml:"Id,attr"`
			Target     string `xml:"Target,attr"`
			TargetMode string `xml:"TargetMode,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal(data, &doc); err != nil {
		return nil, err
	}
	out := make(map[string]relEntry, len(doc.Relationships))
	for _, rel := range doc.Relationships {
		out[rel.ID] = relEntry{target: rel.Target, external: rel.TargetMode == "External"}
	}
	return out, nil
}
