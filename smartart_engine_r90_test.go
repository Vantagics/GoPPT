package gopresentation

import (
	"testing"
)

// The SmartArt data model walk.
//
// A parse must yield the doc point's content children in srcOrd order with
// their texts, transit points skipped, and the layout family named. The
// corpus probe is fill-color-list (three content nodes under the doc point,
// loTypeId hList1) plus an org-chart hierarchy with nested levels.
func TestSmartArtModelFillList(t *testing.T) {
	model, err := parseSmartArtModel([]byte(smartArtDataXML))
	if err != nil {
		t.Fatalf("parse: %v", err)
	}
	if model.Algo != "lin" {
		t.Errorf("family = %q, want lin (hList1 maps to it)", model.Algo)
	}
	kids := model.Root.Children
	if len(kids) != 3 {
		t.Fatalf("doc children = %d, want the three content nodes", len(kids))
	}
	for i, want := range []string{"A", "B", "C"} {
		if kids[i].Text != want {
			t.Errorf("child %d text = %q, want %q", i, kids[i].Text, want)
		}
	}
}

func TestSmartArtModelHierarchy(t *testing.T) {
	model, err := parseSmartArtModel([]byte(smartArtHierXML))
	if err != nil {
		t.Fatalf("parse: %v", err)
	}
	if model.Algo != "hier" {
		t.Errorf("family = %q, want hier (orgChart maps to it)", model.Algo)
	}
	root := model.Root.Children
	if len(root) != 1 || root[0].Text != "A" {
		t.Fatalf("top = %#v, want one child A", root)
	}
	if len(root[0].Children) != 2 {
		t.Fatalf("A's children = %d, want B1 and B2", len(root[0].Children))
	}
}

func TestSmartArtFamilyMapping(t *testing.T) {
	for lo, want := range map[string]string{
		"hList1":    "lin",
		"cycle2":    "cycle",
		"arrow5":    "cycle",
		"orgChart1": "hier",
		"":          "snake",
	} {
		if got := smartArtFamily(lo); got != want {
			t.Errorf("family(%q) = %q, want %q", lo, got, want)
		}
	}
}

// smartArtDataXML is fill-color-list's data part reduced to its points and
// links: a doc point naming hList1 and three content nodes A/B/C.
const smartArtDataXML = `<?xml version="1.0"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dgm:ptLst>
    <dgm:pt modelId="{root}" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/hList1" loCatId="list" csCatId="colorful"/></dgm:pt>
    <dgm:pt modelId="{n1}"><dgm:t><a:p><a:r><a:t>A</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="{n2}"><dgm:t><a:p><a:r><a:t>B</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="{n3}"><dgm:t><a:p><a:r><a:t>C</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="{t1}" type="parTrans" cxnId="{c1}"/>
  </dgm:ptLst>
  <dgm:cxnLst>
    <dgm:cxn modelId="{c1}" srcId="{root}" destId="{n1}" srcOrd="0"/>
    <dgm:cxn modelId="{c2}" srcId="{root}" destId="{n2}" srcOrd="1"/>
    <dgm:cxn modelId="{c3}" srcId="{root}" destId="{n3}" srcOrd="2"/>
  </dgm:cxnLst>
</dgm:dataModel>`

// smartArtHierXML is a mini org chart: A over B1/B2.
const smartArtHierXML = `<?xml version="1.0"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <dgm:ptLst>
    <dgm:pt modelId="{root}" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/orgChart1" loCatId="hierarchy"/></dgm:pt>
    <dgm:pt modelId="{nA}"><dgm:t><a:p><a:r><a:t>A</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="{nB1}"><dgm:t><a:p><a:r><a:t>B1</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="{nB2}"><dgm:t><a:p><a:r><a:t>B2</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst>
    <dgm:cxn modelId="{c0}" srcId="{root}" destId="{nA}" srcOrd="0"/>
    <dgm:cxn modelId="{c1}" srcId="{nA}" destId="{nB1}" srcOrd="0"/>
    <dgm:cxn modelId="{c2}" srcId="{nA}" destId="{nB2}" srcOrd="1"/>
  </dgm:cxnLst>
</dgm:dataModel>`
