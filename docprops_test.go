package gopresentation

import (
	"strings"
	"testing"
	"time"
)

// Document properties: the parts of docProps/ that carried the caller's data
// into the file and then lost it.
//
// Two defects, both of the "the writer emits it and nobody reads it back" kind:
//
//   - API.md documents SetCustomProperty beside GetCustomPropertyValue as one
//     pair. No code wrote docProps/custom.xml at all, so a custom property was
//     readable back through the in-memory API and gone the moment the
//     presentation was saved — and because the part was also never declared in
//     [Content_Types].xml, there was nothing in the package to notice.
//   - docProps/app.xml has always been written, holding the Company. Nothing
//     read it back, so a company name survived exactly one save.
//
// The custom part is checked structurally as well: it is declared in the
// content types and related from _rels/.rels only when it exists, because a
// declared part a package does not contain is what makes PowerPoint offer to
// repair the file.

// customPropsPresentation returns a presentation with one of every property
// type this library models.
func customPropsPresentation() *Presentation {
	p := New()
	props := p.GetDocumentProperties()
	props.SetCustomProperty("version", "1.0", PropertyTypeString)
	props.SetCustomProperty("draft", true, PropertyTypeBoolean)
	props.SetCustomProperty("count", 7, PropertyTypeInteger)
	props.SetCustomProperty("ratio", 1.5, PropertyTypeFloat)
	props.SetCustomProperty("released", time.Date(2026, 9, 12, 8, 30, 0, 0, time.UTC), PropertyTypeDate)
	props.SetCustomProperty("a & b", "escaped", PropertyTypeString)
	return p
}

// TestCustomPropertiesReachThePackage pins the part, its content type and its
// relationship, and the value element each PropertyType produces.
func TestCustomPropertiesReachThePackage(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, customPropsPresentation()))

	part, ok := parts["docProps/custom.xml"]
	if !ok {
		t.Fatal("no docProps/custom.xml in the package: the custom properties were not written")
	}
	custom := string(part)

	if !strings.Contains(string(parts["[Content_Types].xml"]), "/docProps/custom.xml") {
		t.Error("[Content_Types].xml does not declare docProps/custom.xml")
	}
	if !strings.Contains(string(parts["_rels/.rels"]), "docProps/custom.xml") {
		t.Error("_rels/.rels does not relate docProps/custom.xml: nothing can find the part")
	}

	for _, want := range []string{
		"<vt:lpwstr>1.0</vt:lpwstr>",
		"<vt:bool>true</vt:bool>",
		"<vt:i4>7</vt:i4>",
		"<vt:r8>1.5</vt:r8>",
		"<vt:filetime>2026-09-12T08:30:00Z</vt:filetime>",
		// The name is an attribute, so it is escaped like any other.
		`name="a &amp; b"`,
	} {
		if !strings.Contains(custom, want) {
			t.Errorf("docProps/custom.xml has no %s:\n%s", want, custom)
		}
	}
	if !strings.Contains(custom, `fmtid="`+customPropsFmtid+`"`) {
		t.Errorf("a property is missing the format id the spec requires:\n%s", custom)
	}
	// pid 1 is reserved, so the first property is pid 2.
	if !strings.Contains(custom, `pid="2"`) {
		t.Errorf("the first property does not start at pid 2:\n%s", custom)
	}
}

// TestCustomPropertiesAreAbsentWhenUnset is the other half of the declaration
// rule: a package with no custom properties must not declare the part either.
func TestCustomPropertiesAreAbsentWhenUnset(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))

	if _, ok := parts["docProps/custom.xml"]; ok {
		t.Error("a presentation with no custom properties wrote docProps/custom.xml")
	}
	if strings.Contains(string(parts["[Content_Types].xml"]), "/docProps/custom.xml") {
		t.Error("[Content_Types].xml declares a part the package does not contain")
	}
	if strings.Contains(string(parts["_rels/.rels"]), "docProps/custom.xml") {
		t.Error("_rels/.rels relates a part the package does not contain")
	}
}

// TestCustomPropertiesSurviveRoundTrip is the read half: value and declared
// type both have to come back.
func TestCustomPropertiesSurviveRoundTrip(t *testing.T) {
	back := roundTrip(t, customPropsPresentation()).GetDocumentProperties()

	if !back.IsCustomPropertySet("version") {
		t.Fatal("the string property was lost by the round trip")
	}
	if got := back.GetCustomPropertyValue("version"); got != "1.0" {
		t.Errorf("version = %#v, want %q", got, "1.0")
	}
	if got := back.GetCustomPropertyType("version"); got != PropertyTypeString {
		t.Errorf("version type = %d, want %d (string)", got, PropertyTypeString)
	}

	if v, ok := back.GetCustomPropertyValue("draft").(bool); !ok || !v {
		t.Errorf("draft = %#v, want true", back.GetCustomPropertyValue("draft"))
	}
	if got := back.GetCustomPropertyType("draft"); got != PropertyTypeBoolean {
		t.Errorf("draft type = %d, want %d (boolean)", got, PropertyTypeBoolean)
	}

	if v, ok := back.GetCustomPropertyValue("count").(int64); !ok || v != 7 {
		t.Errorf("count = %#v, want int64(7)", back.GetCustomPropertyValue("count"))
	}
	if got := back.GetCustomPropertyType("count"); got != PropertyTypeInteger {
		t.Errorf("count type = %d, want %d (integer)", got, PropertyTypeInteger)
	}

	if v, ok := back.GetCustomPropertyValue("ratio").(float64); !ok || v != 1.5 {
		t.Errorf("ratio = %#v, want 1.5", back.GetCustomPropertyValue("ratio"))
	}
	if got := back.GetCustomPropertyType("ratio"); got != PropertyTypeFloat {
		t.Errorf("ratio type = %d, want %d (float)", got, PropertyTypeFloat)
	}

	want := time.Date(2026, 9, 12, 8, 30, 0, 0, time.UTC)
	if v, ok := back.GetCustomPropertyValue("released").(time.Time); !ok || !v.Equal(want) {
		t.Errorf("released = %#v, want %v", back.GetCustomPropertyValue("released"), want)
	}

	if got := back.GetCustomPropertyValue("a & b"); got != "escaped" {
		t.Errorf("the property whose name needs escaping = %#v, want %q", got, "escaped")
	}
	if n := len(back.GetCustomProperties()); n != 6 {
		t.Errorf("the round trip came back with %d custom properties, want 6", n)
	}
}

// TestCustomPropertyPartIsDeterministic writes the same presentation twice.
//
// The model holds the properties in a map, so walking it directly emits them in
// a different order on every save. A package that differs from itself for no
// reason makes every save look like a change and makes the part useless to diff.
func TestCustomPropertyPartIsDeterministic(t *testing.T) {
	p := customPropsPresentation()
	first := zipParts(t, writeToBytes(t, p))["docProps/custom.xml"]
	second := zipParts(t, writeToBytes(t, p))["docProps/custom.xml"]

	if string(first) != string(second) {
		t.Errorf("two saves of the same presentation produced different custom-property parts:\n%s\n---\n%s", first, second)
	}
}

// TestCompanySurvivesRoundTrip covers docProps/app.xml, which the writer has
// always emitted and nothing read back.
func TestCompanySurvivesRoundTrip(t *testing.T) {
	p := New()
	p.GetDocumentProperties().Company = "Vantagics"

	parts := zipParts(t, writeToBytes(t, p))
	if !strings.Contains(string(parts["docProps/app.xml"]), "<Company>Vantagics</Company>") {
		t.Fatalf("docProps/app.xml does not hold the company:\n%s", parts["docProps/app.xml"])
	}

	back := roundTrip(t, p).GetDocumentProperties()
	if back.Company != "Vantagics" {
		t.Errorf("company after round trip = %q, want %q", back.Company, "Vantagics")
	}
}

// TestSetCustomPropertyOnZeroValueProperties is the guard for a
// DocumentProperties a caller built rather than got from
// NewDocumentProperties: its map is nil, and assigning into a nil map panics.
func TestSetCustomPropertyOnZeroValueProperties(t *testing.T) {
	var dp DocumentProperties
	dp.SetCustomProperty("k", "v", PropertyTypeString)

	if got := dp.GetCustomPropertyValue("k"); got != "v" {
		t.Errorf("value = %#v, want %q", got, "v")
	}
}
