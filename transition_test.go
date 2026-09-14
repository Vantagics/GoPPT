package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// Slide transitions had a model and a setter but no serialisation on either
// side, so SetTransition stored a value that reached neither the file nor the
// reader. These tests hold the element to the shape pml-animationInfo.xsd gives
// it, and they do it by looking at the emitted XML rather than at the model: a
// round trip through this package cannot tell right markup from wrong, because
// the reader matches on local name and never checks element order, so it accepts
// <transition> where PowerPoint demands <p:transition>, and accepts a transition
// placed after </p:sld> if one were ever written there.

// transitionElementNames is the element each type is written as, spelled out
// here rather than read from the table in slide.go so that a typo in that table
// cannot make this test agree with itself.
var transitionElementNames = map[TransitionType]string{
	TransitionBlinds:    "blinds",
	TransitionChecker:   "checker",
	TransitionCircle:    "circle",
	TransitionComb:      "comb",
	TransitionCover:     "cover",
	TransitionCut:       "cut",
	TransitionDiamond:   "diamond",
	TransitionDissolve:  "dissolve",
	TransitionFade:      "fade",
	TransitionNewsflash: "newsflash",
	TransitionPlus:      "plus",
	TransitionPull:      "pull",
	TransitionPush:      "push",
	TransitionRandom:    "random",
	TransitionRandomBar: "randomBar",
	TransitionSplit:     "split",
	TransitionStrips:    "strips",
	TransitionWedge:     "wedge",
	TransitionWheel:     "wheel",
	TransitionWipe:      "wipe",
	TransitionZoom:      "zoom",
}

// slideWithTransition writes a one-slide package and returns the slide's XML.
func slideWithTransition(t *testing.T, tr *Transition) string {
	t.Helper()
	p := New()
	if tr != nil {
		p.GetActiveSlide().SetTransition(tr)
	}
	parts := zipParts(t, writeToBytes(t, p))
	slide, ok := parts["ppt/slides/slide1.xml"]
	if !ok {
		t.Fatal("the written package has no ppt/slides/slide1.xml")
	}
	return string(slide)
}

// transitionPart cuts the transition markup out of a slide, so that a check for
// an absence cannot be satisfied by some other part of the document.
func transitionPart(t *testing.T, slide string) string {
	t.Helper()
	start := strings.Index(slide, "<p:transition")
	if alt := strings.Index(slide, "<mc:AlternateContent"); alt >= 0 && (start < 0 || alt < start) {
		start = alt
	}
	if start < 0 {
		t.Fatalf("the slide carries no transition markup:\n%s", slide)
	}
	end := strings.LastIndex(slide, "</p:transition>")
	if end >= 0 {
		end += len("</p:transition>")
	}
	if altEnd := strings.LastIndex(slide, "</mc:AlternateContent>"); altEnd >= 0 {
		end = altEnd + len("</mc:AlternateContent>")
	}
	if end < start {
		t.Fatalf("the transition markup is malformed:\n%s", slide)
	}
	return slide[start:end]
}

// TestTransitionSitsBetweenClrMapOvrAndTheEndOfSlide pins the position. CT_Slide
// orders its children cSld, clrMapOvr, transition, timing, extLst, and nothing in
// this package's own reader would notice the element being anywhere else.
func TestTransitionSitsBetweenClrMapOvrAndTheEndOfSlide(t *testing.T) {
	slide := slideWithTransition(t, &Transition{Type: TransitionFade})

	clrMapOvr := strings.Index(slide, "</p:clrMapOvr>")
	transition := strings.Index(slide, "<p:transition")
	end := strings.LastIndex(slide, "</p:sld>")
	if clrMapOvr < 0 || transition < 0 || end < 0 {
		t.Fatalf("the slide is missing one of clrMapOvr, transition or the closing p:sld:\n%s", slide)
	}
	if clrMapOvr > transition {
		t.Errorf("p:transition must follow p:clrMapOvr; clrMapOvr closes at %d and the transition starts at %d:\n%s",
			clrMapOvr, transition, slide)
	}
	if transition > end {
		t.Errorf("p:transition must be inside p:sld:\n%s", slide)
	}
}

// TestTransitionElementNamesCarryTheirPrefix is the check a local-name reader
// cannot make. A bare <fade/> reads back perfectly through this package and is
// rejected by PowerPoint.
func TestTransitionElementNamesCarryTheirPrefix(t *testing.T) {
	part := transitionPart(t, slideWithTransition(t, &Transition{Type: TransitionFade}))

	if !strings.Contains(part, "<p:transition") {
		t.Errorf("the transition element is not <p:transition>:\n%s", part)
	}
	if !strings.Contains(part, "<p:fade/>") {
		t.Errorf("the effect element is not <p:fade/>:\n%s", part)
	}
	for _, unprefixed := range []string{"<transition", "<fade", "</transition", "</fade"} {
		if strings.Contains(part, unprefixed) {
			t.Errorf("an element name lost its p: prefix (%q):\n%s", unprefixed, part)
		}
	}
}

// TestEveryTransitionTypeWritesItsSchemaElement covers the whole choice group:
// each type reaches its own element, and comes back as itself.
func TestEveryTransitionTypeWritesItsSchemaElement(t *testing.T) {
	if len(transitionElementNames) != 21 {
		t.Fatalf("the schema's choice group holds 21 effects; this table lists %d", len(transitionElementNames))
	}

	for typ, name := range transitionElementNames {
		p := New()
		p.GetActiveSlide().SetTransition(&Transition{Type: typ, Speed: TransitionSpeedMedium})

		parts := zipParts(t, writeToBytes(t, p))
		xml, ok := parts["ppt/slides/slide1.xml"]
		if !ok {
			t.Fatal("the written package has no ppt/slides/slide1.xml")
		}
		if !strings.Contains(string(xml), "<p:"+name+"/>") {
			t.Errorf("type %d should be written as <p:%s/>:\n%s", typ, name, xml)
		}

		got := roundTrip(t, p).GetAllSlides()[0].GetTransition()
		if got == nil {
			t.Errorf("<p:%s/> was written and read back as no transition at all", name)
			continue
		}
		if got.Type != typ {
			t.Errorf("<p:%s/> read back as type %d, want %d", name, got.Type, typ)
		}
		if got.Speed != TransitionSpeedMedium {
			t.Errorf("<p:%s/> read back with speed %q, want %q", name, got.Speed, TransitionSpeedMedium)
		}
	}
}

// TestTransitionNoneWritesNoElement: the zero value must not add markup, or every
// slide in every deck would grow a <p:transition> it does not want.
func TestTransitionNoneWritesNoElement(t *testing.T) {
	for _, tr := range []*Transition{nil, {Type: TransitionNone}} {
		slide := slideWithTransition(t, tr)
		if strings.Contains(slide, "transition") {
			t.Errorf("no transition was asked for, yet the slide carries one:\n%s", slide)
		}
	}
}

// TestTransitionOmitsWhatTheSchemaAlreadyDefaults covers the attributes that have
// a default in the schema. Stating them costs bytes and, worse, states that the
// document author chose the default on purpose when they said nothing at all.
func TestTransitionOmitsWhatTheSchemaAlreadyDefaults(t *testing.T) {
	part := transitionPart(t, slideWithTransition(t, &Transition{Type: TransitionFade}))
	for _, unwanted := range []string{"spd=", "advClick=", "advTm=", "thruBlk=", "spokes=", "dir=", "orient=", "p14:dur="} {
		if strings.Contains(part, unwanted) {
			t.Errorf("a transition with nothing set still states %q:\n%s", unwanted, part)
		}
	}
	if strings.Contains(part, "mc:AlternateContent") {
		t.Errorf("a transition with no duration must not be wrapped in mc:AlternateContent:\n%s", part)
	}

	// The ones that are set do appear, so the check above is not passing because
	// the writer ignores these fields altogether.
	withThruBlk := transitionPart(t, slideWithTransition(t, &Transition{Type: TransitionFade, ThroughBlack: true}))
	if !strings.Contains(withThruBlk, `thruBlk="1"`) {
		t.Errorf("ThroughBlack never reaches thruBlk:\n%s", withThruBlk)
	}
	withSpokes := transitionPart(t, slideWithTransition(t, &Transition{Type: TransitionWheel, Spokes: 8}))
	if !strings.Contains(withSpokes, `spokes="8"`) {
		t.Errorf("Spokes never reaches the wheel element:\n%s", withSpokes)
	}
	withSplit := transitionPart(t, slideWithTransition(t, &Transition{
		Type:        TransitionSplit,
		Direction:   TransitionDirectionIn,
		Orientation: TransitionOrientationVertical,
	}))
	if !strings.Contains(withSplit, `orient="vert"`) || !strings.Contains(withSplit, `dir="in"`) {
		t.Errorf("split's orient and dir did not both reach the element:\n%s", withSplit)
	}
}

// TestTransitionDurationIsAnAlternateContentPair covers the one field the base
// schema has nowhere to put. p14:dur is the PowerPoint 2010 attribute, so the
// element is wrapped: mc:Choice requires p14 and states the duration, and
// mc:Fallback repeats the transition without it for consumers that cannot.
func TestTransitionDurationIsAnAlternateContentPair(t *testing.T) {
	advClick := false
	part := transitionPart(t, slideWithTransition(t, &Transition{
		Type:             TransitionPush,
		Speed:            TransitionSpeedFast,
		Direction:        TransitionDirectionLeft,
		Duration:         2000,
		AdvanceOnClick:   &advClick,
		AdvanceAfterTime: 3000,
	}))

	for _, want := range []string{
		`<mc:AlternateContent`,
		`xmlns:mc="` + nsMarkupCompat + `"`,
		`xmlns:p14="` + nsPowerPoint2010 + `"`,
		`<mc:Choice Requires="p14">`,
		`<mc:Fallback>`,
		`</mc:Fallback>`,
		`</mc:AlternateContent>`,
		`<p:push dir="l"/>`,
		`spd="fast"`,
		`advClick="0"`,
		`advTm="3000"`,
	} {
		if !strings.Contains(part, want) {
			t.Errorf("the transition markup does not contain %q:\n%s", want, part)
		}
	}

	choiceStart := strings.Index(part, "<mc:Choice")
	choiceEnd := strings.Index(part, "</mc:Choice>")
	fallbackStart := strings.Index(part, "<mc:Fallback>")
	fallbackEnd := strings.Index(part, "</mc:Fallback>")
	if choiceStart < 0 || choiceEnd < 0 || fallbackStart < 0 || fallbackEnd < 0 {
		t.Fatalf("the markup is not a Choice/Fallback pair:\n%s", part)
	}
	if choiceStart > fallbackStart {
		t.Errorf("mc:Choice must come before mc:Fallback:\n%s", part)
	}

	choice := part[choiceStart:choiceEnd]
	fallback := part[fallbackStart:fallbackEnd]
	if !strings.Contains(choice, `p14:dur="2000"`) {
		t.Errorf("mc:Choice must carry the duration:\n%s", choice)
	}
	if strings.Contains(fallback, "p14:dur") {
		t.Errorf("mc:Fallback must not carry p14:dur; it is the copy for readers without p14:\n%s", fallback)
	}
	if !strings.Contains(fallback, "<p:push") {
		t.Errorf("mc:Fallback must repeat the transition itself:\n%s", fallback)
	}
}

// TestTransitionIsReadFromAMarkupCompatibleDocument feeds the reader markup this
// writer would not produce on its own: p14 declared on the root, the Alternate
// Content without namespace declarations of its own, and p14:dur only in the
// choice. The duration has to come from the choice, not from the fallback that
// follows it.
func TestTransitionIsReadFromAMarkupCompatibleDocument(t *testing.T) {
	markup := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="` + nsDrawingML + `" xmlns:r="` + nsOfficeDocRels + `" xmlns:p="` + nsPresentationML +
		`" xmlns:p14="` + nsPowerPoint2010 + `" xmlns:mc="` + nsMarkupCompat + `">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
  <mc:AlternateContent>
    <mc:Choice Requires="p14">
      <p:transition spd="slow" p14:dur="1750" advClick="0" advTm="2000">
        <p:pull dir="rd"/>
      </p:transition>
    </mc:Choice>
    <mc:Fallback>
      <p:transition spd="slow" advClick="0" advTm="2000">
        <p:pull dir="rd"/>
      </p:transition>
    </mc:Fallback>
  </mc:AlternateContent>
</p:sld>`

	tr := readFirstSlideTransition(t, markup)
	if tr == nil {
		t.Fatal("the transition was not read at all")
	}
	if tr.Type != TransitionPull {
		t.Errorf("Type = %d, want TransitionPull (%d)", tr.Type, TransitionPull)
	}
	if tr.Speed != TransitionSpeedSlow {
		t.Errorf("Speed = %q, want %q", tr.Speed, TransitionSpeedSlow)
	}
	if tr.Duration != 1750 {
		t.Errorf("Duration = %d, want 1750 - it comes from mc:Choice, not from the mc:Fallback that repeats the transition without it", tr.Duration)
	}
	if tr.Direction != TransitionDirectionRightDown {
		t.Errorf("Direction = %q, want %q", tr.Direction, TransitionDirectionRightDown)
	}
	if tr.AdvanceOnClick == nil || *tr.AdvanceOnClick {
		t.Errorf("AdvanceOnClick = %v, want a stated false", tr.AdvanceOnClick)
	}
	if tr.AdvanceAfterTime != 2000 {
		t.Errorf("AdvanceAfterTime = %d, want 2000", tr.AdvanceAfterTime)
	}
}

// TestTransitionIsReadFromAFallbackAlone: a document that carries only the
// fallback has no choice to prefer, so the fallback is the transition.
func TestTransitionIsReadFromAFallbackAlone(t *testing.T) {
	markup := slideMarkup(`  <mc:AlternateContent>
    <mc:Fallback>
      <p:transition spd="fast">
        <p:wipe dir="u"/>
      </p:transition>
    </mc:Fallback>
  </mc:AlternateContent>`)

	tr := readFirstSlideTransition(t, markup)
	if tr == nil {
		t.Fatal("a transition that exists only as a fallback was dropped")
	}
	if tr.Type != TransitionWipe {
		t.Errorf("Type = %d, want TransitionWipe (%d)", tr.Type, TransitionWipe)
	}
	if tr.Direction != TransitionDirectionUp {
		t.Errorf("Direction = %q, want %q", tr.Direction, TransitionDirectionUp)
	}
}

// TestTransitionReadsSelfClosingAndUnmodelledElements covers the two element
// shapes that name no modelled effect. Neither may be reported as a transition,
// or a slide that has none would acquire one on the next save.
func TestTransitionReadsSelfClosingAndUnmodelledElements(t *testing.T) {
	cases := map[string]string{
		"self-closing":          `<p:transition spd="fast"/>`,
		"no children":           `<p:transition spd="fast"></p:transition>`,
		"a p14 effect":          `<p:transition spd="fast"><p14:ripple/></p:transition>`,
		"an element of its own": `<p:transition><p:somethingElse/></p:transition>`,
		"nothing in it at all":  `<p:transition/>`,
	}
	for name, element := range cases {
		tr := readFirstSlideTransition(t, slideMarkup("  "+element))
		if tr != nil {
			t.Errorf("%s: read back as %#v, want no transition", name, tr)
		}
	}
}

// TestUncoverAndPullAreOneValue: the schema has no <p:uncover>. What the model
// called TransitionUncover is <p:pull>, the eight-direction counterpart of
// <p:cover>, so the two names are one value and one element.
func TestUncoverAndPullAreOneValue(t *testing.T) {
	if TransitionUncover != TransitionPull {
		t.Errorf("TransitionUncover (%d) and TransitionPull (%d) must be the same value",
			TransitionUncover, TransitionPull)
	}
	part := transitionPart(t, slideWithTransition(t, &Transition{Type: TransitionUncover}))
	if !strings.Contains(part, "<p:pull/>") {
		t.Errorf("TransitionUncover should be written as <p:pull/>:\n%s", part)
	}
	if strings.Contains(strings.ToLower(part), "uncover") {
		t.Errorf("there is no cover-reversing element in the schema, yet one was written:\n%s", part)
	}
}

// TestEarlierTransitionValuesAreUnchanged: these eight predate the writer, and a
// caller may have stored one as an integer (in a database column, say), so their
// values must not shift when the enum grows.
func TestEarlierTransitionValuesAreUnchanged(t *testing.T) {
	want := map[TransitionType]int{
		TransitionNone:     0,
		TransitionFade:     1,
		TransitionPush:     2,
		TransitionWipe:     3,
		TransitionSplit:    4,
		TransitionCover:    5,
		TransitionUncover:  6,
		TransitionDissolve: 7,
	}
	for typ, n := range want {
		if int(typ) != n {
			t.Errorf("a transition type that existed before the writer now has value %d, want %d", int(typ), n)
		}
	}
}

// TestSlideWithATransitionStillRenders: a transition is markup the renderer has
// no use for, but a deck that carries one — including the mc:AlternateContent
// wrapper a duration brings — has to keep rendering.
func TestSlideWithATransitionStillRenders(t *testing.T) {
	p := New()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.SetOffsetX(100000).SetOffsetY(100000)
	shape.SetWidth(4000000).SetHeight(1000000)
	shape.CreateTextRun("HELLO")
	p.GetActiveSlide().SetTransition(&Transition{
		Type:      TransitionPull,
		Speed:     TransitionSpeedSlow,
		Duration:  1500,
		Direction: TransitionDirectionRightDown,
	})

	opts := DefaultRenderOptions()
	opts.Width = 320
	imgs, err := p.SlidesToImages(opts)
	if err != nil {
		t.Fatalf("render a slide that carries a transition: %v", err)
	}
	if len(imgs) != 1 {
		t.Fatalf("rendered %d images, want 1", len(imgs))
	}
}

// slideMarkup wraps slide-level content in the rest of a slide, so a test can
// state only the part it cares about.
func slideMarkup(content string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="` + nsDrawingML + `" xmlns:r="` + nsOfficeDocRels + `" xmlns:p="` + nsPresentationML +
		`" xmlns:p14="` + nsPowerPoint2010 + `" xmlns:mc="` + nsMarkupCompat + `">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
` + content + `
</p:sld>`
}

// readFirstSlideTransition puts the given slide XML into an otherwise ordinary
// package and reports the transition on the first slide.
func readFirstSlideTransition(t *testing.T, markup string) *Transition {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	if _, ok := parts["ppt/slides/slide1.xml"]; !ok {
		t.Fatal("the written package has no ppt/slides/slide1.xml")
	}
	parts["ppt/slides/slide1.xml"] = []byte(markup)
	pkg := buildZip(t, parts)

	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package back: %v", err)
	}
	return pres.GetAllSlides()[0].GetTransition()
}
