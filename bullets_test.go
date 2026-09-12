package gopresentation

import (
	"strings"
	"testing"
)

// Paragraph bullets.
//
// bullet.go declares four bullet kinds and the reader handles all of them, but
// the writer's switch had a case for BulletTypeChar and one for
// BulletTypeNumeric and nothing else. BulletTypeAutoNum fell through and wrote
// no bullet element at all, while buildBulletRun — the renderer's half — treats
// Numeric and AutoNum alike and draws a number. The preview therefore showed a
// numbered item and the file held a paragraph with no bullet, which is the same
// disagreement between two halves of the library that the table writer had.
//
// Two empty values produced markup PowerPoint offers to repair rather than
// markup it accepts: <a:buAutoNum type=""> and <a:buChar char="">.

// bulletedPresentation returns a presentation with one paragraph in one shape,
// and the paragraph so the caller can set its bullet.
func bulletedPresentation() (*Presentation, *Paragraph) {
	p := New()
	shape := p.GetActiveSlide().CreateRichTextShape()
	shape.CreateTextRun("item")
	shape.SetHeight(1000000)
	return p, shape.GetParagraphs()[0]
}

// bulletPart returns the slide part of a written presentation.
func bulletPart(t *testing.T, p *Presentation) string {
	t.Helper()
	return string(zipParts(t, writeToBytes(t, p))["ppt/slides/slide1.xml"])
}

// TestAutoNumBulletIsWritten covers the kind that reached the file as nothing.
func TestAutoNumBulletIsWritten(t *testing.T) {
	b := NewBullet()
	b.Type = BulletTypeAutoNum
	b.StartAt = 1

	p, para := bulletedPresentation()
	para.SetBullet(b)

	part := bulletPart(t, p)
	if !strings.Contains(part, "<a:buAutoNum") {
		t.Errorf("a BulletTypeAutoNum paragraph wrote no bullet:\n%s", part)
	}
	if !strings.Contains(part, `type="arabicPeriod"`) {
		t.Errorf("the auto-numbered bullet has no number format:\n%s", part)
	}
	// A buNone where a number belongs is the silent version of the same bug.
	if strings.Contains(part, "<a:buNone/>") {
		t.Errorf("an auto-numbered paragraph was written as unbilleted:\n%s", part)
	}
}

// TestNumericBulletIsWritten is the control: the kind that already worked has
// to keep working, so the fix cannot be "write buAutoNum for everything".
func TestNumericBulletIsWritten(t *testing.T) {
	p, para := bulletedPresentation()
	para.SetBullet(NewBullet().SetNumericBullet(NumFormatRomanUcPeriod, 4))

	part := bulletPart(t, p)
	if !strings.Contains(part, `<a:buAutoNum type="romanUcPeriod" startAt="4"/>`) {
		t.Errorf("the numbered bullet is wrong:\n%s", part)
	}
}

// TestBulletEmptiesAreNotWrittenAsEmptyMarkup covers the two values that make
// the element invalid. An empty type is not a number format and an empty char
// is not a bullet; PowerPoint asks to repair a part that contains either.
func TestBulletEmptiesAreNotWrittenAsEmptyMarkup(t *testing.T) {
	p, para := bulletedPresentation()
	para.SetBullet(NewBullet().SetNumericBullet(""))

	part := bulletPart(t, p)
	if strings.Contains(part, `type=""`) {
		t.Errorf("an empty number format reached the file:\n%s", part)
	}
	if !strings.Contains(part, `type="`+NumFormatArabicPeriod+`"`) {
		t.Errorf("an empty number format was not replaced by the default:\n%s", part)
	}

	p, para = bulletedPresentation()
	para.SetBullet(NewBullet().SetCharBullet(""))
	part = bulletPart(t, p)
	if strings.Contains(part, `char=""`) {
		t.Errorf("an empty bullet character reached the file:\n%s", part)
	}
	if !strings.Contains(part, `<a:buChar char="•"/>`) {
		t.Errorf("an empty bullet character was not replaced by the default:\n%s", part)
	}

	// The schema's start number starts at 1, and the zero value of Bullet is 0.
	p, para = bulletedPresentation()
	para.SetBullet(NewBullet().SetNumericBullet(NumFormatArabicPeriod, 0))
	part = bulletPart(t, p)
	if !strings.Contains(part, `startAt="1"`) {
		t.Errorf("a start number of 0 was not brought into range:\n%s", part)
	}
}

// TestBulletSurvivesRoundTrip is what the reader already promised: every kind
// the writer emits has to come back as a bullet rather than as plain text.
func TestBulletSurvivesRoundTrip(t *testing.T) {
	b := NewBullet()
	b.Type = BulletTypeAutoNum
	b.StartAt = 2

	p, para := bulletedPresentation()
	para.SetBullet(b)

	back := roundTrip(t, p)
	bullet := firstBullet(t, back)
	if bullet == nil {
		t.Fatal("the round trip lost the bullet: the paragraph came back unbilleted")
	}
	// The file has one element for a numbered bullet, so AutoNum and Numeric
	// are the same thing to PowerPoint and the reader reports Numeric.
	if bullet.Type != BulletTypeNumeric {
		t.Errorf("bullet type = %d, want %d (numeric)", bullet.Type, BulletTypeNumeric)
	}
	if bullet.NumFormat != NumFormatArabicPeriod {
		t.Errorf("number format = %q, want %q", bullet.NumFormat, NumFormatArabicPeriod)
	}
	if bullet.StartAt != 2 {
		t.Errorf("start number = %d, want 2", bullet.StartAt)
	}
}

// firstBullet returns the bullet of the first paragraph with one.
func firstBullet(t *testing.T, pres *Presentation) *Bullet {
	t.Helper()
	for _, slide := range pres.GetAllSlides() {
		for _, shape := range flattenShapes(slide.GetShapes()) {
			for _, para := range shapeParagraphs(shape) {
				if para.GetBullet() != nil && para.GetBullet().Type != BulletTypeNone {
					return para.GetBullet()
				}
			}
		}
	}
	return nil
}
