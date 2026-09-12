package gopresentation

// Bullet represents a paragraph bullet style.
type Bullet struct {
	Type      BulletType
	Style     string // character for BulletChar, e.g. "•", "–"
	Font      string // font name for BulletChar
	StartAt   int    // starting number for BulletNumeric
	NumFormat string // numeric format: "arabicPeriod", "romanUcPeriod", etc.
	Color     *Color
	Size      int // percentage of text size (25-400)
}

// BulletType represents the type of bullet.
type BulletType int

const (
	BulletTypeNone    BulletType = iota
	BulletTypeChar               // character bullet
	BulletTypeNumeric            // numbered bullet
	BulletTypeAutoNum            // auto-numbered
)

// Numeric format constants.
const (
	NumFormatArabicPeriod  = "arabicPeriod"
	NumFormatArabicParen   = "arabicParenR"
	NumFormatRomanUcPeriod = "romanUcPeriod"
	NumFormatRomanLcPeriod = "romanLcPeriod"
	NumFormatAlphaUcPeriod = "alphaUcPeriod"
	NumFormatAlphaLcPeriod = "alphaLcPeriod"
	NumFormatAlphaLcParen  = "alphaLcParenR"
)

// NewBullet creates a new bullet with no bullet type.
func NewBullet() *Bullet {
	return &Bullet{
		Type:    BulletTypeNone,
		StartAt: 1,
		Size:    100,
	}
}

// SetCharBullet sets a character bullet.
func (b *Bullet) SetCharBullet(char string, font ...string) *Bullet {
	b.Type = BulletTypeChar
	b.Style = char
	if len(font) > 0 {
		b.Font = font[0]
	}
	return b
}

// SetNumericBullet sets a numeric bullet.
func (b *Bullet) SetNumericBullet(format string, startAt ...int) *Bullet {
	b.Type = BulletTypeNumeric
	b.NumFormat = format
	if len(startAt) > 0 {
		b.StartAt = startAt[0]
	}
	return b
}

// SetColor sets the bullet color.
func (b *Bullet) SetColor(c Color) *Bullet {
	b.Color = &c
	return b
}

// SetSize sets the bullet size as percentage of text size.
func (b *Bullet) SetSize(pct int) *Bullet {
	if pct < 25 {
		pct = 25
	}
	if pct > 400 {
		pct = 400
	}
	b.Size = pct
	return b
}
