package pdf

import (
	"encoding/binary"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"unicode/utf16"
)

type FontSpec struct {
	Resource string
	Builtin  string
	TrueType *TrueTypeFont
}

type TrueTypeFont struct {
	PostScriptName string
	Data           []byte
	Widths         [224]int
	MissingWidth   int
	Ascent         int
	Descent        int
	CapHeight      int
	ItalicAngle    float64
	BBox           [4]int
	Flags          int
	StemV          int
}

func LoadTrueTypeFont(path string, fallbackName string) (*TrueTypeFont, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	font, err := parseTrueType(data)
	if err != nil {
		return nil, fmt.Errorf("parse %s: %w", path, err)
	}
	if font.PostScriptName == "" {
		font.PostScriptName = sanitizeFontName(fallbackName)
	}
	if font.PostScriptName == "" {
		font.PostScriptName = sanitizeFontName(strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)))
	}
	return font, nil
}

func (f *TrueTypeFont) MeasureText(text string, fontSize float64) float64 {
	if f == nil {
		return 0
	}
	encoded := EncodeWinAnsi(text)
	total := 0
	for _, b := range encoded {
		if b < 32 {
			continue
		}
		total += f.Widths[int(b)-32]
	}
	return float64(total) * fontSize / 1000
}

func EncodeWinAnsi(text string) []byte {
	out := make([]byte, 0, len(text))
	for _, r := range text {
		if b, ok := winAnsiEncode[r]; ok {
			out = append(out, b)
			continue
		}
		if r >= 0 && r <= 255 {
			out = append(out, byte(r))
			continue
		}
		out = append(out, '?')
	}
	return out
}

func EscapeLiteralBytes(data []byte) string {
	var builder strings.Builder
	for _, b := range data {
		switch b {
		case '\\', '(', ')':
			builder.WriteByte('\\')
			builder.WriteByte(b)
		case '\r':
			builder.WriteString(`\r`)
		case '\n':
			builder.WriteString(`\n`)
		case '\t':
			builder.WriteString(`\t`)
		case '\b':
			builder.WriteString(`\b`)
		case '\f':
			builder.WriteString(`\f`)
		default:
			if b < 32 || b > 126 {
				builder.WriteString(fmt.Sprintf(`\%03o`, b))
			} else {
				builder.WriteByte(b)
			}
		}
	}
	return builder.String()
}

func sanitizeFontName(name string) string {
	name = strings.TrimSpace(name)
	if name == "" {
		return ""
	}
	var builder strings.Builder
	for _, r := range name {
		switch {
		case r >= 'A' && r <= 'Z':
			builder.WriteRune(r)
		case r >= 'a' && r <= 'z':
			builder.WriteRune(r)
		case r >= '0' && r <= '9':
			builder.WriteRune(r)
		case r == '-' || r == '_':
			builder.WriteRune(r)
		case r == ' ':
			builder.WriteRune('-')
		}
	}
	return builder.String()
}

type sfntTable struct {
	offset uint32
	length uint32
}

func parseTrueType(data []byte) (*TrueTypeFont, error) {
	if len(data) < 12 {
		return nil, fmt.Errorf("invalid sfnt header")
	}
	numTables := int(u16(data, 4))
	if len(data) < 12+numTables*16 {
		return nil, fmt.Errorf("truncated table directory")
	}

	tables := map[string]sfntTable{}
	for i := 0; i < numTables; i++ {
		entry := 12 + i*16
		tag := string(data[entry : entry+4])
		tables[tag] = sfntTable{
			offset: u32(data, entry+8),
			length: u32(data, entry+12),
		}
	}

	head, err := tableBytes(data, tables, "head")
	if err != nil {
		return nil, err
	}
	hhea, err := tableBytes(data, tables, "hhea")
	if err != nil {
		return nil, err
	}
	maxp, err := tableBytes(data, tables, "maxp")
	if err != nil {
		return nil, err
	}
	hmtx, err := tableBytes(data, tables, "hmtx")
	if err != nil {
		return nil, err
	}
	cmap, err := tableBytes(data, tables, "cmap")
	if err != nil {
		return nil, err
	}

	unitsPerEm := int(u16(head, 18))
	if unitsPerEm == 0 {
		return nil, fmt.Errorf("invalid unitsPerEm")
	}
	scale := func(value int) int {
		return int(math.Round(float64(value) * 1000 / float64(unitsPerEm)))
	}

	numGlyphs := int(u16(maxp, 4))
	numberOfHMetrics := int(u16(hhea, 34))
	if numberOfHMetrics < 1 || numberOfHMetrics > numGlyphs {
		return nil, fmt.Errorf("invalid hmetric count")
	}

	widthsRaw, err := parseWidths(hmtx, numGlyphs, numberOfHMetrics)
	if err != nil {
		return nil, err
	}
	cmapMap, err := parseCmap(cmap)
	if err != nil {
		return nil, err
	}

	font := &TrueTypeFont{
		Data:         append([]byte(nil), data...),
		MissingWidth: scale(int(widthsRaw[0])),
		Ascent:       scale(int(i16(hhea, 4))),
		Descent:      scale(int(i16(hhea, 6))),
		CapHeight:    scale(int(i16(hhea, 4))),
		ItalicAngle:  0,
		BBox: [4]int{
			scale(int(i16(head, 36))),
			scale(int(i16(head, 38))),
			scale(int(i16(head, 40))),
			scale(int(i16(head, 42))),
		},
		Flags: 32,
		StemV: 80,
	}

	if font.BBox[0] == font.BBox[2] {
		font.BBox[2] = 1000
	}
	if font.BBox[1] == font.BBox[3] {
		font.BBox[3] = 1000
	}

	if os2, err := tableBytes(data, tables, "OS/2"); err == nil {
		version := int(u16(os2, 0))
		font.Flags = 32
		if fsType := u16(os2, 8); fsType != 0 {
			_ = fsType
		}
		font.Ascent = scale(int(i16(os2, 68)))
		font.Descent = scale(int(i16(os2, 70)))
		if version >= 2 && len(os2) >= 90 {
			font.CapHeight = scale(int(i16(os2, 88)))
		}
		weightClass := int(u16(os2, 4))
		if weightClass >= 700 {
			font.StemV = 120
		}
	}
	if post, err := tableBytes(data, tables, "post"); err == nil && len(post) >= 16 {
		font.ItalicAngle = fixed32(post, 4)
		if font.ItalicAngle != 0 {
			font.Flags |= 64
		}
		if u32(post, 12) != 0 {
			font.Flags |= 1
		}
	}
	if name, err := tableBytes(data, tables, "name"); err == nil {
		font.PostScriptName = parsePostScriptName(name)
	}

	for b := 32; b <= 255; b++ {
		codepoint := winAnsiDecode[byte(b)]
		glyph := cmapMap[codepoint]
		width := font.MissingWidth
		if int(glyph) < len(widthsRaw) {
			width = scale(int(widthsRaw[glyph]))
		}
		font.Widths[b-32] = width
	}

	return font, nil
}

func parseWidths(hmtx []byte, numGlyphs, numberOfHMetrics int) ([]uint16, error) {
	widths := make([]uint16, numGlyphs)
	required := numberOfHMetrics*4 + (numGlyphs-numberOfHMetrics)*2
	if len(hmtx) < required {
		return nil, fmt.Errorf("truncated hmtx table")
	}

	lastAdvance := uint16(0)
	for i := 0; i < numberOfHMetrics; i++ {
		lastAdvance = u16(hmtx, i*4)
		widths[i] = lastAdvance
	}
	for i := numberOfHMetrics; i < numGlyphs; i++ {
		widths[i] = lastAdvance
	}
	return widths, nil
}

func parseCmap(cmap []byte) (map[rune]uint16, error) {
	if len(cmap) < 4 {
		return nil, fmt.Errorf("truncated cmap table")
	}
	numTables := int(u16(cmap, 2))
	bestOffset := uint32(0)
	bestFormat := uint16(0)
	for i := 0; i < numTables; i++ {
		entry := 4 + i*8
		if len(cmap) < entry+8 {
			return nil, fmt.Errorf("truncated cmap encoding record")
		}
		platform := u16(cmap, entry)
		encoding := u16(cmap, entry+2)
		offset := u32(cmap, entry+4)
		if int(offset)+2 > len(cmap) {
			continue
		}
		format := u16(cmap, int(offset))
		score := -1
		switch {
		case platform == 3 && encoding == 10 && (format == 12 || format == 4):
			score = 4
		case platform == 3 && encoding == 1 && (format == 4 || format == 12):
			score = 3
		case platform == 0 && (format == 4 || format == 12):
			score = 2
		}
		if score > 0 && (bestOffset == 0 || score > formatScore(bestFormat)) {
			bestOffset = offset
			bestFormat = format
		}
	}
	if bestOffset == 0 {
		return nil, fmt.Errorf("supported cmap subtable not found")
	}

	subtable := cmap[bestOffset:]
	switch bestFormat {
	case 4:
		return parseCmapFormat4(subtable)
	case 12:
		return parseCmapFormat12(subtable)
	default:
		return nil, fmt.Errorf("unsupported cmap format %d", bestFormat)
	}
}

func formatScore(format uint16) int {
	switch format {
	case 12:
		return 12
	case 4:
		return 4
	default:
		return 0
	}
}

func parseCmapFormat4(data []byte) (map[rune]uint16, error) {
	if len(data) < 16 {
		return nil, fmt.Errorf("truncated cmap format 4")
	}
	segCount := int(u16(data, 6) / 2)
	if segCount == 0 {
		return nil, fmt.Errorf("empty cmap segments")
	}
	if len(data) < 16+segCount*8 {
		return nil, fmt.Errorf("truncated cmap segment arrays")
	}

	endCodes := make([]uint16, segCount)
	startCodes := make([]uint16, segCount)
	idDeltas := make([]int16, segCount)
	idRangeOffsets := make([]uint16, segCount)

	offset := 14
	for i := 0; i < segCount; i++ {
		endCodes[i] = u16(data, offset+i*2)
	}
	offset += segCount * 2
	offset += 2
	for i := 0; i < segCount; i++ {
		startCodes[i] = u16(data, offset+i*2)
	}
	offset += segCount * 2
	for i := 0; i < segCount; i++ {
		idDeltas[i] = i16(data, offset+i*2)
	}
	offset += segCount * 2
	rangeOffsetPos := offset
	for i := 0; i < segCount; i++ {
		idRangeOffsets[i] = u16(data, offset+i*2)
	}

	result := map[rune]uint16{}
	for _, cp := range sortedWinAnsiRunes() {
		code := uint16(cp)
		for i := 0; i < segCount; i++ {
			if code < startCodes[i] || code > endCodes[i] {
				continue
			}
			if idRangeOffsets[i] == 0 {
				result[cp] = uint16((int(code) + int(idDeltas[i])) & 0xFFFF)
				break
			}

			glyphOffset := rangeOffsetPos + i*2 + int(idRangeOffsets[i]) + int(code-startCodes[i])*2
			if glyphOffset+2 > len(data) {
				break
			}
			glyph := u16(data, glyphOffset)
			if glyph != 0 {
				glyph = uint16((int(glyph) + int(idDeltas[i])) & 0xFFFF)
			}
			result[cp] = glyph
			break
		}
	}
	return result, nil
}

func parseCmapFormat12(data []byte) (map[rune]uint16, error) {
	if len(data) < 16 {
		return nil, fmt.Errorf("truncated cmap format 12")
	}
	groupCount := int(u32(data, 12))
	if len(data) < 16+groupCount*12 {
		return nil, fmt.Errorf("truncated cmap format 12 groups")
	}

	result := map[rune]uint16{}
	for _, cp := range sortedWinAnsiRunes() {
		code := uint32(cp)
		for i := 0; i < groupCount; i++ {
			groupOffset := 16 + i*12
			start := u32(data, groupOffset)
			end := u32(data, groupOffset+4)
			startGlyph := u32(data, groupOffset+8)
			if code < start || code > end {
				continue
			}
			result[cp] = uint16(startGlyph + (code - start))
			break
		}
	}
	return result, nil
}

func parsePostScriptName(data []byte) string {
	if len(data) < 6 {
		return ""
	}
	count := int(u16(data, 2))
	stringOffset := int(u16(data, 4))
	best := ""
	for i := 0; i < count; i++ {
		record := 6 + i*12
		if len(data) < record+12 {
			break
		}
		platformID := u16(data, record)
		encodingID := u16(data, record+2)
		languageID := u16(data, record+4)
		nameID := u16(data, record+6)
		length := int(u16(data, record+8))
		offset := int(u16(data, record+10))
		if nameID != 6 || stringOffset+offset+length > len(data) {
			continue
		}
		raw := data[stringOffset+offset : stringOffset+offset+length]
		switch {
		case platformID == 3 && (encodingID == 1 || encodingID == 10):
			value := decodeUTF16BE(raw)
			if languageID == 0x409 {
				return sanitizeFontName(value)
			}
			best = sanitizeFontName(value)
		case platformID == 1:
			best = sanitizeFontName(string(raw))
		}
	}
	return best
}

func decodeUTF16BE(data []byte) string {
	if len(data)%2 != 0 {
		data = data[:len(data)-1]
	}
	words := make([]uint16, 0, len(data)/2)
	for i := 0; i < len(data); i += 2 {
		words = append(words, binary.BigEndian.Uint16(data[i:i+2]))
	}
	return string(utf16.Decode(words))
}

func tableBytes(data []byte, tables map[string]sfntTable, tag string) ([]byte, error) {
	table, ok := tables[tag]
	if !ok {
		return nil, fmt.Errorf("missing %s table", tag)
	}
	start := int(table.offset)
	end := start + int(table.length)
	if start < 0 || end > len(data) || start > end {
		return nil, fmt.Errorf("invalid %s table bounds", tag)
	}
	return data[start:end], nil
}

func sortedWinAnsiRunes() []rune {
	runes := make([]rune, 0, len(winAnsiEncode))
	seen := map[rune]bool{}
	for _, r := range winAnsiDecode {
		if seen[r] {
			continue
		}
		seen[r] = true
		runes = append(runes, r)
	}
	sort.Slice(runes, func(i, j int) bool { return runes[i] < runes[j] })
	return runes
}

func u16(data []byte, offset int) uint16 {
	return binary.BigEndian.Uint16(data[offset : offset+2])
}

func i16(data []byte, offset int) int16 {
	return int16(u16(data, offset))
}

func u32(data []byte, offset int) uint32 {
	return binary.BigEndian.Uint32(data[offset : offset+4])
}

func fixed32(data []byte, offset int) float64 {
	raw := int32(u32(data, offset))
	return float64(raw) / 65536
}

var winAnsiDecode = [256]rune{
	0x0000, 0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0007,
	0x0008, 0x0009, 0x000A, 0x000B, 0x000C, 0x000D, 0x000E, 0x000F,
	0x0010, 0x0011, 0x0012, 0x0013, 0x0014, 0x0015, 0x0017, 0x0017,
	0x02D8, 0x02C7, 0x02C6, 0x02D9, 0x02DD, 0x02DB, 0x02DA, 0x02DC,
	0x0020, 0x0021, 0x0022, 0x0023, 0x0024, 0x0025, 0x0026, 0x0027,
	0x0028, 0x0029, 0x002A, 0x002B, 0x002C, 0x002D, 0x002E, 0x002F,
	0x0030, 0x0031, 0x0032, 0x0033, 0x0034, 0x0035, 0x0036, 0x0037,
	0x0038, 0x0039, 0x003A, 0x003B, 0x003C, 0x003D, 0x003E, 0x003F,
	0x0040, 0x0041, 0x0042, 0x0043, 0x0044, 0x0045, 0x0046, 0x0047,
	0x0048, 0x0049, 0x004A, 0x004B, 0x004C, 0x004D, 0x004E, 0x004F,
	0x0050, 0x0051, 0x0052, 0x0053, 0x0054, 0x0055, 0x0056, 0x0057,
	0x0058, 0x0059, 0x005A, 0x005B, 0x005C, 0x005D, 0x005E, 0x005F,
	0x0060, 0x0061, 0x0062, 0x0063, 0x0064, 0x0065, 0x0066, 0x0067,
	0x0068, 0x0069, 0x006A, 0x006B, 0x006C, 0x006D, 0x006E, 0x006F,
	0x0070, 0x0071, 0x0072, 0x0073, 0x0074, 0x0075, 0x0076, 0x0077,
	0x0078, 0x0079, 0x007A, 0x007B, 0x007C, 0x007D, 0x007E, 0x007F,
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
	0x00A0, 0x00A1, 0x00A2, 0x00A3, 0x00A4, 0x00A5, 0x00A6, 0x00A7,
	0x00A8, 0x00A9, 0x00AA, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x00AF,
	0x00B0, 0x00B1, 0x00B2, 0x00B3, 0x00B4, 0x00B5, 0x00B6, 0x00B7,
	0x00B8, 0x00B9, 0x00BA, 0x00BB, 0x00BC, 0x00BD, 0x00BE, 0x00BF,
	0x00C0, 0x00C1, 0x00C2, 0x00C3, 0x00C4, 0x00C5, 0x00C6, 0x00C7,
	0x00C8, 0x00C9, 0x00CA, 0x00CB, 0x00CC, 0x00CD, 0x00CE, 0x00CF,
	0x00D0, 0x00D1, 0x00D2, 0x00D3, 0x00D4, 0x00D5, 0x00D6, 0x00D7,
	0x00D8, 0x00D9, 0x00DA, 0x00DB, 0x00DC, 0x00DD, 0x00DE, 0x00DF,
	0x00E0, 0x00E1, 0x00E2, 0x00E3, 0x00E4, 0x00E5, 0x00E6, 0x00E7,
	0x00E8, 0x00E9, 0x00EA, 0x00EB, 0x00EC, 0x00ED, 0x00EE, 0x00EF,
	0x00F0, 0x00F1, 0x00F2, 0x00F3, 0x00F4, 0x00F5, 0x00F6, 0x00F7,
	0x00F8, 0x00F9, 0x00FA, 0x00FB, 0x00FC, 0x00FD, 0x00FE, 0x00FF,
}

var winAnsiEncode = func() map[rune]byte {
	result := map[rune]byte{}
	for i, r := range winAnsiDecode {
		result[r] = byte(i)
	}
	return result
}()
