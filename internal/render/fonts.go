package render

import (
	"io/fs"
	"os"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"sync"

	"vsdx2pdf/internal/pdf"
)

type fontVariant int

const (
	fontRegular fontVariant = iota
	fontBold
	fontItalic
	fontBoldItalic
)

type familyFonts struct {
	regular    *pdf.FontSpec
	bold       *pdf.FontSpec
	italic     *pdf.FontSpec
	boldItalic *pdf.FontSpec
}

type fontLibrary struct {
	defaultFamily string
	families      map[string]*familyFonts
	attempted     map[string]bool
	loadedByPath  map[string]*pdf.FontSpec
	loaded        []*pdf.FontSpec
	fallback      *pdf.FontSpec
	nextResource  int
}

func newFontLibrary(defaultFamily string) *fontLibrary {
	library := &fontLibrary{
		defaultFamily: strings.TrimSpace(defaultFamily),
		families:      map[string]*familyFonts{},
		attempted:     map[string]bool{},
		loadedByPath:  map[string]*pdf.FontSpec{},
		nextResource:  1,
	}
	library.fallback = &pdf.FontSpec{
		Resource: library.allocateResource(),
		Builtin:  "Helvetica",
	}
	return library
}

func (library *fontLibrary) Resolve(family string, bold, italic bool) *pdf.FontSpec {
	family = strings.TrimSpace(family)
	if family == "" {
		family = library.defaultFamily
	}
	if family != "" {
		normalized := normalizeFontFamily(family)
		if normalized != "" {
			library.ensureFamily(family)
			if fonts := library.families[normalized]; fonts != nil {
				if resolved := fonts.resolve(bold, italic); resolved != nil {
					return resolved
				}
			}
		}
	}
	return library.fallback
}

func (library *fontLibrary) Specs() []pdf.FontSpec {
	specs := make([]pdf.FontSpec, 0, len(library.loaded)+1)
	specs = append(specs, *library.fallback)
	for _, spec := range library.loaded {
		specs = append(specs, *spec)
	}
	return specs
}

func (library *fontLibrary) ensureFamily(family string) {
	normalized := normalizeFontFamily(family)
	if normalized == "" || library.attempted[normalized] {
		return
	}
	library.attempted[normalized] = true

	fonts := &familyFonts{}
	fonts.regular = library.loadVariant(family, fontRegular)
	fonts.bold = library.loadVariant(family, fontBold)
	fonts.italic = library.loadVariant(family, fontItalic)
	fonts.boldItalic = library.loadVariant(family, fontBoldItalic)
	library.families[normalized] = fonts
}

func (library *fontLibrary) loadVariant(family string, variant fontVariant) *pdf.FontSpec {
	path := findSystemFontPath(family, variant)
	if path == "" {
		return nil
	}
	if existing := library.loadedByPath[path]; existing != nil {
		return existing
	}

	font, err := pdf.LoadTrueTypeFont(path, family)
	if err != nil {
		return nil
	}

	spec := &pdf.FontSpec{
		Resource: library.allocateResource(),
		TrueType: font,
	}
	library.loadedByPath[path] = spec
	library.loaded = append(library.loaded, spec)
	return spec
}

func (library *fontLibrary) allocateResource() string {
	resource := "F" + strconv.Itoa(library.nextResource)
	library.nextResource++
	return resource
}

func (fonts *familyFonts) resolve(bold, italic bool) *pdf.FontSpec {
	switch {
	case bold && italic:
		if fonts.boldItalic != nil {
			return fonts.boldItalic
		}
		if fonts.bold != nil {
			return fonts.bold
		}
		if fonts.italic != nil {
			return fonts.italic
		}
		return fonts.regular
	case bold:
		if fonts.bold != nil {
			return fonts.bold
		}
		return fonts.regular
	case italic:
		if fonts.italic != nil {
			return fonts.italic
		}
		return fonts.regular
	default:
		return fonts.regular
	}
}

func findSystemFontPath(family string, variant fontVariant) string {
	for _, candidateFamily := range fontFamilyFallbacks(family) {
		for _, stem := range fontStemCandidates(candidateFamily, variant) {
			if path, ok := systemFontIndex()[normalizeFontFamily(stem)]; ok {
				return path
			}
		}
	}
	return ""
}

func fontFamilyFallbacks(family string) []string {
	base := normalizeFontFamily(family)
	switch base {
	case "", "calibri":
		return []string{"Calibri", "Carlito", "Arial", "Liberation Sans", "DejaVu Sans"}
	case "helvetica":
		return []string{"Helvetica", "Arial", "Liberation Sans", "DejaVu Sans"}
	case "arial":
		return []string{"Arial", "Liberation Sans", "DejaVu Sans"}
	case "arialnarrow", "helveticanarrow":
		return []string{"Arial Narrow", "Arial", "Liberation Sans", "DejaVu Sans"}
	case "bookantiqua":
		return []string{"Book Antiqua", "Palatino Linotype", "Bookman Old Style"}
	default:
		return []string{family}
	}
}

func fontStemCandidates(family string, variant fontVariant) []string {
	switch normalizeFontFamily(family) {
	case "calibri":
		switch variant {
		case fontBold:
			return []string{"calibrib"}
		case fontItalic:
			return []string{"calibrii"}
		case fontBoldItalic:
			return []string{"calibriz"}
		default:
			return []string{"calibri"}
		}
	case "carlito":
		switch variant {
		case fontBold:
			return []string{"Carlito-Bold", "CarlitoBold"}
		case fontItalic:
			return []string{"Carlito-Italic", "CarlitoItalic"}
		case fontBoldItalic:
			return []string{"Carlito-BoldItalic", "CarlitoBoldItalic"}
		default:
			return []string{"Carlito-Regular", "Carlito"}
		}
	case "arial":
		switch variant {
		case fontBold:
			return []string{"arialbd"}
		case fontItalic:
			return []string{"ariali"}
		case fontBoldItalic:
			return []string{"arialbi"}
		default:
			return []string{"arial"}
		}
	case "arialnarrow":
		switch variant {
		case fontBold:
			return []string{"arialnb"}
		case fontItalic:
			return []string{"arialni"}
		case fontBoldItalic:
			return []string{"arialnbi"}
		default:
			return []string{"arialn", "ArialNarrow"}
		}
	case "palatinolinotype":
		switch variant {
		case fontBold:
			return []string{"pala"}
		case fontItalic:
			return []string{"palai"}
		case fontBoldItalic:
			return []string{"palabi"}
		default:
			return []string{"pala"}
		}
	case "bookantiqua":
		switch variant {
		case fontBold:
			return []string{"bookosb"}
		case fontItalic:
			return []string{"bookosi"}
		case fontBoldItalic:
			return []string{"bookosbi"}
		default:
			return []string{"bookos", "bookantiqua"}
		}
	case "liberationsans":
		switch variant {
		case fontBold:
			return []string{"LiberationSans-Bold"}
		case fontItalic:
			return []string{"LiberationSans-Italic"}
		case fontBoldItalic:
			return []string{"LiberationSans-BoldItalic"}
		default:
			return []string{"LiberationSans-Regular", "LiberationSans"}
		}
	case "dejavusans":
		switch variant {
		case fontBold:
			return []string{"DejaVuSans-Bold"}
		case fontItalic:
			return []string{"DejaVuSans-Oblique", "DejaVuSans-Italic"}
		case fontBoldItalic:
			return []string{"DejaVuSans-BoldOblique", "DejaVuSans-BoldItalic"}
		default:
			return []string{"DejaVuSans"}
		}
	}

	base := strings.TrimSpace(family)
	switch variant {
	case fontBold:
		return []string{base + " Bold", base + "-Bold", base + "Bd", base + "bd"}
	case fontItalic:
		return []string{base + " Italic", base + "-Italic", base + "It", base + "i"}
	case fontBoldItalic:
		return []string{base + " Bold Italic", base + "-BoldItalic", base + "Bi", base + "bi"}
	default:
		return []string{base, base + " Regular", base + "-Regular"}
	}
}

func normalizeFontFamily(value string) string {
	var builder strings.Builder
	for _, r := range strings.ToLower(strings.TrimSpace(value)) {
		switch {
		case r >= 'a' && r <= 'z':
			builder.WriteRune(r)
		case r >= '0' && r <= '9':
			builder.WriteRune(r)
		}
	}
	return builder.String()
}

var (
	loadSystemFontsOnce sync.Once
	cachedSystemFonts   map[string]string
)

func systemFontIndex() map[string]string {
	loadSystemFontsOnce.Do(func() {
		cachedSystemFonts = map[string]string{}
		for _, root := range systemFontDirs() {
			info, err := os.Stat(root)
			if err != nil || !info.IsDir() {
				continue
			}
			_ = filepath.WalkDir(root, func(path string, entry fs.DirEntry, err error) error {
				if err != nil || entry.IsDir() {
					return nil
				}
				if strings.ToLower(filepath.Ext(entry.Name())) != ".ttf" {
					return nil
				}
				name := normalizeFontFamily(strings.TrimSuffix(entry.Name(), filepath.Ext(entry.Name())))
				if name == "" {
					return nil
				}
				if _, exists := cachedSystemFonts[name]; !exists {
					cachedSystemFonts[name] = path
				}
				return nil
			})
		}
	})
	return cachedSystemFonts
}

func systemFontDirs() []string {
	dirs := []string{}
	if runtime.GOOS == "windows" {
		if windir := os.Getenv("WINDIR"); windir != "" {
			dirs = append(dirs, filepath.Join(windir, "Fonts"))
		}
		dirs = append(dirs, `C:\Windows\Fonts`)
	}
	dirs = append(dirs,
		`/System/Library/Fonts`,
		`/Library/Fonts`,
		`/usr/share/fonts`,
		`/usr/local/share/fonts`,
		filepath.Join(os.Getenv("HOME"), ".fonts"),
		filepath.Join(os.Getenv("HOME"), "Library", "Fonts"),
	)
	return dirs
}
