package pdf

import (
	"bytes"
	"compress/zlib"
	"fmt"
	"io"
	"strconv"
	"strings"
)

type Page struct {
	Width   float64
	Height  float64
	Content []byte
}

type Document struct {
	Pages []Page
	Fonts []FontSpec
}

func (d *Document) AddPage(width, height float64, content []byte) {
	d.Pages = append(d.Pages, Page{
		Width:   width,
		Height:  height,
		Content: append([]byte(nil), content...),
	})
}

func (d *Document) Bytes() ([]byte, error) {
	writer := newObjectWriter()

	pagesID := writer.reserve()
	catalogID := writer.reserve()
	fontResources := d.fontResources()
	fontObjectIDs := make(map[string]int, len(fontResources))
	for _, font := range fontResources {
		fontObjectIDs[font.Resource] = writer.reserve()
	}

	contentIDs := make([]int, len(d.Pages))
	pageIDs := make([]int, len(d.Pages))
	for i := range d.Pages {
		contentIDs[i] = writer.reserve()
		pageIDs[i] = writer.reserve()
	}

	for _, font := range fontResources {
		if err := writeFontObject(writer, fontObjectIDs[font.Resource], font); err != nil {
			return nil, err
		}
	}

	resourceObject := buildFontResourceObject(fontResources, fontObjectIDs)

	for i, page := range d.Pages {
		stream := append([]byte(nil), page.Content...)
		contentObject := fmt.Sprintf("<< /Length %d >>\nstream\n%s\nendstream", len(stream), stream)
		writer.set(contentIDs[i], []byte(contentObject))

		pageObject := fmt.Sprintf("<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %s %s] /Resources %s /Contents %d 0 R >>",
			pagesID,
			FormatNumber(page.Width),
			FormatNumber(page.Height),
			resourceObject,
			contentIDs[i],
		)
		writer.set(pageIDs[i], []byte(pageObject))
	}

	kids := make([]string, 0, len(pageIDs))
	for _, id := range pageIDs {
		kids = append(kids, fmt.Sprintf("%d 0 R", id))
	}
	writer.set(pagesID, []byte(fmt.Sprintf("<< /Type /Pages /Count %d /Kids [%s] >>", len(pageIDs), strings.Join(kids, " "))))
	writer.set(catalogID, []byte(fmt.Sprintf("<< /Type /Catalog /Pages %d 0 R >>", pagesID)))

	return writer.bytes(catalogID)
}

func (d *Document) fontResources() []FontSpec {
	if len(d.Fonts) == 0 {
		return []FontSpec{{
			Resource: "F1",
			Builtin:  "Helvetica",
		}}
	}

	resources := make([]FontSpec, 0, len(d.Fonts))
	seen := map[string]bool{}
	for _, font := range d.Fonts {
		if font.Resource == "" || seen[font.Resource] {
			continue
		}
		if font.Builtin == "" && font.TrueType == nil {
			continue
		}
		seen[font.Resource] = true
		resources = append(resources, font)
	}
	if len(resources) == 0 {
		return []FontSpec{{
			Resource: "F1",
			Builtin:  "Helvetica",
		}}
	}
	return resources
}

func buildFontResourceObject(fonts []FontSpec, objectIDs map[string]int) string {
	entries := make([]string, 0, len(fonts))
	for _, font := range fonts {
		objectID, ok := objectIDs[font.Resource]
		if !ok {
			continue
		}
		entries = append(entries, fmt.Sprintf("/%s %d 0 R", font.Resource, objectID))
	}
	return fmt.Sprintf("<< /Font << %s >> >>", strings.Join(entries, " "))
}

func writeFontObject(writer *objectWriter, fontID int, font FontSpec) error {
	if font.TrueType == nil {
		writer.set(fontID, []byte(fmt.Sprintf("<< /Type /Font /Subtype /Type1 /BaseFont /%s >>", sanitizeFontName(font.Builtin))))
		return nil
	}

	fontFileID := writer.reserve()
	descriptorID := writer.reserve()

	fontStream, err := makeFontStream(font.TrueType.Data)
	if err != nil {
		return fmt.Errorf("embed font %s: %w", font.Resource, err)
	}
	writer.set(fontFileID, fontStream)

	descriptor := fmt.Sprintf("<< /Type /FontDescriptor /FontName /%s /Flags %d /FontBBox [%d %d %d %d] /ItalicAngle %s /Ascent %d /Descent %d /CapHeight %d /StemV %d /MissingWidth %d /FontFile2 %d 0 R >>",
		font.TrueType.PostScriptName,
		font.TrueType.Flags,
		font.TrueType.BBox[0],
		font.TrueType.BBox[1],
		font.TrueType.BBox[2],
		font.TrueType.BBox[3],
		FormatNumber(font.TrueType.ItalicAngle),
		font.TrueType.Ascent,
		font.TrueType.Descent,
		font.TrueType.CapHeight,
		font.TrueType.StemV,
		font.TrueType.MissingWidth,
		fontFileID,
	)
	writer.set(descriptorID, []byte(descriptor))

	widths := make([]string, 0, len(font.TrueType.Widths))
	for _, width := range font.TrueType.Widths {
		widths = append(widths, strconv.Itoa(width))
	}
	fontObject := fmt.Sprintf("<< /Type /Font /Subtype /TrueType /BaseFont /%s /FirstChar 32 /LastChar 255 /Widths [%s] /Encoding /WinAnsiEncoding /FontDescriptor %d 0 R >>",
		font.TrueType.PostScriptName,
		strings.Join(widths, " "),
		descriptorID,
	)
	writer.set(fontID, []byte(fontObject))
	return nil
}

func makeFontStream(data []byte) ([]byte, error) {
	var compressed bytes.Buffer
	zipWriter := zlib.NewWriter(&compressed)
	if _, err := zipWriter.Write(data); err != nil {
		return nil, err
	}
	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	var object bytes.Buffer
	_, _ = io.WriteString(&object, fmt.Sprintf("<< /Length %d /Length1 %d /Filter /FlateDecode >>\nstream\n", compressed.Len(), len(data)))
	object.Write(compressed.Bytes())
	_, _ = io.WriteString(&object, "\nendstream")
	return object.Bytes(), nil
}

func EscapeText(value string) string {
	replacer := strings.NewReplacer(
		"\\", "\\\\",
		"(", "\\(",
		")", "\\)",
	)
	return replacer.Replace(value)
}

func FormatNumber(value float64) string {
	return strconv.FormatFloat(value, 'f', -1, 64)
}

type objectWriter struct {
	nextID  int
	objects map[int][]byte
}

func newObjectWriter() *objectWriter {
	return &objectWriter{
		nextID:  1,
		objects: map[int][]byte{},
	}
}

func (w *objectWriter) reserve() int {
	id := w.nextID
	w.nextID++
	return id
}

func (w *objectWriter) set(id int, body []byte) {
	w.objects[id] = body
}

func (w *objectWriter) bytes(rootID int) ([]byte, error) {
	var buffer bytes.Buffer
	buffer.WriteString("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n")

	offsets := make([]int, w.nextID)
	for id := 1; id < w.nextID; id++ {
		body, ok := w.objects[id]
		if !ok {
			return nil, fmt.Errorf("missing PDF object %d", id)
		}
		offsets[id] = buffer.Len()
		fmt.Fprintf(&buffer, "%d 0 obj\n", id)
		buffer.Write(body)
		buffer.WriteString("\nendobj\n")
	}

	xrefOffset := buffer.Len()
	fmt.Fprintf(&buffer, "xref\n0 %d\n", w.nextID)
	buffer.WriteString("0000000000 65535 f \n")
	for id := 1; id < w.nextID; id++ {
		fmt.Fprintf(&buffer, "%010d 00000 n \n", offsets[id])
	}

	fmt.Fprintf(&buffer, "trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n", w.nextID, rootID, xrefOffset)
	return buffer.Bytes(), nil
}
