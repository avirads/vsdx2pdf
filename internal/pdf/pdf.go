package pdf

import (
	"bytes"
	"fmt"
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
	fontID := writer.reserve()

	contentIDs := make([]int, len(d.Pages))
	pageIDs := make([]int, len(d.Pages))
	for i := range d.Pages {
		contentIDs[i] = writer.reserve()
		pageIDs[i] = writer.reserve()
	}

	writer.set(fontID, []byte("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"))

	for i, page := range d.Pages {
		stream := append([]byte(nil), page.Content...)
		contentObject := fmt.Sprintf("<< /Length %d >>\nstream\n%s\nendstream", len(stream), stream)
		writer.set(contentIDs[i], []byte(contentObject))

		pageObject := fmt.Sprintf("<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %s %s] /Resources << /Font << /F1 %d 0 R >> >> /Contents %d 0 R >>",
			pagesID,
			FormatNumber(page.Width),
			FormatNumber(page.Height),
			fontID,
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
