package vsdx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"sort"
	"strconv"
	"strings"
)

const (
	relTypeDocument = "http://schemas.microsoft.com/visio/2010/relationships/document"
	relTypePages    = "http://schemas.microsoft.com/visio/2010/relationships/pages"
	relTypeMasters  = "http://schemas.microsoft.com/visio/2010/relationships/masters"
)

type Document struct {
	Pages       []*Page
	DefaultFont string
	masters     map[string]*Master
	pageByID    map[string]*Page
}

type Page struct {
	ID           string
	Name         string
	NameU        string
	BackgroundID string
	IsBackground bool
	Width        float64
	Height       float64
	Shapes       []*Shape
}

type Master struct {
	ID        string
	Name      string
	NameU     string
	Shapes    []*Shape
	RootShape *Shape
	shapeByID map[string]*Shape
}

type Shape struct {
	ID               string
	Name             string
	NameU            string
	Type             string
	MasterID         string
	MasterShapeID    string
	resolvedMasterID string
	OneD             bool
	Text             string
	Cells            map[string]Cell
	Sections         map[string][]*Section
	Shapes           []*Shape
}

type Section struct {
	Name  string
	Index string
	Cells map[string]Cell
	Rows  []*Row
}

type Row struct {
	Name  string
	Type  string
	Index string
	Cells map[string]Cell
}

type Cell struct {
	Name    string
	Value   string
	Unit    string
	Formula string
}

func Open(filePath string) (*Document, error) {
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil, err
	}
	return Parse(data)
}

func Parse(data []byte) (*Document, error) {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}

	files := make(map[string]*zip.File, len(reader.File))
	for _, file := range reader.File {
		files[normalizeZipPath(file.Name)] = file
	}

	rootRels, err := readRelationships(files, "_rels/.rels")
	if err != nil {
		return nil, fmt.Errorf("read package relationships: %w", err)
	}

	documentPart, ok := relationshipTarget("", rootRels, relTypeDocument)
	if !ok {
		return nil, fmt.Errorf("Visio document relationship not found")
	}

	documentRels, err := readRelationships(files, relsPath(documentPart))
	if err != nil {
		return nil, fmt.Errorf("read document relationships: %w", err)
	}

	pagesPart, ok := relationshipTarget(documentPart, documentRels, relTypePages)
	if !ok {
		return nil, fmt.Errorf("Visio pages relationship not found")
	}

	document := &Document{
		masters:  map[string]*Master{},
		pageByID: map[string]*Page{},
	}
	document.DefaultFont = parseDefaultFont(files, documentPart)

	if mastersPart, ok := relationshipTarget(documentPart, documentRels, relTypeMasters); ok {
		masters, err := parseMasters(files, mastersPart)
		if err != nil {
			return nil, err
		}
		document.masters = masters
	}

	pages, err := parsePages(files, pagesPart)
	if err != nil {
		return nil, err
	}
	document.Pages = pages
	for _, page := range pages {
		document.pageByID[page.ID] = page
	}

	return document, nil
}

func (d *Document) Master(id string) *Master {
	return d.masters[id]
}

func (d *Document) BackgroundPage(id string) *Page {
	return d.pageByID[id]
}

func (p *Page) NameOrFallback() string {
	switch {
	case strings.TrimSpace(p.Name) != "":
		return p.Name
	case strings.TrimSpace(p.NameU) != "":
		return p.NameU
	default:
		return "Page " + p.ID
	}
}

func (m *Master) Shape(id string) *Shape {
	if id == "" {
		return m.RootShape
	}
	return m.shapeByID[id]
}

func (s *Shape) Cell(name string) (Cell, bool) {
	cell, ok := s.Cells[name]
	return cell, ok
}

func (s *Shape) ResolvedMasterID() string {
	if s.resolvedMasterID != "" {
		return s.resolvedMasterID
	}
	return s.MasterID
}

func (s *Shape) SectionsNamed(name string) []*Section {
	return s.Sections[name]
}

func (sec *Section) Cell(name string) (Cell, bool) {
	cell, ok := sec.Cells[name]
	return cell, ok
}

func (row *Row) Cell(name string) (Cell, bool) {
	cell, ok := row.Cells[name]
	return cell, ok
}

func parsePages(files map[string]*zip.File, pagesPart string) ([]*Page, error) {
	var pagesXML pagesListXML
	if err := readXMLFile(files, pagesPart, &pagesXML); err != nil {
		return nil, fmt.Errorf("read pages list: %w", err)
	}

	rels, err := readRelationships(files, relsPath(pagesPart))
	if err != nil {
		return nil, fmt.Errorf("read pages list relationships: %w", err)
	}

	pages := make([]*Page, 0, len(pagesXML.Pages))
	for _, entry := range pagesXML.Pages {
		target, ok := targetByID(pagesPart, rels, entry.Rel.ID)
		if !ok {
			return nil, fmt.Errorf("page %s relationship %q not found", entry.ID, entry.Rel.ID)
		}

		var contents pageContentsXML
		if err := readXMLFile(files, target, &contents); err != nil {
			return nil, fmt.Errorf("read page %s: %w", entry.ID, err)
		}

		page := &Page{
			ID:           entry.ID,
			Name:         entry.Name,
			NameU:        entry.NameU,
			BackgroundID: entry.BackPage,
			IsBackground: entry.Background == "1",
			Width:        readLength(entry.PageSheet.Cells, "PageWidth", readLength(contents.PageSheet.Cells, "PageWidth", 8.5)),
			Height:       readLength(entry.PageSheet.Cells, "PageHeight", readLength(contents.PageSheet.Cells, "PageHeight", 11)),
			Shapes:       parseShapes(contents.Shapes.Shapes),
		}
		pages = append(pages, page)
	}

	return pages, nil
}

func parseMasters(files map[string]*zip.File, mastersPart string) (map[string]*Master, error) {
	var mastersXML mastersListXML
	if err := readXMLFile(files, mastersPart, &mastersXML); err != nil {
		return nil, fmt.Errorf("read masters list: %w", err)
	}

	rels, err := readRelationships(files, relsPath(mastersPart))
	if err != nil {
		return nil, fmt.Errorf("read masters relationships: %w", err)
	}

	masters := make(map[string]*Master, len(mastersXML.Masters))
	for _, entry := range mastersXML.Masters {
		target, ok := targetByID(mastersPart, rels, entry.Rel.ID)
		if !ok {
			continue
		}

		var contents pageContentsXML
		if err := readXMLFile(files, target, &contents); err != nil {
			return nil, fmt.Errorf("read master %s: %w", entry.ID, err)
		}

		parsedShapes := parseMasterShapes(contents.Shapes.Shapes)
		master := &Master{
			ID:        entry.ID,
			Name:      entry.Name,
			NameU:     entry.NameU,
			Shapes:    parsedShapes,
			shapeByID: map[string]*Shape{},
		}
		switch len(parsedShapes) {
		case 0:
			master.RootShape = nil
		case 1:
			master.RootShape = parsedShapes[0]
		default:
			master.RootShape = &Shape{
				ID:       entry.ID,
				Name:     entry.Name,
				NameU:    entry.NameU,
				Cells:    map[string]Cell{},
				Sections: map[string][]*Section{},
				Shapes:   parsedShapes,
			}
		}
		indexShapes(master.shapeByID, master.RootShape)
		masters[master.ID] = master
	}

	return masters, nil
}

func indexShapes(index map[string]*Shape, shape *Shape) {
	if shape == nil {
		return
	}
	if shape.ID != "" {
		index[shape.ID] = shape
	}
	for _, child := range shape.Shapes {
		indexShapes(index, child)
	}
}

func parseShapes(raw []shapeXML) []*Shape {
	return parseShapesWithInheritedMaster(raw, "")
}

func parseMasterShapes(raw []shapeXML) []*Shape {
	return parseShapesWithInheritedMaster(raw, "")
}

func parseShapesWithInheritedMaster(raw []shapeXML, inheritedMasterID string) []*Shape {
	shapes := make([]*Shape, 0, len(raw))
	for _, item := range raw {
		effectiveMasterID := inheritedMasterID
		if item.Master != "" {
			effectiveMasterID = item.Master
		}
		shape := &Shape{
			ID:               item.ID,
			Name:             item.Name,
			NameU:            item.NameU,
			Type:             item.Type,
			MasterID:         item.Master,
			MasterShapeID:    item.MasterShape,
			resolvedMasterID: effectiveMasterID,
			OneD:             item.OneD == "1" || strings.EqualFold(item.OneD, "true"),
			Text:             flattenText(item.Text.Raw),
			Cells:            makeCellMap(item.Cells),
			Sections:         makeSectionMap(item.Sections),
			Shapes:           parseShapesWithInheritedMaster(item.Shapes.Shapes, effectiveMasterID),
		}
		shapes = append(shapes, shape)
	}
	return shapes
}

func makeCellMap(raw []cellXML) map[string]Cell {
	cells := make(map[string]Cell, len(raw))
	for _, cell := range raw {
		cells[cell.Name] = Cell{
			Name:    cell.Name,
			Value:   cell.Value,
			Unit:    cell.Unit,
			Formula: cell.Formula,
		}
	}
	return cells
}

func makeSectionMap(raw []sectionXML) map[string][]*Section {
	sections := map[string][]*Section{}
	for _, item := range raw {
		section := &Section{
			Name:  item.Name,
			Index: item.Index,
			Cells: makeCellMap(item.Cells),
			Rows:  makeRows(item.Rows),
		}
		sections[section.Name] = append(sections[section.Name], section)
	}
	for name := range sections {
		sort.SliceStable(sections[name], func(i, j int) bool {
			return sections[name][i].Index < sections[name][j].Index
		})
	}
	return sections
}

func makeRows(raw []rowXML) []*Row {
	rows := make([]*Row, 0, len(raw))
	for _, item := range raw {
		rows = append(rows, &Row{
			Name:  item.Name,
			Type:  item.Type,
			Index: item.Index,
			Cells: makeCellMap(item.Cells),
		})
	}
	return rows
}

func readRelationships(files map[string]*zip.File, partPath string) ([]relationshipXML, error) {
	var rels relationshipsXML
	if err := readXMLFile(files, partPath, &rels); err != nil {
		return nil, err
	}
	return rels.Relationships, nil
}

func relationshipTarget(ownerPart string, rels []relationshipXML, relType string) (string, bool) {
	for _, rel := range rels {
		if rel.Type == relType {
			return resolveTarget(ownerPart, rel.Target), true
		}
	}
	return "", false
}

func targetByID(ownerPart string, rels []relationshipXML, id string) (string, bool) {
	for _, rel := range rels {
		if rel.ID == id {
			return resolveTarget(ownerPart, rel.Target), true
		}
	}
	return "", false
}

func resolveTarget(ownerPart string, target string) string {
	target = strings.ReplaceAll(target, "\\", "/")
	if strings.HasPrefix(target, "/") {
		return normalizeZipPath(strings.TrimPrefix(target, "/"))
	}
	baseDir := path.Dir(ownerPart)
	if baseDir == "." {
		baseDir = ""
	}
	return normalizeZipPath(path.Clean(path.Join(baseDir, target)))
}

func relsPath(ownerPart string) string {
	dir, file := path.Split(ownerPart)
	return normalizeZipPath(path.Join(dir, "_rels", file+".rels"))
}

func normalizeZipPath(p string) string {
	return strings.TrimPrefix(strings.ReplaceAll(p, "\\", "/"), "./")
}

func readXMLFile(files map[string]*zip.File, filePath string, out any) error {
	file, ok := files[normalizeZipPath(filePath)]
	if !ok {
		return fmt.Errorf("zip entry %q not found", filePath)
	}
	reader, err := file.Open()
	if err != nil {
		return err
	}
	defer reader.Close()

	data, err := io.ReadAll(reader)
	if err != nil {
		return err
	}

	if err := xml.Unmarshal(data, out); err != nil {
		return fmt.Errorf("unmarshal %s: %w", filePath, err)
	}
	return nil
}

func flattenText(raw string) string {
	if strings.TrimSpace(raw) == "" {
		return ""
	}

	decoder := xml.NewDecoder(strings.NewReader("<root>" + raw + "</root>"))
	var builder strings.Builder
	appendNewline := func() {
		if builder.Len() == 0 || strings.HasSuffix(builder.String(), "\n") {
			return
		}
		builder.WriteByte('\n')
	}

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return strings.TrimSpace(raw)
		}

		switch token := token.(type) {
		case xml.CharData:
			builder.WriteString(string(token))
		case xml.StartElement:
			switch token.Name.Local {
			case "pp", "tp":
				appendNewline()
			}
		}
	}

	text := strings.ReplaceAll(builder.String(), "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	lines := strings.Split(text, "\n")
	for i, line := range lines {
		lines[i] = strings.TrimSpace(line)
	}
	return strings.TrimSpace(strings.Join(lines, "\n"))
}

func readLength(cells []cellXML, name string, fallback float64) float64 {
	for _, cell := range cells {
		if cell.Name != name {
			continue
		}
		if value, ok := ParseLength(cell.Value); ok {
			return value
		}
	}
	return fallback
}

func ParseBool(value string) (bool, bool) {
	trimmed := strings.TrimSpace(strings.Trim(value, "\""))
	switch strings.ToLower(trimmed) {
	case "1", "true", "yes":
		return true, true
	case "0", "false", "no":
		return false, true
	}
	number, ok := ParseNumber(trimmed)
	if !ok {
		return false, false
	}
	return number != 0, true
}

func ParseNumber(value string) (float64, bool) {
	number, _ := splitNumericValue(value)
	if number == "" {
		return 0, false
	}
	parsed, err := strconv.ParseFloat(number, 64)
	if err != nil {
		return 0, false
	}
	return parsed, true
}

func ParseLength(value string) (float64, bool) {
	number, unit := splitNumericValue(value)
	if number == "" {
		return 0, false
	}
	parsed, err := strconv.ParseFloat(number, 64)
	if err != nil {
		return 0, false
	}

	switch normalizeUnit(unit) {
	case "", "in", "inch", "inches":
		return parsed, true
	case "ft", "foot", "feet":
		return parsed * 12, true
	case "yd", "yard", "yards":
		return parsed * 36, true
	case "mi", "mile", "miles":
		return parsed * 63360, true
	case "mm":
		return parsed / 25.4, true
	case "cm":
		return parsed / 2.54, true
	case "m":
		return parsed * 39.3700787402, true
	case "km":
		return parsed * 39370.0787402, true
	case "pt":
		return parsed / 72, true
	case "pc":
		return parsed / 6, true
	default:
		return parsed, true
	}
}

func ParseAngle(value string) (float64, bool) {
	number, unit := splitNumericValue(value)
	if number == "" {
		return 0, false
	}
	parsed, err := strconv.ParseFloat(number, 64)
	if err != nil {
		return 0, false
	}

	switch normalizeUnit(unit) {
	case "", "rad", "radian", "radians":
		return parsed, true
	case "deg", "degree", "degrees":
		return parsed * (3.141592653589793 / 180), true
	case "grad":
		return parsed * (3.141592653589793 / 200), true
	default:
		return parsed, true
	}
}

func splitNumericValue(value string) (string, string) {
	s := strings.TrimSpace(strings.Trim(value, "\""))
	if s == "" {
		return "", ""
	}

	start := 0
	if s[0] == '+' || s[0] == '-' {
		start = 1
	}

	end := start
	seenDigit := false
	seenDot := false
	seenExp := false

	for end < len(s) {
		ch := s[end]
		switch {
		case ch >= '0' && ch <= '9':
			seenDigit = true
			end++
		case ch == '.' && !seenDot && !seenExp:
			seenDot = true
			end++
		case (ch == 'e' || ch == 'E') && seenDigit && !seenExp:
			seenExp = true
			seenDigit = false
			end++
			if end < len(s) && (s[end] == '+' || s[end] == '-') {
				end++
			}
		default:
			goto done
		}
	}

done:
	if !seenDigit {
		return "", ""
	}
	return s[:end], strings.TrimSpace(s[end:])
}

func normalizeUnit(unit string) string {
	return strings.Trim(strings.ToLower(strings.TrimSpace(unit)), ".")
}

func parseDefaultFont(files map[string]*zip.File, documentPart string) string {
	var doc visioDocumentXML
	if err := readXMLFile(files, documentPart, &doc); err != nil {
		return ""
	}
	for _, face := range doc.FaceNames {
		if name := strings.TrimSpace(face.NameU); name != "" {
			return name
		}
		if name := strings.TrimSpace(face.Name); name != "" {
			return name
		}
	}
	return ""
}

type relationshipsXML struct {
	Relationships []relationshipXML `xml:"Relationship"`
}

type relationshipXML struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

type pagesListXML struct {
	Pages []pageEntryXML `xml:"Page"`
}

type pageEntryXML struct {
	ID         string          `xml:"ID,attr"`
	Name       string          `xml:"Name,attr"`
	NameU      string          `xml:"NameU,attr"`
	Background string          `xml:"Background,attr"`
	BackPage   string          `xml:"BackPage,attr"`
	PageSheet  sheetXML        `xml:"PageSheet"`
	Rel        relationshipRef `xml:"Rel"`
}

type mastersListXML struct {
	Masters []masterEntryXML `xml:"Master"`
}

type masterEntryXML struct {
	ID    string          `xml:"ID,attr"`
	Name  string          `xml:"Name,attr"`
	NameU string          `xml:"NameU,attr"`
	Rel   relationshipRef `xml:"Rel"`
}

type relationshipRef struct {
	ID string `xml:"id,attr"`
}

type pageContentsXML struct {
	PageSheet sheetXML  `xml:"PageSheet"`
	Shapes    shapesXML `xml:"Shapes"`
}

type sheetXML struct {
	Cells []cellXML `xml:"Cell"`
}

type shapesXML struct {
	Shapes []shapeXML `xml:"Shape"`
}

type shapeXML struct {
	ID          string       `xml:"ID,attr"`
	Name        string       `xml:"Name,attr"`
	NameU       string       `xml:"NameU,attr"`
	Type        string       `xml:"Type,attr"`
	Master      string       `xml:"Master,attr"`
	MasterShape string       `xml:"MasterShape,attr"`
	OneD        string       `xml:"OneD,attr"`
	Text        innerXML     `xml:"Text"`
	Cells       []cellXML    `xml:"Cell"`
	Sections    []sectionXML `xml:"Section"`
	Shapes      shapesXML    `xml:"Shapes"`
}

type innerXML struct {
	Raw string `xml:",innerxml"`
}

type sectionXML struct {
	Name  string    `xml:"N,attr"`
	Index string    `xml:"IX,attr"`
	Cells []cellXML `xml:"Cell"`
	Rows  []rowXML  `xml:"Row"`
}

type rowXML struct {
	Name  string    `xml:"N,attr"`
	Type  string    `xml:"T,attr"`
	Index string    `xml:"IX,attr"`
	Cells []cellXML `xml:"Cell"`
}

type cellXML struct {
	Name    string `xml:"N,attr"`
	Value   string `xml:"V,attr"`
	Unit    string `xml:"U,attr"`
	Formula string `xml:"F,attr"`
}

type visioDocumentXML struct {
	FaceNames []faceNameXML `xml:"FaceNames>FaceName"`
}

type faceNameXML struct {
	Name  string `xml:"Name,attr"`
	NameU string `xml:"NameU,attr"`
}
