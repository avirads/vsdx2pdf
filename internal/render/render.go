package render

import (
	"bytes"
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"unicode"

	"vsdx2pdf/internal/pdf"
	"vsdx2pdf/internal/vsdx"
)

type point struct {
	X float64
	Y float64
}

type matrix struct {
	A float64
	B float64
	C float64
	D float64
	E float64
	F float64
}

type color struct {
	R float64
	G float64
	B float64
}

type pathOp struct {
	Kind   string
	Points []point
}

type vectorPath struct {
	Ops []pathOp
}

type geometryRender struct {
	Path   vectorPath
	NoFill bool
	NoLine bool
}

type context struct {
	document *vsdx.Document
	fonts    *fontLibrary
}

func Convert(document *vsdx.Document) ([]byte, error) {
	ctx := &context{
		document: document,
		fonts:    newFontLibrary(document.DefaultFont),
	}
	result := &pdf.Document{}
	maxWidth := 0.0
	maxHeight := 0.0

	for _, page := range document.Pages {
		if page.IsBackground {
			continue
		}
		if page.Width > maxWidth {
			maxWidth = page.Width
		}
		if page.Height > maxHeight {
			maxHeight = page.Height
		}
	}
	if maxWidth <= 0 {
		maxWidth = 8.5
	}
	if maxHeight <= 0 {
		maxHeight = 11
	}

	for _, page := range document.Pages {
		if page.IsBackground {
			continue
		}
		content, err := ctx.renderPage(page, maxHeight-page.Height)
		if err != nil {
			return nil, fmt.Errorf("render page %s: %w", page.NameOrFallback(), err)
		}
		width := maxWidth * 72
		height := maxHeight * 72
		result.AddPage(width, height, content)
	}

	if len(result.Pages) == 0 {
		result.AddPage(612, 792, []byte{})
	}
	result.Fonts = ctx.fonts.Specs()

	return result.Bytes()
}

func (ctx *context) renderPage(page *vsdx.Page, offsetY float64) ([]byte, error) {
	var buffer bytes.Buffer
	buffer.WriteString("1 J\n1 j\n")
	if offsetY != 0 {
		fmt.Fprintf(&buffer, "q\n1 0 0 1 0 %s cm\n", pdf.FormatNumber(offsetY*72))
	}

	visited := map[string]bool{}
	ctx.renderBackground(&buffer, page, visited)
	for _, shape := range page.Shapes {
		ctx.renderShape(&buffer, shape, identityMatrix())
	}
	if offsetY != 0 {
		buffer.WriteString("Q\n")
	}

	return buffer.Bytes(), nil
}

func (ctx *context) renderBackground(buffer *bytes.Buffer, page *vsdx.Page, visited map[string]bool) {
	if page.BackgroundID == "" || visited[page.BackgroundID] {
		return
	}
	visited[page.BackgroundID] = true

	background := ctx.document.BackgroundPage(page.BackgroundID)
	if background == nil {
		return
	}
	ctx.renderBackground(buffer, background, visited)
	for _, shape := range background.Shapes {
		ctx.renderShape(buffer, shape, identityMatrix())
	}
}

func (ctx *context) renderShape(buffer *bytes.Buffer, shape *vsdx.Shape, parent matrix) {
	if shape == nil || ctx.shapeHidden(shape) || strings.EqualFold(shape.Type, "Guide") {
		return
	}

	xform := multiply(parent, ctx.shapeTransform(shape))
	width := ctx.lengthCell(shape, "Width", 0)
	height := ctx.lengthCell(shape, "Height", 0)

	if geoms := ctx.geometrySections(shape); len(geoms) > 0 {
		for _, geom := range buildGeometry(width, height, geoms) {
			ctx.writeGeometry(buffer, geom, shape, xform)
		}
	} else if path, ok := ctx.fallbackPath(shape, width, height); ok {
		ctx.writeGeometry(buffer, geometryRender{Path: path}, shape, xform)
	} else if shape.OneD {
		ctx.writeOneDLine(buffer, shape, parent)
	}

	ctx.writeText(buffer, shape, xform, width, height)

	for _, child := range ctx.children(shape) {
		ctx.renderShape(buffer, child, xform)
	}
}

func (ctx *context) writeGeometry(buffer *bytes.Buffer, geom geometryRender, shape *vsdx.Shape, xform matrix) {
	if len(geom.Path.Ops) == 0 {
		return
	}

	strokeColor, hasStroke := ctx.strokeStyle(shape)
	fillColor, hasFill := ctx.fillStyle(shape)
	lineWidth := ctx.lineWidth(shape)
	linePattern := ctx.linePattern(shape)

	if geom.NoLine {
		hasStroke = false
	}
	if geom.NoFill {
		hasFill = false
	}
	if !hasFill && !hasStroke {
		return
	}

	buffer.WriteString("q\n")
	if hasFill {
		fmt.Fprintf(buffer, "%s %s %s rg\n", pdf.FormatNumber(fillColor.R), pdf.FormatNumber(fillColor.G), pdf.FormatNumber(fillColor.B))
	}
	if hasStroke {
		fmt.Fprintf(buffer, "%s %s %s RG\n", pdf.FormatNumber(strokeColor.R), pdf.FormatNumber(strokeColor.G), pdf.FormatNumber(strokeColor.B))
		fmt.Fprintf(buffer, "%s w\n", pdf.FormatNumber(lineWidth*72))
		ctx.writeDashPattern(buffer, linePattern, lineWidth*72)
	}

	for _, op := range geom.Path.Ops {
		switch op.Kind {
		case "m":
			p := applyInPoints(xform, op.Points[0])
			fmt.Fprintf(buffer, "%s %s m\n", pdf.FormatNumber(p.X), pdf.FormatNumber(p.Y))
		case "l":
			p := applyInPoints(xform, op.Points[0])
			fmt.Fprintf(buffer, "%s %s l\n", pdf.FormatNumber(p.X), pdf.FormatNumber(p.Y))
		case "c":
			c1 := applyInPoints(xform, op.Points[0])
			c2 := applyInPoints(xform, op.Points[1])
			end := applyInPoints(xform, op.Points[2])
			fmt.Fprintf(buffer, "%s %s %s %s %s %s c\n",
				pdf.FormatNumber(c1.X), pdf.FormatNumber(c1.Y),
				pdf.FormatNumber(c2.X), pdf.FormatNumber(c2.Y),
				pdf.FormatNumber(end.X), pdf.FormatNumber(end.Y),
			)
		case "h":
			buffer.WriteString("h\n")
		}
	}

	switch {
	case hasFill && hasStroke:
		buffer.WriteString("B\n")
	case hasFill:
		buffer.WriteString("f\n")
	case hasStroke:
		buffer.WriteString("S\n")
	}
	buffer.WriteString("Q\n")
}

func (ctx *context) writeOneDLine(buffer *bytes.Buffer, shape *vsdx.Shape, parent matrix) {
	beginX, okBeginX := ctx.lengthCellMaybe(shape, "BeginX")
	beginY, okBeginY := ctx.lengthCellMaybe(shape, "BeginY")
	endX, okEndX := ctx.lengthCellMaybe(shape, "EndX")
	endY, okEndY := ctx.lengthCellMaybe(shape, "EndY")
	if !(okBeginX && okBeginY && okEndX && okEndY) {
		return
	}

	strokeColor, hasStroke := ctx.strokeStyle(shape)
	if !hasStroke {
		return
	}

	lineWidth := ctx.lineWidth(shape)
	linePattern := ctx.linePattern(shape)
	start := applyInPoints(parent, point{X: beginX, Y: beginY})
	end := applyInPoints(parent, point{X: endX, Y: endY})

	buffer.WriteString("q\n")
	fmt.Fprintf(buffer, "%s %s %s RG\n", pdf.FormatNumber(strokeColor.R), pdf.FormatNumber(strokeColor.G), pdf.FormatNumber(strokeColor.B))
	fmt.Fprintf(buffer, "%s w\n", pdf.FormatNumber(lineWidth*72))
	ctx.writeDashPattern(buffer, linePattern, lineWidth*72)
	fmt.Fprintf(buffer, "%s %s m\n", pdf.FormatNumber(start.X), pdf.FormatNumber(start.Y))
	fmt.Fprintf(buffer, "%s %s l\n", pdf.FormatNumber(end.X), pdf.FormatNumber(end.Y))
	buffer.WriteString("S\nQ\n")
}

func (ctx *context) writeText(buffer *bytes.Buffer, shape *vsdx.Shape, shapeXform matrix, width, height float64) {
	text := strings.TrimSpace(ctx.text(shape))
	if text == "" {
		return
	}
	if hidden, ok := ctx.boolCell(shape, "HideText"); ok && hidden {
		return
	}

	textBox := ctx.textMatrix(shape, width, height)
	world := multiply(shapeXform, textBox)
	angle := math.Atan2(world.B, world.A)

	textWidth := ctx.textLength(shape, "TxtWidth", width)
	textHeight := ctx.textLength(shape, "TxtHeight", height)
	leftMargin := ctx.textLength(shape, "LeftMargin", 0.08)
	rightMargin := ctx.textLength(shape, "RightMargin", 0.08)
	topMargin := ctx.textLength(shape, "TopMargin", 0.04)
	bottomMargin := ctx.textLength(shape, "BottomMargin", 0.04)
	if textWidth <= leftMargin+rightMargin {
		leftMargin, rightMargin = 0, 0
	}
	if textHeight <= topMargin+bottomMargin {
		topMargin, bottomMargin = 0, 0
	}

	fontSize := ctx.fontSize(shape)
	fontSpec := ctx.textFont(shape)
	textColor := ctx.textColor(shape)
	availableWidth := math.Max(textWidth-leftMargin-rightMargin, 0.1)
	availableHeightPoints := math.Max((textHeight-topMargin-bottomMargin)*72, 0.1)
	lines, fontSize := fitTextToBox(text, availableWidth*72, availableHeightPoints, fontSize, func(value string, size float64) float64 {
		return measureTextWidth(value, fontSpec, size)
	})
	if len(lines) == 0 {
		return
	}

	lineHeight := textLineHeight(fontSize)
	contentHeight := float64(len(lines)) * lineHeight
	innerHeight := math.Max(availableHeightPoints, lineHeight)
	topGap := 0.0
	switch ctx.verticalAlign(shape) {
	case 0:
		topGap = 0
	case 2:
		topGap = math.Max(innerHeight-contentHeight, 0)
	default:
		topGap = math.Max(innerHeight-contentHeight, 0) / 2
	}

	buffer.WriteString("q\n")
	fmt.Fprintf(buffer, "%s %s %s rg\n", pdf.FormatNumber(textColor.R), pdf.FormatNumber(textColor.G), pdf.FormatNumber(textColor.B))

	startYPoints := textHeight*72 - topMargin*72 - topGap - fontSize
	innerWidthPoints := math.Max((textWidth-leftMargin-rightMargin)*72, 0.1)
	cosAngle := math.Cos(angle)
	sinAngle := math.Sin(angle)

	for index, line := range lines {
		lineWidth := measureTextWidth(line, fontSpec, fontSize)
		xPoints := leftMargin * 72
		switch ctx.horizontalAlign(shape) {
		case 1:
			xPoints = leftMargin*72 + (innerWidthPoints-lineWidth)/2
		case 2:
			xPoints = textWidth*72 - rightMargin*72 - lineWidth
		}
		if xPoints < 0 {
			xPoints = 0
		}

		yPoints := startYPoints - float64(index)*lineHeight
		if yPoints < 0 {
			break
		}

		position := applyInPoints(world, point{X: xPoints / 72, Y: yPoints / 72})
		buffer.WriteString("BT\n")
		fmt.Fprintf(buffer, "/%s %s Tf\n", fontSpec.Resource, pdf.FormatNumber(fontSize))
		fmt.Fprintf(buffer, "%s %s %s %s %s %s Tm\n",
			pdf.FormatNumber(cosAngle),
			pdf.FormatNumber(sinAngle),
			pdf.FormatNumber(-sinAngle),
			pdf.FormatNumber(cosAngle),
			pdf.FormatNumber(position.X),
			pdf.FormatNumber(position.Y),
		)
		fmt.Fprintf(buffer, "(%s) Tj\n", pdf.EscapeLiteralBytes(pdf.EncodeWinAnsi(line)))
		buffer.WriteString("ET\n")
	}

	buffer.WriteString("Q\n")
}

func (ctx *context) writeDashPattern(buffer *bytes.Buffer, pattern int, lineWidth float64) {
	scale := math.Max(lineWidth, 1)
	switch pattern {
	case 2:
		fmt.Fprintf(buffer, "[%s %s] 0 d\n", pdf.FormatNumber(scale*4), pdf.FormatNumber(scale*2))
	case 3:
		fmt.Fprintf(buffer, "[%s %s] 0 d\n", pdf.FormatNumber(scale), pdf.FormatNumber(scale*2))
	case 4:
		fmt.Fprintf(buffer, "[%s %s %s %s] 0 d\n", pdf.FormatNumber(scale*6), pdf.FormatNumber(scale*2), pdf.FormatNumber(scale), pdf.FormatNumber(scale*2))
	case 5:
		fmt.Fprintf(buffer, "[%s %s] 0 d\n", pdf.FormatNumber(scale*8), pdf.FormatNumber(scale*3))
	default:
		buffer.WriteString("[] 0 d\n")
	}
}

func (ctx *context) shapeHidden(shape *vsdx.Shape) bool {
	if hidden, ok := ctx.boolCell(shape, "NoShow"); ok && hidden {
		return true
	}
	return false
}

func (ctx *context) shapeTransform(shape *vsdx.Shape) matrix {
	width := ctx.lengthCell(shape, "Width", 0)
	height := ctx.lengthCell(shape, "Height", 0)
	pinX := ctx.lengthCell(shape, "PinX", 0)
	pinY := ctx.lengthCell(shape, "PinY", 0)
	locPinX := ctx.lengthCell(shape, "LocPinX", width/2)
	locPinY := ctx.lengthCell(shape, "LocPinY", height/2)
	angle := ctx.angleCell(shape, "Angle", 0)
	flipX, _ := ctx.boolCell(shape, "FlipX")
	flipY, _ := ctx.boolCell(shape, "FlipY")

	scaleX := 1.0
	scaleY := 1.0
	if flipX {
		scaleX = -1
	}
	if flipY {
		scaleY = -1
	}

	return multiply(
		translate(pinX, pinY),
		multiply(
			rotate(angle),
			multiply(scale(scaleX, scaleY), translate(-locPinX, -locPinY)),
		),
	)
}

func (ctx *context) textMatrix(shape *vsdx.Shape, width, height float64) matrix {
	txtWidth := ctx.textLength(shape, "TxtWidth", width)
	txtHeight := ctx.textLength(shape, "TxtHeight", height)
	txtPinX := ctx.textLength(shape, "TxtPinX", width/2)
	txtPinY := ctx.textLength(shape, "TxtPinY", height/2)
	txtLocPinX := ctx.textLength(shape, "TxtLocPinX", txtWidth/2)
	txtLocPinY := ctx.textLength(shape, "TxtLocPinY", txtHeight/2)
	txtAngle := ctx.angleCell(shape, "TxtAngle", 0)

	return multiply(
		translate(txtPinX, txtPinY),
		multiply(rotate(txtAngle), translate(-txtLocPinX, -txtLocPinY)),
	)
}

func (ctx *context) children(shape *vsdx.Shape) []*vsdx.Shape {
	base := ctx.baseShape(shape)
	if base == nil || len(base.Shapes) == 0 {
		if len(shape.Shapes) > 0 {
			return shape.Shapes
		}
		return nil
	}
	if len(shape.Shapes) == 0 {
		return base.Shapes
	}

	overrides := make(map[string]*vsdx.Shape, len(shape.Shapes))
	used := make(map[*vsdx.Shape]bool, len(shape.Shapes))
	merged := make([]*vsdx.Shape, 0, len(base.Shapes)+len(shape.Shapes))

	for _, child := range shape.Shapes {
		if child.MasterShapeID != "" {
			overrides[child.MasterShapeID] = child
		}
	}

	for _, baseChild := range base.Shapes {
		if override, ok := overrides[baseChild.ID]; ok {
			merged = append(merged, override)
			used[override] = true
			continue
		}
		merged = append(merged, baseChild)
	}

	for _, child := range shape.Shapes {
		if used[child] {
			continue
		}
		merged = append(merged, child)
	}

	return merged
}

func (ctx *context) text(shape *vsdx.Shape) string {
	if strings.TrimSpace(shape.Text) != "" {
		return shape.Text
	}
	if base := ctx.baseShape(shape); base != nil {
		return ctx.text(base)
	}
	return ""
}

func (ctx *context) geometrySections(shape *vsdx.Shape) []*vsdx.Section {
	sections := shape.SectionsNamed("Geometry")
	base := ctx.baseShape(shape)
	if base == nil {
		return ctx.resolveGeometrySectionBooleans(shape, cloneSections(sections))
	}
	baseSections := ctx.geometrySections(base)
	if len(sections) == 0 {
		return baseSections
	}
	if len(baseSections) == 0 {
		return ctx.resolveGeometrySectionBooleans(shape, cloneSections(sections))
	}
	return ctx.resolveGeometrySectionBooleans(shape, mergeSections(baseSections, sections))
}

func (ctx *context) baseShape(shape *vsdx.Shape) *vsdx.Shape {
	masterID := shape.ResolvedMasterID()
	if masterID == "" {
		return nil
	}
	master := ctx.document.Master(masterID)
	if master == nil {
		return nil
	}
	if shape.MasterShapeID != "" {
		return master.Shape(shape.MasterShapeID)
	}
	return master.RootShape
}

func (ctx *context) cell(shape *vsdx.Shape, name string) (string, bool) {
	cell, ok := ctx.cellData(shape, name)
	if !ok {
		return "", false
	}
	return cell.Value, true
}

func (ctx *context) cellData(shape *vsdx.Shape, name string) (vsdx.Cell, bool) {
	if cell, ok := shape.Cell(name); ok && (strings.TrimSpace(cell.Value) != "" || strings.TrimSpace(cell.Formula) != "") {
		return cell, true
	}
	if base := ctx.baseShape(shape); base != nil {
		return ctx.cellData(base, name)
	}
	return vsdx.Cell{}, false
}

func (ctx *context) sectionCell(shape *vsdx.Shape, sectionName, cellName string) (string, bool) {
	if sections := shape.SectionsNamed(sectionName); len(sections) > 0 {
		if cell, ok := sections[0].Cell(cellName); ok && strings.TrimSpace(cell.Value) != "" {
			return cell.Value, true
		}
	}
	if base := ctx.baseShape(shape); base != nil {
		return ctx.sectionCell(base, sectionName, cellName)
	}
	return "", false
}

func (ctx *context) rowCell(shape *vsdx.Shape, sectionName, cellName string) (string, bool) {
	cell, ok := ctx.rowCellData(shape, sectionName, cellName)
	if !ok {
		return "", false
	}
	return cell.Value, true
}

func (ctx *context) rowCellData(shape *vsdx.Shape, sectionName, cellName string) (vsdx.Cell, bool) {
	if sections := shape.SectionsNamed(sectionName); len(sections) > 0 && len(sections[0].Rows) > 0 {
		if cell, ok := sections[0].Rows[0].Cell(cellName); ok && (strings.TrimSpace(cell.Value) != "" || strings.TrimSpace(cell.Formula) != "") {
			return cell, true
		}
	}
	if base := ctx.baseShape(shape); base != nil {
		return ctx.rowCellData(base, sectionName, cellName)
	}
	return vsdx.Cell{}, false
}

func (ctx *context) boolCell(shape *vsdx.Shape, name string) (bool, bool) {
	cell, ok := ctx.cellData(shape, name)
	if !ok {
		return false, false
	}
	if value, ok := ctx.evalBoolFormula(shape, cell.Formula); ok {
		return value, true
	}
	return vsdx.ParseBool(cell.Value)
}

func (ctx *context) lengthCell(shape *vsdx.Shape, name string, fallback float64) float64 {
	if value, ok := ctx.lengthCellMaybe(shape, name); ok {
		return value
	}
	return fallback
}

func (ctx *context) lengthCellMaybe(shape *vsdx.Shape, name string) (float64, bool) {
	value, ok := ctx.cell(shape, name)
	if !ok {
		return 0, false
	}
	return vsdx.ParseLength(value)
}

func (ctx *context) angleCell(shape *vsdx.Shape, name string, fallback float64) float64 {
	value, ok := ctx.cell(shape, name)
	if !ok {
		return fallback
	}
	if parsed, ok := vsdx.ParseAngle(value); ok {
		return parsed
	}
	return fallback
}

func (ctx *context) textLength(shape *vsdx.Shape, name string, fallback float64) float64 {
	if value, ok := ctx.cell(shape, name); ok {
		if parsed, ok := vsdx.ParseLength(value); ok {
			return parsed
		}
	}
	if value, ok := ctx.sectionCell(shape, "TextXForm", name); ok {
		if parsed, ok := vsdx.ParseLength(value); ok {
			return parsed
		}
	}
	if value, ok := ctx.sectionCell(shape, "TextBlock", name); ok {
		if parsed, ok := vsdx.ParseLength(value); ok {
			return parsed
		}
	}
	return fallback
}

func (ctx *context) lineWidth(shape *vsdx.Shape) float64 {
	if value, ok := ctx.cell(shape, "LineWeight"); ok {
		if parsed, ok := vsdx.ParseLength(value); ok {
			return math.Max(parsed, 0.01)
		}
		if number, ok := vsdx.ParseNumber(value); ok {
			return math.Max(number, 0.01)
		}
	}
	return 1.0 / 72.0
}

func (ctx *context) linePattern(shape *vsdx.Shape) int {
	value, ok := ctx.cell(shape, "LinePattern")
	if !ok {
		return 1
	}
	number, ok := vsdx.ParseNumber(value)
	if !ok {
		return 1
	}
	return int(math.Round(number))
}

func (ctx *context) fillStyle(shape *vsdx.Shape) (color, bool) {
	if hidden, ok := ctx.boolCell(shape, "NoFill"); ok && hidden {
		return color{}, false
	}
	if value, ok := ctx.cell(shape, "FillPattern"); ok {
		if number, ok := vsdx.ParseNumber(value); ok && int(math.Round(number)) == 0 {
			return color{}, false
		}
	}
	if value, ok := ctx.cell(shape, "FillForegndTrans"); ok && transparencyValue(value) >= 1 {
		return color{}, false
	}
	if cell, ok := ctx.cellData(shape, "FillForegnd"); ok {
		if fill, ok := parseColorCell(cell.Value, cell.Formula); ok {
			return fill, true
		}
	}
	if fill, ok := parseColorValue(ctx.cellValue(shape, "FillForegnd")); ok {
		return fill, true
	}
	return color{R: 1, G: 1, B: 1}, true
}

func (ctx *context) strokeStyle(shape *vsdx.Shape) (color, bool) {
	if hidden, ok := ctx.boolCell(shape, "NoLine"); ok && hidden {
		return color{}, false
	}
	if ctx.linePattern(shape) == 0 {
		return color{}, false
	}
	if value, ok := ctx.cell(shape, "LineColorTrans"); ok && transparencyValue(value) >= 1 {
		return color{}, false
	}
	if cell, ok := ctx.cellData(shape, "LineColor"); ok {
		if stroke, ok := parseColorCell(cell.Value, cell.Formula); ok {
			return stroke, true
		}
		if strings.EqualFold(strings.TrimSpace(cell.Value), "Themed") {
			if fill, ok := ctx.fillStyle(shape); ok {
				return darkenColor(fill, 0.75), true
			}
		}
	}
	if stroke, ok := parseColorValue(ctx.cellValue(shape, "LineColor")); ok {
		return stroke, true
	}
	return color{R: 0, G: 0, B: 0}, true
}

func (ctx *context) textColor(shape *vsdx.Shape) color {
	if cell, ok := ctx.rowCellData(shape, "Character", "Color"); ok {
		if parsed, ok := parseColorCell(cell.Value, cell.Formula); ok {
			if strings.EqualFold(strings.TrimSpace(cell.Value), "Themed") {
				if fill, ok := ctx.fillStyle(shape); ok && colorLuminance(fill) < 0.45 {
					return color{R: 1, G: 1, B: 1}
				}
			}
			return parsed
		}
	}
	if fill, ok := ctx.fillStyle(shape); ok && colorLuminance(fill) < 0.45 {
		return color{R: 1, G: 1, B: 1}
	}
	return color{R: 0, G: 0, B: 0}
}

func (ctx *context) fontSize(shape *vsdx.Shape) float64 {
	if value, ok := ctx.rowCell(shape, "Character", "Size"); ok {
		if size, ok := parseFontSize(value); ok {
			return size
		}
	}
	return 12
}

func (ctx *context) textFont(shape *vsdx.Shape) *pdf.FontSpec {
	bold, italic := ctx.fontStyle(shape)
	return ctx.fonts.Resolve(ctx.fontFamily(shape), bold, italic)
}

func (ctx *context) fontFamily(shape *vsdx.Shape) string {
	if value, ok := ctx.rowCell(shape, "Character", "Font"); ok {
		trimmed := strings.TrimSpace(strings.Trim(value, "\""))
		switch {
		case trimmed == "":
		case strings.EqualFold(trimmed, "Themed"):
		default:
			if _, ok := vsdx.ParseNumber(trimmed); !ok {
				return trimmed
			}
		}
	}
	if strings.TrimSpace(ctx.document.DefaultFont) != "" {
		return ctx.document.DefaultFont
	}
	return "Calibri"
}

func (ctx *context) fontStyle(shape *vsdx.Shape) (bold bool, italic bool) {
	value, ok := ctx.rowCell(shape, "Character", "Style")
	if !ok {
		return false, false
	}
	trimmed := strings.TrimSpace(strings.Trim(value, "\""))
	if trimmed == "" || strings.EqualFold(trimmed, "Themed") {
		return false, false
	}
	parsed, ok := vsdx.ParseNumber(trimmed)
	if !ok {
		return false, false
	}
	flags := int(math.Round(parsed))
	return flags&1 != 0, flags&2 != 0
}

func (ctx *context) horizontalAlign(shape *vsdx.Shape) int {
	if value, ok := ctx.rowCell(shape, "Paragraph", "HorzAlign"); ok {
		if parsed, ok := vsdx.ParseNumber(value); ok {
			return int(math.Round(parsed))
		}
	}
	return 0
}

func (ctx *context) verticalAlign(shape *vsdx.Shape) int {
	if value, ok := ctx.cell(shape, "VerticalAlign"); ok {
		if parsed, ok := vsdx.ParseNumber(value); ok {
			return int(math.Round(parsed))
		}
	}
	if value, ok := ctx.sectionCell(shape, "TextBlock", "VerticalAlign"); ok {
		if parsed, ok := vsdx.ParseNumber(value); ok {
			return int(math.Round(parsed))
		}
	}
	return 1
}

func (ctx *context) cellValue(shape *vsdx.Shape, name string) string {
	value, _ := ctx.cell(shape, name)
	return value
}

func (ctx *context) fallbackPath(shape *vsdx.Shape, width, height float64) (vectorPath, bool) {
	if len(ctx.children(shape)) > 0 || width <= 0 || height <= 0 || shape.OneD {
		return vectorPath{}, false
	}

	name := strings.ToLower(shape.Name + " " + shape.NameU)
	switch {
	case strings.Contains(name, "ellipse"), strings.Contains(name, "circle"), strings.Contains(name, "oval"):
		center := point{X: width / 2, Y: height / 2}
		return ellipsePath(center, point{X: width / 2, Y: 0}, point{X: 0, Y: height / 2}), true
	default:
		path := vectorPath{}
		path.Move(point{X: 0, Y: 0})
		path.Line(point{X: width, Y: 0})
		path.Line(point{X: width, Y: height})
		path.Line(point{X: 0, Y: height})
		path.Close()
		return path, true
	}
}

func buildGeometry(width, height float64, sections []*vsdx.Section) []geometryRender {
	renders := make([]geometryRender, 0, len(sections))
	sorted := append([]*vsdx.Section(nil), sections...)
	sort.SliceStable(sorted, func(i, j int) bool {
		return sorted[i].Index < sorted[j].Index
	})

	for _, section := range sorted {
		if hidden, ok := sectionBool(section, "NoShow"); ok && hidden {
			continue
		}
		path := vectorPath{}
		current := point{}
		hasCurrent := false
		startPoint := point{}
		hasStart := false

		for _, row := range section.Rows {
			rowType := row.Type
			if rowType == "" {
				rowType = row.Name
			}

			switch rowType {
			case "MoveTo":
				pt := point{X: rowLength(row, "X"), Y: rowLength(row, "Y")}
				path.Move(pt)
				current = pt
				startPoint = pt
				hasCurrent = true
				hasStart = true
			case "RelMoveTo":
				pt := point{X: rowNumber(row, "X") * width, Y: rowNumber(row, "Y") * height}
				path.Move(pt)
				current = pt
				startPoint = pt
				hasCurrent = true
				hasStart = true
			case "LineTo":
				if !hasCurrent {
					continue
				}
				pt := point{X: rowLength(row, "X"), Y: rowLength(row, "Y")}
				path.Line(pt)
				current = pt
			case "RelLineTo":
				if !hasCurrent {
					continue
				}
				pt := point{X: rowNumber(row, "X") * width, Y: rowNumber(row, "Y") * height}
				path.Line(pt)
				current = pt
			case "ArcTo":
				if !hasCurrent {
					continue
				}
				end := point{X: rowLength(row, "X"), Y: rowLength(row, "Y")}
				addCircularArc(&path, current, end, rowLength(row, "A"))
				current = end
			case "RelArcTo":
				if !hasCurrent {
					continue
				}
				end := point{X: rowNumber(row, "X") * width, Y: rowNumber(row, "Y") * height}
				addCircularArc(&path, current, end, rowLength(row, "A"))
				current = end
			case "EllipticalArcTo":
				if !hasCurrent {
					continue
				}
				end := point{X: rowLength(row, "X"), Y: rowLength(row, "Y")}
				control := point{X: rowLength(row, "A"), Y: rowLength(row, "B")}
				addEllipticalArc(&path, current, control, end, rowAngle(row, "C"), rowNumberWithDefault(row, "D", 1))
				current = end
			case "RelEllipticalArcTo":
				if !hasCurrent {
					continue
				}
				end := point{X: rowNumber(row, "X") * width, Y: rowNumber(row, "Y") * height}
				control := point{X: rowNumber(row, "A") * width, Y: rowNumber(row, "B") * height}
				addEllipticalArc(&path, current, control, end, rowAngle(row, "C"), rowNumberWithDefault(row, "D", 1))
				current = end
			case "Ellipse":
				center := point{X: rowLength(row, "X"), Y: rowLength(row, "Y")}
				axisA := point{X: rowLength(row, "A") - center.X, Y: rowLength(row, "B") - center.Y}
				axisB := point{X: rowLength(row, "C") - center.X, Y: rowLength(row, "D") - center.Y}
				path = ellipsePath(center, axisA, axisB)
				current = point{X: center.X + axisA.X, Y: center.Y + axisA.Y}
				startPoint = current
				hasCurrent = true
				hasStart = true
			case "RelEllipse":
				center := point{X: rowNumber(row, "X") * width, Y: rowNumber(row, "Y") * height}
				axisA := point{X: rowNumber(row, "A")*width - center.X, Y: rowNumber(row, "B")*height - center.Y}
				axisB := point{X: rowNumber(row, "C")*width - center.X, Y: rowNumber(row, "D")*height - center.Y}
				path = ellipsePath(center, axisA, axisB)
				current = point{X: center.X + axisA.X, Y: center.Y + axisA.Y}
				startPoint = current
				hasCurrent = true
				hasStart = true
			case "NURBSTo", "SplineKnot", "PolylineTo", "RelCubBezTo", "CubBezTo", "RelQuadBezTo", "QuadBezTo":
				if !hasCurrent {
					continue
				}
				end := endpointFromRow(row, rowType, width, height)
				path.Line(end)
				current = end
			case "Close":
				if hasStart {
					path.Close()
					current = startPoint
				}
			}
		}

		if len(path.Ops) == 0 {
			continue
		}
		renders = append(renders, geometryRender{
			Path:   path,
			NoFill: sectionBoolValue(section, "NoFill"),
			NoLine: sectionBoolValue(section, "NoLine"),
		})
	}

	return renders
}

func endpointFromRow(row *vsdx.Row, rowType string, width, height float64) point {
	switch rowType {
	case "RelCubBezTo", "RelQuadBezTo", "PolylineTo", "SplineKnot", "NURBSTo":
		return point{X: rowNumber(row, "X") * width, Y: rowNumber(row, "Y") * height}
	default:
		return point{X: rowLength(row, "X"), Y: rowLength(row, "Y")}
	}
}

func addCircularArc(path *vectorPath, start, end point, bulge float64) {
	if nearlyEqual(bulge, 0) {
		path.Line(end)
		return
	}

	chord := subtract(end, start)
	chordLen := length(chord)
	if nearlyEqual(chordLen, 0) {
		return
	}

	sagitta := math.Abs(bulge)
	radius := (chordLen*chordLen)/(8*sagitta) + sagitta/2
	mid := point{X: (start.X + end.X) / 2, Y: (start.Y + end.Y) / 2}
	leftNormal := normalize(point{X: -chord.Y, Y: chord.X})
	sign := 1.0
	if bulge < 0 {
		sign = -1
	}
	centerOffset := radius - sagitta
	center := point{
		X: mid.X - leftNormal.X*sign*centerOffset,
		Y: mid.Y - leftNormal.Y*sign*centerOffset,
	}

	startAngle := math.Atan2(start.Y-center.Y, start.X-center.X)
	endAngle := math.Atan2(end.Y-center.Y, end.X-center.X)
	if bulge > 0 {
		for endAngle <= startAngle {
			endAngle += 2 * math.Pi
		}
	} else {
		for endAngle >= startAngle {
			endAngle -= 2 * math.Pi
		}
	}

	for _, segment := range arcSegments(center, radius, radius, 0, startAngle, endAngle) {
		path.Curve(segment[0], segment[1], segment[2])
	}
}

func addEllipticalArc(path *vectorPath, start, control, end point, rotation, ratio float64) {
	if nearlyEqual(ratio, 0) {
		path.Line(end)
		return
	}

	transform := func(p point) point {
		rotated := rotatePoint(p, -rotation)
		return point{X: rotated.X, Y: rotated.Y * ratio}
	}
	inverse := func(p point) point {
		scaled := point{X: p.X, Y: p.Y / ratio}
		return rotatePoint(scaled, rotation)
	}

	s := transform(start)
	c := transform(control)
	e := transform(end)
	center, radius, ok := circleFromThreePoints(s, c, e)
	if !ok {
		path.Line(end)
		return
	}

	startAngle := math.Atan2(s.Y-center.Y, s.X-center.X)
	controlAngle := math.Atan2(c.Y-center.Y, c.X-center.X)
	endAngle := math.Atan2(e.Y-center.Y, e.X-center.X)
	ccw := angleOnSweep(startAngle, controlAngle, endAngle, true)
	if ccw {
		for endAngle <= startAngle {
			endAngle += 2 * math.Pi
		}
	} else {
		for endAngle >= startAngle {
			endAngle -= 2 * math.Pi
		}
	}

	for _, segment := range arcSegments(center, radius, radius, 0, startAngle, endAngle) {
		path.Curve(inverse(segment[0]), inverse(segment[1]), inverse(segment[2]))
	}
}

func ellipsePath(center, axisA, axisB point) vectorPath {
	path := vectorPath{}
	start := point{X: center.X + axisA.X, Y: center.Y + axisA.Y}
	path.Move(start)

	step := math.Pi / 2
	for i := 0; i < 4; i++ {
		t0 := float64(i) * step
		t1 := float64(i+1) * step
		p0, c1, c2, p3 := ellipseBezier(center, axisA, axisB, t0, t1)
		if i == 0 {
			path.Ops[len(path.Ops)-1].Points[0] = p0
		}
		path.Curve(c1, c2, p3)
	}
	path.Close()
	return path
}

func ellipseBezier(center, axisA, axisB point, t0, t1 float64) (point, point, point, point) {
	p := func(t float64) point {
		return point{
			X: center.X + axisA.X*math.Cos(t) + axisB.X*math.Sin(t),
			Y: center.Y + axisA.Y*math.Cos(t) + axisB.Y*math.Sin(t),
		}
	}
	d := func(t float64) point {
		return point{
			X: -axisA.X*math.Sin(t) + axisB.X*math.Cos(t),
			Y: -axisA.Y*math.Sin(t) + axisB.Y*math.Cos(t),
		}
	}

	alpha := 4.0 / 3.0 * math.Tan((t1-t0)/4.0)
	p0 := p(t0)
	p3 := p(t1)
	d0 := d(t0)
	d1 := d(t1)
	c1 := point{X: p0.X + d0.X*alpha, Y: p0.Y + d0.Y*alpha}
	c2 := point{X: p3.X - d1.X*alpha, Y: p3.Y - d1.Y*alpha}
	return p0, c1, c2, p3
}

func arcSegments(center point, radiusX, radiusY, rotation, startAngle, endAngle float64) [][3]point {
	delta := endAngle - startAngle
	segments := int(math.Ceil(math.Abs(delta) / (math.Pi / 2)))
	if segments < 1 {
		segments = 1
	}
	step := delta / float64(segments)
	results := make([][3]point, 0, segments)

	for i := 0; i < segments; i++ {
		t0 := startAngle + float64(i)*step
		t1 := t0 + step
		alpha := 4.0 / 3.0 * math.Tan((t1-t0)/4.0)

		p0 := ellipsePoint(center, radiusX, radiusY, rotation, t0)
		p3 := ellipsePoint(center, radiusX, radiusY, rotation, t1)
		d0 := ellipseDerivative(radiusX, radiusY, rotation, t0)
		d1 := ellipseDerivative(radiusX, radiusY, rotation, t1)
		c1 := point{X: p0.X + d0.X*alpha, Y: p0.Y + d0.Y*alpha}
		c2 := point{X: p3.X - d1.X*alpha, Y: p3.Y - d1.Y*alpha}
		results = append(results, [3]point{c1, c2, p3})
	}

	return results
}

func ellipsePoint(center point, radiusX, radiusY, rotation, t float64) point {
	cosT := math.Cos(t)
	sinT := math.Sin(t)
	x := radiusX * cosT
	y := radiusY * sinT
	return rotateAndTranslate(point{X: x, Y: y}, rotation, center)
}

func ellipseDerivative(radiusX, radiusY, rotation, t float64) point {
	x := -radiusX * math.Sin(t)
	y := radiusY * math.Cos(t)
	return rotatePoint(point{X: x, Y: y}, rotation)
}

func circleFromThreePoints(a, b, c point) (point, float64, bool) {
	d := 2 * (a.X*(b.Y-c.Y) + b.X*(c.Y-a.Y) + c.X*(a.Y-b.Y))
	if nearlyEqual(d, 0) {
		return point{}, 0, false
	}
	ax2ay2 := a.X*a.X + a.Y*a.Y
	bx2by2 := b.X*b.X + b.Y*b.Y
	cx2cy2 := c.X*c.X + c.Y*c.Y
	center := point{
		X: (ax2ay2*(b.Y-c.Y) + bx2by2*(c.Y-a.Y) + cx2cy2*(a.Y-b.Y)) / d,
		Y: (ax2ay2*(c.X-b.X) + bx2by2*(a.X-c.X) + cx2cy2*(b.X-a.X)) / d,
	}
	return center, distance(center, a), true
}

func angleOnSweep(start, mid, end float64, ccw bool) bool {
	if ccw {
		for end < start {
			end += 2 * math.Pi
		}
		for mid < start {
			mid += 2 * math.Pi
		}
		return mid > start && mid < end
	}
	for end > start {
		end -= 2 * math.Pi
	}
	for mid > start {
		mid -= 2 * math.Pi
	}
	return mid < start && mid > end
}

func sectionBool(section *vsdx.Section, name string) (bool, bool) {
	cell, ok := section.Cell(name)
	if !ok {
		return false, false
	}
	return vsdx.ParseBool(cell.Value)
}

func sectionBoolValue(section *vsdx.Section, name string) bool {
	value, ok := sectionBool(section, name)
	return ok && value
}

func rowLength(row *vsdx.Row, name string) float64 {
	if cell, ok := row.Cell(name); ok {
		if parsed, ok := vsdx.ParseLength(cell.Value); ok {
			return parsed
		}
		if parsed, ok := vsdx.ParseNumber(cell.Value); ok {
			return parsed
		}
	}
	return 0
}

func rowNumber(row *vsdx.Row, name string) float64 {
	if cell, ok := row.Cell(name); ok {
		if parsed, ok := vsdx.ParseNumber(cell.Value); ok {
			return parsed
		}
	}
	return 0
}

func rowNumberWithDefault(row *vsdx.Row, name string, fallback float64) float64 {
	if cell, ok := row.Cell(name); ok {
		if parsed, ok := vsdx.ParseNumber(cell.Value); ok {
			return parsed
		}
	}
	return fallback
}

func rowAngle(row *vsdx.Row, name string) float64 {
	if cell, ok := row.Cell(name); ok {
		if parsed, ok := vsdx.ParseAngle(cell.Value); ok {
			return parsed
		}
		if parsed, ok := vsdx.ParseNumber(cell.Value); ok {
			return parsed
		}
	}
	return 0
}

func parseColorValue(value string) (color, bool) {
	trimmed := strings.TrimSpace(strings.Trim(value, "\""))
	if trimmed == "" {
		return color{}, false
	}
	if strings.HasPrefix(trimmed, "#") && len(trimmed) == 7 {
		r, err1 := parseHexPair(trimmed[1:3])
		g, err2 := parseHexPair(trimmed[3:5])
		b, err3 := parseHexPair(trimmed[5:7])
		if err1 == nil && err2 == nil && err3 == nil {
			return color{R: float64(r) / 255, G: float64(g) / 255, B: float64(b) / 255}, true
		}
	}
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "rgb(") && strings.HasSuffix(lower, ")") {
		parts := strings.Split(lower[4:len(lower)-1], ",")
		if len(parts) == 3 {
			values := [3]float64{}
			for i, part := range parts {
				number, ok := vsdx.ParseNumber(part)
				if !ok {
					return color{}, false
				}
				values[i] = number / 255
			}
			return color{R: values[0], G: values[1], B: values[2]}, true
		}
	}
	if number, ok := vsdx.ParseNumber(trimmed); ok {
		palette := []color{
			{0, 0, 0}, {1, 1, 1}, {1, 0, 0}, {0, 1, 0}, {0, 0, 1},
			{1, 1, 0}, {1, 0, 1}, {0, 1, 1}, {0.75, 0.75, 0.75}, {0.5, 0.5, 0.5},
		}
		index := int(math.Round(number))
		if index >= 0 && index < len(palette) {
			return palette[index], true
		}
	}
	return color{}, false
}

func parseColorCell(value, formula string) (color, bool) {
	if parsed, ok := parseColorValue(value); ok {
		return parsed, true
	}

	formula = strings.TrimSpace(formula)
	if formula == "" {
		return color{}, false
	}

	if index := strings.Index(formula, "#"); index >= 0 && index+7 <= len(formula) {
		if parsed, ok := parseColorValue(formula[index : index+7]); ok {
			return parsed, true
		}
	}

	upper := strings.ToUpper(formula)
	if start := strings.Index(upper, "RGB("); start >= 0 {
		end := strings.Index(upper[start:], ")")
		if end > 0 {
			if parsed, ok := parseColorValue(formula[start : start+end+1]); ok {
				return parsed, true
			}
		}
	}

	return color{}, false
}

func parseFontSize(value string) (float64, bool) {
	number, unit := splitValue(value)
	if number == "" {
		return 0, false
	}
	parsed, ok := vsdx.ParseNumber(number)
	if !ok {
		return 0, false
	}
	switch strings.Trim(strings.ToLower(unit), ".") {
	case "", "pt":
		if unit == "" && parsed < 3 {
			return parsed * 72, true
		}
		return parsed, true
	case "in":
		return parsed * 72, true
	default:
		if length, ok := vsdx.ParseLength(value); ok {
			return length * 72, true
		}
	}
	return parsed, true
}

func splitValue(value string) (string, string) {
	s := strings.TrimSpace(strings.Trim(value, "\""))
	if s == "" {
		return "", ""
	}
	i := 0
	if s[0] == '+' || s[0] == '-' {
		i++
	}
	for i < len(s) {
		ch := s[i]
		if (ch >= '0' && ch <= '9') || ch == '.' || ch == 'e' || ch == 'E' || ch == '+' || ch == '-' {
			i++
			continue
		}
		break
	}
	return strings.TrimSpace(s[:i]), strings.TrimSpace(s[i:])
}

func transparencyValue(value string) float64 {
	if parsed, ok := vsdx.ParseNumber(value); ok {
		if parsed > 1 {
			return parsed / 100
		}
		return parsed
	}
	return 0
}

func parseHexPair(value string) (int64, error) {
	return strconv.ParseInt(value, 16, 64)
}

func colorLuminance(value color) float64 {
	return value.R*0.299 + value.G*0.587 + value.B*0.114
}

func darkenColor(value color, factor float64) color {
	if factor < 0 {
		factor = 0
	}
	if factor > 1 {
		factor = 1
	}
	return color{
		R: value.R * factor,
		G: value.G * factor,
		B: value.B * factor,
	}
}

func wrapText(text string, maxWidth, fontSize float64, measure func(string, float64) float64) []string {
	if strings.TrimSpace(text) == "" {
		return nil
	}

	var lines []string
	paragraphs := strings.Split(text, "\n")
	for _, paragraph := range paragraphs {
		words := strings.Fields(paragraph)
		if len(words) == 0 {
			lines = append(lines, "")
			continue
		}
		current := words[0]
		for _, word := range words[1:] {
			candidate := current + " " + word
			if measure(candidate, fontSize) <= maxWidth {
				current = candidate
				continue
			}
			if measure(word, fontSize) > maxWidth {
				lines = append(lines, hardWrap(word, maxWidth, fontSize, measure)...)
				current = ""
				continue
			}
			lines = append(lines, current)
			current = word
		}
		if current != "" {
			lines = append(lines, current)
		}
	}
	return lines
}

func hardWrap(word string, maxWidth, fontSize float64, measure func(string, float64) float64) []string {
	if word == "" {
		return nil
	}
	var lines []string
	var builder strings.Builder
	for _, r := range word {
		next := builder.String() + string(r)
		if builder.Len() > 0 && measure(next, fontSize) > maxWidth {
			lines = append(lines, builder.String())
			builder.Reset()
		}
		builder.WriteRune(r)
	}
	if builder.Len() > 0 {
		lines = append(lines, builder.String())
	}
	return lines
}

func fitTextToBox(text string, maxWidth, maxHeight, initialFontSize float64, measure func(string, float64) float64) ([]string, float64) {
	if initialFontSize <= 0 {
		initialFontSize = 12
	}

	fontSize := initialFontSize
	minFontSize := math.Min(initialFontSize, 4)
	if minFontSize <= 0 {
		minFontSize = 4
	}
	if maxHeight <= 0 {
		return wrapText(text, maxWidth, fontSize, measure), fontSize
	}

	for {
		lines := wrapText(text, maxWidth, fontSize, measure)
		if len(lines) == 0 {
			return nil, fontSize
		}
		if float64(len(lines))*textLineHeight(fontSize) <= maxHeight+0.01 {
			return lines, fontSize
		}
		if fontSize <= minFontSize {
			return lines, fontSize
		}
		fontSize -= 0.5
		if fontSize < minFontSize {
			fontSize = minFontSize
		}
	}
}

func textLineHeight(fontSize float64) float64 {
	if fontSize <= 8 {
		return fontSize * 1.05
	}
	return fontSize * 1.15
}

func estimateTextWidth(text string, fontSize float64) float64 {
	total := 0.0
	for _, r := range text {
		switch {
		case r == ' ':
			total += fontSize * 0.28
		case unicode.IsUpper(r):
			total += fontSize * 0.62
		case unicode.IsDigit(r):
			total += fontSize * 0.55
		case strings.ContainsRune("il.,:;'|!", r):
			total += fontSize * 0.25
		default:
			total += fontSize * 0.52
		}
	}
	return total
}

func measureTextWidth(text string, font *pdf.FontSpec, fontSize float64) float64 {
	if font != nil && font.TrueType != nil {
		return font.TrueType.MeasureText(text, fontSize)
	}
	return estimateTextWidth(text, fontSize)
}

func (ctx *context) resolveGeometrySectionBooleans(shape *vsdx.Shape, sections []*vsdx.Section) []*vsdx.Section {
	for _, section := range sections {
		if section == nil {
			continue
		}
		for _, name := range []string{"NoShow", "NoFill", "NoLine"} {
			cell, ok := section.Cells[name]
			if !ok {
				continue
			}
			if value, ok := ctx.evalBoolFormula(shape, cell.Formula); ok {
				if value {
					cell.Value = "1"
				} else {
					cell.Value = "0"
				}
				section.Cells[name] = cell
			}
		}
	}
	return sections
}

func (ctx *context) evalBoolFormula(shape *vsdx.Shape, formula string) (bool, bool) {
	value, ok := ctx.evalFormula(shape, formula)
	if !ok {
		return false, false
	}
	return parseFormulaBool(value)
}

func (ctx *context) evalFormula(shape *vsdx.Shape, formula string) (string, bool) {
	expr := trimFormula(formula)
	if expr == "" || strings.EqualFold(expr, "Inh") || strings.EqualFold(expr, "No Formula") {
		return "", false
	}

	if value, ok := parseFormulaLiteral(expr); ok {
		return value, true
	}

	if name, args, ok := parseFormulaCall(expr); ok {
		switch strings.ToUpper(name) {
		case "NOT":
			if len(args) != 1 {
				return "", false
			}
			value, ok := ctx.evalBoolFormula(shape, args[0])
			if !ok {
				return "", false
			}
			return boolString(!value), true
		case "OR":
			if len(args) == 0 {
				return "", false
			}
			for _, arg := range args {
				value, ok := ctx.evalBoolFormula(shape, arg)
				if !ok {
					return "", false
				}
				if value {
					return "1", true
				}
			}
			return "0", true
		case "AND":
			if len(args) == 0 {
				return "", false
			}
			for _, arg := range args {
				value, ok := ctx.evalBoolFormula(shape, arg)
				if !ok {
					return "", false
				}
				if !value {
					return "0", true
				}
			}
			return "1", true
		case "IF":
			if len(args) != 3 {
				return "", false
			}
			condition, ok := ctx.evalBoolFormula(shape, args[0])
			if !ok {
				return "", false
			}
			if condition {
				return ctx.evalFormula(shape, args[1])
			}
			return ctx.evalFormula(shape, args[2])
		default:
			return "", false
		}
	}

	if index, op, ok := formulaComparison(expr); ok {
		leftValue, ok := ctx.evalFormula(shape, expr[:index])
		if !ok {
			return "", false
		}
		rightValue, ok := ctx.evalFormula(shape, expr[index+len(op):])
		if !ok {
			return "", false
		}
		return boolString(compareFormulaValues(leftValue, rightValue, op)), true
	}

	if value, ok := ctx.resolveFormulaReference(shape, expr); ok {
		return value, true
	}

	return strings.Trim(strings.TrimSpace(expr), `"`), true
}

func (ctx *context) resolveFormulaReference(shape *vsdx.Shape, expr string) (string, bool) {
	trimmed := strings.TrimSpace(expr)
	upper := strings.ToUpper(trimmed)
	switch {
	case strings.HasPrefix(upper, "USER."):
		return ctx.userValue(shape, trimmed[5:])
	case strings.Contains(upper, "!USER."):
		parts := strings.SplitN(trimmed, "!", 2)
		if len(parts) != 2 {
			return "", false
		}
		target := ctx.shapeByFormulaRef(shape, parts[0])
		if target == nil {
			return "", false
		}
		return ctx.userValue(target, strings.TrimPrefix(parts[1], "User."))
	default:
		return "", false
	}
}

func (ctx *context) userValue(shape *vsdx.Shape, rowName string) (string, bool) {
	for _, section := range shape.SectionsNamed("User") {
		for _, row := range section.Rows {
			if strings.EqualFold(row.Name, rowName) {
				if cell, ok := row.Cell("Value"); ok && strings.TrimSpace(cell.Value) != "" {
					return cell.Value, true
				}
			}
		}
	}
	if base := ctx.baseShape(shape); base != nil {
		return ctx.userValue(base, rowName)
	}
	return "", false
}

func (ctx *context) shapeByFormulaRef(current *vsdx.Shape, ref string) *vsdx.Shape {
	trimmed := strings.TrimSpace(ref)
	upper := strings.ToUpper(trimmed)
	if strings.HasPrefix(upper, "SHEET.") {
		id := strings.TrimSpace(trimmed[len("Sheet."):])
		for _, page := range ctx.document.Pages {
			if found := findShapeByID(page.Shapes, id); found != nil {
				return found
			}
		}
		for _, masterID := range []string{current.ResolvedMasterID()} {
			if masterID == "" {
				continue
			}
			if master := ctx.document.Master(masterID); master != nil {
				if found := master.Shape(id); found != nil {
					return found
				}
			}
		}
	}
	return nil
}

func findShapeByID(shapes []*vsdx.Shape, id string) *vsdx.Shape {
	for _, shape := range shapes {
		if shape == nil {
			continue
		}
		if shape.ID == id {
			return shape
		}
		if found := findShapeByID(shape.Shapes, id); found != nil {
			return found
		}
	}
	return nil
}

func parseFormulaLiteral(expr string) (string, bool) {
	trimmed := strings.Trim(strings.TrimSpace(expr), `"`)
	switch strings.ToUpper(trimmed) {
	case "TRUE":
		return "1", true
	case "FALSE":
		return "0", true
	}
	if _, ok := vsdx.ParseNumber(trimmed); ok {
		return trimmed, true
	}
	return "", false
}

func parseFormulaCall(expr string) (string, []string, bool) {
	expr = trimFormula(expr)
	open := strings.Index(expr, "(")
	if open <= 0 || !strings.HasSuffix(expr, ")") {
		return "", nil, false
	}
	name := strings.TrimSpace(expr[:open])
	body := expr[open+1 : len(expr)-1]
	if name == "" {
		return "", nil, false
	}
	return name, splitFormulaArgs(body), true
}

func splitFormulaArgs(expr string) []string {
	if strings.TrimSpace(expr) == "" {
		return nil
	}
	args := []string{}
	depth := 0
	start := 0
	for index := 0; index < len(expr); index++ {
		switch expr[index] {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case ',':
			if depth == 0 {
				args = append(args, strings.TrimSpace(expr[start:index]))
				start = index + 1
			}
		}
	}
	args = append(args, strings.TrimSpace(expr[start:]))
	return args
}

func formulaComparison(expr string) (int, string, bool) {
	depth := 0
	for index := 0; index < len(expr); index++ {
		switch expr[index] {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case '<':
			if depth == 0 && index+1 < len(expr) && expr[index+1] == '>' {
				return index, "<>", true
			}
		case '=':
			if depth == 0 {
				return index, "=", true
			}
		}
	}
	return 0, "", false
}

func compareFormulaValues(left, right, op string) bool {
	if leftNumber, ok := vsdx.ParseNumber(left); ok {
		if rightNumber, ok := vsdx.ParseNumber(right); ok {
			switch op {
			case "=":
				return math.Abs(leftNumber-rightNumber) < 1e-9
			case "<>":
				return math.Abs(leftNumber-rightNumber) >= 1e-9
			}
		}
	}
	if leftBool, ok := parseFormulaBool(left); ok {
		if rightBool, ok := parseFormulaBool(right); ok {
			switch op {
			case "=":
				return leftBool == rightBool
			case "<>":
				return leftBool != rightBool
			}
		}
	}
	switch op {
	case "=":
		return strings.EqualFold(strings.Trim(left, `"`), strings.Trim(right, `"`))
	case "<>":
		return !strings.EqualFold(strings.Trim(left, `"`), strings.Trim(right, `"`))
	default:
		return false
	}
}

func parseFormulaBool(value string) (bool, bool) {
	if parsed, ok := vsdx.ParseBool(value); ok {
		return parsed, true
	}
	if number, ok := vsdx.ParseNumber(value); ok {
		return math.Abs(number) > 1e-9, true
	}
	return false, false
}

func boolString(value bool) string {
	if value {
		return "1"
	}
	return "0"
}

func trimFormula(expr string) string {
	expr = strings.TrimSpace(expr)
	for strings.HasPrefix(expr, "(") && strings.HasSuffix(expr, ")") {
		depth := 0
		whole := true
		for index := 0; index < len(expr); index++ {
			switch expr[index] {
			case '(':
				depth++
			case ')':
				depth--
				if depth == 0 && index != len(expr)-1 {
					whole = false
				}
			}
		}
		if !whole {
			break
		}
		expr = strings.TrimSpace(expr[1 : len(expr)-1])
	}
	return expr
}

func cloneSections(sections []*vsdx.Section) []*vsdx.Section {
	if len(sections) == 0 {
		return nil
	}
	cloned := make([]*vsdx.Section, 0, len(sections))
	for _, section := range sections {
		cloned = append(cloned, cloneSection(section))
	}
	return cloned
}

func cloneSection(section *vsdx.Section) *vsdx.Section {
	if section == nil {
		return nil
	}
	cloned := &vsdx.Section{
		Name:  section.Name,
		Index: section.Index,
		Cells: cloneCellMap(section.Cells),
		Rows:  make([]*vsdx.Row, 0, len(section.Rows)),
	}
	for _, row := range section.Rows {
		cloned.Rows = append(cloned.Rows, cloneRow(row))
	}
	return cloned
}

func cloneRow(row *vsdx.Row) *vsdx.Row {
	if row == nil {
		return nil
	}
	return &vsdx.Row{
		Name:  row.Name,
		Type:  row.Type,
		Index: row.Index,
		Cells: cloneCellMap(row.Cells),
	}
}

func cloneCellMap(source map[string]vsdx.Cell) map[string]vsdx.Cell {
	if len(source) == 0 {
		return map[string]vsdx.Cell{}
	}
	cloned := make(map[string]vsdx.Cell, len(source))
	for key, cell := range source {
		cloned[key] = cell
	}
	return cloned
}

func mergeSections(base, override []*vsdx.Section) []*vsdx.Section {
	merged := make([]*vsdx.Section, 0, maxInt(len(base), len(override)))
	used := make([]bool, len(override))

	for index, baseSection := range base {
		overrideIndex := matchingSection(baseSection, override, used, index)
		if overrideIndex == -1 {
			merged = append(merged, cloneSection(baseSection))
			continue
		}
		used[overrideIndex] = true
		merged = append(merged, mergeSection(baseSection, override[overrideIndex]))
	}

	for index, section := range override {
		if used[index] {
			continue
		}
		merged = append(merged, cloneSection(section))
	}

	return merged
}

func matchingSection(baseSection *vsdx.Section, override []*vsdx.Section, used []bool, fallbackIndex int) int {
	if baseSection != nil && baseSection.Index != "" {
		for index, section := range override {
			if used[index] || section == nil {
				continue
			}
			if section.Index == baseSection.Index {
				return index
			}
		}
		return -1
	}
	if anySectionHasIndex(override) {
		return -1
	}
	if fallbackIndex < len(override) && !used[fallbackIndex] {
		return fallbackIndex
	}
	return -1
}

func mergeSection(baseSection, overrideSection *vsdx.Section) *vsdx.Section {
	if baseSection == nil {
		return cloneSection(overrideSection)
	}
	if overrideSection == nil {
		return cloneSection(baseSection)
	}

	merged := &vsdx.Section{
		Name:  firstNonEmpty(overrideSection.Name, baseSection.Name),
		Index: firstNonEmpty(overrideSection.Index, baseSection.Index),
		Cells: cloneCellMap(baseSection.Cells),
		Rows:  mergeRows(baseSection.Rows, overrideSection.Rows),
	}
	for key, cell := range overrideSection.Cells {
		merged.Cells[key] = cell
	}
	return merged
}

func mergeRows(baseRows, overrideRows []*vsdx.Row) []*vsdx.Row {
	merged := make([]*vsdx.Row, 0, maxInt(len(baseRows), len(overrideRows)))
	used := make([]bool, len(overrideRows))

	for index, baseRow := range baseRows {
		overrideIndex := matchingRow(baseRow, overrideRows, used, index)
		if overrideIndex == -1 {
			merged = append(merged, cloneRow(baseRow))
			continue
		}
		used[overrideIndex] = true
		merged = append(merged, mergeRow(baseRow, overrideRows[overrideIndex]))
	}

	for index, row := range overrideRows {
		if used[index] {
			continue
		}
		merged = append(merged, cloneRow(row))
	}

	return merged
}

func matchingRow(baseRow *vsdx.Row, overrideRows []*vsdx.Row, used []bool, fallbackIndex int) int {
	if baseRow != nil && baseRow.Index != "" {
		for index, row := range overrideRows {
			if used[index] || row == nil {
				continue
			}
			if row.Index == baseRow.Index {
				return index
			}
		}
		return -1
	}
	if anyRowHasIndex(overrideRows) {
		return -1
	}
	if fallbackIndex < len(overrideRows) && !used[fallbackIndex] {
		return fallbackIndex
	}
	return -1
}

func mergeRow(baseRow, overrideRow *vsdx.Row) *vsdx.Row {
	if baseRow == nil {
		return cloneRow(overrideRow)
	}
	if overrideRow == nil {
		return cloneRow(baseRow)
	}

	merged := &vsdx.Row{
		Name:  firstNonEmpty(overrideRow.Name, baseRow.Name),
		Type:  firstNonEmpty(overrideRow.Type, baseRow.Type),
		Index: firstNonEmpty(overrideRow.Index, baseRow.Index),
		Cells: cloneCellMap(baseRow.Cells),
	}
	for key, cell := range overrideRow.Cells {
		merged.Cells[key] = cell
	}
	return merged
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func anySectionHasIndex(sections []*vsdx.Section) bool {
	for _, section := range sections {
		if section != nil && strings.TrimSpace(section.Index) != "" {
			return true
		}
	}
	return false
}

func anyRowHasIndex(rows []*vsdx.Row) bool {
	for _, row := range rows {
		if row != nil && strings.TrimSpace(row.Index) != "" {
			return true
		}
	}
	return false
}

func identityMatrix() matrix {
	return matrix{A: 1, D: 1}
}

func translate(x, y float64) matrix {
	return matrix{A: 1, D: 1, E: x, F: y}
}

func scale(x, y float64) matrix {
	return matrix{A: x, D: y}
}

func rotate(angle float64) matrix {
	sin, cos := math.Sincos(angle)
	return matrix{A: cos, B: sin, C: -sin, D: cos}
}

func multiply(left, right matrix) matrix {
	return matrix{
		A: left.A*right.A + left.C*right.B,
		B: left.B*right.A + left.D*right.B,
		C: left.A*right.C + left.C*right.D,
		D: left.B*right.C + left.D*right.D,
		E: left.A*right.E + left.C*right.F + left.E,
		F: left.B*right.E + left.D*right.F + left.F,
	}
}

func apply(m matrix, p point) point {
	return point{
		X: m.A*p.X + m.C*p.Y + m.E,
		Y: m.B*p.X + m.D*p.Y + m.F,
	}
}

func applyInPoints(m matrix, p point) point {
	p = apply(m, p)
	return point{X: p.X * 72, Y: p.Y * 72}
}

func rotatePoint(p point, angle float64) point {
	sin, cos := math.Sincos(angle)
	return point{
		X: p.X*cos - p.Y*sin,
		Y: p.X*sin + p.Y*cos,
	}
}

func rotateAndTranslate(p point, angle float64, offset point) point {
	rotated := rotatePoint(p, angle)
	return point{X: rotated.X + offset.X, Y: rotated.Y + offset.Y}
}

func subtract(a, b point) point {
	return point{X: a.X - b.X, Y: a.Y - b.Y}
}

func distance(a, b point) float64 {
	return length(subtract(a, b))
}

func length(p point) float64 {
	return math.Hypot(p.X, p.Y)
}

func normalize(p point) point {
	l := length(p)
	if nearlyEqual(l, 0) {
		return point{}
	}
	return point{X: p.X / l, Y: p.Y / l}
}

func nearlyEqual(a, b float64) bool {
	return math.Abs(a-b) < 1e-9
}

func (p *vectorPath) Move(pt point) {
	p.Ops = append(p.Ops, pathOp{Kind: "m", Points: []point{pt}})
}

func (p *vectorPath) Line(pt point) {
	p.Ops = append(p.Ops, pathOp{Kind: "l", Points: []point{pt}})
}

func (p *vectorPath) Curve(c1, c2, pt point) {
	p.Ops = append(p.Ops, pathOp{Kind: "c", Points: []point{c1, c2, pt}})
}

func (p *vectorPath) Close() {
	p.Ops = append(p.Ops, pathOp{Kind: "h"})
}
