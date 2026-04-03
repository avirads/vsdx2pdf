package vsdx_test

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"

	"vsdx2pdf/internal/render"
	"vsdx2pdf/internal/vsdx"
)

func TestParseAndRenderSimplePackage(t *testing.T) {
	data := buildFixturePackage(t)

	document, err := vsdx.Parse(data)
	if err != nil {
		t.Fatalf("Parse() error = %v", err)
	}

	if len(document.Pages) != 1 {
		t.Fatalf("expected 1 page, got %d", len(document.Pages))
	}

	page := document.Pages[0]
	if page.Name != "Page-1" {
		t.Fatalf("page name = %q", page.Name)
	}
	if page.Width != 20 || page.Height != 30 {
		t.Fatalf("page size = %vx%v, want 20x30", page.Width, page.Height)
	}
	if len(page.Shapes) != 1 {
		t.Fatalf("expected 1 shape, got %d", len(page.Shapes))
	}

	pdfBytes, err := render.Convert(document)
	if err != nil {
		t.Fatalf("Convert() error = %v", err)
	}

	output := string(pdfBytes)
	if !strings.Contains(output, "%PDF-1.4") {
		t.Fatalf("output is not a PDF")
	}
	if !strings.Contains(output, "(Hello world) Tj") {
		t.Fatalf("expected rendered text in output PDF")
	}
	if !strings.Contains(output, " m\n") || !strings.Contains(output, " l\n") {
		t.Fatalf("expected path commands in output PDF")
	}
}

func buildFixturePackage(t *testing.T) []byte {
	t.Helper()

	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)

	files := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.document.main+xml"/>
  <Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>
  <Override PartName="/visio/pages/page1.xml" ContentType="application/vnd.ms-visio.page+xml"/>
</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document" Target="visio/document.xml"/>
</Relationships>`,
		"visio/document.xml": `<?xml version="1.0" encoding="UTF-8"?>
<VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main"/>`,
		"visio/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/pages" Target="pages/pages.xml"/>
</Relationships>`,
		"visio/pages/pages.xml": `<?xml version="1.0" encoding="UTF-8"?>
<Pages xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Page ID="1" Name="Page-1" NameU="Page-1">
    <PageSheet>
      <Cell N="PageWidth" V="20" U="IN"/>
      <Cell N="PageHeight" V="30" U="IN"/>
    </PageSheet>
    <Rel r:id="rId1"/>
  </Page>
</Pages>`,
		"visio/pages/_rels/pages.xml.rels": `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page1.xml"/>
</Relationships>`,
		"visio/pages/page1.xml": `<?xml version="1.0" encoding="UTF-8"?>
<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main">
  <Shapes>
    <Shape ID="1" Name="Rect" NameU="Rectangle">
      <Cell N="PinX" V="4.25"/>
      <Cell N="PinY" V="5.5"/>
      <Cell N="Width" V="4"/>
      <Cell N="Height" V="2"/>
      <Cell N="LocPinX" V="2"/>
      <Cell N="LocPinY" V="1"/>
      <Cell N="FillForegnd" V="#D9EAF7"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="LineColor" V="#1F4E79"/>
      <Cell N="LinePattern" V="1"/>
      <Cell N="LineWeight" V="0.02"/>
      <Text>Hello world</Text>
      <Section N="Character">
        <Row IX="0">
          <Cell N="Size" V="12" U="PT"/>
          <Cell N="Color" V="#102030"/>
        </Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0">
          <Cell N="HorzAlign" V="1"/>
        </Row>
      </Section>
      <Section N="Geometry" IX="0">
        <Row T="MoveTo">
          <Cell N="X" V="0"/>
          <Cell N="Y" V="0"/>
        </Row>
        <Row T="LineTo">
          <Cell N="X" V="4"/>
          <Cell N="Y" V="0"/>
        </Row>
        <Row T="LineTo">
          <Cell N="X" V="4"/>
          <Cell N="Y" V="2"/>
        </Row>
        <Row T="LineTo">
          <Cell N="X" V="0"/>
          <Cell N="Y" V="2"/>
        </Row>
        <Row T="Close"/>
      </Section>
    </Shape>
  </Shapes>
</PageContents>`,
	}

	for name, content := range files {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatalf("Create(%q) error = %v", name, err)
		}
		if _, err := entry.Write([]byte(content)); err != nil {
			t.Fatalf("Write(%q) error = %v", name, err)
		}
	}

	if err := writer.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	return buffer.Bytes()
}
