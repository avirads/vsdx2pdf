# About

`vsdx2pdf` exists to solve one specific problem: convert modern Visio documents to PDF fully offline, from a single executable, with no dependency on Visio or any external conversion service.

## Design

- Language: Go
- Runtime dependencies: none
- Rendering approach: direct `.vsdx` package parsing plus custom PDF generation
- Packaging model: single binary per operating system

The converter reads the zipped Office/Visio package, resolves pages and master relationships, interprets the Visio shape tree, and writes PDF drawing commands directly.

## Supported Features

- Multi-page documents
- Background pages
- Grouped shapes
- Common master-based shapes
- Solid fills and strokes
- Basic dashed line patterns
- Text extraction and basic text layout
- Shape transforms: translation, rotation, flip
- Common geometry rows:
  - `MoveTo`
  - `LineTo`
  - `ArcTo`
  - `EllipticalArcTo`
  - `Ellipse`
  - Relative variants of those rows where commonly used

## Known Limitations

- Advanced Visio theme evaluation is only partially covered
- Raster images and OLE/foreign objects are not rendered yet
- Complex spline/NURBS/polyline semantics currently fall back to simpler line output
- Transparency is treated conservatively rather than fully modeled in PDF graphics state
- Text styling is intentionally basic and currently uses a standard PDF font
- Pattern fills are simplified to solid foreground fills

## Release Strategy

The repository includes a GitHub Actions release workflow. When a semantic version tag is pushed, GitHub Actions builds the project on Windows, Linux, and macOS, bundles the binaries, and attaches them to the release.
