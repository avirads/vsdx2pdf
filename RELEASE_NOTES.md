# v1.0.1

This release improves PDF fidelity for complex Visio documents and updates the release workflow to publish explicit notes with each tagged build.

## Rendering Improvements

- Embed TrueType fonts in generated PDFs instead of forcing all text through Helvetica.
- Detect and use local Calibri-compatible font variants for regular, bold, italic, and bold-italic text when available.
- Use actual font metrics for PDF text measurement and wrapping.
- Read default document font metadata from the VSDX package and use it during rendering.
- Respect page sizes defined in `visio/pages/pages.xml`, not just per-page content parts.
- Preserve inherited `MasterID` context for nested shapes and improve child/master shape merging.
- Merge inherited geometry sections and rows by Visio indices so partial overrides do not corrupt paths.
- Evaluate formula-driven boolean visibility/style gates such as `NoShow`, `NoFill`, and `NoLine` for master-based geometry variants.
- Improve themed color fallback handling for fill, stroke, and text in common master-based shapes.

## Verification

- `go test ./...`
- Re-checked rendering locally against a private LibreOffice-generated reference PDF used during parity work.

## Notes

- Output is materially closer on large multi-page documents, but this is still a native renderer and not yet full LibreOffice parity.
- No LibreOffice runtime has been bundled in this release.
