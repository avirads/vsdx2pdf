# Usage

## Command

```text
vsdx2pdf [flags] <input.vsdx>
```

## Flags

- `-o <path>`: write the generated PDF to a specific path
- `-list-pages`: print page names and IDs without converting
- `-version`: print the build version

## Examples

Convert `diagram.vsdx` to `diagram.pdf`:

```powershell
vsdx2pdf .\diagram.vsdx
```

Convert to a custom output file:

```powershell
vsdx2pdf -o .\out\diagram.pdf .\diagram.vsdx
```

Inspect pages before converting:

```powershell
vsdx2pdf -list-pages .\diagram.vsdx
```

## Behavior

- If `-o` is omitted, the output path defaults to the input file name with a `.pdf` extension.
- Each Visio page becomes one PDF page.
- Background pages are rendered before the foreground page content.
- The command exits with a non-zero code on parse, render, or write errors.

## Building

Local build:

```powershell
go build -o dist/vsdx2pdf.exe ./cmd/vsdx2pdf
```

Cross-platform builds:

```powershell
$env:CGO_ENABLED = "0"
$env:GOOS = "linux"
$env:GOARCH = "amd64"
go build -o dist/vsdx2pdf-linux-amd64 ./cmd/vsdx2pdf
```
