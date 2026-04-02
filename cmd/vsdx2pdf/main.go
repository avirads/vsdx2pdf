package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"

	"vsdx2pdf/internal/render"
	"vsdx2pdf/internal/vsdx"
)

var version = "dev"

func main() {
	var outputPath string
	var listPages bool
	var showVersion bool

	flag.StringVar(&outputPath, "o", "", "Output PDF path. Defaults to the input file name with a .pdf extension.")
	flag.BoolVar(&listPages, "list-pages", false, "List pages in the input document without rendering.")
	flag.BoolVar(&showVersion, "version", false, "Print the build version and exit.")
	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "Usage: %s [flags] <input.vsdx>\n\n", filepath.Base(os.Args[0]))
		flag.PrintDefaults()
	}
	flag.Parse()

	if showVersion {
		fmt.Println(version)
		return
	}

	if flag.NArg() != 1 {
		flag.Usage()
		os.Exit(2)
	}

	inputPath := flag.Arg(0)
	document, err := vsdx.Open(inputPath)
	if err != nil {
		exitf("open VSDX: %v", err)
	}

	if listPages {
		for _, page := range document.Pages {
			background := ""
			if page.IsBackground {
				background = " [background]"
			}
			fmt.Printf("%s (%s)%s\n", page.NameOrFallback(), page.ID, background)
		}
		return
	}

	if outputPath == "" {
		ext := filepath.Ext(inputPath)
		outputPath = inputPath[:len(inputPath)-len(ext)] + ".pdf"
	}

	pdfBytes, err := render.Convert(document)
	if err != nil {
		exitf("render PDF: %v", err)
	}

	outputDir := filepath.Dir(outputPath)
	if outputDir != "." {
		if err := os.MkdirAll(outputDir, 0o755); err != nil {
			exitf("create output directory: %v", err)
		}
	}
	if err := os.WriteFile(outputPath, pdfBytes, 0o644); err != nil {
		exitf("write PDF: %v", err)
	}

	fmt.Printf("Wrote %s\n", outputPath)
}

func exitf(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
