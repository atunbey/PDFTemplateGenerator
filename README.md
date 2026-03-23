# PDFTemplateGenerator

A cross-platform **.NET MAUI Blazor** application for generating Word and Excel documents by merging template files with CSV data.

## Features

- Fill Word (`.docx`) templates: replace `{{Placeholder}}` tokens with values from the first CSV row
- Append table rows to Word templates from all CSV rows
- Fill Excel (`.xlsx`) templates from CSV data
- Append Excel table rows from all CSV rows
- Share or open generated files directly from the app
- Runs on Windows, Android, iOS, and macOS Catalyst

## Getting started

See the **[Documentation site](docs/)** for full setup and usage instructions, or read the [Getting Started guide](docs/content/docs/guides/getting-started.md) directly.

## Documentation (Doks)

The `docs/` directory contains a [Doks](https://getdoks.org/) documentation site powered by [Hugo](https://gohugo.io/).

### Prerequisites

- [Node.js](https://nodejs.org/) ≥ 24
- [Hugo Extended](https://gohugo.io/installation/) ≥ 0.123.7

### Install and run the docs site locally

```bash
cd docs
npm install
npm run dev
```

Then open <http://localhost:1313/> in your browser.

### Build the docs site for production

```bash
cd docs
npm install
npm run build
# Output is in docs/public/
```

## Development

1. Open `PDFTemplateGenerator.sln` in Visual Studio 2022 (with the .NET MAUI workload).
2. Restore NuGet packages.
3. Select your target platform and press **F5**.

## License

MIT