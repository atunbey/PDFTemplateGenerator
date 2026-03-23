---
title: "Getting Started"
description: "Learn how to clone, build, and run PDFTemplateGenerator on your machine, add your own Word and Excel templates, and produce merged documents from CSV data."
summary: "Install prerequisites, clone the repository, and launch the app in minutes."
date: 2024-01-01T00:00:00+00:00
lastmod: 2024-01-01T00:00:00+00:00
draft: false
weight: 810
toc: true
params:
  seo:
    title: ""
    description: ""
    canonical: ""
    robots: ""
---

## Prerequisites

Before you begin, make sure the following tools are installed on your machine:

- [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) (or later)
- [Visual Studio 2022](https://visualstudio.microsoft.com/) with the **.NET MAUI** workload enabled, **or**
  [Visual Studio Code](https://code.visualstudio.com/) with the [C# Dev Kit](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.csdevkit) extension

## Clone the repository

```bash
git clone https://github.com/atunbey/PDFTemplateGenerator.git
cd PDFTemplateGenerator
```

## Open in Visual Studio

1. Open `PDFTemplateGenerator.sln` in Visual Studio 2022.
2. Restore NuGet packages (Visual Studio does this automatically on first open).
3. Select your target platform from the toolbar (e.g. **Windows Machine** or an Android emulator).
4. Press **F5** to build and run the application.

## Open in Visual Studio Code

```bash
cd PDFTemplateGenerator
dotnet restore
dotnet build
```

To run on Windows:

```bash
dotnet run --framework net8.0-windows10.0.19041.0
```

## Add your own templates

Place your Word (`.docx`) or Excel (`.xlsx`) templates and CSV data files inside:

```
PDFTemplateGenerator/Resources/Raw/
```

The app will detect files in this folder and make them available through the UI for template merging.

## Further reading

- [Template Reference](/docs/reference/templates/) — learn about placeholder syntax and CSV column mapping
- [NPOI Word merge service](/docs/reference/word-merge/) — API reference for `WordMergeServiceNPOI`
