---
title: "Template Reference"
description: "Reference for Word and Excel placeholder syntax, CSV column mapping, fill-first-row mode, append-table mode, and output file handling in PDFTemplateGenerator."
summary: "Understand how placeholders are defined in templates and how CSV columns map to them."
date: 2024-01-01T00:00:00+00:00
lastmod: 2024-01-01T00:00:00+00:00
draft: false
weight: 910
toc: true
params:
  seo:
    title: ""
    description: ""
    canonical: ""
    robots: ""
---

## Placeholder syntax

Placeholders in Word (`.docx`) and Excel (`.xlsx`) templates use double curly-brace notation:

```
{{ColumnName}}
```

The placeholder name must **exactly** match the CSV column header (case-sensitive).

### Example

Given a CSV file with the following header row:

| Make  | Model  | Year | VIN               |
|-------|--------|------|-------------------|
| Honda | Accord | 2023 | 1HGBH41JXMN109186 |

Your Word template should contain placeholders like:

```
Vehicle: {{Make}} {{Model}} ({{Year}})
VIN: {{VIN}}
```

After merging, the first CSV row produces:

```
Vehicle: Honda Accord (2023)
VIN: 1HGBH41JXMN109186
```

## Fill placeholders (first row only)

The **Fill placeholders** action reads the **first data row** from the CSV and replaces every matching placeholder in the template with the corresponding value.

Use this mode when you need to produce a single, personalised document.

## Append table (all rows)

The **Append table** action iterates over **every CSV row** and appends a new row to an existing table inside the template document.

The template must contain a table whose header cells match the CSV column names. PDFTemplateGenerator locates the table by matching the header row.

## Word merge service

The `WordMergeServiceNPOI` class (in `Services/WordMergeServiceNPOI.cs`) exposes two public async methods:

| Method | Description |
|--------|-------------|
| `FillDocxPlaceholdersFromCsvAsync(templateAsset, csvAsset, outputFileName)` | Replaces placeholders in the template with values from the first CSV row |
| `FillDocxTableFromCsvAsync(templateAsset, csvAsset, outputFileName, matchTableByHeader)` | Appends one row per CSV record to a matching table in the template |

## Excel merge service

The `ExcelMergeService` class (in `Services/ExcelMergeService.cs`) exposes:

| Method | Description |
|--------|-------------|
| `FillTemplateFromCsvAsync(templateAssetFileName, csvAssetFileName, outputFileName)` | Fills named-range or placeholder cells in the template with the first CSV row |
| `AppendTableFromCsvAsync(templateAssetFileName, csvAssetFileName, outputFileName, sheetName, headerRowIdx)` | Appends all CSV rows as data rows beneath the template's header row |

## Output files

Generated documents are saved to the app's local data directory. After a successful merge the UI displays the output path and offers **Open** and **Share** buttons.
