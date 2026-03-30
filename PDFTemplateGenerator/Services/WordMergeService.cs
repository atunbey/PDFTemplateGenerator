using NPOI.XWPF.UserModel;

namespace PDFTemplateGenerator.Services;

public sealed class WordMergeService(
    ICsvReaderService csvReader,
    IWordTemplateRenderer templateRenderer) : IWordMergeService
{
    public async Task<string> FillDocxPlaceholdersFromCsvAsync(
        string templateAsset = "Template.docx",
        string csvAsset = "Data.csv",
        string outputFileName = "Output_Filled.docx")
    {
        var templateBytes = await LoadTemplateBytesAsync(templateAsset);

        var csvData = await csvReader.ReadFromAppPackageAsync(csvAsset);
        var header = csvData.Header;
        var rows = csvData.Rows;

        if (rows.Count == 0)
        {
            throw new InvalidOperationException("CSV has no data rows.");
        }

        var lastOutputPath = string.Empty;

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            using var ms = new MemoryStream(templateBytes);
            using var document = new XWPFDocument(ms);

            var row = rows[rowIndex];
            var outputName = BuildOutputName(row, outputFileName);
            var data = RowToDict(header, row);

            templateRenderer.ReplacePlaceholdersEverywhere(document, data);

            var outPath = Path.Combine(FileSystem.AppDataDirectory, outputName);
            await using var outFs = new FileStream(outPath, FileMode.Create, FileAccess.Write);
            document.Write(outFs);
            lastOutputPath = outPath;
        }

        return lastOutputPath;
    }

    public async Task<string> FillDocxTableFromCsvAsync(
        string templateAsset = "Template.docx",
        string csvAsset = "Data.csv",
        string outputFileName = "Output_Table.docx",
        bool matchTableByHeader = true)
    {
        await using var templateStream = await FileSystem.OpenAppPackageFileAsync(templateAsset);
        using var document = new XWPFDocument(templateStream);

        var csvData = await csvReader.ReadFromAppPackageAsync(csvAsset);
        var csvHeader = csvData.Header;
        var dataRows = csvData.Rows;

        if (csvHeader.Count == 0)
        {
            throw new InvalidOperationException("CSV has no header row.");
        }

        XWPFTable? table = matchTableByHeader
            ? templateRenderer.FindTableByHeader(document, csvHeader)
            : document.Tables.FirstOrDefault();

        if (table == null)
        {
            throw new InvalidOperationException("Target table was not found in the document.");
        }

        if (table.Rows.Count == 0)
        {
            throw new InvalidOperationException("Target table has no rows (need at least a header row).");
        }

        var tableHeaderNames = table.Rows[0].GetTableCells()
            .Select(c => (c.Paragraphs.FirstOrDefault()?.Text ?? string.Empty).Trim())
            .ToList();

        foreach (var csvRow in dataRows)
        {
            var dict = RowToDict(csvHeader, csvRow);
            var newRow = table.CreateRow();

            while (newRow.GetTableCells().Count < tableHeaderNames.Count)
            {
                newRow.AddNewTableCell();
            }

            for (int c = 0; c < tableHeaderNames.Count; c++)
            {
                var key = tableHeaderNames[c];
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                var normalizedKey = key.StartsWith("«") && key.EndsWith("»") && key.Length > 2
                    ? key[1..^1]
                    : key;

                dict.TryGetValue(normalizedKey, out var raw);
                var cell = newRow.GetCell(c);
                var paragraph = cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
                templateRenderer.ClearParagraph(paragraph);
                paragraph.CreateRun().SetText(raw ?? string.Empty);
            }
        }

        var outPath = Path.Combine(FileSystem.AppDataDirectory, outputFileName);
        await using var outFs = new FileStream(outPath, FileMode.Create, FileAccess.Write);
        document.Write(outFs);

        return outPath;
    }

    private static async Task<byte[]> LoadTemplateBytesAsync(string templateAsset)
    {
        await using var stream = await FileSystem.OpenAppPackageFileAsync(templateAsset);
        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms);
        return ms.ToArray();
    }

    private static string BuildOutputName(List<string> row, string fallback)
    {
        if (row.Count > 9)
        {
            return $"{row[1]}_{row[2]}_{row[4]}_{row[9]}.docx";
        }

        return fallback;
    }

    private static Dictionary<string, string> RowToDict(List<string> header, List<string> row)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < header.Count; i++)
        {
            var key = header[i];
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            dict[key] = i < row.Count ? row[i] : string.Empty;
        }

        return dict;
    }
}
