using NPOI.XWPF.UserModel;

namespace PDFTemplateGenerator.Services;

public sealed class WordTemplateRenderer : IWordTemplateRenderer
{
    public void ReplacePlaceholdersEverywhere(XWPFDocument document, IReadOnlyDictionary<string, string> data)
    {
        foreach (var paragraph in document.Paragraphs)
        {
            ReplaceInParagraph(paragraph, data);
        }

        foreach (var table in document.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        ReplaceInParagraph(paragraph, data);
                    }
                }
            }
        }

        foreach (var headerParagraph in document.HeaderList.SelectMany(h => h.Paragraphs))
        {
            ReplaceInParagraph(headerParagraph, data);
        }
    }

    public XWPFTable? FindTableByHeader(XWPFDocument document, List<string> csvHeader)
    {
        foreach (var table in document.Tables)
        {
            if (table.Rows.Count == 0)
            {
                continue;
            }

            var firstRow = table.Rows[0];
            var headers = firstRow.GetTableCells()
                .Select(c => (c.Paragraphs.FirstOrDefault()?.Text ?? string.Empty).Trim())
                .ToList();

            if (HeadersEqual(headers, csvHeader))
            {
                return table;
            }
        }

        return null;
    }

    public void ClearParagraph(XWPFParagraph paragraph)
    {
        for (int i = paragraph.Runs.Count - 1; i >= 0; i--)
        {
            paragraph.RemoveRun(i);
        }
    }

    private void ReplaceInParagraph(XWPFParagraph paragraph, IReadOnlyDictionary<string, string> data)
    {
        var original = paragraph.Text ?? string.Empty;
        var replaced = ReplacePlaceholders(original, data);

        if (replaced == original)
        {
            return;
        }

        ClearParagraph(paragraph);

        var run = paragraph.CreateRun();
        run.SetText(replaced);
    }

    private static string ReplacePlaceholders(string text, IReadOnlyDictionary<string, string> data)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        foreach (var kv in data)
        {
            text = text.Replace($"«{kv.Key}»", kv.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            text = text.Replace($"${{{kv.Key}}}", kv.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        return text;
    }

    private static bool HeadersEqual(IList<string> documentHeaders, IList<string> csvHeader)
    {
        var normalized = documentHeaders
            .Where(h => !string.IsNullOrWhiteSpace(h) && h.Length >= 2)
            .Select(h => h.StartsWith("«") && h.EndsWith("»") ? h[1..^1] : h)
            .ToList();

        return normalized.Any(h => csvHeader.Contains(h, StringComparer.OrdinalIgnoreCase));
    }
}
