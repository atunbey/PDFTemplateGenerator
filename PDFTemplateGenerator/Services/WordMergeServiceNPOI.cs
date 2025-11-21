using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFTemplateGenerator.Services
{
    internal class WordMergeServiceNPOI
    {

        /// <summary>
        /// Replaces ${Placeholders} in the DOCX template using the FIRST data row in the CSV.
        /// Saves to AppDataDirectory and returns output file path.
        /// </summary>
        public async Task<string> FillDocxPlaceholdersFromCsvAsync(
            string templateAsset = "Template.docx",
            string csvAsset = "Data.csv",
            string outputFileName = "Output_Filled.docx")
        {
            // Load template.docx from Resources/Raw
            //added to load template into byte[] for reuse in loop
            byte[] templateBytes = File.ReadAllBytes(templateAsset);
            using var templateStream = await FileSystem.OpenAppPackageFileAsync(templateAsset);
            //using var doc = new XWPFDocument(templateStream);
            //changed to be able to reload doc. need to load template into memory using the below will hit the disk every creation
            
            //XWPFDocument doc;

            // Read CSV → first data row
            var (header, rows) = await ReadCsvAsync(csvAsset);
            if (rows.Count == 0)
                throw new InvalidOperationException("CSV has no data rows.");

            //var dict = RowToDict(header, rows[0]);

            int rowProc = 0;
            var filePath = "";
            while (rowProc < rows.Count)
            {
                using (var ms = new MemoryStream(templateBytes))
                {
                    //doc = new XWPFDocument(templateStream);
                    var doc = new XWPFDocument(ms);
                    outputFileName = rows[rowProc][1] + "_" + rows[rowProc][2] + "_" + rows[rowProc][4] + "_" + rows[rowProc][9] + ".docx";
                    var dict = RowToDict(header, rows[rowProc]);

                    var optionHeader = new string[40];
                    var optionRow = new string[40];
                    for (int i = 0; i <= optionHeader.Length - 1; i++)
                    {
                        optionHeader[i] = "Options" + (i + 1).ToString();
                        optionRow[i] = rows[0][10].Split("|").Length > i ? rows[0][10].Split("|")[i] : "";
                    }
                    var optionDict = RowToDict(optionHeader.ToList(), optionRow.ToList());

                    // Replace in all paragraphs (body)
                    foreach (var p in doc.Paragraphs)
                        ReplaceInParagraph(p, dict);

                    // Replace inside all tables (cells have their own paragraphs)
                    foreach (var table in doc.Tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {
                                foreach (var p in cell.Paragraphs)
                                    ReplaceInParagraph(p, table.Rows.Count == 20 ? optionDict : dict);
                            }
                        }
                    }

                    // Save
                    var outPath = Path.Combine(FileSystem.AppDataDirectory, outputFileName);
                    using (var outFs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                        doc.Write(outFs);
                    filePath = outPath;
                    rowProc++;
                }
            }

            return filePath;
        }

        /// <summary>
        /// Appends data rows from CSV into a table in the template.
        /// If tableHeaderMatch is provided, finds the first table whose FIRST ROW text matches CSV headers.
        /// Otherwise, uses the first table.
        /// </summary>
        public async Task<string> FillDocxTableFromCsvAsync(
            string templateAsset = "Template.docx",
            string csvAsset = "Data.csv",
            string outputFileName = "Output_Table.docx",
            bool matchTableByHeader = true)
        {
            using var templateStream = await FileSystem.OpenAppPackageFileAsync(templateAsset);
            using var doc = new XWPFDocument(templateStream);

            var (csvHeader, dataRows) = await ReadCsvAsync(csvAsset);
            if (csvHeader.Count == 0)
                throw new InvalidOperationException("CSV has no header row.");

            // Find target table
            XWPFTable? table = null;

            if (matchTableByHeader)
            {
                table = FindTableByHeader(doc, csvHeader);
                if (table == null)
                    throw new InvalidOperationException("No table found whose first row matches the CSV header.");
            }
            else
            {
                table = doc.Tables.FirstOrDefault()
                    ?? throw new InvalidOperationException("No tables found in the document.");
            }

            if (table.Rows.Count == 0)
                throw new InvalidOperationException("Target table has no rows (need at least a header row).");

            // Build a map from header name → column index (from table's first row)
            var headerRow = table.Rows[0];
            var headerCells = headerRow.GetTableCells();
            var tableHeaderNames = headerCells
                .Select(c => (c.Paragraphs.FirstOrDefault()?.Text ?? "").Trim())
                .ToList();

            // Append one row per CSV record
            foreach (var csvRow in dataRows)
            {
                var dict = RowToDict(csvHeader, csvRow);

                // Create a new row at the end (copies basic table structure, not styles)
                var newRow = table.CreateRow();

                // Ensure cell count matches header count
                while (newRow.GetTableCells().Count < tableHeaderNames.Count)
                    newRow.AddNewTableCell();

                for (int c = 0; c < tableHeaderNames.Count; c++)
                {
                    var key = tableHeaderNames[c];
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    dict.TryGetValue(key.Substring(1,key.Length - 2), out var raw);
                    var text = raw ?? "";

                    var cell = newRow.GetCell(c);
                    // Clear existing paragraphs/content
                    // (XWPF creates a paragraph by default; re-use it)
                    var p = cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
                    ClearParagraph(p);
                    var r = p.CreateRun();
                    r.SetText(text);
                }
            }

            // Save
            var outPath = Path.Combine(FileSystem.AppDataDirectory, outputFileName);
            using (var outFs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                doc.Write(outFs);

            return outPath;
        }

        // ----------------- Internals -----------------

        // Replaces placeholders in a paragraph by rebuilding runs.
        private static void ReplaceInParagraph(XWPFParagraph p, Dictionary<string, string> data)
        {
            if (p == null) return;

            // Concatenate current text (across runs)
            var original = p.Text ?? string.Empty;
            var replaced = ReplacePlaceholders(original, data);

            if (replaced == original) return;

            // Clear existing runs
            for (int i = p.Runs.Count - 1; i >= 0; i--)
                p.RemoveRun(i);

            // Create a single run with the replaced text
            var run = p.CreateRun();
            run.SetText(replaced);
            // NOTE: We don't preserve per-segment formatting. If placeholders are plain text,
            // this is typically fine. To preserve formatting, keep placeholders within a single run.
            run.IsBold = true;

            switch (original)
            {
                case "«Price»":
                    run.FontSize = 28;
                    break;
            }
        }

        private static void ClearParagraph(XWPFParagraph p)
        {
            for (int i = p.Runs.Count - 1; i >= 0; i--)
                p.RemoveRun(i);
        }

        private static string ReplacePlaceholders(string text, Dictionary<string, string> data)
        {
            if (string.IsNullOrEmpty(text)) return text;
            foreach (var kv in data)
                text = text.Replace("«" + kv.Key + "»", kv.Value ?? "");
            return text;
        }

        private static XWPFTable? FindTableByHeader(XWPFDocument doc, List<string> csvHeader)
        {
            foreach (var t in doc.Tables)
            {
                if (t.Rows.Count == 0) continue;
                var firstRow = t.Rows[0];
                var headers = firstRow.GetTableCells()
                                      .Select(c => (c.Paragraphs.FirstOrDefault()?.Text ?? "").Trim())
                                      .ToList();

                if (HeadersEqual(headers, csvHeader))
                    return t;
            }
            return null;
        }

        private static bool HeadersEqual(IList<string> a, IList<string> b)
        {
            //if (a.Count != b.Count) return false;
            //for (int i = 0; i < a.Count; i++)
            //    if (!string.Equals(a[i], b[i], StringComparison.OrdinalIgnoreCase))
            //        return false;
            //    return true;
            a = a.Select(x => x.Substring(1, x.Length - 2)).ToList();
            if (a.Any(aa => b.Contains(aa)))
                return true;
            return false;
        }

        private static async Task<(List<string> header, List<List<string>> rows)> ReadCsvAsync(
            string csvAssetFileName, char sep = ',')
        {
            using var s = await FileSystem.OpenAppPackageFileAsync(csvAssetFileName);
            using var reader = new StreamReader(s);

            var lines = new List<string>();
            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync();
                if (line != null) lines.Add(line);
            }

            var parsed = lines.Select(l => ParseCsvLine(l, sep)).ToList();
            if (parsed.Count == 0) return (new List<string>(), new List<List<string>>());

            var header = parsed[0].Select(h => (h ?? "").Trim()).ToList();
            var rows = parsed
                .Skip(1)
                .Where(r => r.Any(v => !string.IsNullOrWhiteSpace(v)))
                .ToList();
            return (header, rows);
        }

        private static List<string> ParseCsvLine(string line, char sep)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var cur = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; }
                        else { inQuotes = false; }
                    }
                    else cur.Append(ch);
                }
                else
                {
                    if (ch == '"') inQuotes = true;
                    else if (ch == sep) { result.Add(cur.ToString()); cur.Clear(); }
                    else cur.Append(ch);
                }
            }
            result.Add(cur.ToString());
            return result;
        }

        private static Dictionary<string, string> RowToDict(List<string> header, List<string> row)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < header.Count; i++)
            {
                var key = header[i];
                if (string.IsNullOrWhiteSpace(key)) continue;
                dict[key] = i < row.Count ? row[i] : "";
            }
            return dict;
        }

    }
}
