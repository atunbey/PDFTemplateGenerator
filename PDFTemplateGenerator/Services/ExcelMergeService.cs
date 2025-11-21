using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFTemplateGenerator.Services
{
    internal class ExcelMergeService
    {

        /// <summary>
        /// Fills ${Placeholders} in the template using the FIRST data row in the CSV.
        /// Saves to AppDataDirectory and returns the full output file path.
        /// </summary>
        public async Task<string> FillTemplateFromCsvAsync(
            string templateAssetFileName = "Template.xlsx",
            string csvAssetFileName = "Data.csv",
            string outputFileName = "Output_Filled.xlsx")
        {
            // 1) Load template from Resources/Raw as a stream
            using var templateStream = await FileSystem.OpenAppPackageFileAsync(templateAssetFileName);
            IWorkbook wb = new XSSFWorkbook(templateStream);

            // 2) Read CSV: first row = header, next row = data
            var (header, rows) = await ReadCsvAssetAsync(csvAssetFileName);
            if (rows.Count == 0)
                throw new InvalidOperationException("CSV has no data rows.");

            var data = RowToDict(header, rows[0]);

            // 3) Replace ${Key} placeholders across all sheets
            for (int s = 0; s < wb.NumberOfSheets; s++)
            {
                var sheet = wb.GetSheetAt(s);
                foreach (IRow row in sheet)
                {
                    if (row == null) continue;
                    foreach (var cell in row.Cells)
                    {
                        if (cell == null || cell.CellType != CellType.String) continue;

                        var text = cell.StringCellValue;
                        var replaced = ReplacePlaceholders(text, data);
                        if (!ReferenceEquals(replaced, text))
                            cell.SetCellValue(replaced);
                    }
                }
                sheet.ForceFormulaRecalculation = true;
            }

            // 4) Save output
            var outPath = Path.Combine(FileSystem.AppDataDirectory, outputFileName);
            using (var outStream = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                wb.Write(outStream);

            wb.Close();
            return outPath;
        }


        /// <summary>
        /// Appends CSV rows under the header row in a template sheet.
        /// Header names in the template's first row must match CSV headers.
        /// Saves to AppDataDirectory and returns the full output file path.
        /// </summary>
        public async Task<string> AppendTableFromCsvAsync(
            string templateAssetFileName = "Template.xlsx",
            string csvAssetFileName = "Data.csv",
            string outputFileName = "Output_Table.xlsx",
            string? sheetName = "Report",
            int headerRowIdx = 0)
        {
            // 1) Load template
            using var templateStream = await FileSystem.OpenAppPackageFileAsync(templateAssetFileName);
            IWorkbook wb = new XSSFWorkbook(templateStream);

            // 2) Resolve sheet & model row (next row after header)
            var sheet = sheetName != null ? wb.GetSheet(sheetName) : wb.GetSheetAt(0);
            if (sheet == null)
                throw new InvalidOperationException($"Sheet '{sheetName}' not found in template.");

            int firstDataRowIdx = headerRowIdx + 1;
            var modelRow = sheet.GetRow(firstDataRowIdx) ?? sheet.CreateRow(firstDataRowIdx);

            // 3) Read CSV
            var (csvHeader, dataRows) = await ReadCsvAssetAsync(csvAssetFileName);
            if (csvHeader.Count == 0)
                throw new InvalidOperationException("CSV has no header row.");

            // 4) Map by template header names
            var templateHeaderRow = sheet.GetRow(headerRowIdx)
                ?? throw new InvalidOperationException($"Template header row {headerRowIdx} not found.");
            int templateColCount = templateHeaderRow.LastCellNum;

            string[] templateCols = new string[templateColCount];
            for (int c = 0; c < templateColCount; c++)
                templateCols[c] = (templateHeaderRow.GetCell(c)?.ToString() ?? "").Trim();

            // 5) Write rows
            int writeRowIdx = firstDataRowIdx;
            foreach (var csvRow in dataRows)
            {
                var dict = RowToDict(csvHeader, csvRow);

                var row = sheet.GetRow(writeRowIdx) ?? sheet.CreateRow(writeRowIdx);
                CopyRowStyle(modelRow, row);

                for (int c = 0; c < templateCols.Length; c++)
                {
                    var key = templateCols[c];
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    dict.TryGetValue(key, out var raw);
                    WriteSmart(row, c, raw ?? "");
                }
                writeRowIdx++;
            }

            sheet.ForceFormulaRecalculation = true;

            // 6) Save
            var outPath = Path.Combine(FileSystem.AppDataDirectory, outputFileName);
            using (var outStream = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                wb.Write(outStream);

            wb.Close();
            return outPath;
        }


        // ----------------- Helpers -----------------

        private static async Task<(List<string> header, List<List<string>> rows)> ReadCsvAssetAsync(
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
            var rows = parsed.Skip(1)
                             .Where(r => r.Any(v => !string.IsNullOrWhiteSpace(v)))
                             .ToList();

            return (header, rows);
        }

        private static List<string> ParseCsvLine(string line, char sep)
        {
            // Handles quoted fields, commas within quotes, and escaped quotes ("")
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

        private static void CopyRowStyle(IRow from, IRow to)
        {
            to.Height = from.Height;
            int max = Math.Max(from.LastCellNum, (short) 1);
            for (int c = 0; c < max; c++)
            {
                var src = from.GetCell(c);
                var dst = to.GetCell(c) ?? to.CreateCell(c);
                if (src == null) continue;

                ICellStyle newStyle = to.Sheet.Workbook.CreateCellStyle();
                newStyle.CloneStyleFrom(src.CellStyle);
                dst.CellStyle = newStyle;
            }
        }

        private static void WriteSmart(IRow row, int col, string raw)
        {
            var cell = row.GetCell(col) ?? row.CreateCell(col);

            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
            {
                cell.SetCellValue(num);
                return;
            }
            if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
            {
                cell.SetCellValue(dt);
                return;
            }
            if (bool.TryParse(raw, out var b))
            {
                cell.SetCellValue(b);
                return;
            }
            cell.SetCellValue(raw ?? "");
        }

        private static string ReplacePlaceholders(string text, Dictionary<string, string> data)
        {
            foreach (var kv in data)
                text = text.Replace("${" + kv.Key + "}", kv.Value ?? "");
            return text;
        }

    }
}
