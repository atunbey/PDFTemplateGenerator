using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public sealed class DealerInventoryReportService(
    ICsvReaderService csvReader,
    DealerInventoryReportOptions options) : IDealerInventoryReportService
{
    public async Task<DealerInventoryReportResult> BuildReportAsync()
    {
        if (string.IsNullOrWhiteSpace(options.WorkingDirectory))
        {
            throw new InvalidOperationException("Dealer inventory working directory is not configured.");
        }

        var completeInventoryPath = Path.Combine(options.WorkingDirectory, options.CompleteInventoryRelativePath);
        var websiteInventoryPath = Path.Combine(options.WorkingDirectory, options.WebsiteInventoryRelativePath);

        var completeRows = await csvReader.ReadRowsFromFileAsync(completeInventoryPath);
        if (completeRows.Count == 0)
        {
            return new DealerInventoryReportResult();
        }

        var websiteRows = await csvReader.ReadRowsFromFileAsync(websiteInventoryPath);

        var completeParsed = ReadHeaderRows(completeRows, isHeader: false);
        var websiteParsed = ReadHeaderRows(websiteRows, isHeader: true);

        var websiteStockSet = new HashSet<string>(
            websiteParsed.rows
                .Select(r => GetValue(r, 9))
                .Where(v => !string.IsNullOrWhiteSpace(v)),
            StringComparer.OrdinalIgnoreCase);

        var vehicles = completeParsed.rows
            .Select(row =>
            {
                var stockNumber = GetValue(row, 2);
                return new DealerInventoryVehicle
                {
                    StockNumber = stockNumber,
                    Field6 = GetValue(row, 6),
                    Field7 = GetValue(row, 7),
                    Field8 = GetValue(row, 8),
                    Field9 = GetValue(row, 9),
                    WebsiteStatus = websiteStockSet.Contains(stockNumber) ? "Online" : "Needs Work"
                };
            })
            .ToList();

        return new DealerInventoryReportResult
        {
            Vehicles = vehicles
        };
    }

    private static (List<string> header, List<List<string>> rows) ReadHeaderRows(List<List<string>> lines, bool isHeader)
    {
        var header = isHeader && lines.Count > 0
            ? lines[0].Select(h => (h ?? string.Empty).Trim()).ToList()
            : new List<string>();

        var rows = lines
            .Skip(isHeader ? 1 : 0)
            .Where(r => r.Any(v => !string.IsNullOrWhiteSpace(v)))
            .ToList();

        return (header, rows);
    }

    private static string GetValue(List<string> row, int index)
    {
        return index >= 0 && index < row.Count
            ? row[index] ?? string.Empty
            : string.Empty;
    }
}
