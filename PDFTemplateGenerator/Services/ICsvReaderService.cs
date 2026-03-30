namespace PDFTemplateGenerator.Services;

public interface ICsvReaderService
{
    Task<CsvReadResult> ReadFromAppPackageAsync(string assetFileName, char separator = ',', CancellationToken cancellationToken = default);
    Task<List<List<string>>> ReadRowsFromFileAsync(string absolutePath, char separator = ',', CancellationToken cancellationToken = default);
}
