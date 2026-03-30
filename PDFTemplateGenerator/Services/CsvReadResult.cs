namespace PDFTemplateGenerator.Services;

public sealed class CsvReadResult
{
    public List<string> Header { get; init; } = new();
    public List<List<string>> Rows { get; init; } = new();
}
