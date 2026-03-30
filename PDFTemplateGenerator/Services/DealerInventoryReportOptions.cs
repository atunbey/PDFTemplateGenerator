namespace PDFTemplateGenerator.Services;

public sealed class DealerInventoryReportOptions
{
    public string WorkingDirectory { get; init; } = string.Empty;
    public string CompleteInventoryRelativePath { get; init; } = "CSVinventory\\comsoftInventoryCSI2JTZ.CSV";
    public string WebsiteInventoryRelativePath { get; init; } = "CSVinventory\\WebsiteInventory.csv";
}
