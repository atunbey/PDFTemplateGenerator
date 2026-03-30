namespace PDFTemplateGenerator.Models;

public sealed class DealerInventoryVehicle
{
    public string StockNumber { get; init; } = string.Empty;
    public string Field6 { get; init; } = string.Empty;
    public string Field7 { get; init; } = string.Empty;
    public string Field8 { get; init; } = string.Empty;
    public string Field9 { get; init; } = string.Empty;
    public string WebsiteStatus { get; init; } = string.Empty;

    public string Summary =>
        $"{WebsiteStatus} {StockNumber} {Field6} {Field7} {Field8} {Field9}".Trim();
}
