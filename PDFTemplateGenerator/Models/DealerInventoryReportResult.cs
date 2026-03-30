namespace PDFTemplateGenerator.Models;

public sealed class DealerInventoryReportResult
{
    public List<DealerInventoryVehicle> Vehicles { get; init; } = new();
}
