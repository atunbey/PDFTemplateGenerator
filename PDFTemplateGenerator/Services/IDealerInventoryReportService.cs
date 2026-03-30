using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public interface IDealerInventoryReportService
{
    Task<DealerInventoryReportResult> BuildReportAsync();
}
