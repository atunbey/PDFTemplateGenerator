using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public interface IGristClientService
{
    Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesAsync(CancellationToken cancellationToken = default);
    Task<BeneficiaryClient?> GetBeneficiaryByIdAsync(string clientId, CancellationToken cancellationToken = default);
}
