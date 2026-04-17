using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public interface IGristClientService
{
    Task<CounselUser?> AuthenticateCounselAsync(string userName, string password, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<string>> GetAssociatedBeneficiaryIdsAsync(string counselClientId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesForCounselAsync(string counselClientId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesAsync(CancellationToken cancellationToken = default);
    Task<BeneficiaryClient?> GetBeneficiaryByIdAsync(string clientId, CancellationToken cancellationToken = default);
}
