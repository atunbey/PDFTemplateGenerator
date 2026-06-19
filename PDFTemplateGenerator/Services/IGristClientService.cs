using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public interface IGristClientService
{
    Task<CounselUser?> AuthenticateCounselAsync(string userName, string password, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<string>> GetAssociatedBeneficiaryIdsAsync(string counselClientId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesForCounselAsync(string counselClientId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesAsync(CancellationToken cancellationToken = default);
    Task<BeneficiaryClient?> GetBeneficiaryByIdAsync(string clientId, CancellationToken cancellationToken = default);

    /// <summary>
    /// Builds a flat merge dictionary for a selected client and document.
    /// Reads Moor.Document.dataRequired JSON to fetch rows from related tables dynamically.
    /// </summary>
    Task<IReadOnlyDictionary<string, string>> GetMergedFieldsForDocumentAsync(
        string clientId,
        string documentName,
        CancellationToken cancellationToken = default);

    /// <summary>Returns all Trustor rows linked to <paramref name="beneficiaryId"/>, ordered by isPrimary descending then insertion order.</summary>
    Task<IReadOnlyList<IReadOnlyDictionary<string, string>>> GetTrustorsForBeneficiaryAsync(string beneficiaryId, CancellationToken cancellationToken = default);

    /// <summary>Returns all Trustee rows linked to <paramref name="beneficiaryId"/>, ordered by sortOrder then role.</summary>
    Task<IReadOnlyList<IReadOnlyDictionary<string, string>>> GetTrusteesForBeneficiaryAsync(string beneficiaryId, CancellationToken cancellationToken = default);

    /// <summary>Returns the DocumentExecution row for <paramref name="beneficiaryId"/> and <paramref name="templateName"/>, or null if none exists.</summary>
    Task<IReadOnlyDictionary<string, string>?> GetDocumentExecutionAsync(string beneficiaryId, string templateName, CancellationToken cancellationToken = default);
}
