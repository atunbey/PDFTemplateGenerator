namespace PDFTemplateGenerator.Services;

public interface ILegalDocumentService
{
    Task<IReadOnlyList<string>> GetAvailableTemplatesAsync(CancellationToken cancellationToken = default);
    Task<string> GenerateCertificateOfTrustAsync(string clientId, string templateFileName, CancellationToken cancellationToken = default);
}
