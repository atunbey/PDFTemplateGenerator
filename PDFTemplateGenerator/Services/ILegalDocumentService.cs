namespace PDFTemplateGenerator.Services;

public interface ILegalDocumentService
{
    IEnumerable<string> GetAvailableTemplates();
    Task<string> GenerateCertificateOfTrustAsync(string clientId, string templateFileName, CancellationToken cancellationToken = default);
}
