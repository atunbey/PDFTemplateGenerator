namespace PDFTemplateGenerator.Services;

public sealed class LegalDocumentOptions
{
    public string CertificateTemplateFolder { get; init; } = string.Empty;
    public string NextcloudFolderUrl { get; init; } = string.Empty;
    public string NextcloudWebDavFolderUrl { get; init; } = string.Empty;
    public string NextcloudUsername { get; init; } = string.Empty;
    public string NextcloudAppPassword { get; init; } = string.Empty;
    public bool EnableNextcloudTemplates { get; init; } = true;
    public bool EnableLocalTemplateFallback { get; init; } = false;
}
