namespace PDFTemplateGenerator.Services;

public sealed class GristApiOptions
{
    public string BeneficiaryRecordsUrl { get; init; } = string.Empty;
    public string CounselRecordsUrl { get; init; } = string.Empty;
    public string AssociationsRecordsUrl { get; init; } = string.Empty;
    public string ApiKey { get; init; } = string.Empty;
}
