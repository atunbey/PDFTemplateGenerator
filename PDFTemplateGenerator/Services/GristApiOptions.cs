namespace PDFTemplateGenerator.Services;

public sealed class GristApiOptions
{
    public string BeneficiaryRecordsUrl { get; init; } = string.Empty;
    public string CounselRecordsUrl { get; init; } = string.Empty;
    public string AssociationsRecordsUrl { get; init; } = string.Empty;
    public string MoorDocumentRecordsUrl { get; init; } = string.Empty;
    public string ApiKey { get; init; } = string.Empty;

    // Relational trust tables — column names map directly to document merge tags.
    // Leave empty to disable lookup (document generation still works from Beneficiary data alone).
    public string TrustorRecordsUrl { get; init; } = string.Empty;
    public string TrusteeRecordsUrl { get; init; } = string.Empty;
    public string DocumentExecutionRecordsUrl { get; init; } = string.Empty;
}
