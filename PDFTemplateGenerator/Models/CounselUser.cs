namespace PDFTemplateGenerator.Models;

public sealed class CounselUser
{
    public string RecordId { get; init; } = string.Empty;
    public string UserName { get; init; } = string.Empty;
    public string ClientId { get; init; } = string.Empty;
    public bool IsActive { get; init; }
    public string FirstName { get; init; } = string.Empty;
    public string MiddleName { get; init; } = string.Empty;
    public string LastName { get; init; } = string.Empty;
    public string ZipCode { get; init; } = string.Empty;
}
