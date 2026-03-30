using System.Collections.ObjectModel;

namespace PDFTemplateGenerator.Models;

public sealed class BeneficiaryClient
{
    public string RecordId { get; init; } = string.Empty;
    public string LastName { get; init; } = string.Empty;
    public string FirstName { get; init; } = string.Empty;
    public string MiddleName { get; init; } = string.Empty;
    public IReadOnlyDictionary<string, string> Fields { get; init; } =
        new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));

    public string DisplayName
    {
        get
        {
            var firstMiddle = string.Join(" ", new[] { FirstName, MiddleName }
                .Where(s => !string.IsNullOrWhiteSpace(s)));

            if (string.IsNullOrWhiteSpace(LastName) && string.IsNullOrWhiteSpace(firstMiddle))
            {
                return string.IsNullOrWhiteSpace(RecordId) ? "Client" : $"Client #{RecordId}";
            }

            if (string.IsNullOrWhiteSpace(firstMiddle))
            {
                return LastName;
            }

            if (string.IsNullOrWhiteSpace(LastName))
            {
                return firstMiddle;
            }

            return $"{LastName}, {firstMiddle}";
        }
    }
}
