using System.Text.Json;
using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public sealed class GristClientService(HttpClient httpClient, GristApiOptions options) : IGristClientService
{
    public async Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesAsync(CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(options.BeneficiaryRecordsUrl))
        {
            throw new InvalidOperationException("Grist beneficiary records URL is not configured.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Get, options.BeneficiaryRecordsUrl);

        if (!string.IsNullOrWhiteSpace(options.ApiKey))
        {
            request.Headers.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", options.ApiKey);
        }

        using var response = await httpClient.SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync(cancellationToken);
        using var doc = JsonDocument.Parse(json);

        if (!doc.RootElement.TryGetProperty("records", out var records) || records.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<BeneficiaryClient>();
        }

        var clients = new List<BeneficiaryClient>();

        foreach (var record in records.EnumerateArray())
        {
            var recordId = record.TryGetProperty("id", out var idElement)
                ? idElement.ToString()
                : string.Empty;

            if (!record.TryGetProperty("fields", out var fields) || fields.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in fields.EnumerateObject())
            {
                dict[field.Name] = field.Value.ValueKind switch
                {
                    JsonValueKind.Null => string.Empty,
                    JsonValueKind.String => field.Value.GetString() ?? string.Empty,
                    _ => field.Value.ToString()
                };
            }

            dict.TryGetValue("lName", out var lName);
            dict.TryGetValue("fName", out var fName);
            dict.TryGetValue("mName", out var mName);

            clients.Add(new BeneficiaryClient
            {
                RecordId = recordId,
                LastName = lName ?? string.Empty,
                FirstName = fName ?? string.Empty,
                MiddleName = mName ?? string.Empty,
                Fields = dict
            });
        }

        return clients
            .OrderBy(c => c.LastName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(c => c.FirstName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(c => c.MiddleName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<BeneficiaryClient?> GetBeneficiaryByIdAsync(string clientId, CancellationToken cancellationToken = default)
    {
        var clients = await GetBeneficiariesAsync(cancellationToken);
        return clients.FirstOrDefault(c =>
            string.Equals(c.RecordId, clientId, StringComparison.OrdinalIgnoreCase));
    }
}
