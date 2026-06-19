using System.Text.Json;
using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public sealed class GristClientService(HttpClient httpClient, GristApiOptions options) : IGristClientService
{
    public async Task<CounselUser?> AuthenticateCounselAsync(string userName, string password, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(password))
        {
            return null;
        }

        var records = await GetRecordsAsync(options.CounselRecordsUrl, cancellationToken);
        foreach (var record in records)
        {
            var recordUserName = GetFieldAsString(record.Fields, "userName", "username", "email", "login");
            var recordPassword = GetFieldAsString(record.Fields, "passWord", "password", "pass");
            var isActive = GetFieldAsBoolean(record.Fields, "isactive", "isActive", "active");

            if (!isActive)
                continue;

            if (!string.Equals(recordUserName, userName, StringComparison.OrdinalIgnoreCase))
                continue;

            if (!string.Equals(recordPassword, password, StringComparison.Ordinal))
                continue;

            var rawClientId = GetFieldAsString(record.Fields, "clientId", "clientid", "counsel", "counselClientId").Trim();
            var resolvedClientId = string.IsNullOrWhiteSpace(rawClientId) ? record.Id : rawClientId;
            var profile = await GetBeneficiaryProfileAsync(resolvedClientId, cancellationToken);

            return new CounselUser
            {
                RecordId = record.Id,
                UserName = recordUserName,
                ClientId = resolvedClientId,
                IsActive = true,
                FirstName = profile.FirstName,
                MiddleName = profile.MiddleName,
                LastName = profile.LastName,
                ZipCode = profile.ZipCode
            };
        }

        return null;
    }

    private async Task<(string FirstName, string MiddleName, string LastName, string ZipCode)> GetBeneficiaryProfileAsync(
        string clientId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(clientId))
            return (string.Empty, string.Empty, string.Empty, string.Empty);

        try
        {
            var beneficiary = await GetBeneficiaryByIdAsync(clientId, cancellationToken);
            if (beneficiary is null)
                return (string.Empty, string.Empty, string.Empty, string.Empty);

            beneficiary.Fields.TryGetValue("zip", out var zip1);
            beneficiary.Fields.TryGetValue("zipCode", out var zip2);
            beneficiary.Fields.TryGetValue("postalCode", out var zip3);
            var zip = zip1 ?? zip2 ?? zip3 ?? string.Empty;

            return (beneficiary.FirstName, beneficiary.MiddleName, beneficiary.LastName, zip);
        }
        catch
        {
            return (string.Empty, string.Empty, string.Empty, string.Empty);
        }
    }

    public async Task<IReadOnlyList<string>> GetAssociatedBeneficiaryIdsAsync(string counselClientId, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(counselClientId))
        {
            return Array.Empty<string>();
        }

        var records = await GetRecordsAsync(options.AssociationsRecordsUrl, cancellationToken);
        var clientIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var record in records)
        {
            var associationCounselId = GetFieldAsString(record.Fields, "counsel", "counselId", "counselClientId").Trim();
            if (!string.Equals(associationCounselId, counselClientId, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var clientId = GetFieldAsString(record.Fields, "clientId", "clientid", "client", "beneficiary", "beneficiaryId").Trim();
            if (string.IsNullOrWhiteSpace(clientId))
            {
                continue;
            }

            clientIds.Add(clientId);
        }

        return clientIds.OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public async Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesForCounselAsync(string counselClientId, CancellationToken cancellationToken = default)
    {
        var clientIds = await GetAssociatedBeneficiaryIdsAsync(counselClientId, cancellationToken);
        if (clientIds.Count == 0)
        {
            return Array.Empty<BeneficiaryClient>();
        }

        var allowed = new HashSet<string>(clientIds, StringComparer.OrdinalIgnoreCase);
        var allClients = await GetBeneficiariesAsync(cancellationToken);

        return allClients
            .Where(client => allowed.Contains(client.RecordId))
            .OrderBy(c => c.LastName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(c => c.FirstName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(c => c.MiddleName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<IReadOnlyList<BeneficiaryClient>> GetBeneficiariesAsync(CancellationToken cancellationToken = default)
    {
        var records = await GetRecordsAsync(options.BeneficiaryRecordsUrl, cancellationToken);
        var clients = new List<BeneficiaryClient>();

        foreach (var record in records)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in record.Fields)
            {
                dict[field.Key] = ConvertJsonToString(field.Value);
            }

            dict.TryGetValue("lName", out var lName);
            dict.TryGetValue("fName", out var fName);
            dict.TryGetValue("mName", out var mName);

            clients.Add(new BeneficiaryClient
            {
                RecordId = record.Id,
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

    public async Task<IReadOnlyDictionary<string, string>> GetMergedFieldsForDocumentAsync(
        string clientId,
        string documentName,
        CancellationToken cancellationToken = default)
    {
        var client = await GetBeneficiaryByIdAsync(clientId, cancellationToken)
            ?? throw new InvalidOperationException($"Client with id {clientId} was not found in Grist.");

        // Beneficiary fields are always included as base merge data.
        var merged = new Dictionary<string, string>(client.Fields, StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(options.MoorDocumentRecordsUrl))
            return merged;

        var documentRecord = await GetMoorDocumentRecordAsync(clientId, documentName, cancellationToken);
        if (documentRecord is null)
            return merged;

        // Moor.Document columns can contribute direct merge tags too.
        foreach (var field in documentRecord.Fields)
            merged[field.Key] = ConvertJsonToString(field.Value);

        var dataRequiredRaw = GetFieldAsString(documentRecord.Fields, "dataRequired", "datarequired", "data_required");
        if (string.IsNullOrWhiteSpace(dataRequiredRaw))
            return merged;

        var successorTrusteeIndex = 1;

        foreach (var reference in ParseDataRequiredReferences(dataRequiredRaw))
        {
            var related = await GetRelatedRowAsFlatDictionaryAsync(reference.Table, reference.RecordId, clientId, cancellationToken);
            if (related is null)
                continue;

            if (IsTrusteeTable(reference.Table) && related.TryGetValue("trusteeFullName", out var trusteeName))
            {
                if (!merged.ContainsKey("trusteeFullName") || string.IsNullOrWhiteSpace(merged["trusteeFullName"]))
                {
                    merged["trusteeFullName"] = trusteeName;
                }
                else
                {
                    merged[$"successorTrustee{successorTrusteeIndex}"] = trusteeName;
                    successorTrusteeIndex++;
                }
            }

            foreach (var kv in related)
            {
                if (IsTrusteeTable(reference.Table) && string.Equals(kv.Key, "trusteeFullName", StringComparison.OrdinalIgnoreCase))
                    continue;

                merged[kv.Key] = kv.Value;
            }
        }

        return merged;
    }

    public async Task<IReadOnlyList<IReadOnlyDictionary<string, string>>> GetTrustorsForBeneficiaryAsync(
        string beneficiaryId, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(options.TrustorRecordsUrl) || string.IsNullOrWhiteSpace(beneficiaryId))
            return Array.Empty<IReadOnlyDictionary<string, string>>();

        var records = await GetRecordsAsync(options.TrustorRecordsUrl, cancellationToken);
        var result = new List<Dictionary<string, string>>();

        foreach (var record in records)
        {
            var beneficiaryField = GetFieldAsString(record.Fields, "beneficiaryId", "beneficiary", "clientId", "client");
            if (!string.Equals(beneficiaryField, beneficiaryId, StringComparison.OrdinalIgnoreCase))
                continue;

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in record.Fields)
                dict[field.Key] = ConvertJsonToString(field.Value);
            result.Add(dict);
        }

        // Primary trustors first
        return result
            .OrderByDescending(d => d.TryGetValue("isPrimary", out var v) && !string.Equals(v, "false", StringComparison.OrdinalIgnoreCase) && !string.Equals(v, "0", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(v))
            .ThenBy(d => d.TryGetValue("sortOrder", out var s) ? s : "9999")
            .Cast<IReadOnlyDictionary<string, string>>()
            .ToList();
    }

    public async Task<IReadOnlyList<IReadOnlyDictionary<string, string>>> GetTrusteesForBeneficiaryAsync(
        string beneficiaryId, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(options.TrusteeRecordsUrl) || string.IsNullOrWhiteSpace(beneficiaryId))
            return Array.Empty<IReadOnlyDictionary<string, string>>();

        var records = await GetRecordsAsync(options.TrusteeRecordsUrl, cancellationToken);
        var result = new List<Dictionary<string, string>>();

        foreach (var record in records)
        {
            var beneficiaryField = GetFieldAsString(record.Fields, "beneficiaryId", "beneficiary", "clientId", "client");
            if (!string.Equals(beneficiaryField, beneficiaryId, StringComparison.OrdinalIgnoreCase))
                continue;

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in record.Fields)
                dict[field.Key] = ConvertJsonToString(field.Value);
            result.Add(dict);
        }

        // Primary role first, then by sortOrder
        return result
            .OrderBy(d => { d.TryGetValue("role", out var r); return string.Equals(r, "Primary", StringComparison.OrdinalIgnoreCase) ? 0 : 1; })
            .ThenBy(d => d.TryGetValue("sortOrder", out var s) ? s : "9999")
            .Cast<IReadOnlyDictionary<string, string>>()
            .ToList();
    }

    public async Task<IReadOnlyDictionary<string, string>?> GetDocumentExecutionAsync(
        string beneficiaryId, string templateName, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(options.DocumentExecutionRecordsUrl) || string.IsNullOrWhiteSpace(beneficiaryId))
            return null;

        var records = await GetRecordsAsync(options.DocumentExecutionRecordsUrl, cancellationToken);

        foreach (var record in records)
        {
            var beneficiaryField = GetFieldAsString(record.Fields, "beneficiaryId", "beneficiary", "clientId", "client");
            if (!string.Equals(beneficiaryField, beneficiaryId, StringComparison.OrdinalIgnoreCase))
                continue;

            // Match by template name if the row has one; otherwise match any row for this beneficiary
            var docTemplate = GetFieldAsString(record.Fields, "documentTemplate", "template", "templateName");
            if (!string.IsNullOrWhiteSpace(docTemplate) &&
                !string.Equals(docTemplate, templateName, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(docTemplate, Path.GetFileNameWithoutExtension(templateName), StringComparison.OrdinalIgnoreCase))
                continue;

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in record.Fields)
                dict[field.Key] = ConvertJsonToString(field.Value);
            return dict;
        }

        return null;
    }

    private async Task<IReadOnlyList<GristRecord>> GetRecordsAsync(string url, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            throw new InvalidOperationException("A required Grist records URL is not configured.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Get, url);

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
            return Array.Empty<GristRecord>();
        }

        var result = new List<GristRecord>();
        foreach (var record in records.EnumerateArray())
        {
            var recordId = record.TryGetProperty("id", out var idElement)
                ? idElement.ToString()
                : string.Empty;

            if (!record.TryGetProperty("fields", out var fields) || fields.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var dict = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in fields.EnumerateObject())
            {
                dict[field.Name] = field.Value.Clone();
            }

            result.Add(new GristRecord(recordId, dict));
        }

        return result;
    }

    private async Task<GristRecord?> GetMoorDocumentRecordAsync(string clientId, string documentName, CancellationToken cancellationToken)
    {
        var records = await GetRecordsAsync(options.MoorDocumentRecordsUrl, cancellationToken);
        var docNameNoExt = Path.GetFileNameWithoutExtension(documentName);

        foreach (var record in records)
        {
            var rowClientId = GetFieldAsString(record.Fields, "clientId", "client", "beneficiaryId", "beneficiary").Trim();
            if (!string.Equals(rowClientId, clientId, StringComparison.OrdinalIgnoreCase))
                continue;

            var rowDocName = GetFieldAsString(record.Fields, "documentName", "document", "templateName", "documentTemplate").Trim();
            if (string.IsNullOrWhiteSpace(rowDocName))
                return record;

            if (string.Equals(rowDocName, documentName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(rowDocName, docNameNoExt, StringComparison.OrdinalIgnoreCase))
                return record;
        }

        return null;
    }

    private async Task<IReadOnlyDictionary<string, string>?> GetRelatedRowAsFlatDictionaryAsync(
        string tableName,
        string? recordId,
        string clientId,
        CancellationToken cancellationToken)
    {
        var records = await GetRecordsForDynamicTableAsync(tableName, cancellationToken);
        if (records.Count == 0)
            return null;

        GristRecord? match = null;

        if (!string.IsNullOrWhiteSpace(recordId))
        {
            match = records.FirstOrDefault(r =>
                string.Equals(r.Id, recordId, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(GetFieldAsString(r.Fields, "id", "recordId", "rowId"), recordId, StringComparison.OrdinalIgnoreCase));
        }

        if (match is null)
        {
            match = records.FirstOrDefault(r =>
                string.Equals(GetFieldAsString(r.Fields, "clientId", "client", "beneficiaryId", "beneficiary"), clientId, StringComparison.OrdinalIgnoreCase));
        }

        if (match is null)
            return null;

        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in match.Fields)
            dict[field.Key] = ConvertJsonToString(field.Value);
        return dict;
    }

    private async Task<IReadOnlyList<GristRecord>> GetRecordsForDynamicTableAsync(string tableName, CancellationToken cancellationToken)
    {
        var primaryUrl = ResolveTableRecordsUrl(tableName);
        if (string.IsNullOrWhiteSpace(primaryUrl))
            return Array.Empty<GristRecord>();

        try
        {
            return await GetRecordsAsync(primaryUrl, cancellationToken);
        }
        catch
        {
            // Allow Moor.Trustee style names in JSON even when Grist table IDs are Moor_Trustee.
            if (!tableName.Contains('.', StringComparison.Ordinal))
                return Array.Empty<GristRecord>();

            var fallbackUrl = ResolveTableRecordsUrl(tableName.Replace('.', '_'));
            if (string.IsNullOrWhiteSpace(fallbackUrl) || string.Equals(fallbackUrl, primaryUrl, StringComparison.OrdinalIgnoreCase))
                return Array.Empty<GristRecord>();

            try
            {
                return await GetRecordsAsync(fallbackUrl, cancellationToken);
            }
            catch
            {
                return Array.Empty<GristRecord>();
            }
        }
    }

    private string ResolveTableRecordsUrl(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            return string.Empty;

        if (Uri.TryCreate(tableName, UriKind.Absolute, out _))
            return tableName;

        var seedUrl = options.MoorDocumentRecordsUrl;
        if (string.IsNullOrWhiteSpace(seedUrl))
            seedUrl = options.BeneficiaryRecordsUrl;

        if (string.IsNullOrWhiteSpace(seedUrl))
            return string.Empty;

        var marker = "/tables/";
        var idx = seedUrl.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
            return string.Empty;

        var prefix = seedUrl[..(idx + marker.Length)];
        return $"{prefix}{Uri.EscapeDataString(tableName)}/records";
    }

    private static IReadOnlyList<TableReference> ParseDataRequiredReferences(string raw)
    {
        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;
            var result = new List<TableReference>();

            if (root.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in root.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object)
                        continue;

                    var table = ReadStringProperty(item, "table", "tableName", "sourceTable");
                    if (string.IsNullOrWhiteSpace(table))
                        continue;

                    var id = ReadStringProperty(item, "id", "recordId", "rowId", "clientId", "beneficiaryId");
                    result.Add(new TableReference(table, string.IsNullOrWhiteSpace(id) ? null : id));
                }

                return result;
            }

            if (root.ValueKind == JsonValueKind.Object)
            {
                // Object style: { "Moor.Trustee": "12", "Moor.Witness": ["31","32"] }
                foreach (var prop in root.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in prop.Value.EnumerateArray())
                        {
                            var id = ConvertJsonToString(item);
                            result.Add(new TableReference(prop.Name, string.IsNullOrWhiteSpace(id) ? null : id));
                        }
                    }
                    else
                    {
                        var id = ConvertJsonToString(prop.Value);
                        result.Add(new TableReference(prop.Name, string.IsNullOrWhiteSpace(id) ? null : id));
                    }
                }

                return result;
            }

            return Array.Empty<TableReference>();
        }
        catch
        {
            // Invalid JSON should not block document generation.
            return Array.Empty<TableReference>();
        }
    }

    private static string ReadStringProperty(JsonElement obj, params string[] aliases)
    {
        foreach (var alias in aliases)
        {
            foreach (var prop in obj.EnumerateObject())
            {
                if (string.Equals(NormalizeKey(prop.Name), NormalizeKey(alias), StringComparison.Ordinal))
                    return ConvertJsonToString(prop.Value);
            }
        }

        return string.Empty;
    }

    private static bool IsTrusteeTable(string tableName)
    {
        var normalized = NormalizeKey(tableName);
        return normalized.Contains("trustee", StringComparison.Ordinal);
    }

    private static string GetFieldAsString(IReadOnlyDictionary<string, JsonElement> fields, params string[] aliases)
    {
        foreach (var alias in aliases)
        {
            if (TryFindField(fields, alias, out var value))
            {
                return ConvertJsonToString(value);
            }
        }

        return string.Empty;
    }

    private static bool GetFieldAsBoolean(IReadOnlyDictionary<string, JsonElement> fields, params string[] aliases)
    {
        foreach (var alias in aliases)
        {
            if (!TryFindField(fields, alias, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.True)
            {
                return true;
            }

            if (value.ValueKind == JsonValueKind.False || value.ValueKind == JsonValueKind.Null)
            {
                return false;
            }

            var stringValue = ConvertJsonToString(value).Trim();
            if (string.IsNullOrWhiteSpace(stringValue))
            {
                return false;
            }

            if (bool.TryParse(stringValue, out var boolValue))
            {
                return boolValue;
            }

            if (int.TryParse(stringValue, out var intValue))
            {
                return intValue != 0;
            }

            if (string.Equals(stringValue, "yes", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(stringValue, "y", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(stringValue, "no", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(stringValue, "n", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return false;
    }

    private static bool TryFindField(IReadOnlyDictionary<string, JsonElement> fields, string alias, out JsonElement value)
    {
        if (fields.TryGetValue(alias, out value))
        {
            return true;
        }

        var normalizedAlias = NormalizeKey(alias);
        foreach (var item in fields)
        {
            if (string.Equals(NormalizeKey(item.Key), normalizedAlias, StringComparison.Ordinal))
            {
                value = item.Value;
                return true;
            }
        }

        value = default;
        return false;
    }

    private static string NormalizeKey(string key)
    {
        var chars = key.Where(char.IsLetterOrDigit).ToArray();
        return new string(chars).ToLowerInvariant();
    }

    private static string ConvertJsonToString(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.Null => string.Empty,
            JsonValueKind.String => value.GetString() ?? string.Empty,
            _ => value.ToString()
        };
    }

    private sealed record TableReference(string Table, string? RecordId);

    private sealed record GristRecord(string Id, IReadOnlyDictionary<string, JsonElement> Fields);
}
