using System.Net.Http.Headers;
using System.Text;
using System.Xml.Linq;
using NPOI.XWPF.UserModel;

namespace PDFTemplateGenerator.Services;

public sealed class LegalDocumentService(
    IGristClientService gristClientService,
    LegalDocumentOptions options,
    HttpClient httpClient) : ILegalDocumentService
{
    public async Task<IReadOnlyList<string>> GetAvailableTemplatesAsync(CancellationToken cancellationToken = default)
    {
        if (options.EnableNextcloudTemplates)
        {
            var nextcloudTemplates = await GetNextcloudTemplatesAsync(cancellationToken);
            if (nextcloudTemplates.Count > 0)
            {
                return nextcloudTemplates;
            }
        }

        if (options.EnableLocalTemplateFallback)
        {
            return GetLocalTemplates().ToList();
        }

        return Array.Empty<string>();
    }

    public async Task<string> GenerateCertificateOfTrustAsync(string clientId, string templateFileName, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(templateFileName))
            throw new InvalidOperationException("No template file was selected.");

        if (string.IsNullOrWhiteSpace(clientId))
            throw new InvalidOperationException("No beneficiary record ID was selected.");

        var client = await gristClientService.GetBeneficiaryByIdAsync(clientId, cancellationToken)
            ?? throw new InvalidOperationException($"Client with id {clientId} was not found in Grist.");

        var templateBytes = await ResolveTemplateBytesAsync(templateFileName, cancellationToken);
        using var ms = new MemoryStream(templateBytes);
        using var doc = new XWPFDocument(ms);

        foreach (var headerParagraph in doc.HeaderList.SelectMany(h => h.Paragraphs))
        {
            ReplaceInParagraph(headerParagraph, client.Fields);
        }

        foreach (var paragraph in doc.Paragraphs)
        {
            ReplaceInParagraph(paragraph, client.Fields);
        }

        foreach (var table in doc.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        ReplaceInParagraph(paragraph, client.Fields);
                    }
                }
            }
        }

        var outputFileName = BuildOutputFileName(client, templateFileName);
        var outputDirectory = GetPreferredOutputDirectory();
        var outPath = Path.Combine(outputDirectory, outputFileName);
        await using var outFs = new FileStream(outPath, FileMode.Create, FileAccess.Write);
        doc.Write(outFs);

        return outPath;
    }

    private async Task<byte[]> ResolveTemplateBytesAsync(string templateFileName, CancellationToken cancellationToken)
    {
        if (options.EnableNextcloudTemplates)
        {
            try
            {
                return await DownloadTemplateFromNextcloudAsync(templateFileName, cancellationToken);
            }
            catch when (options.EnableLocalTemplateFallback)
            {
                // Fallback handled below.
            }
        }

        if (options.EnableLocalTemplateFallback)
        {
            return await LoadLocalTemplateBytesAsync(templateFileName, cancellationToken);
        }

        throw new FileNotFoundException(
            "Template was not found from Nextcloud and local fallback is disabled.",
            templateFileName);
    }

    private async Task<IReadOnlyList<string>> GetNextcloudTemplatesAsync(CancellationToken cancellationToken)
    {
        var folderUrl = ResolveNextcloudWebDavFolderUrl();
        if (string.IsNullOrWhiteSpace(folderUrl))
        {
            return Array.Empty<string>();
        }

        using var response = await SendPropfindAsync(folderUrl, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return Array.Empty<string>();
        }

        var xml = await response.Content.ReadAsStringAsync(cancellationToken);
        var xdoc = XDocument.Parse(xml);
        XNamespace dav = "DAV:";

        var templates = xdoc.Descendants(dav + "response")
            .Select(r => r.Element(dav + "href")?.Value)
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Select(h => ExtractFileNameFromHref(folderUrl, h!))
            .Where(name => !string.IsNullOrWhiteSpace(name) && name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return templates;
    }

    private async Task<HttpResponseMessage> SendPropfindAsync(string folderUrl, CancellationToken cancellationToken)
    {
        const string propfindBody = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><d:propfind xmlns:d=\"DAV:\"><d:prop><d:resourcetype/></d:prop></d:propfind>";

        try
        {
            using var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), folderUrl);
            request.Headers.Add("Depth", "1");
            ApplyNextcloudAuth(request);
            request.Content = new StringContent(propfindBody, Encoding.UTF8, "application/xml");

            return await httpClient.SendAsync(request, cancellationToken);
        }
        catch (Exception ex) when (IsUnsupportedHttpMethod(ex))
        {
            // Android handlers can reject custom verbs like PROPFIND.
            using var fallbackRequest = new HttpRequestMessage(HttpMethod.Post, folderUrl);
            fallbackRequest.Headers.Add("Depth", "1");
            fallbackRequest.Headers.Add("X-HTTP-Method-Override", "PROPFIND");
            ApplyNextcloudAuth(fallbackRequest);
            fallbackRequest.Content = new StringContent(propfindBody, Encoding.UTF8, "application/xml");

            return await httpClient.SendAsync(fallbackRequest, cancellationToken);
        }
    }

    private static bool IsUnsupportedHttpMethod(Exception ex)
    {
        Exception? current = ex;
        while (current is not null)
        {
            var message = current.Message ?? string.Empty;
            if (message.Contains("but was PROPFIND", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("invalid http method", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("unsupported", StringComparison.OrdinalIgnoreCase) &&
                message.Contains("method", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            current = current.InnerException;
        }

        return false;
    }

    private async Task<byte[]> DownloadTemplateFromNextcloudAsync(string templateFileName, CancellationToken cancellationToken)
    {
        var folderUrl = ResolveNextcloudWebDavFolderUrl();
        if (string.IsNullOrWhiteSpace(folderUrl))
        {
            throw new InvalidOperationException("Nextcloud WebDAV folder URL is not configured.");
        }

        var requestUrl = BuildFileUrl(folderUrl, templateFileName);
        using var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
        ApplyNextcloudAuth(request);

        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw new FileNotFoundException(
                $"Template '{templateFileName}' was not found in Nextcloud (status {(int)response.StatusCode}).",
                requestUrl);
        }

        return await response.Content.ReadAsByteArrayAsync(cancellationToken);
    }

    private IEnumerable<string> GetLocalTemplates()
    {
        if (string.IsNullOrWhiteSpace(options.CertificateTemplateFolder) || !Directory.Exists(options.CertificateTemplateFolder))
            return Enumerable.Empty<string>();

        return Directory.EnumerateFiles(options.CertificateTemplateFolder, "*.docx", SearchOption.TopDirectoryOnly)
            .Select(Path.GetFileName)
            .Where(f => f is not null)
            .Cast<string>()
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase);
    }

    private async Task<byte[]> LoadLocalTemplateBytesAsync(string templateFileName, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(options.CertificateTemplateFolder))
            throw new InvalidOperationException("Certificate template folder is not configured.");

        var fullTemplatePath = Path.Combine(options.CertificateTemplateFolder, templateFileName);
        if (!File.Exists(fullTemplatePath))
            throw new FileNotFoundException("Certificate template file was not found.", fullTemplatePath);

        return await File.ReadAllBytesAsync(fullTemplatePath, cancellationToken);
    }

    private string ResolveNextcloudWebDavFolderUrl()
    {
        if (!string.IsNullOrWhiteSpace(options.NextcloudWebDavFolderUrl))
        {
            return EnsureTrailingSlash(options.NextcloudWebDavFolderUrl);
        }

        if (string.IsNullOrWhiteSpace(options.NextcloudFolderUrl))
        {
            return string.Empty;
        }

        if (!Uri.TryCreate(options.NextcloudFolderUrl, UriKind.Absolute, out var browseUri))
        {
            return string.Empty;
        }

        if (TryGetShareTokenFromFolderUrl(browseUri, out var shareToken))
        {
            var authorityForShare = browseUri.IsDefaultPort
                ? browseUri.Host
                : $"{browseUri.Host}:{browseUri.Port}";

            var publicWebDav = $"{browseUri.Scheme}://{authorityForShare}/nextcloud/public.php/dav/files/{Uri.EscapeDataString(shareToken)}";
            return EnsureTrailingSlash(publicWebDav);
        }

        if (string.IsNullOrWhiteSpace(options.NextcloudUsername))
        {
            return string.Empty;
        }

        var dir = GetQueryParam(browseUri.Query, "dir");
        if (string.IsNullOrWhiteSpace(dir))
        {
            dir = "/";
        }

        var authority = browseUri.IsDefaultPort
            ? browseUri.Host
            : $"{browseUri.Host}:{browseUri.Port}";

        var encodedDir = EncodePath(dir);
        var webDav = $"{browseUri.Scheme}://{authority}/remote.php/dav/files/{Uri.EscapeDataString(options.NextcloudUsername)}{encodedDir}";
        return EnsureTrailingSlash(webDav);
    }

    private void ApplyNextcloudAuth(HttpRequestMessage request)
    {
        if (!string.IsNullOrWhiteSpace(options.NextcloudUsername))
        {
            var token = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{options.NextcloudUsername}:{options.NextcloudAppPassword}"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", token);
            return;
        }

        if (!Uri.TryCreate(options.NextcloudFolderUrl, UriKind.Absolute, out var browseUri))
        {
            return;
        }

        if (!TryGetShareTokenFromFolderUrl(browseUri, out var shareToken))
        {
            return;
        }

        // Public Nextcloud share WebDAV uses share token as username and optional share password.
        var publicToken = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{shareToken}:{options.NextcloudAppPassword}"));
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", publicToken);
    }

    private static string BuildFileUrl(string folderUrl, string fileName)
    {
        var trimmed = EnsureTrailingSlash(folderUrl);
        return trimmed + Uri.EscapeDataString(fileName);
    }

    private static string ExtractFileNameFromHref(string folderUrl, string href)
    {
        var baseUri = new Uri(EnsureTrailingSlash(folderUrl));
        var absolute = Uri.TryCreate(baseUri, href, out var absoluteUri) ? absoluteUri : baseUri;
        var fileName = Path.GetFileName(Uri.UnescapeDataString(absolute.AbsolutePath));
        return fileName ?? string.Empty;
    }

    private static string EnsureTrailingSlash(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        return value.EndsWith('/') ? value : value + "/";
    }

    private static string GetQueryParam(string query, string key)
    {
        var trimmed = query.StartsWith('?') ? query[1..] : query;
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        foreach (var part in trimmed.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = part.IndexOf('=');
            var rawKey = idx >= 0 ? part[..idx] : part;
            if (!string.Equals(Uri.UnescapeDataString(rawKey), key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var rawValue = idx >= 0 ? part[(idx + 1)..] : string.Empty;
            return Uri.UnescapeDataString(rawValue);
        }

        return string.Empty;
    }

    private static string EncodePath(string path)
    {
        var normalized = string.IsNullOrWhiteSpace(path) ? "/" : path;
        if (!normalized.StartsWith('/'))
        {
            normalized = "/" + normalized;
        }

        var segments = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Select(Uri.EscapeDataString);
        return "/" + string.Join('/', segments);
    }

    private static bool TryGetShareTokenFromFolderUrl(Uri browseUri, out string token)
    {
        token = string.Empty;

        var segments = browseUri.AbsolutePath
            .Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Select(Uri.UnescapeDataString)
            .ToArray();

        for (var i = 0; i < segments.Length - 1; i++)
        {
            if (!string.Equals(segments[i], "s", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            token = segments[i + 1];
            return !string.IsNullOrWhiteSpace(token);
        }

        return false;
    }

    private static string BuildOutputFileName(Models.BeneficiaryClient client, string templateFileName)
    {
        var parts = new[] { client.LastName, client.FirstName, client.MiddleName }
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(SanitizeFileNamePart)
            .ToArray();

        var clientName = parts.Length > 0
            ? string.Join("_", parts)
            : $"Client_{client.RecordId}";

        var templateName = SanitizeFileNamePart(Path.GetFileNameWithoutExtension(templateFileName) ?? string.Empty);
        var baseName = string.IsNullOrWhiteSpace(templateName)
            ? clientName
            : $"{clientName}_{templateName}";

        return $"{baseName}.docx";
    }

    private static string GetPreferredOutputDirectory()
    {
        var candidates = new[]
        {
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            Environment.GetFolderPath(Environment.SpecialFolder.Personal),
            FileSystem.AppDataDirectory
        };

        foreach (var candidate in candidates.Where(c => !string.IsNullOrWhiteSpace(c)))
        {
            try
            {
                Directory.CreateDirectory(candidate);
                return candidate;
            }
            catch
            {
                // Try the next writable location.
            }
        }

        return FileSystem.AppDataDirectory;
    }

    private static string SanitizeFileNamePart(string value)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var clean = new string(value.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());
        return clean.Trim();
    }

    private static void ReplaceInParagraph(XWPFParagraph paragraph, IReadOnlyDictionary<string, string> data)
    {
        var original = paragraph.Text ?? string.Empty;
        var replaced = ReplacePlaceholders(original, data);

        if (replaced == original)
        {
            return;
        }

        for (int i = paragraph.Runs.Count - 1; i >= 0; i--)
        {
            paragraph.RemoveRun(i);
        }

        var run = paragraph.CreateRun();
        run.SetText(replaced);
    }

    private static string ReplacePlaceholders(string text, IReadOnlyDictionary<string, string> data)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        foreach (var kv in data)
        {
            text = text.Replace($"«{kv.Key}»", kv.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            text = text.Replace($"${{{kv.Key}}}", kv.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        return text;
    }
}
