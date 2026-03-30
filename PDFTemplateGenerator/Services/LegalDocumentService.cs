using NPOI.XWPF.UserModel;

namespace PDFTemplateGenerator.Services;

public sealed class LegalDocumentService(
    IGristClientService gristClientService,
    LegalDocumentOptions options) : ILegalDocumentService
{
    public IEnumerable<string> GetAvailableTemplates()
    {
        if (string.IsNullOrWhiteSpace(options.CertificateTemplateFolder) || !Directory.Exists(options.CertificateTemplateFolder))
            return Enumerable.Empty<string>();

        return Directory.EnumerateFiles(options.CertificateTemplateFolder, "*.docx", SearchOption.TopDirectoryOnly)
            .Select(Path.GetFileName)
            .Where(f => f is not null)
            .Cast<string>()
            .OrderBy(f => f);
    }

    public async Task<string> GenerateCertificateOfTrustAsync(string clientId, string templateFileName, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(options.CertificateTemplateFolder))
            throw new InvalidOperationException("Certificate template folder is not configured.");

        if (string.IsNullOrWhiteSpace(templateFileName))
            throw new InvalidOperationException("No template file was selected.");

        if (string.IsNullOrWhiteSpace(clientId))
            throw new InvalidOperationException("No beneficiary record ID was selected.");

        var fullTemplatePath = Path.Combine(options.CertificateTemplateFolder, templateFileName);

                var client = await gristClientService.GetBeneficiaryByIdAsync(clientId, cancellationToken)
                        ?? throw new InvalidOperationException($"Client with id {clientId} was not found in Grist.");

                var templateBytes = await LoadTemplateBytesAsync(fullTemplatePath, cancellationToken);
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
        var outPath = Path.Combine(FileSystem.AppDataDirectory, outputFileName);
        await using var outFs = new FileStream(outPath, FileMode.Create, FileAccess.Write);
        doc.Write(outFs);

        return outPath;
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

    private static async Task<byte[]> LoadTemplateBytesAsync(string templatePathOrAsset, CancellationToken cancellationToken)
    {
        if (Path.IsPathRooted(templatePathOrAsset))
        {
            if (!File.Exists(templatePathOrAsset))
            {
                throw new FileNotFoundException("Certificate template file was not found.", templatePathOrAsset);
            }

            return await File.ReadAllBytesAsync(templatePathOrAsset, cancellationToken);
        }

        await using var packageStream = await FileSystem.OpenAppPackageFileAsync(templatePathOrAsset);
        using var ms = new MemoryStream();
        await packageStream.CopyToAsync(ms, cancellationToken);
        return ms.ToArray();
    }
}
