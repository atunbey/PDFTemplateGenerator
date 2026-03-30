namespace PDFTemplateGenerator.Services;

public sealed class CsvReaderService : ICsvReaderService
{
    public async Task<CsvReadResult> ReadFromAppPackageAsync(
        string assetFileName,
        char separator = ',',
        CancellationToken cancellationToken = default)
    {
        await using var stream = await FileSystem.OpenAppPackageFileAsync(assetFileName);
        var parsed = await ReadParsedLinesAsync(stream, separator, cancellationToken);

        if (parsed.Count == 0)
        {
            return new CsvReadResult();
        }

        var header = parsed[0].Select(h => (h ?? string.Empty).Trim()).ToList();
        var rows = parsed
            .Skip(1)
            .Where(r => r.Any(v => !string.IsNullOrWhiteSpace(v)))
            .ToList();

        return new CsvReadResult
        {
            Header = header,
            Rows = rows
        };
    }

    public async Task<List<List<string>>> ReadRowsFromFileAsync(
        string absolutePath,
        char separator = ',',
        CancellationToken cancellationToken = default)
    {
        await using var stream = File.OpenRead(absolutePath);
        return await ReadParsedLinesAsync(stream, separator, cancellationToken);
    }

    private static async Task<List<List<string>>> ReadParsedLinesAsync(
        Stream stream,
        char separator,
        CancellationToken cancellationToken)
    {
        using var reader = new StreamReader(stream);
        var lines = new List<string>();

        while (!reader.EndOfStream)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync();
            if (line != null)
            {
                lines.Add(line);
            }
        }

        return lines.Select(l => ParseCsvLine(l, separator)).ToList();
    }

    private static List<string> ParseCsvLine(string line, char separator)
    {
        var result = new List<string>();
        var current = new System.Text.StringBuilder();
        var inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            var ch = line[i];

            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(ch);
                }
            }
            else
            {
                if (ch == '"')
                {
                    inQuotes = true;
                }
                else if (ch == separator)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(ch);
                }
            }
        }

        result.Add(current.ToString());
        return result;
    }
}
