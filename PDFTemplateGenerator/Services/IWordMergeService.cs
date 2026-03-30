namespace PDFTemplateGenerator.Services;

public interface IWordMergeService
{
    Task<string> FillDocxPlaceholdersFromCsvAsync(
        string templateAsset = "Template.docx",
        string csvAsset = "Data.csv",
        string outputFileName = "Output_Filled.docx");

    Task<string> FillDocxTableFromCsvAsync(
        string templateAsset = "Template.docx",
        string csvAsset = "Data.csv",
        string outputFileName = "Output_Table.docx",
        bool matchTableByHeader = true);
}
