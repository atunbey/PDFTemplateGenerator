using NPOI.XWPF.UserModel;

namespace PDFTemplateGenerator.Services;

public interface IWordTemplateRenderer
{
    void ReplacePlaceholdersEverywhere(XWPFDocument document, IReadOnlyDictionary<string, string> data);
    XWPFTable? FindTableByHeader(XWPFDocument document, List<string> csvHeader);
    void ClearParagraph(XWPFParagraph paragraph);
}
