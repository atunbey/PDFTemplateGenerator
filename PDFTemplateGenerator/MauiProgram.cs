using Microsoft.Extensions.Logging;
using PDFTemplateGenerator.Services;
using System.Net;

namespace PDFTemplateGenerator
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder
                .UseMauiApp<App>()
                .ConfigureFonts(fonts =>
                {
                    fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                });

            builder.Services.AddMauiBlazorWebView();
            builder.Services.AddSingleton<ICsvReaderService, CsvReaderService>();
            builder.Services.AddSingleton<IWordTemplateRenderer, WordTemplateRenderer>();
            builder.Services.AddSingleton<IWordMergeService, WordMergeService>();
            builder.Services.AddSingleton<IDealerInventoryReportService, DealerInventoryReportService>();
            builder.Services.AddSingleton<ExcelMergeService>();
            builder.Services.AddSingleton(new HttpClient(new SocketsHttpHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            }));
            builder.Services.AddSingleton(new AutomotiveApiOptions
            {
                RecordsUrl = Environment.GetEnvironmentVariable("AUTOMOTIVE_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/czGVryH9XzdH5qBAXzD4ib/tables/AutomotiveHeaders/records",
                ApiKey = Environment.GetEnvironmentVariable("AUTOMOTIVE_API_KEY")
                    ?? "863a5652184fa2a988f217019a3ebf751f7d3fc7"
            });
            builder.Services.AddSingleton(new GristApiOptions
            {
                BeneficiaryRecordsUrl = Environment.GetEnvironmentVariable("GRIST_BENEFICIARY_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Beneficiary/records",
                ApiKey = Environment.GetEnvironmentVariable("GRIST_API_KEY")
                    ?? "863a5652184fa2a988f217019a3ebf751f7d3fc7"
            });
            builder.Services.AddSingleton(new LegalDocumentOptions
            {
                CertificateTemplateFolder = Environment.GetEnvironmentVariable("LEGALDOC_CERT_TEMPLATE_PATH")
                    ?? string.Empty,
                NextcloudFolderUrl = Environment.GetEnvironmentVariable("LEGALDOC_NEXTCLOUD_FOLDER_URL")
                    ?? "https://tools.kushkurriculum.org/nextcloud/s/oPsT24RoZGssbMN",
                NextcloudWebDavFolderUrl = Environment.GetEnvironmentVariable("LEGALDOC_NEXTCLOUD_WEBDAV_FOLDER_URL")
                    ?? string.Empty,
                NextcloudUsername = Environment.GetEnvironmentVariable("LEGALDOC_NEXTCLOUD_USERNAME")
                    ?? string.Empty,
                NextcloudAppPassword = Environment.GetEnvironmentVariable("LEGALDOC_NEXTCLOUD_APP_PASSWORD")
                    ?? string.Empty,
                EnableNextcloudTemplates = !string.Equals(
                    Environment.GetEnvironmentVariable("LEGALDOC_ENABLE_NEXTCLOUD_TEMPLATES"),
                    "false",
                    StringComparison.OrdinalIgnoreCase),
                EnableLocalTemplateFallback = string.Equals(
                    Environment.GetEnvironmentVariable("LEGALDOC_ENABLE_LOCAL_TEMPLATE_FALLBACK"),
                    "true",
                    StringComparison.OrdinalIgnoreCase)
            });
            builder.Services.AddSingleton(new DealerInventoryReportOptions
            {
                WorkingDirectory = Environment.GetEnvironmentVariable("DEALER_REPORT_WORKING_DIRECTORY")
                    ?? "C:\\Users\\atunbey\\OneDrive - Afuraka Technology Services\\Documents\\ParkersDocumentProcessing\\",
                CompleteInventoryRelativePath = Environment.GetEnvironmentVariable("DEALER_COMPLETE_INVENTORY_PATH")
                    ?? "CSVinventory\\comsoftInventoryCSI2JTZ.CSV",
                WebsiteInventoryRelativePath = Environment.GetEnvironmentVariable("DEALER_WEBSITE_INVENTORY_PATH")
                    ?? "CSVinventory\\WebsiteInventory.csv"
            });
            builder.Services.AddSingleton<IGristClientService, GristClientService>();
            builder.Services.AddSingleton<ILegalDocumentService, LegalDocumentService>();

#if DEBUG
            builder.Services.AddBlazorWebViewDeveloperTools();
    		builder.Logging.AddDebug();
#endif

            return builder.Build();
        }
    }
}
