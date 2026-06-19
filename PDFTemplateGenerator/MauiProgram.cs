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
            builder.Services.AddSingleton(new HttpClient(new SocketsHttpHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            }));
            builder.Services.AddSingleton(new GristApiOptions
            {
                BeneficiaryRecordsUrl = Environment.GetEnvironmentVariable("GRIST_BENEFICIARY_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Beneficiary/records",
                CounselRecordsUrl = Environment.GetEnvironmentVariable("GRIST_COUNSEL_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Moor_Counsel/records",
                AssociationsRecordsUrl = Environment.GetEnvironmentVariable("GRIST_ASSOCIATIONS_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Moor_Associations/records",
                MoorDocumentRecordsUrl = Environment.GetEnvironmentVariable("GRIST_MOOR_DOCUMENT_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Moor_Document/records",
                ApiKey = Environment.GetEnvironmentVariable("GRIST_API_KEY")
                    ?? "863a5652184fa2a988f217019a3ebf751f7d3fc7",

                // New relational tables — create these in Grist and set the env vars, or replace the empty defaults below.
                TrustorRecordsUrl = Environment.GetEnvironmentVariable("GRIST_TRUSTOR_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Trustor/records",
                TrusteeRecordsUrl = Environment.GetEnvironmentVariable("GRIST_TRUSTEE_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/Moor_Trustee/records",
                DocumentExecutionRecordsUrl = Environment.GetEnvironmentVariable("GRIST_DOCUMENT_EXECUTION_RECORDS_URL")
                    ?? "https://onlinedata.kushkurriculum.org/api/docs/193hm5A4YK9FczhVGXxtgo/tables/DocumentExecution/records",
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
            builder.Services.AddSingleton<IGristClientService, GristClientService>();
            builder.Services.AddSingleton<ICounselSessionService, CounselSessionService>();
            builder.Services.AddSingleton<ILegalDocumentService, LegalDocumentService>();

#if DEBUG
            builder.Services.AddBlazorWebViewDeveloperTools();
    		builder.Logging.AddDebug();
#endif

            return builder.Build();
        }
    }
}
