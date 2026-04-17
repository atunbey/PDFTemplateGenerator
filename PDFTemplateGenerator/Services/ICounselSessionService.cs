using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public interface ICounselSessionService
{
    bool IsAuthenticated { get; }
    CounselUser? CurrentCounsel { get; }
    IReadOnlyCollection<string> AuthorizedClientIds { get; }

    event Action OnChange;

    void SignIn(CounselUser counsel, IEnumerable<string> authorizedClientIds);
    void SignOut();
}
