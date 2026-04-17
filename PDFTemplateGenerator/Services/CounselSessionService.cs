using PDFTemplateGenerator.Models;

namespace PDFTemplateGenerator.Services;

public sealed class CounselSessionService : ICounselSessionService
{
    private readonly HashSet<string> _authorizedClientIds = new(StringComparer.OrdinalIgnoreCase);

    public bool IsAuthenticated => CurrentCounsel is not null;
    public CounselUser? CurrentCounsel { get; private set; }
    public IReadOnlyCollection<string> AuthorizedClientIds => _authorizedClientIds;

    public event Action OnChange = delegate { };

    public void SignIn(CounselUser counsel, IEnumerable<string> authorizedClientIds)
    {
        CurrentCounsel = counsel;
        _authorizedClientIds.Clear();

        foreach (var clientId in authorizedClientIds)
        {
            if (!string.IsNullOrWhiteSpace(clientId))
            {
                _authorizedClientIds.Add(clientId.Trim());
            }
        }

        OnChange.Invoke();
    }

    public void SignOut()
    {
        CurrentCounsel = null;
        _authorizedClientIds.Clear();
        OnChange.Invoke();
    }
}
