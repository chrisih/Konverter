using System.IO;
using System.Security.Cryptography;
using Konverter.Models;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Desktop;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Services
{
  public interface IOnedriveService
  {
    Task<AuthenticationResult> Login();
    Task Save(Presentation presentation);
  }

  public class OnedriveService : IOnedriveService
  {
    private readonly IOptionsMonitor<OnedriveConfig> _config;

    public OnedriveService(IOptionsMonitor<OnedriveConfig> config)
    {
      _config = config;
    }

    public async Task<AuthenticationResult> Login()
    {
      var app = PublicClientApplicationBuilder
                .Create(_config.CurrentValue.ClientId)
                .WithWindowsDesktopFeatures(new BrokerOptions(BrokerOptions.OperatingSystems.Windows) {ListOperatingSystemAccounts = true})
                .WithDefaultRedirectUri()
                .Build();
      
      var atp = new IntegratedWindowsTokenProvider(_config.CurrentValue.ClientId, _config.CurrentValue.TenantId, _config.CurrentValue.Scopes);
      var provider = new BaseBearerTokenAuthenticationProvider(atp);
      
      TokenCacheHelper.EnableSerialization(app.UserTokenCache);

      var _client = new GraphServiceClient(provider);

      var accounts = await app.GetAccountsAsync();
      
      AuthenticationResult token = null;

      try
      {
        token = await app.AcquireTokenSilent(_config.CurrentValue.Scopes, accounts.FirstOrDefault()).ExecuteAsync();
      }
      catch (MsalServiceException ex)
      {
        token = await app.AcquireTokenInteractive(_config.CurrentValue.Scopes).ExecuteAsync();
      }
      return token;
    }

    public async Task Save(Presentation presentation)
    {

    }
  }

  public class IntegratedWindowsTokenProvider : IAccessTokenProvider
{
    private readonly IPublicClientApplication publicClient;
    private readonly string[] _scopes;
    
    public IntegratedWindowsTokenProvider(string clientId, string tenantId, IEnumerable<string> scopes)
    { 
      _scopes = scopes.ToArray();

      publicClient = PublicClientApplicationBuilder
          .Create(clientId)
          .WithTenantId(tenantId)
          .Build();

        AllowedHostsValidator = new AllowedHostsValidator();
    }

    /// <summary>
    /// Gets an <see cref="AllowedHostsValidator"/> that validates if the
    /// target host of a request is allowed for authentication.
    /// </summary>
    public AllowedHostsValidator AllowedHostsValidator { get; }

    /// <inheritdoc/>
    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        var result = await publicClient
            .AcquireTokenByIntegratedWindowsAuth(_scopes)
            .ExecuteAsync(cancellationToken);
        return result.AccessToken;
    }
}

  static class TokenCacheHelper
  {
    static TokenCacheHelper()
    {
      try
      {
        CacheFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ".msalcache.bin3");
      }
      catch (InvalidOperationException)
      {
        CacheFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location + ".msalcache.bin3";
      }
    }

    public static string CacheFilePath { get; private set; }

    private static readonly object FileLock = new object();

    public static void BeforeAccessNotification(TokenCacheNotificationArgs args)
    {
      lock (FileLock)
      {
        args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath)
                ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath), null, DataProtectionScope.CurrentUser)
                : null);
      }
    }

    public static void AfterAccessNotification(TokenCacheNotificationArgs args)
    {
      if (args.HasStateChanged)
      {
        lock (FileLock)
        {
          File.WriteAllBytes(CacheFilePath, ProtectedData.Protect(args.TokenCache.SerializeMsalV3(), null, DataProtectionScope.CurrentUser));
        }
      }
    }

    internal static void EnableSerialization(ITokenCache tokenCache)
    {
      tokenCache.SetBeforeAccess(BeforeAccessNotification);
      tokenCache.SetAfterAccess(AfterAccessNotification);
    }
  }
}
