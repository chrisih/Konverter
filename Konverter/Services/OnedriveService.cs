using Konverter.Models;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using Konverter.Services.Abstraction;

namespace Konverter.Services
{
  public class OnedriveService : IOnedriveService
  {
    public OnedriveService(IOptionsMonitor<OnedriveConfig> config) 
    {
      WindowsBrokerOptions brokerOptions = new WindowsBrokerOptions();

      var _clientApp = PublicClientApplicationBuilder.Create(config.CurrentValue.ClientId)
          .WithAuthority($"https://login.microsoftonline.com/{config.CurrentValue.TenantId}")
          .WithRedirectUri("http://localhost:54448")
          .WithWindowsBrokerOptions(brokerOptions)
          .Build();

      MsalCacheHelper cacheHelper = CreateCacheHelperAsync().GetAwaiter().GetResult();

      // Let the cache helper handle MSAL's cache, otherwise the user will be prompted to sign-in every time.
      cacheHelper.RegisterCache(_clientApp.UserTokenCache);
      
      var accounts = _clientApp.GetAccountsAsync().GetAwaiter().GetResult();
      AuthenticationResult authResult = null;

      if(!accounts.Any())
      {
        authResult = _clientApp.AcquireTokenInteractive(config.CurrentValue.Scopes).ExecuteAsync().GetAwaiter().GetResult();
      }
      else
      {
        authResult = _clientApp.AcquireTokenSilent(config.CurrentValue.Scopes, accounts.First()).ExecuteAsync().GetAwaiter().GetResult();
      }

    }

    private static async Task<MsalCacheHelper> CreateCacheHelperAsync()
    {
      // Since this is a WPF application, only Windows storage is configured
      var storageProperties = new StorageCreationPropertiesBuilder(
                        System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".msalcache.bin",
                        MsalCacheHelper.UserRootDirectory)
                          .Build();

      MsalCacheHelper cacheHelper = await MsalCacheHelper.CreateAsync(
                  storageProperties,
                  new TraceSource("MSAL.CacheTrace"))
               .ConfigureAwait(false);

      return cacheHelper;
    }

    public Task Logon()
    {
      throw new NotImplementedException();
    }

    public async Task Save(Presentation presentation)
    {

    }
  }
}
