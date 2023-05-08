using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Windows;

namespace Konverter
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application, IAuthenticationProvider
  {
    public App()
    {
      
    }

    private async void InitGraph()
    {
      var tenant = "5b7bb46f-d979-4f15-939c-8ed3211bcefc";
      var clientid = "f16e659d-3182-4f18-b51b-553ecaefc5ca";
      var clientsecret = "6S38Q~At9sLWfYV4PVz.IvRboN-ApH798TGS-b~1";

      var csc = new InteractiveBrowserCredential(tenant, clientid, new InteractiveBrowserCredentialOptions
      {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        RedirectUri = new Uri("https://login.microsoftonline.com/common/oauth2/nativeclient")
      });

      
      var c = new ChainedTokenCredential(csc);
      
      var client = new GraphServiceClient(c, new string[] { "https://graph.microsoft.com/.default" });

      try
      {
        var drive = await client.Me.Drive.GetAsync();
      }
      catch (ODataError err)
      {
        var mainError = err.Error;
        if (mainError != null)
        {
          var inner = mainError.Innererror;
          if(inner != null)
          {
            
          }
        }
        var e = err.AdditionalData.ToDictionary(k => k.Key, k => k.Value);
      }
    }

    private AuthenticationResult _result;

    public async Task AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
    {
      if (request.URI.Host == "graph.microsoft.com")
      {        
        request.Headers.Add("Authorization", $"Bearer {_result.AccessToken}");
      }
    }
  }
}













/*
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Desktop;
using Microsoft.Identity.Client.Broker;
using System.Windows;
using active_directory_wpf_msgraph_v2;

namespace Konverter
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application
  {
    public App()
    {
      CreateApplication(true, false);
    }

    public static void CreateApplication(bool useWam, bool useBrokerPreview)
    {
      var builder = PublicClientApplicationBuilder.Create(ClientId)
          .WithAuthority($"{Instance}{Tenant}")
          .WithDefaultRedirectUri();

      //Use of Broker Requires redirect URI "ms-appx-web://microsoft.aad.brokerplugin/{client_id}" in app registration
      if (useWam && !useBrokerPreview)
      {
        builder.WithWindowsBroker(true);
      }
      else if (useWam && useBrokerPreview)
      {
        builder.WithBrokerPreview(true);
      }
      _clientApp = builder.Build();
      TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
    }

    // Below are the clientId (Application Id) of your app registration and the tenant information. 
    // You have to replace:
    // - the content of ClientID with the Application Id for your app registration
    // - The content of Tenant by the information about the accounts allowed to sign-in in your application:
    //   - For Work or School account in your org, use your tenant ID, or domain
    //   - for any Work or School accounts, use organizations
    //   - for any Work or School accounts, or Microsoft personal account, use organizations
    //   - for Microsoft Personal account, use consumers
    private static string ClientId = "23b77883-36ba-4ca2-b9cc-fba7bd0139d1";

    // Note: Tenant is important for the quickstart.
    private static string Tenant = "organizations";
    private static string Instance = "https://login.microsoftonline.com/";
    private static IPublicClientApplication _clientApp;

    public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
  }
}
 */