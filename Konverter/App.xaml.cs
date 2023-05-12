using Dropbox.Api.Files;
using Dropbox.Api;
using System.Windows;
using static Dropbox.Api.TeamLog.EventCategory;
using System.Net.Http;

namespace Konverter
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application
  {
    #region Variables  
    private DropboxClient DBClient;
    private ListFolderArg DBFolders;
    private string oauth2State;
    private const string RedirectUri = "https://localhost/authorize"; // Same as we have configured Under [Application] -> settings -> redirect URIs.  
    #endregion

    public App()
    {
      Uri authorizeUri = DropboxOAuth2Helper.GetAuthorizeUri(OAuthResponseType.Token, "00k3px1id0wizbq", RedirectUri, state: oauth2State);
      var AuthenticationURL = authorizeUri.AbsoluteUri.ToString();

      //var login = new Login("00k3px1id0wizbq", AuthenticationURL, oauth2State);
      //login.ShowDialog();
      //if (login.Result)
      {
        //var AccessToken = login.AccessToken;
        //DropboxClientConfig CC = new DropboxClientConfig("kghuelben", 1);
        //HttpClient HTC = new HttpClient();
        //HTC.Timeout = TimeSpan.FromMinutes(10); // set timeout for each ghttp request to Dropbox API.  
        //CC.HttpClient = HTC;

        //var AccessToken = "sl.BeGn-B1CjyJJ3ofDoH22prqpPcJFaAbUTSQtMEX81Jd_g7H8fysjNOt_RX-U_kcsL2ILBYscviWEkenYNnhVJM5sFGN9GwSQejc0j6uBceLQWhaeeMuE0YxsfABvahaiM4KsmIM";

        //DBClient = new DropboxClient(AccessToken, CC);
        //GetFolders();
      }
    }

    private async void GetFolders()
    {
      var sharedLink = new SharedLink("https://www.dropbox.com/sh/c4tuhbjz4p0npv4/AACURQzuxX8rj8RFRZSfSlzRa?dl=0");
      var sharedFiles = await DBClient.Files.ListFolderAsync(path: "", sharedLink: sharedLink);
    }
  }
}