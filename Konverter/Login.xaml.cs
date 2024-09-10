using Dropbox.Api;
using System.Windows;

namespace Konverter
{
  /// <summary>
  /// Interaktionslogik für Login.xaml
  /// </summary>
  public partial class Login : Window
  {
    #region Variables  
    private const string RedirectUri = "https://localhost/authorize";
    private string DBAppKey = string.Empty;
    private string DBAuthenticationURL = string.Empty;
    private string DBoauth2State = string.Empty;
    #endregion

    #region Properties  
    public string AccessToken { get; private set; }

    public string UserId { get; private set; }

    public bool Result { get; private set; }
    #endregion


    public Login(string AppKey, string AuthenticationURL, string oauth2State)
    {
      InitializeComponent();
      DBAppKey = AppKey;
      DBAuthenticationURL = AuthenticationURL;
      DBoauth2State = oauth2State;
    }

    public void Navigate()
    {
      try
      {
        if (!string.IsNullOrEmpty(DBAppKey))
        {
          Uri authorizeUri = new Uri(DBAuthenticationURL);
          Browser.Source = authorizeUri;
        }
      }
      catch (Exception)
      {
        throw;
      }
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      Dispatcher.BeginInvoke(new Action(Navigate));
      // Navigate();  
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
      try
      {
        this.Close();
      }
      catch (Exception)
      {
        throw;
      }

    }

    private void Browser_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
    {
      if (!Browser.Source.ToString().StartsWith(RedirectUri.ToString(), StringComparison.OrdinalIgnoreCase))
      {
        // we need to ignore all navigation that isn't to the redirect uri.  
        return;
      }


      try
      {

        OAuth2Response result = DropboxOAuth2Helper.ParseTokenFragment(Browser.Source);
        if (result.State != DBoauth2State)
        {
          return;
        }

        this.AccessToken = result.AccessToken;
        this.Uid = result.Uid;
        this.Result = true;
      }

      catch (ArgumentException ex)
      {
      }

      finally
      {
        this.Close();
      }
    }
  }
}
