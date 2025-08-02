using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;

namespace Konverter.Extensions
{
  public static class MSALExtensions
  {
    public static PublicClientApplicationBuilder AddWindowsBroker(this PublicClientApplicationBuilder app)
    {
      return app.WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows));
    }
  }
}
