using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class GraphClientManager : IDisposable
    {
        private static AuthenticationResult AuthToken = null;
        private static HttpClient client = null;

        private async static Task<string> GetAuthTokenAsync()
        {
            if (AuthToken == null || AuthToken.ExpiresOn < DateTimeOffset.UtcNow)
            {
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(Settings.ClientId)
                .WithClientSecret(Settings.ClientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{Settings.TenantId}"))
                .Build();
                string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
                AuthToken = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            }

            return AuthToken.AccessToken;
        }

        public HttpClient GetGraphHttpClient()
        {
            if (client == null)
            {
                client = new HttpClient();
            }

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",  GetAuthTokenAsync().Result);
            return client;
        }

        public void Dispose()
        {
            client?.Dispose();
        }
    }
}
