using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;

namespace SitePermissions
{
    class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            var scopes = new string[] { "https://graph.microsoft.com/.default" };

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(Globals.clientId)
            .WithTenantId(Globals.tenantId)
            .WithClientSecret(Globals.clientSecret)
            .Build();

            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
            {
                var authResult = await confidentialClientApplication
                .AcquireTokenForClient(scopes)
                .ExecuteAsync();
                
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            }));

            return graphServiceClient;
        }
    }
}
