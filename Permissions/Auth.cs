using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System;

namespace SitePermissions
{
    class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            SecretClientOptions options = new SecretClientOptions()
            {
                Retry =
                {
                    Delay = TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = Azure.Core.RetryMode.Exponential
                }
            };

            var client = new SecretClient(new System.Uri(Globals.keyVaultUrl), new DefaultAzureCredential(), options);
            KeyVaultSecret secret = client.GetSecret(Globals.secretName);
            var secretValue = secret.Value;

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(Globals.clientId)
            .WithTenantId(Globals.tenantId)
            .WithClientSecret(secretValue)
            .Build();

            var scopes = new string[] { "https://graph.microsoft.com/.default" };

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
