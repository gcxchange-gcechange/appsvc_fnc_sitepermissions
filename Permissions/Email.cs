using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static System.Formats.Asn1.AsnWriter;

namespace SitePermissions
{
    public static class Email
    {
        // Go through all the sites, find the owner emails, and inform them their site settings have changed.
        public static async Task<List<Tuple<Microsoft.Graph.User, bool>>> InformOwners(ICollection<Site> sites, GraphServiceClient graphAPIAuth,  ILogger log)
        {
            var results = new List<Tuple<Microsoft.Graph.User, bool>>();
            var siteOwners = new List<Microsoft.Graph.User>();
            foreach (var site in sites)
            {
                var groupQueryOptions = new List<QueryOption>()
                {
                    new QueryOption("$search", "\"mailNickname:" + site.Name +"\"")
                };

                var groups = await graphAPIAuth.Groups
                .Request(groupQueryOptions)
                .Header("ConsistencyLevel", "eventual")
                .GetAsync();

                do
                {
                    foreach (var group in groups)
                    {
                        var owners = await graphAPIAuth.Groups[group.Id].Owners
                        .Request()
                        .GetAsync();

                        do
                        {
                            foreach (var owner in owners)
                            {
                                var user = await graphAPIAuth.Users[owner.Id]
                                .Request()
                                .Select("displayName,mail")
                                .GetAsync();

                                if (user != null)
                                {
                                    siteOwners.Add(user);
                                }
                            }
                        }
                        while (owners.NextPageRequest != null && (owners = await owners.NextPageRequest.GetAsync()).Count > 0);
                    }
                }
                while (groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0);
            }

            foreach (var owner in siteOwners.Distinct())
            {
                var result = await SendMisconfiguredEmail(owner.DisplayName, owner.Mail, log);
                results.Add(new Tuple<Microsoft.Graph.User, bool>(owner, result));
            }
            return results;
        }

        // Returns true if the email was sent successfully 
        public static async Task<bool> SendMisconfiguredEmail(string Username, string UserEmail, ILogger log)
        {
            var res = true;
            var scopes = new[] { "user.read mail.send" };
            ROPCConfidentialTokenCredential authdelegated = new ROPCConfidentialTokenCredential();
            var graphClient_delegated = new GraphServiceClient(authdelegated, scopes);
 
            try
            {
                var message = new Message
                {
                    Subject = "An important message from GCXchange | Un message important de GCÉchange",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = @$"(La version française suit)

Dear { Username }, 

Thank you for your interest and participation as a Community Owner on the GCXchange platform. 

Communities continue to be our most popular feature of the platform, and we look forward to seeing new and interesting communities appear as the platform grows.

One of the key features of GCXchange is the open by default model which allows for seamless cross-departmental collaboration. We encourage all site owners to maintain the open by default practice as outlined in our Terms of Use. 

The GCXchange Team received a notification that the permissions were modified to the Community space in which you are Owner. This action goes against the open by default model of the platform. Please note that the default settings will now be reapplied, and we would like to give you a friendly reminder to please no longer change these permissions. 

Please let us know if you have any questions or concerns, and once again thank you for being a valued member of GCXchange.  
 
Regards, 
The GCX Team 

--------------------------------------

Bonjour { Username }, 

Nous vous remercions de l’intérêt que vous manifestez à l’égard de la plateforme GCÉchange, de même que de votre participation en tant que responsable de l’une de ses collectivités. 

Les collectivités demeurent la fonction la plus prisée de la plateforme, et nous sommes impatients de voir de nouvelles collectivités intéressantes se créer, au fur et à mesure que la plateforme évolue. 

L’une des principales caractéristiques de GCÉchange est son modèle ouvert par défaut, qui permet une collaboration permanente entre les ministères. Nous encourageons tous les responsables à maintenir la pratique d’ouverture par défaut, comme indiqué dans nos conditions d’utilisation. 

L’équipe de GCÉchange a reçu un avis concernant la modification des autorisations relatives à la collectivité dont vous êtes responsable. Cette modification va à l’encontre du modèle ouvert par défaut de la plateforme. Veuillez prendre note que les paramètres par défaut seront rétablis, et nous vous prions de ne plus modifier les autorisations. 

N’hésitez pas à nous faire part de vos questions et de vos préoccupations. Vous êtes un membre important de GCÉchange et nous souhaitons réitérer notre gratitude.  
 
Nous vous prions d’agréer l’expression de nos sentiments les meilleurs. 
Équipe de GCÉchange"
                    },
                    ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = UserEmail
                            }
                        }
                    }
                };

                await graphClient_delegated.Users[Globals.emailSenderId]
                .SendMail(message, null)
                .Request()
                .PostAsync();

                log.LogInformation($"Email sent to {UserEmail}");
            }
            catch (Exception ex)
            {
                log.LogError($"Error sending email to {UserEmail}: {ex.Message}");
                res = false;
            }

            return res;
        }
    }
}
