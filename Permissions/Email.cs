using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Threading.Tasks;

namespace SitePermissions
{
    public static class Email
    {
        // Returns true if the email was sent successfully 
        public static async Task<bool> SendMisconfiguredEmail(string Username, string UserEmail, ILogger log)
        {
            var res = true;
            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            try
            {
                var message = new Message
                {
                    Subject = "English Subject | French Subject",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = @$"
                        (La version française suit)

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

                        Nous vous remercions de l’intérêt que vous portez à la plateforme GCÉchange et de votre participation en tant que responsable de l’une de ses collectivités. 

                        Les collectivités demeurent la fonction la plus prisée de la plateforme, et nous sommes impatients de voir de nouvelles collectivités intéressantes se créer à mesure que la plateforme évolue. 

                        L’une des principales caractéristiques de GCÉchange est son modèle ouvert par défaut, qui permet une collaboration permanente entre les ministères. Nous encourageons tous les responsables à maintenir la pratique d’ouverture par défaut, comme indiqué dans nos conditions d’utilisation. 

                        L’équipe de GCÉchange a reçu un avis concernant la modification des autorisations relatives à la collectivité dont vous êtes responsable. Cette modification va à l’encontre du modèle ouvert par défaut de la plateforme. Veuillez noter que les paramètres par défaut seront rétablis, et nous vous prions de ne plus modifier les autorisations. 

                        N’hésitez pas à nous faire part de vos questions et de vos préoccupations. Vous êtes un membre important de GCÉchange et nous souhaitons réitérer notre gratitude.  
                         
                        Cordialement, 
                        L’équipe de GCÉchange"
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

                await graphAPIAuth.Users[Globals.emailSenderId]
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
