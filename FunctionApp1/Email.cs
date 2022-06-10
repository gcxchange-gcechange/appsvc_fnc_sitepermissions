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
        public static async Task<bool> SendMisconfiguredEmail(string SiteName, string Username, string UserEmail, ILogger log)
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

                        Hi { Username },

                        We've detected the site permissions for { SiteName } have been misconfigured. We're fixing them now. 

                        Yabadabadoo

                        --------------------------------------

                        (The English version precedes)

                        Bonjour { Username },

                        We've forgotten to write a french version of this email. Refer to the one above. 

                        Yabadabadoo"
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
