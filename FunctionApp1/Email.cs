using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Net.Mail;
using System.Threading.Tasks;

namespace SitePermissions
{
    public static class Email
    {
        private static readonly int smtp_port = Int16.Parse(Globals.smtp_port);
        private static readonly string smtp_link = Globals.GetSMTP_link();
        private static readonly string smtp_username = Globals.GetSMTP_username();
        private static readonly string smtp_password = Globals.GetSMTP_password();

        public static async Task<bool> SendMisconfiguredEmail(string SiteName, string Username, string UserEmail, ILogger log)
        {
            var result = false;
            string EmailSender = Globals.UserSender;

            var Body = @$"
                        (La version française suit)<br><br>
                        Hi { Username },<br><br>
                        We've detected the site permissions for { SiteName } have been misconfigured. We're fixing them now. Yabadabadoo<br>
                        <hr/>
                        (The English version precedes)<br><br>
                        Bonjour { Username },<br><br>
                        We've forgotten to write a french version of this email. Refer to the one above. Yabadabadoo<br>";

            MailMessage mail = new MailMessage();

            mail.From = new MailAddress(EmailSender);
            mail.To.Add(UserEmail);
            mail.Subject = "English Subject | French Subject";
            mail.Body = Body;
            mail.IsBodyHtml = true;

            SmtpClient SmtpServer = new SmtpClient(smtp_link);
            SmtpServer.Port = smtp_port;
            SmtpServer.Credentials = new System.Net.NetworkCredential(smtp_username, smtp_password);
            SmtpServer.EnableSsl = true;

            log.LogInformation($"UserEmail : {UserEmail}");

            try
            {
                SmtpServer.Send(mail);
                log.LogInformation("mail sent");
                result = true;
            }
            catch (ServiceException ex)
            {
                log.LogInformation($"Error sending email for {SiteName}: {ex.Message}");
                result = false;
            }

            return result;
        }
    }
}
