using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.Graph;
using System.Net.Mail;
using PnP.Framework;

namespace SitePermissions
{
    public static class Permissions
    {
        [FunctionName("GetMisconfigured")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            List<Site> misconfiguredSites = new List<Site>();

            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var sitesQueryOptions = new List<QueryOption>()
            {
                new QueryOption("search", "DepartmentId:{" + Globals.hubId + "}"),
            };

            var allSites = await graphAPIAuth.Sites
            .Request(sitesQueryOptions)
            .Header("ConsistencyLevel", "eventual")
            .GetAsync();

            do
            {
                foreach (var site in allSites)
                {
                    var permissions = await graphAPIAuth.Sites[site.Id].Permissions
                    .Request()
                    .Header("ConsistencyLevel", "eventual")
                    .GetAsync();

                    var misconfigured = false;

                    var ctx = new AuthenticationManager().GetACSAppOnlyContext(site.WebUrl, Globals.appOnlyId, Globals.appOnlySecret, AzureEnvironment.Production);
                    var web = ctx.Web;

                    foreach(var group in Globals.groups)
                    {
                        try
                        {
                            var adGroup = ctx.Web.EnsureUserByObjectId(Guid.Parse(group.GroupId), Guid.Parse(Globals.tenantId), Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup);
                            ctx.Load(adGroup);
                            var spGroup = ctx.Web.AssociatedMemberGroup;
                            spGroup.Users.AddUser(adGroup);

                            var writeDefinition = ctx.Web.RoleDefinitions.GetByName(group.PermissionLevel);
                            var roleDefCollection = new Microsoft.SharePoint.Client.RoleDefinitionBindingCollection(ctx);
                            roleDefCollection.Add(writeDefinition);
                            var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                            ctx.Load(spGroup, x => x.Users);
                            ctx.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            log.LogInformation($"Error adding {group.GroupName} to {site.WebUrl} - {ex.Message}");
                        }
                    }


                    do
                    {
                        foreach (var permission in permissions)
                        {
                            // TODO: Look for misconfig here
                            var p = permission;
                        }
                    }
                    while (permissions.NextPageRequest != null && (permissions = await permissions.NextPageRequest.GetAsync()).Count > 0);

                    if (misconfigured)
                    {
                        misconfiguredSites.Add(site);
                    }
                }
            }
            while (allSites.NextPageRequest != null && (allSites = await allSites.NextPageRequest.GetAsync()).Count > 0);

            var res = await InformOwners(misconfiguredSites, graphAPIAuth, log);

            return new OkObjectResult(misconfiguredSites);
        }

        public static async Task<List<Tuple<User, bool>>> InformOwners(ICollection<Site> sites, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var results = new List<Tuple<User, bool>>();

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
                                    var result = await SendEmail(site.DisplayName, user.DisplayName, user.Mail, log);
                                    results.Add(new Tuple<User, bool>(user, result));
                                }
                            }
                        }
                        while (owners.NextPageRequest != null && (owners = await owners.NextPageRequest.GetAsync()).Count > 0);
                    }
                }
                while (groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0);
            }

            return results;
        }

        public static async Task<bool> SendEmail(string SiteName, string Username, string UserEmail, ILogger log)
        {
            var result = false;
            string EmailSender = Globals.UserSender;
            int smtp_port = Int16.Parse(Globals.smtp_port);
            string smtp_link = Globals.GetSMTP_link();
            string smtp_username = Globals.GetSMTP_username();
            string smtp_password = Globals.GetSMTP_password();

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
                log.LogInformation($"Error sending email: {ex.Message}");
                result = false;
            }

            return result;
        }
    }
}
