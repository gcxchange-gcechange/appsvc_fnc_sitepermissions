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
        [FunctionName("HandleMisconfigured")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            List<Site> misconfiguredSites = new List<Site>();

            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            // Get subsites
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
                    var ctx = new AuthenticationManager().GetACSAppOnlyContext(site.WebUrl, Globals.appOnlyId, Globals.appOnlySecret);
                    bool misconfigured = false;

                    // Go through each group defined in local.settings.json
                    foreach (var group in Globals.groups)
                    {
                        try
                        {
                            var actvdirGroup = ctx.Web.EnsureUser(group.GroupName);
                            ctx.Load(actvdirGroup);
                            ctx.ExecuteQuery();

                            var permissions = ctx.Web.GetUserEffectivePermissions(actvdirGroup.LoginName);
                            ctx.ExecuteQuery();

                            // https://pnp.github.io/pnpcore/api/PnP.Core.Model.SharePoint.PermissionKind.html
                            switch (group.PermissionLevel)
                            {
                                case "Read":

                                    if (!permissions.Value.Has(Microsoft.SharePoint.Client.PermissionKind.ViewPages))
                                    {
                                        await AddGroup(ctx, group, log);

                                        misconfigured = true;
                                    }

                                    break;

                                case "Edit":

                                    if (!permissions.Value.Has(Microsoft.SharePoint.Client.PermissionKind.EditListItems))
                                    { 
                                        await AddGroup(ctx, group, log);

                                        misconfigured = true;
                                    }

                                    break;

                                case "Full Control":

                                    if (!permissions.Value.Has(Microsoft.SharePoint.Client.PermissionKind.ManagePermissions))
                                    {
                                        await AddGroup(ctx, group, log);

                                        misconfigured = true;
                                    }
                                    break;

                                case "Site Collection Administrator":

                                    if (!permissions.Value.Has(Microsoft.SharePoint.Client.PermissionKind.ManageWeb))
                                    {
                                        await RemoveSiteCollectionAdministrator(ctx, log);
                                        await AddSiteCollectionAdministrator(group, ctx, log);

                                        misconfigured = true;
                                    }

                                    break;

                                default:

                                    log.LogInformation($"Error parsing group permission level - {group.PermissionLevel}");

                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogInformation($"Error adding {group.GroupName} to {site.WebUrl} - {ex.Source}: {ex.Message} | {ex.InnerException}");
                        }
                    }

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

        public static async Task<bool> AddGroup(Microsoft.SharePoint.Client.ClientContext ctx, Globals.Group group, ILogger log)
        {
            var result = true;

            try
            {
                var adGroup = ctx.Web.EnsureUser(group.GroupName);
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
                log.LogInformation($"Error adding {group.GroupName} to {ctx.Site.Url} - {ex.Source}: {ex.Message} | {ex.InnerException}");
            }

            return result;
        }

        private static async Task<bool> AddSiteCollectionAdministrator(Globals.Group group, Microsoft.SharePoint.Client.ClientContext ctx, ILogger log)
        {
            var result = true;

            // TODO

            return result;
        }

        private static async Task<IActionResult> RemoveSiteCollectionAdministrator(Microsoft.SharePoint.Client.ClientContext ctx, ILogger log)
        {
            ctx.Load(ctx.Web);
            ctx.Load(ctx.Site);
            ctx.Load(ctx.Site.RootWeb);
            ctx.ExecuteQuery();

            var users = ctx.Site.RootWeb.SiteUsers;
            ctx.Load(users);
            ctx.ExecuteQuery();

            foreach (var user in users)
            {
                user.IsSiteAdmin = false;
                user.Update();
                ctx.Load(user);
                ctx.ExecuteQuery();

                log.LogInformation($"Removed {user.UserPrincipalName} as Site Collection Administrators from {ctx.Site.Url}");
            }

            return new OkObjectResult(users);
        }

        private static async Task<List<Tuple<User, bool>>> InformOwners(ICollection<Site> sites, GraphServiceClient graphAPIAuth, ILogger log)
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

        private static async Task<bool> SendEmail(string SiteName, string Username, string UserEmail, ILogger log)
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
                log.LogInformation($"Error sending email for {SiteName}: {ex.Message}");
                result = false;
            }

            return result;
        }
    }
}
