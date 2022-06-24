using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.Graph;
using PnP.Framework;
using Microsoft.SharePoint.Client;
using Site = Microsoft.Graph.Site;
using PnP.Framework.Entities;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.Linq;

namespace SitePermissions
{
    public static class Permissions
    {
        [FunctionName("HandleMisconfigured")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log, ExecutionContext executionContext)
        {
            log.LogInformation($"Site permissions function executed at: {DateTime.Now}");

            var allPermissionLevels = new List<string>() { PermissionLevel.Read, PermissionLevel.Contribute, PermissionLevel.Design, PermissionLevel.Edit, PermissionLevel.FullControl };
            var misconfiguredSites = new List<Site>();
            var reports = new List<Report>();

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

            var excludeSiteIds = Globals.GetExcludedSiteIds();

            do
            {
                foreach (var site in allSites)
                {
                    var siteId = site.Id.Split(",")[1];

                    if (excludeSiteIds.Contains(siteId))
                        continue;

                    var misconfigured = false;

                    var ctx = auth.appOnlyAuth(site.WebUrl, log);

                    // Create a report of the site before we make any changes.
                    reports.Add(new Report(site, ctx));

                    var readGroups = new List<Globals.Group>();
                    var editGroups = new List<Globals.Group>();
                    var fullControlGroups = new List<Globals.Group>();
                    var designGroups = new List<Globals.Group>();
                    var contributeGroups = new List<Globals.Group>();
                    var siteCollectionAdminGroups = new List<Globals.Group>();

                    // Validate the default role definitions (Read, Edit, Full Control) have the required base permissions
                    misconfigured = !await ValidateRoleDefinitions(ctx, log);

                    // Go through each group defined in local.settings.json
                    foreach (var group in Globals.groups)
                    {
                        var hasRead = await HasPermissionLevel(new Globals.Group(group.GroupName, group.Id, PermissionLevel.Read), ctx, log);
                        var hasEdit = await HasPermissionLevel(new Globals.Group(group.GroupName, group.Id, PermissionLevel.Edit), ctx, log);
                        var hasFullControl = await HasPermissionLevel(new Globals.Group(group.GroupName, group.Id, PermissionLevel.FullControl), ctx, log);
                        
                        try
                        {
                            switch (group.PermissionLevel)
                            {
                                case PermissionLevel.Read:

                                    if (!hasRead || hasEdit || hasFullControl)
                                    {
                                        await RemovePermissionLevels(group, allPermissionLevels, ctx, log);
                                        await GrantPermissionLevel(group, group.PermissionLevel, ctx, log);

                                        misconfigured = true;
                                        log.LogWarning($"{group.GroupName} didn't pass {PermissionLevel.Read} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.GroupName} passed {PermissionLevel.Read} check");
                                    }

                                    readGroups.Add(group);

                                    break;

                                case PermissionLevel.Edit:

                                    if (!hasEdit || hasRead || hasFullControl)
                                    {
                                        await RemovePermissionLevels(group, allPermissionLevels, ctx, log);
                                        await GrantPermissionLevel(group, group.PermissionLevel, ctx, log);

                                        misconfigured = true;

                                        log.LogWarning($"{group.GroupName} didn't pass {PermissionLevel.Edit} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.GroupName} passed {PermissionLevel.Edit} check");
                                    }

                                    editGroups.Add(group);

                                    break;

                                case PermissionLevel.FullControl:

                                    if (!hasFullControl || hasRead || hasEdit)
                                    {
                                        await RemovePermissionLevels(group, allPermissionLevels, ctx, log);
                                        await GrantPermissionLevel(group, group.PermissionLevel, ctx, log);

                                        misconfigured = true;

                                        log.LogWarning($"{group.GroupName} didn't pass {PermissionLevel.FullControl} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.GroupName} passed {PermissionLevel.FullControl} check");
                                    }

                                    fullControlGroups.Add(group);

                                    break;

                                case PermissionLevel.SiteCollectionAdministrator:

                                    if (!await IsSiteCollectionAdministrator(group, ctx, log))
                                    {
                                        misconfigured = true;

                                        log.LogWarning($"{group.GroupName} didn't pass {PermissionLevel.SiteCollectionAdministrator} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.GroupName} passed {PermissionLevel.SiteCollectionAdministrator} check");
                                    }

                                    siteCollectionAdminGroups.Add(group);

                                    break;

                                default:

                                    log.LogError($"Error parsing group permission level - {group.PermissionLevel}");

                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogError($"Error adding {group.GroupName} to {site.WebUrl} - {ex.Source}: {ex.Message} | {ex.InnerException}");
                        }
                    }

                    await RemoveSiteCollectionAdministrators(siteCollectionAdminGroups, ctx, log);
                    if (misconfigured)
                    {
                        foreach (var group in siteCollectionAdminGroups)
                        {
                            await AddSiteCollectionAdministrator(group, ctx, log);
                        }
                    }

                    var expectedRead = await RemoveUnknownPermissionLevels(readGroups, PermissionLevel.Read, ctx, log);
                    var expectedEdit = await RemoveUnknownPermissionLevels(editGroups, PermissionLevel.Edit, ctx, log);
                    var expectedFullControl = await RemoveUnknownPermissionLevels(fullControlGroups, PermissionLevel.FullControl, ctx, log);
                    var expectedContribute = await RemoveUnknownPermissionLevels(contributeGroups, PermissionLevel.Contribute, ctx, log);
                    var expectedDesign = await RemoveUnknownPermissionLevels(designGroups, PermissionLevel.Design, ctx, log);

                    var expectedGroups = await CleanseGroups(siteCollectionAdminGroups, ctx, log);

                    misconfigured = !misconfigured ? !(expectedRead && expectedEdit && expectedFullControl && expectedGroups && expectedContribute && expectedDesign) : misconfigured;

                    if (misconfigured)
                    {
                        misconfiguredSites.Add(site);

                        log.LogWarning($"Found misconfigured site: {site.Name} - {site.WebUrl}");
                    }
                }
            }
            while (allSites.NextPageRequest != null && (allSites = await allSites.NextPageRequest.GetAsync()).Count > 0);

            await StoreData.StoreReports(executionContext, reports, "reports", log);

            await InformOwners(misconfiguredSites, graphAPIAuth, log);

            return new OkObjectResult(misconfiguredSites);
        }

        // Returns true if the group has the expected permission level
        private static async Task<bool> HasPermissionLevel(Globals.Group group, ClientContext ctx, ILogger log)
        {
            var result = false;

            try
            {
                var roleAssignments = ctx.Web.RoleAssignments;

                ctx.Load(roleAssignments, r => r.Include(i => i.Member, i => i.RoleDefinitionBindings));
                ctx.ExecuteQuery();

                for (var i = 0; i < roleAssignments.Count && !result; i++)
                {
                    var ra = roleAssignments[i];
                    
                    if (ra.Member is Microsoft.SharePoint.Client.User && GetObjectId(((Microsoft.SharePoint.Client.User)ra.Member).LoginName) == group.Id)
                    {
                        foreach (var role in ra.RoleDefinitionBindings)
                        {
                            if (role.Name == group.PermissionLevel)
                            {
                                result = true;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error validating {group.GroupName} has {group.PermissionLevel}: {ex.Message}");
            }

            return result;
        }

        // Goes through site permissions and removes any that are not in the expectedGroups for the permissionLevel
        // Returns false if any were found and removed.
        private static async Task<bool> RemoveUnknownPermissionLevels(List<Globals.Group> expectedGroups, string permissionLevel, ClientContext ctx, ILogger log)
        {
            var result = true;

            try
            {
                var roleAssignments = ctx.Web.RoleAssignments;

                ctx.Load(roleAssignments, r => r.Include(i => i.Member, i => i.RoleDefinitionBindings));
                ctx.ExecuteQuery();

                for (var i = 0; i < roleAssignments.Count; i++)
                {
                    var ra = roleAssignments[i];

                    if (ra.Member is Microsoft.SharePoint.Client.User)
                    {
                        if (expectedGroups.Count > 0)
                        {
                            foreach (var group in expectedGroups)
                            {
                                if (GetObjectId(((Microsoft.SharePoint.Client.User)ra.Member).LoginName) != group.Id)
                                {
                                    result = !await RemoveAllSpecificPermissionLevel(ra, permissionLevel, ctx, log) == false && result ? false : result;
                                }
                            }
                        }
                        else
                        {
                            result = !await RemoveAllSpecificPermissionLevel(ra, permissionLevel, ctx, log) == false && result ? false : result;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message}");
            }

            return result;
        }

        // This function will remove any memebers of the site's sharepoint groups that aren't supposed to be there by default
        // Returns false if any were found and removed.
        private static async Task<bool> CleanseGroups(List<Globals.Group> approvedAdminGroups, ClientContext ctx, ILogger log)
        {
            var result = true;

            try
            {
                var roleAssignments = ctx.Web.RoleAssignments;

                ctx.Load(roleAssignments, r => r.Include(i => i.Member, i => i.RoleDefinitionBindings));
                ctx.ExecuteQuery();

                for (var i = 0; i < roleAssignments.Count; i++)
                {
                    var ra = roleAssignments[i];

                    var isOwnersGroup = ra.Member.Title == ctx.Web.Title + " Owners";
                    var isMembersGroup = ra.Member.Title == ctx.Web.Title + " Members";
                    var isVisitorsGroup = ra.Member.Title == ctx.Web.Title + " Visitors";

                    // Only look through the Owners, Members and Visitors SharePoint groups.
                    if (ra.Member is Microsoft.SharePoint.Client.Group && (isOwnersGroup || isMembersGroup || isVisitorsGroup))
                    {
                        var oGroup = ctx.Web.SiteGroups.GetByName(ra.Member.LoginName);

                        var oUserCollection = oGroup.Users;

                        ctx.Load(oUserCollection);
                        ctx.ExecuteQuery();

                        foreach (var user in oUserCollection)
                        {
                            var isMembersUser = user.Title == ctx.Web.Title + " Members";
                            var isAdmin = approvedAdminGroups.Any(x => x.GroupName == user.Title);

                            // Ignore system accounts, approved admins in the owners group, and members in the member group
                            if (user.Title == "System Account" || (isOwnersGroup && isAdmin) || (isMembersGroup && isMembersUser))
                                continue;

                            oGroup.Users.RemoveByLoginName(user.LoginName);
                            ctx.ExecuteQuery();

                            result = false;

                            log.LogWarning($"Removing {user.Title} from {ra.Member.Title}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"{ex.Message}");
            }

            return result;
        }

        // Adds the group to the sites permission list at the level defined in group.PermissionLevel
        private static async Task<bool> GrantPermissionLevel(Globals.Group group, string permissionLevel, ClientContext ctx, ILogger log)
        {
            var result = true;

            try
            {

                // TODO: Figure out why we get an unauthorized access error when trying to ensure by Id
                //var adGroup = ctx.Web.EnsureUserByObjectId(Guid.Parse(group.Id), Guid.Parse(Globals.tenantId), Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup);
                var adGroup = ctx.Web.EnsureUser(group.GroupName);
                ctx.Load(adGroup);
                var spGroup = ctx.Web.AssociatedMemberGroup;
                spGroup.Users.AddUser(adGroup);

                var writeDefinition = ctx.Web.RoleDefinitions.GetByName(permissionLevel);
                var roleDefCollection = new RoleDefinitionBindingCollection(ctx);
                roleDefCollection.Add(writeDefinition);
                var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                ctx.Load(spGroup, x => x.Users);
                ctx.ExecuteQuery();

                log.LogInformation($"Gave {group.GroupName} {permissionLevel} on {ctx.Site.Url}");
            }
            catch (Exception ex)
            {
                log.LogError($"Error adding {group.GroupName} to {ctx.Site.Url} - {ex.Source}: {ex.Message} | {ex.InnerException}");
            }

            return result;
        }

        // Removes all role definition bindings for the group. Returns true if it was successful.
        private static async Task<bool> RemovePermissionLevels(Globals.Group group, List<string> levelsToRemove, ClientContext ctx, ILogger log)
        {
            var result = false;

            try
            {
                var roleAssignments = ctx.Web.RoleAssignments;

                ctx.Load(roleAssignments, r => r.Include(i => i.Member, i => i.RoleDefinitionBindings));
                ctx.ExecuteQuery();

                for (var i = 0; i < roleAssignments.Count; i++)
                {
                    var ra = roleAssignments[i];
                    if (ra.Member is Microsoft.SharePoint.Client.User && GetObjectId(((Microsoft.SharePoint.Client.User)ra.Member).LoginName) == group.Id)
                    {
                        foreach (var role in ra.RoleDefinitionBindings.ToArray())
                        {
                            if (levelsToRemove.Any(x => x.Equals(role.Name)))
                            {
                                ra.RoleDefinitionBindings.Remove(role);
                                result = true;

                                log.LogWarning($"Removing {role.Name} from {((Microsoft.SharePoint.Client.User)ra.Member).Title}");
                            }
                        }

                        ra.Update();
                    }
                }

                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.LogError($"Error removing {group.GroupName} from {ctx.Site.Url} - {ex.Source}: {ex.Message} | {ex.InnerException}");
            }

            return result;
        }

        // Adds the group to the site collection administrator list
        private static async Task<bool> AddSiteCollectionAdministrator(Globals.Group group, ClientContext ctx, ILogger log)
        {
            var result = true;

            try
            {
                ctx.Load(ctx.Web);
                ctx.Load(ctx.Site);
                ctx.Load(ctx.Site.RootWeb);
                ctx.ExecuteQuery();

                // TODO: Figure out how to use ID instead of groupName 

                List<UserEntity> admins = new List<UserEntity>();
                UserEntity adminUserEntity = new UserEntity();

                adminUserEntity.LoginName = group.GroupName;
                admins.Add(adminUserEntity);

                if (admins.Count > 0)
                {
                    ctx.Site.RootWeb.AddAdministrators(admins, true);
                }

                log.LogInformation($"Added {group.GroupName} as Site Collection Administrators to {ctx.Site.Url}");
            }
            catch (Exception ex)
            {
                log.LogError($"Error addming site collection admin to {ctx.Site.Url}: {ex.Message}");
            }
            

            return result;
        }

        // Removes all site collection administrators.
        private static async Task<IActionResult> RemoveSiteCollectionAdministrators(List<Globals.Group> approvedAdminGroups, ClientContext ctx, ILogger log)
        {
            var removedUsers = new List<Microsoft.SharePoint.Client.User>();

            ctx.Load(ctx.Web);
            ctx.Load(ctx.Site);
            ctx.Load(ctx.Site.RootWeb);
            ctx.ExecuteQuery();

            var users = ctx.Site.RootWeb.SiteUsers;
            ctx.Load(users);
            ctx.ExecuteQuery();

            foreach (var user in users)
            {
                var isOwnersGroup = user.Title == ctx.Web.Title + " Owners";
                var isAdmin = approvedAdminGroups.Any(x => x.GroupName == user.Title);

                if (user.IsSiteAdmin && !isOwnersGroup && !isAdmin)
                {
                    try
                    {
                        user.IsSiteAdmin = false;
                        user.Update();
                        ctx.Load(user);
                        ctx.ExecuteQuery();

                        removedUsers.Add(user);

                        log.LogWarning($"Removed {user.Title} from Site Collection Administrators for {ctx.Site.Url}");
                    }
                    catch (Exception ex)
                    {
                        log.LogError($"Error removing {user.Title} from site collection administrator: {ex.Message}");
                    }                        
                }
            }

            return new OkObjectResult(removedUsers);
        }

        // Returns true if the group is found in the site collections administrator list
        private static async Task<bool> IsSiteCollectionAdministrator(Globals.Group group, ClientContext ctx, ILogger log)
        {
            try
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
                    if (user.IsSiteAdmin && GetObjectId(user.LoginName) == group.Id)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error verifying site collection administrator for {group.GroupName}: {ex.Message}");
            }

            return false;
        }

        // Goes through the Read, Edit, and Full Control role definitions of the site to determine if they have been changed 
        // If any were changed it will change them back to the default (defined in PermissionLevel.cs)
        private static async Task<bool> ValidateRoleDefinitions(ClientContext ctx, ILogger log)
        {
            var isValid = true;
            
            try
            {
                var readRoleDef = ctx.Web.RoleDefinitions.GetByName(PermissionLevel.Read);
                ctx.Load(readRoleDef);
                ctx.ExecuteQuery();

                if (!PermissionLevel.HasRead(readRoleDef.BasePermissions))
                {
                    var newPermissions = new BasePermissions();

                    foreach (var perm in PermissionLevel.ReadPermissions)
                    {
                        newPermissions.Set(perm);
                    }

                    readRoleDef.BasePermissions = newPermissions;

                    readRoleDef.Update();
                    ctx.Load(readRoleDef);
                    ctx.ExecuteQuery();

                    isValid = false;

                    log.LogWarning($"{PermissionLevel.Read} permission level definition is invalid");
                }
                else
                {
                    log.LogInformation($"{PermissionLevel.Read} permission level definition is valid");
                }

                var editRoleDef = ctx.Web.RoleDefinitions.GetByName(PermissionLevel.Edit);
                ctx.Load(editRoleDef);
                ctx.ExecuteQuery();

                if (!PermissionLevel.HasEdit(editRoleDef.BasePermissions))
                {
                    var newPermissions = new BasePermissions();

                    foreach (var perm in PermissionLevel.EditPermissions)
                    {
                        newPermissions.Set(perm);
                    }

                    editRoleDef.BasePermissions = newPermissions;

                    editRoleDef.Update();
                    ctx.Load(editRoleDef);
                    ctx.ExecuteQuery();

                    isValid = false;

                    log.LogWarning($"{PermissionLevel.Edit} permission level definition is invalid");
                }
                else
                {
                    log.LogInformation($"{PermissionLevel.Edit} permission level definition is valid");
                }

                var fullControlRoleDef = ctx.Web.RoleDefinitions.GetByName(PermissionLevel.FullControl);
                ctx.Load(fullControlRoleDef);
                ctx.ExecuteQuery();

                if (!PermissionLevel.HasFullControl(fullControlRoleDef.BasePermissions))
                {
                    var newPermissions = new BasePermissions();

                    foreach (var perm in PermissionLevel.FullControlPermissions)
                    {
                        newPermissions.Set(perm);
                    }

                    fullControlRoleDef.BasePermissions = newPermissions;

                    fullControlRoleDef.Update();
                    ctx.Load(fullControlRoleDef);
                    ctx.ExecuteQuery();

                    isValid = false;

                    log.LogWarning($"{PermissionLevel.FullControl} permission level definition is invalid");
                }
                else 
                { 
                    log.LogInformation($"{PermissionLevel.FullControl} permission level definition is valid"); 
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error while validating role definitions: {ex}");
            }

            return isValid;
        }

        // Go through all the sites, find the owner emails, and inform them their site settings have changed.
        private static async Task<List<Tuple<Microsoft.Graph.User, bool>>> InformOwners(ICollection<Site> sites, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var results = new List<Tuple<Microsoft.Graph.User, bool>>();

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
                                    var result = await Email.SendMisconfiguredEmail(user.DisplayName, user.Mail, log);
                                    results.Add(new Tuple<Microsoft.Graph.User, bool>(user, result));
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

        public static async Task<bool> RemoveAllSpecificPermissionLevel(Microsoft.SharePoint.Client.RoleAssignment ra, string permissionLevel, ClientContext ctx, ILogger log)
        {
            var result = false;

            foreach (var role in ra.RoleDefinitionBindings)
            {
                if (role.Name == permissionLevel)
                {
                    result = await RemovePermissionLevels(new Globals.Group(((Microsoft.SharePoint.Client.User)ra.Member).Title, GetObjectId(((Microsoft.SharePoint.Client.User)ra.Member).LoginName), permissionLevel), new List<string>() { permissionLevel }, ctx, log);
                    break;
                }
            }

            return result;
        }

        public static string GetObjectId(string loginName)
        {
            var split = loginName.Split('|');
            return split.Length == 3 ? split[2] : "-1";
        } 
    }
}
