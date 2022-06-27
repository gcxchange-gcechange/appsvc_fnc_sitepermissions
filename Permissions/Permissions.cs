using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Site = Microsoft.Graph.Site;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Extensions.Http;

namespace SitePermissions
{
    public static class Permissions
    {
        [FunctionName("HandleMisconfigured")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log, ExecutionContext executionContext)
        {
            log.LogInformation($"Site permissions function executed at: {DateTime.Now}");

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

                    var readGroups = new List<Group>();
                    var editGroups = new List<Group>();
                    var fullControlGroups = new List<Group>();
                    var designGroups = new List<Group>();
                    var contributeGroups = new List<Group>();
                    var siteCollectionAdminGroups = new List<Group>();

                    // Validate the default role definitions (Read, Edit, Full Control) have the required base permissions
                    misconfigured = await ValidateRoleDefinitions(ctx, log) == false;

                    // Go through each group defined in local.settings.json
                    foreach (var group in Globals.groups)
                    {
                        var hasRead = await group.HasPermissionLevel(PermissionLevel.Read, ctx, log);
                        var hasEdit = await group.HasPermissionLevel(PermissionLevel.Edit, ctx, log);
                        var hasFullControl = await group.HasPermissionLevel(PermissionLevel.FullControl, ctx, log);

                        try
                        {
                            switch (group.AssignedPermissionLevel)
                            {
                                case PermissionLevel.Read:

                                    if (!hasRead || hasEdit || hasFullControl)
                                    {
                                        await group.RemovePermissionLevels(new List<string>() { PermissionLevel.Contribute, PermissionLevel.Design, PermissionLevel.Edit, PermissionLevel.FullControl }, ctx, log);
                                        await group.AddPermissionLevel(group.AssignedPermissionLevel, ctx, log);

                                        misconfigured = true;
                                        log.LogWarning($"{group.Name} didn't pass {PermissionLevel.Read} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.Name} passed {PermissionLevel.Read} check");
                                    }

                                    readGroups.Add(group);

                                    break;

                                case PermissionLevel.Edit:

                                    if (!hasEdit || hasRead || hasFullControl)
                                    {
                                        await group.RemovePermissionLevels(new List<string>() { PermissionLevel.Read, PermissionLevel.Contribute, PermissionLevel.Design, PermissionLevel.FullControl }, ctx, log);
                                        await group.AddPermissionLevel(group.AssignedPermissionLevel, ctx, log);

                                        misconfigured = true;

                                        log.LogWarning($"{group.Name} didn't pass {PermissionLevel.Edit} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.Name} passed {PermissionLevel.Edit} check");
                                    }

                                    editGroups.Add(group);

                                    break;

                                case PermissionLevel.FullControl:

                                    if (!hasFullControl || hasRead || hasEdit)
                                    {
                                        await group.RemovePermissionLevels(new List<string>() { PermissionLevel.Read, PermissionLevel.Contribute, PermissionLevel.Design, PermissionLevel.Edit }, ctx, log);
                                        await group.AddPermissionLevel(group.AssignedPermissionLevel, ctx, log);

                                        misconfigured = true;

                                        log.LogWarning($"{group.Name} didn't pass {PermissionLevel.FullControl} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.Name} passed {PermissionLevel.FullControl} check");
                                    }

                                    fullControlGroups.Add(group);

                                    break;

                                case PermissionLevel.SiteCollectionAdministrator:

                                    if (!await group.IsSiteCollectionAdministrator(ctx, log))
                                    {
                                        misconfigured = true;

                                        log.LogWarning($"{group.Name} didn't pass {PermissionLevel.SiteCollectionAdministrator} check");
                                    }
                                    else
                                    {
                                        log.LogInformation($"{group.Name} passed {PermissionLevel.SiteCollectionAdministrator} check");
                                    }

                                    siteCollectionAdminGroups.Add(group);

                                    break;

                                default:

                                    log.LogError($"Error parsing group permission level - {group.AssignedPermissionLevel}");

                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            log.LogError($"Error adding {group.Name} to {site.WebUrl} - {ex.Source}: {ex.Message} | {ex.InnerException}");
                        }
                    }

                    misconfigured = !await Group.Helpers.RemoveSiteCollectionAdministrators(siteCollectionAdminGroups, ctx, log) ? true : misconfigured;

                    if (misconfigured)
                    {
                        foreach (var group in siteCollectionAdminGroups)
                        {
                            await Group.Helpers.AddSiteCollectionAdministrator(group, ctx, log);
                        }
                    }

                    var validRead = await Group.Helpers.RemoveUnknownPermissionLevels(readGroups, PermissionLevel.Read, ctx, log);
                    var validEdit = await Group.Helpers.RemoveUnknownPermissionLevels(editGroups, PermissionLevel.Edit, ctx, log);
                    var validFullControl = await Group.Helpers.RemoveUnknownPermissionLevels(fullControlGroups, PermissionLevel.FullControl, ctx, log);
                    var validContribute = await Group.Helpers.RemoveUnknownPermissionLevels(contributeGroups, PermissionLevel.Contribute, ctx, log);
                    var validDesign = await Group.Helpers.RemoveUnknownPermissionLevels(designGroups, PermissionLevel.Design, ctx, log);

                    var validSharePointGroups = await Group.Helpers.CleanSharePointGroups(siteCollectionAdminGroups, ctx, log);

                    misconfigured = !misconfigured ? !(validRead && validEdit && validFullControl && validSharePointGroups && validContribute && validDesign) : misconfigured;

                    if (misconfigured)
                    {
                        misconfiguredSites.Add(site);

                        log.LogWarning($"Found misconfigured site: {site.Name} - {site.WebUrl}");
                    }
                }
            }
            while (allSites.NextPageRequest != null && (allSites = await allSites.NextPageRequest.GetAsync()).Count > 0);

            await StoreData.StoreReports(executionContext, reports, "reports", log);

            await Email.InformOwners(misconfiguredSites, graphAPIAuth, log);

            return new OkObjectResult(misconfiguredSites);
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
    }
}
