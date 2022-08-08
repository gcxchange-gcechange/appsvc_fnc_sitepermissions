using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using PnP.Framework.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SitePermissions
{
    public class Group
    {
        public Group(string groupName, string id, string permissionLevel)
        {
            Name = groupName;
            Id = id;
            AssignedPermissionLevel = permissionLevel;
        }

        public string Name { get; set; }
        public string Id { get; set; }
        public string AssignedPermissionLevel { get; set; }

        public async Task<bool> HasPermissionLevel(string permissionLevel, ClientContext ctx, ILogger log)
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

                    if (ra.Member is Microsoft.SharePoint.Client.User && Helpers.GetObjectId(((Microsoft.SharePoint.Client.User)ra.Member).LoginName) == Id)
                    {
                        foreach (var role in ra.RoleDefinitionBindings)
                        {
                            if (role.Name == permissionLevel)
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
                log.LogError($"Error validating {Name} has {permissionLevel}: {ex.Message}");
            }

            return result;
        }

        // Returns true if a level was removed.
        public async Task<bool> RemovePermissionLevels(List<string> levelsToRemove, ClientContext ctx, ILogger log)
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

                    if (ra.Member is Microsoft.SharePoint.Client.User && Helpers.GetObjectId((ra.Member).LoginName) == Id)
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
                        break;
                    }
                }

                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.LogError($"Error removing {Name} from {ctx.Site.Url} - {ex.Source}: {ex.Message} | {ex.InnerException}");
            }

            return result;
        }

        public async Task<bool> AddPermissionLevel(string permissionLevel, ClientContext ctx, ILogger log)
        {
            var result = true;

            try
            {
                var user = Helpers.GetUserById(Id, ctx, log);

                var adGroup = ctx.Web.EnsureUser(user.Title);
                ctx.Load(adGroup);
                var spGroup = ctx.Web.AssociatedMemberGroup;
                spGroup.Users.AddUser(adGroup);

                var writeDefinition = ctx.Web.RoleDefinitions.GetByName(permissionLevel);
                var roleDefCollection = new RoleDefinitionBindingCollection(ctx);
                roleDefCollection.Add(writeDefinition);
                var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                ctx.Load(spGroup, x => x.Users);
                ctx.ExecuteQuery();

                log.LogInformation($"Gave {Name} {permissionLevel} on {ctx.Site.Url}");
            }
            catch (Exception ex)
            {
                log.LogError($"Error adding {Name} to {ctx.Site.Url} - {ex.Source}: {ex.Message} | {ex.InnerException}");
            }

            return result;
        }

        // Returns true if the group is found in the site collections administrator list
        public async Task<bool> IsSiteCollectionAdministrator(ClientContext ctx, ILogger log)
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
                    if (user.IsSiteAdmin && Helpers.GetObjectId(user.LoginName) == Id)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error verifying site collection administrator for {Name}: {ex.Message}");
            }

            return false;
        }

        public static class Helpers
        {
            //Goes through site permissions and removes any that are not in the expectedGroups for the permissionLevel
            //Returns false if any were found and removed.
            public static async Task<bool> RemoveUnknownPermissionLevels(List<Group> expectedGroups, string permissionLevel, ClientContext ctx, ILogger log)
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

                        if (ra.Member is Microsoft.SharePoint.Client.User && !expectedGroups.Any(x => x.Id == GetObjectId((ra.Member).LoginName)))
                        {

                            foreach (var role in ra.RoleDefinitionBindings)
                            {
                                if (role.Name == permissionLevel)
                                {
                                    var group = new Group(((Microsoft.SharePoint.Client.User)ra.Member).Title, GetObjectId(((Microsoft.SharePoint.Client.User)ra.Member).LoginName), permissionLevel);
                                    result = !await group.RemovePermissionLevels(new List<string>() { permissionLevel }, ctx, log) == false && result ? false : result;
                                }
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
            public static async Task<bool> CleanSharePointGroups(ClientContext ctx, GraphServiceClient graphAPIAuth, ILogger log)
            {
                var result = true;

                try
                {
                    var memberGroupFound = false;
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
                                var isMembersUser = user.Title == ctx.Web.Title + " Members" || user.Title == ctx.Web.Description + " Members";

                                if (!memberGroupFound)
                                    memberGroupFound = (isMembersGroup && isMembersUser);

                                // Ignore system accounts and members in the member group
                                if (user.Title == "System Account" || (isMembersGroup && isMembersUser))
                                    continue;

                                oGroup.Users.RemoveByLoginName(user.LoginName);
                                ctx.ExecuteQuery();

                                result = false;

                                log.LogWarning($"Removing {user.Title} from {ra.Member.Title}");
                            }

                            // No member group found, add it
                            if (isMembersGroup && !memberGroupFound)
                            {
                                log.LogWarning($"{ctx.Web.Title} Members user is missing from the {ctx.Web.Title} Members SharePoint Group");

                                var groups = await graphAPIAuth.Groups
                                .Request()
                                .Header("ConsistencyLevel", "eventual")
                                .GetAsync();

                                var groupFound = false;
                                do
                                {
                                    foreach (var group in groups)
                                    {
                                        if (group.DisplayName == ctx.Web.Title)
                                        {
                                            var user = ctx.Web.EnsureUser($"c:0o.c|federateddirectoryclaimprovider|{group.Id}");
                                            ctx.Load(user);
                                            ctx.ExecuteQuery();
                                            
                                            var addUser = oGroup.Users.AddUser(user);
                                            ctx.Load(addUser);
                                            ctx.ExecuteQuery();

                                            groupFound = true;
                                            result = false;

                                            log.LogInformation($"Added {group.DisplayName}");
                                            
                                            break;
                                        }
                                    }
                                } while (!groupFound && groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0);
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

            // Adds the group to the site collection administrator list
            public static async Task<bool> AddSiteCollectionAdministrator(Group group, ClientContext ctx, ILogger log)
            {
                var result = true;

                try
                {
                    ctx.Load(ctx.Web);
                    ctx.Load(ctx.Site);
                    ctx.Load(ctx.Site.RootWeb);
                    ctx.ExecuteQuery();

                    List<UserEntity> admins = new List<UserEntity>();
                    UserEntity adminUserEntity = new UserEntity();

                    var user = GetUserById(group.Id, ctx, log);

                    adminUserEntity.LoginName = user.LoginName;

                    admins.Add(adminUserEntity);

                    ctx.Site.RootWeb.AddAdministrators(admins, true);

                    log.LogInformation($"Added {group.Name} as Site Collection Administrators to {ctx.Site.Url}");
                }
                catch (Exception ex)
                {
                    log.LogError($"Error addming site collection admin to {ctx.Site.Url}: {ex.Message}");
                }


                return result;
            }

            // Removes all site collection administrators.
            public static async Task<bool> RemoveSiteCollectionAdministrators(List<Group> approvedAdminGroups, ClientContext ctx, ILogger log)
            {
                var isValid = true;
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
                    var isAdmin = approvedAdminGroups.Any(x => x.Name == user.Title);

                    if (user.IsSiteAdmin && !isAdmin)
                    {
                        try
                        {
                            user.IsSiteAdmin = false;
                            user.Update();
                            ctx.Load(user);
                            ctx.ExecuteQuery();

                            removedUsers.Add(user);
                            isValid = false;

                            log.LogWarning($"Removed {user.Title} from Site Collection Administrators for {ctx.Site.Url}");
                        }
                        catch (Exception ex)
                        {
                            log.LogError($"Error removing {user.Title} from site collection administrator: {ex.Message}");
                        }
                    }
                }

                return isValid;
            }

            public static Microsoft.SharePoint.Client.User GetUserById(string id, ClientContext ctx, ILogger log)
            {
                Microsoft.SharePoint.Client.User user = null;

                ctx.Load(ctx.Web.SiteUsers);
                ctx.ExecuteQuery();

                foreach (var _user in ctx.Web.SiteUsers)
                {
                    if (GetObjectId(_user.LoginName) == id)
                    {
                        user = _user;
                        break;
                    }
                }

                if (user == null)
                    log.LogError($"Unable to find user by Id {id}");

                return user;
            }

            public static Microsoft.SharePoint.Client.User GetUserByEmail(string email, ClientContext ctx, ILogger log)
            {
                Microsoft.SharePoint.Client.User user = null;

                ctx.Load(ctx.Web.SiteUsers);
                ctx.ExecuteQuery();

                foreach (var _user in ctx.Web.SiteUsers)
                {
                    if (_user.Email == email)
                    {
                        user = _user;
                        break;
                    }
                }

                if (user == null)
                    log.LogError($"Unable to find user by email \"{email}\"");

                return user;
            }

            public static string GetObjectId(string loginName)
            {
                var split = loginName.Split('|');
                return split.Length == 3 ? split[2] : "-1";
            }
        }
    }
}
