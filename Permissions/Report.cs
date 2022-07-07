using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using UserOrGroup = SitePermissions.Group;

namespace SitePermissions
{
    public class Report
    {
        public string Title;
        public string Id;
        public string URL;
        public List<BasePermission> PermissionLevels;
        public List<UserOrGroup> UsersAndGroups;
        public List<SharePointGroup> SharePointGroups;
        public List<User> SiteCollectionAdministrators;

        [JsonIgnore]
        public ClientContext ctx;

        public Report(Microsoft.Graph.Site site, ClientContext context)
        {
            ctx = context;
            Title = site.Name;
            Id = site.Id.Split(",")[1];
            URL = site.WebUrl;
            PermissionLevels = GetPermissionLevelReport();
            UsersAndGroups = GetUserAndGroupReport();
            SharePointGroups = GetSharePointGroupReport();
            SiteCollectionAdministrators = SiteCollectionsAdministrators();
        }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }

        private List<BasePermission> GetPermissionLevelReport()
        {
            var basePermissionsReport = new List<BasePermission>();

            var readRoleDef = ctx.Web.RoleDefinitions.GetById((int)PermissionLevel.RoleDefinitionIds.Read);
            ctx.Load(readRoleDef);
            ctx.ExecuteQuery();

            var readReport = new BasePermission(readRoleDef.Name, PermissionLevel.GetEffectivePermissions(readRoleDef.BasePermissions));
            basePermissionsReport.Add(readReport);

            var contributeRoleDef = ctx.Web.RoleDefinitions.GetById((int)PermissionLevel.RoleDefinitionIds.Contribute);
            ctx.Load(contributeRoleDef);
            ctx.ExecuteQuery();

            var contributeReport = new BasePermission(contributeRoleDef.Name, PermissionLevel.GetEffectivePermissions(contributeRoleDef.BasePermissions));
            basePermissionsReport.Add(contributeReport);

            var editRoleDef = ctx.Web.RoleDefinitions.GetById((int)PermissionLevel.RoleDefinitionIds.Edit);
            ctx.Load(editRoleDef);
            ctx.ExecuteQuery();

            var editReport = new BasePermission(editRoleDef.Name, PermissionLevel.GetEffectivePermissions(editRoleDef.BasePermissions));
            basePermissionsReport.Add(editReport);

            var designRoleDef = ctx.Web.RoleDefinitions.GetById((int)PermissionLevel.RoleDefinitionIds.Design);
            ctx.Load(designRoleDef);
            ctx.ExecuteQuery();

            var designReport = new BasePermission(designRoleDef.Name, PermissionLevel.GetEffectivePermissions(designRoleDef.BasePermissions));
            basePermissionsReport.Add(designReport);

            var fullControlRoleDef = ctx.Web.RoleDefinitions.GetById((int)PermissionLevel.RoleDefinitionIds.FullControl);
            ctx.Load(fullControlRoleDef);
            ctx.ExecuteQuery();

            var fullControlReport = new BasePermission(fullControlRoleDef.Name, PermissionLevel.GetEffectivePermissions(fullControlRoleDef.BasePermissions));
            basePermissionsReport.Add(fullControlReport);

            return basePermissionsReport;
        }

        private List<UserOrGroup> GetUserAndGroupReport()
        {
            var domainGroupReport = new List<UserOrGroup>();

            var roleAssignments = ctx.Web.RoleAssignments;

            ctx.Load(roleAssignments, r => r.Include(i => i.Member, i => i.RoleDefinitionBindings));
            ctx.ExecuteQuery();

            for (var i = 0; i < roleAssignments.Count; i++)
            {
                var ra = roleAssignments[i];

                if (ra.Member is Microsoft.SharePoint.Client.User)
                {
                    foreach (var role in ra.RoleDefinitionBindings)
                    {
                        var name = ((Microsoft.SharePoint.Client.User)ra.Member).Title;
                        var id = ((Microsoft.SharePoint.Client.User)ra.Member).LoginName.Split('|')[2];
                        var permissionLevel = role.Name;

                        var domainGroup = new UserOrGroup(name, id, permissionLevel);

                        domainGroupReport.Add(domainGroup);
                    }
                }
            }

            return domainGroupReport;
        }

        private List<SharePointGroup> GetSharePointGroupReport()
        {
            var sharePointGroupReport = new List<SharePointGroup>();

            var roleAssignments = ctx.Web.RoleAssignments;

            ctx.Load(ctx.Web);
            ctx.Load(roleAssignments, r => r.Include(i => i.Member, i => i.RoleDefinitionBindings));
            ctx.ExecuteQuery();

            for (var i = 0; i < roleAssignments.Count; i++)
            {
                var ra = roleAssignments[i];

                if (ra.Member is Microsoft.SharePoint.Client.Group && (ra.Member.Title == ctx.Web.Title + " Owners" || ra.Member.Title == ctx.Web.Title + " Members" || ra.Member.Title == ctx.Web.Title + " Visitors"))
                {
                    var oGroup = ctx.Web.SiteGroups.GetByName(ra.Member.LoginName);

                    var oUserCollection = oGroup.Users;

                    ctx.Load(oUserCollection);
                    ctx.ExecuteQuery();

                    var sharePointGroup = new SharePointGroup();
                    sharePointGroup.Title = ra.Member.Title;

                    var userList = new List<User>();

                    foreach (var user in oUserCollection)
                    {
                        var reportUser = new User();

                        reportUser.Title = user.Title;

                        if (user.Title != "System Account")
                            reportUser.Id = user.LoginName.Split('|')[2];

                        userList.Add(reportUser);
                    }

                    sharePointGroup.Users = userList;

                    sharePointGroupReport.Add(sharePointGroup);
                }
            }

            return sharePointGroupReport;
        }

        private List<User> SiteCollectionsAdministrators()
        {
            var siteCollectionsAdministrators = new List<User>();

            ctx.Load(ctx.Web);
            ctx.Load(ctx.Site);
            ctx.Load(ctx.Site.RootWeb);
            ctx.ExecuteQuery();

            var users = ctx.Site.RootWeb.SiteUsers;
            ctx.Load(users);
            ctx.ExecuteQuery();

            foreach (var user in users)
            {
                if (user.IsSiteAdmin)
                {
                    var reportUser = new User();

                    reportUser.Title = user.Title;
                    reportUser.Id = user.LoginName.Split('|')[2];

                    siteCollectionsAdministrators.Add(reportUser);
                }
            }

            return siteCollectionsAdministrators;
        }
    }

    public class BasePermission
    {
        public string Title;
        public List<string> Permissions;

        public BasePermission(string title, List<PermissionKind> permissions)
        {
            Title = title;

            var permissionString = new List<string>();
            foreach (var permission in permissions)
            {
                permissionString.Add(Enum.GetName(permission));
            }

            Permissions = permissionString;
        }
    }

    public class SharePointGroup
    {
        public string Title;
        public List<User> Users;
    }

    public class User
    {
        public string Title;
        public string Id;
    }
}
