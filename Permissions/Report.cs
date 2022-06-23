using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

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

            var readRoleDef = ctx.Web.RoleDefinitions.GetByName("Read");
            ctx.Load(readRoleDef);
            ctx.ExecuteQuery();

            var readReport = new BasePermission("Read", PermissionLevel.GetEffectivePermissions(readRoleDef.BasePermissions));
            basePermissionsReport.Add(readReport);

            var editRoleDef = ctx.Web.RoleDefinitions.GetByName("Edit");
            ctx.Load(editRoleDef);
            ctx.ExecuteQuery();

            var editReport = new BasePermission("Edit", PermissionLevel.GetEffectivePermissions(editRoleDef.BasePermissions));
            basePermissionsReport.Add(editReport);

            var fullControlRoleDef = ctx.Web.RoleDefinitions.GetByName("Full Control");
            ctx.Load(fullControlRoleDef);
            ctx.ExecuteQuery();

            var fullControlReport = new BasePermission("Full Control", PermissionLevel.GetEffectivePermissions(fullControlRoleDef.BasePermissions));
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
                        var domainGroup = new UserOrGroup();

                        domainGroup.Title = ((Microsoft.SharePoint.Client.User)ra.Member).Title;
                        domainGroup.Id = ((Microsoft.SharePoint.Client.User)ra.Member).LoginName.Split('|')[2];
                        domainGroup.PermissionLevel = role.Name;

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

                if (ra.Member is Group && (ra.Member.Title == ctx.Web.Title + " Owners" || ra.Member.Title == ctx.Web.Title + " Members" || ra.Member.Title == ctx.Web.Title + " Visitors"))
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

    public class UserOrGroup
    {
        public string Title;
        public string Id;
        public string PermissionLevel;
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
