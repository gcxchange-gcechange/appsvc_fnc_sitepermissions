using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace SitePermissions
{
    public static class PermissionLevel
    {
        public static bool HasRead(ClientResult<BasePermissions> permissions)
        {
            return IsValid(permissions, Read);
        }

        public static bool HasEdit(ClientResult<BasePermissions> permissions)
        {
            return IsValid(permissions, Edit);
        }

        public static bool HasFullControl(ClientResult<BasePermissions> permissions)
        {
            return IsValid(permissions, FullControl);
        }

        private static bool IsValid(ClientResult<BasePermissions> permissions, PermissionKind[] masterKey)
        {
            var retVal = true;

            foreach (var permission in masterKey)
            {
                if (!permissions.Value.Has(permission))
                {
                    retVal = false;
                    break;
                }
            }

            return retVal;
        }

        // https://pnp.github.io/pnpcore/api/PnP.Core.Model.SharePoint.PermissionKind.html
        private static readonly PermissionKind[] Read = {
            // List Permissions
            PermissionKind.ViewListItems,
            PermissionKind.OpenItems,
            PermissionKind.ViewVersions,
            PermissionKind.CreateAlerts,
            PermissionKind.ViewFormPages,
            // Site Permissions
            PermissionKind.CreateSSCSite,
            PermissionKind.ViewPages,
            PermissionKind.BrowseUserInfo,
            PermissionKind.UseRemoteAPIs,
            PermissionKind.UseClientIntegration,
            PermissionKind.Open
        };

        private static readonly PermissionKind[] Edit = {
            // List Permissions
            PermissionKind.ManageLists,
            PermissionKind.AddListItems,
            PermissionKind.EditListItems,
            PermissionKind.DeleteListItems,
            PermissionKind.ViewListItems,
            PermissionKind.OpenItems,
            PermissionKind.ViewVersions,
            PermissionKind.DeleteVersions,
            PermissionKind.CreateAlerts,
            PermissionKind.ViewFormPages,
            // Site Permissions
            PermissionKind.BrowseDirectories,
            PermissionKind.CreateSSCSite,
            PermissionKind.BrowseUserInfo,
            PermissionKind.ViewPages,
            PermissionKind.UseRemoteAPIs,
            PermissionKind.UseClientIntegration,
            PermissionKind.Open,
            PermissionKind.EditMyUserInfo,
            // Personal Permissions
            PermissionKind.ManagePersonalViews,
            PermissionKind.AddDelPrivateWebParts,
            PermissionKind.UpdatePersonalWebParts
        };

        private static readonly PermissionKind[] FullControl = {
            // List Permissions
            PermissionKind.ManageLists,
            PermissionKind.CancelCheckout,
            PermissionKind.AddListItems,
            PermissionKind.EditListItems,
            PermissionKind.DeleteListItems,
            PermissionKind.ViewListItems,
            PermissionKind.ApproveItems,
            PermissionKind.OpenItems,
            PermissionKind.ViewVersions,
            PermissionKind.DeleteVersions,
            PermissionKind.CreateAlerts,
            PermissionKind.ViewFormPages,
            PermissionKind.AnonymousSearchAccessList,
            PermissionKind.AnonymousSearchAccessWebLists,
            // Site Permissions
            PermissionKind.ManagePermissions,
            PermissionKind.ViewUsageData,
            PermissionKind.ManageSubwebs,
            PermissionKind.ManageWeb,
            PermissionKind.ApplyThemeAndBorder,
            PermissionKind.ApplyStyleSheets,
            PermissionKind.CreateGroups,
            PermissionKind.BrowseDirectories,
            PermissionKind.CreateSSCSite,
            PermissionKind.ViewPages,
            PermissionKind.EnumeratePermissions,
            PermissionKind.BrowseUserInfo,
            PermissionKind.ManageAlerts,
            PermissionKind.UseRemoteAPIs,
            PermissionKind.UseClientIntegration,
            PermissionKind.Open,
            PermissionKind.EditMyUserInfo,
            // Personal Permissions
            PermissionKind.ManagePersonalViews,
            PermissionKind.AddDelPrivateWebParts,
            PermissionKind.UpdatePersonalWebParts
        };
    }
}
