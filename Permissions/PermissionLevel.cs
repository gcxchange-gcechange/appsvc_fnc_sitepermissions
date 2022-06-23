using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace SitePermissions
{
    // TODO: add support for Design and Contribute
    public static class PermissionLevel
    {
        public static bool HasRead(BasePermissions permissions)
        {
            return IsValid(permissions, Read);
        }

        public static bool HasEdit(BasePermissions permissions)
        {
            return IsValid(permissions, Edit);
        }

        public static bool HasFullControl(BasePermissions permissions)
        {
            return IsValid(permissions, FullControl);
        }

        private static bool IsValid(BasePermissions permissions, PermissionKind[] masterKey)
        {
            var retVal = true;

            foreach (var permission in masterKey)
            {
                if (!permissions.Has(permission))
                {
                    retVal = false;
                    break;
                }
            }

            // Check for any extra permissions
            if (retVal)
            {
                retVal = GetEffectivePermissions(permissions).Count == masterKey.Length;
            }  

            return retVal;
        }

        public static List<PermissionKind> GetEffectivePermissions(BasePermissions permissions)
        {
            var retVal = new List<PermissionKind>();

            foreach (PermissionKind perm in System.Enum.GetValues(typeof(PermissionKind)))
            {
                var hasPermission = permissions.Has(perm);
                if (hasPermission)
                    retVal.Add(perm);
            }
            
            return retVal;
        }

        // https://pnp.github.io/pnpcore/api/PnP.Core.Model.SharePoint.PermissionKind.html
        public static readonly PermissionKind[] Read = {
            PermissionKind.EmptyMask,
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
        
        public static readonly PermissionKind[] Edit = {
            PermissionKind.EmptyMask,
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

        public static readonly PermissionKind[] FullControl = {
            PermissionKind.EmptyMask,
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
            PermissionKind.AddAndCustomizePages,
            // Personal Permissions
            PermissionKind.ManagePersonalViews,
            PermissionKind.AddDelPrivateWebParts,
            PermissionKind.UpdatePersonalWebParts
        };
    }
}
