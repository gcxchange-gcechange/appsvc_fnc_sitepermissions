using System.Collections.Generic;

namespace SitePermissions
{
    public static class Globals
    {
        public static readonly string hubId = GetEnvironmentVariable("hubId");
        public static readonly string clientId = GetEnvironmentVariable("clientId");
        public static readonly string tenantId = GetEnvironmentVariable("tenantId");
        public static readonly string appOnlyId = GetEnvironmentVariable("appOnlyId");
        public static readonly string emailSenderId = GetEnvironmentVariable("emailSenderId");

        public static readonly string keyVaultUrl = GetEnvironmentVariable("keyVaultUrl");
        public static readonly string secretNameClient = GetEnvironmentVariable("secretNameClient");
        public static readonly string secretNameAppOnly = GetEnvironmentVariable("secretNameAppOnly");

        public static readonly List<Group> groups = GetGroups();

        public static List<string> GetExcludedSiteIds()
        {
            var excludedSiteIds = new List<string>(GetEnvironmentVariable("excludeSiteIds").Replace(" ", "").Split(","));
            excludedSiteIds.Add(hubId);

            return excludedSiteIds;
        }

        private static List<Group> GetGroups()
        {
            var groups = new List<Group>();
            var split = GetEnvironmentVariable("groups").Split(",");

            foreach (var group in split)
            {
                if (group != string.Empty)
                {
                    var props = group.Split("|");
                    groups.Add(new Group(props[0].Trim(), props[1].Trim(), props[2].Trim()));
                }
            }

            return groups;
        }

        private static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, System.EnvironmentVariableTarget.Process);
        }
    }
}
