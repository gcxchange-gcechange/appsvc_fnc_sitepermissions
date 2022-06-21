﻿using System.Collections.Generic;

namespace SitePermissions
{
    public static class Globals
    {
        public static readonly string hubId = GetEnvironmentVariable("hubId");
        public static readonly string clientId = GetEnvironmentVariable("clientId");
        public static readonly string clientSecret = GetEnvironmentVariable("clientSecret");
        public static readonly string tenantId = GetEnvironmentVariable("tenantId");
        public static readonly string appOnlyId = GetEnvironmentVariable("appOnlyId");
        public static readonly string appOnlySecret = GetEnvironmentVariable("appOnlySecret");
        public static readonly string emailSenderId = GetEnvironmentVariable("emailSenderId");
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

        public class Group
        {
            public Group(string groupName, string id, string permissionLevel)
            {
                GroupName = groupName;
                Id = id;
                PermissionLevel = permissionLevel;
            }

            public string GroupName { get; set; }
            public string Id { get; set; }
            public string PermissionLevel { get; set; }
        }
    }
}
