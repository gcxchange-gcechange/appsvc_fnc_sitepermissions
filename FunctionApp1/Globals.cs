using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace SitePermissions
{
    public static class Globals
    {
        static IConfiguration config = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
        .AddEnvironmentVariables()
        .Build();

        public static readonly string hubId = config["hubId"];
        public static readonly string clientId = config["clientId"];
        public static readonly string clientSecret = config["clientSecret"];
        public static readonly string tenantId = config["tenantId"];
        public static readonly string appOnlyId = config["appOnlyId"];
        public static readonly string appOnlySecret = config["appOnlySecret"];

        public static readonly List<Group> groups = GetGroups();

        public static readonly string emailSenderId = config["emailSenderId"];
        
        public static string[] GetExcludedSiteIds()
        {
            var excludedSiteIds = config["excludeSiteIds"].Replace(" ", "").Split(",");
            excludedSiteIds.Prepend(hubId);

            return excludedSiteIds;
        }

        private static List<Group> GetGroups()
        {
            var groups = new List<Group>();

            var array = JArray.Parse(config["groups"]);
            foreach (JObject obj in array.Children<JObject>())
            {
                groups.Add(obj.ToObject<Group>());
            }

            return groups;
        }

        public class Group
        {
            public string GroupName { get; set; }
            public string GroupId { get; set; }
            public string PermissionLevel { get; set; }
        }
    }
}
