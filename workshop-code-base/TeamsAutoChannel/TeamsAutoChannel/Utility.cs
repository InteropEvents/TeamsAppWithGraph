using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class Utility
    {
        private static TeamsTabInfo tabInfo = null;
        public static TeamsTabInfo GetTabInfoFromAppManifest()
        {
            if (tabInfo == null)
            {
                string tabManifestContent = File.ReadAllText("manifest.json");
                TeamsAppManifest manifest = JsonConvert.DeserializeObject<TeamsAppManifest>(tabManifestContent);
                tabInfo = new TeamsTabInfo();
                tabInfo.DisplayName = manifest.DisplayName;
                tabInfo.Configuration = manifest.ConfigList[0];
            }

            return tabInfo;
        }
    }

    public class TeamsAppManifest
    {
        [JsonProperty(PropertyName = "name")]
        public TeamsAppName Name;

        public string DisplayName
        {
            get
            {
                return Name.ShortName;
            }
        }

        [JsonProperty(PropertyName = "staticTabs")]
        public List<TeamsTabConfiguration> ConfigList;
    }

    public class TeamsAppName
    {
        [JsonProperty(PropertyName = "short")]
        public string ShortName;
    }
}
