using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Office.Interop.TeamsAuto
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();
        }

        static async Task MainAsync(string[] args)
        {
            using (GraphClientManager manager = new GraphClientManager())
            {
                string groupId = await GroupHelper.GetDemoGroupId(manager.GetGraphHttpClient());
                string channelName = "TestChannel -" + DateTime.Now.ToString("yyyyMMddmmss");
                string channelId = await TeamsHelper.CreateChannelAsync(groupId, channelName, manager.GetGraphHttpClient());

                TeamsTabInfo tabInfo = Utility.GetTabInfoFromAppManifest();
                IEnumerable<TeamsApp> installedApps = await TeamsHelper.GetInstalledAppsAsync(groupId, manager.GetGraphHttpClient());
                TeamsApp targetApp = installedApps.FirstOrDefault(m => string.Equals(m?.Definition?.DisplayName, tabInfo.DisplayName, StringComparison.OrdinalIgnoreCase));
                if (targetApp != null)
                {
                    tabInfo.Id = targetApp.Definition.TeamsAppId;
                    await TeamsHelper.AddCustomTabAsync(groupId, channelId, tabInfo, manager.GetGraphHttpClient());
                }
                else
                {
                    throw new FocusException($"app {tabInfo.DisplayName} not installed");
                }
            }
        }
    }
}
