using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class TeamsHelper
    {
        //https://docs.microsoft.com/en-us/graph/api/channel-post?view=graph-rest-1.0&tabs=http
        public static async Task<string> CreateChannelAsync(string groupId, string channelName, HttpClient graphHttpClient)
        {
            //Read the graph document to update the api path the the message content
            HttpContent content = new StringContent("Please implement the content");
            var apiPath = "Please update the api path";

            HttpResponseMessage response = await graphHttpClient.PostAsync(apiPath, content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new FocusException($"Create teams channel graph call failed: {responseMsg}");
            }

            return JsonConvert.DeserializeObject<Identity>(responseMsg).Id;
        }

        public static async Task UploadFileAsync(string groupId, string channelId, byte[] fileContent, string fileName, HttpClient graphhttpClient)
        {
            HttpContent content = new StreamContent(new MemoryStream(fileContent));
            HttpResponseMessage response = await graphhttpClient.PutAsync($"{Settings.GraphBaseUri}/groups/{groupId}/drive/root:/{channelId}/{fileName}:/content", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new FocusException($"Update file to channel graph call failed: {responseMsg}");
            }
        }

        //https://docs.microsoft.com/en-us/graph/api/teamstab-add?view=graph-rest-1.0
        public static async Task AddCustomTabAsync(string groupId, string channelId, TeamsTabInfo tabInfo, HttpClient graphHttpClient)
        {
            //Read the graph document to update the api path
            var apiPath = "Please update the api path";

            HttpContent content = new StringContent(JsonConvert.SerializeObject(tabInfo), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await graphHttpClient.PostAsync(apiPath, content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new FocusException($"Create teams channel tab graph call failed: {responseMsg}");
            }
        }

        public static async Task<IEnumerable<TeamsApp>> GetInstalledAppsAsync(string groupId, HttpClient graphhttpClient)
        {
            HttpResponseMessage response = await graphhttpClient.GetAsync($"{Settings.GraphBaseUri}/teams/{groupId}/installedApps?$expand=teamsAppDefinition");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new FocusException($"Get teams installed apps graph call failed: {responseMsg}");
            }

            GraphDataSet<TeamsApp> dataSet = JsonConvert.DeserializeObject<GraphDataSet<TeamsApp>>(responseMsg);
            return dataSet.Value;
        }
    }

    public class TeamsTabInfo
    {
        [JsonProperty(PropertyName = "displayName")]
        public string DisplayName;

        [JsonProperty(PropertyName = "teamsApp@odata.bind")]
        public string TeamsAppId
        {
            get
            {
                return $"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/{Id}";
            }
        }

        public string Id;

        [JsonProperty(PropertyName = "configuration")]
        public TeamsTabConfiguration Configuration;
    }

    public class TeamsTabConfiguration
    {
        [JsonProperty(PropertyName = "entityId")]
        public string EntityId;

        [JsonProperty(PropertyName = "contentUrl")]
        public string ContentUrl;

        [JsonProperty(PropertyName = "websiteUrl")]
        public string WebsiteUrl;

        [JsonProperty(PropertyName = "removeUrl")]
        public string RemoveUrl;
    }

    public class TeamsApp
    {
        [JsonProperty(PropertyName = "id")]
        public string Id;

        [JsonProperty(PropertyName = "teamsAppDefinition")]
        public TeamsAppDefinition Definition;
    }

    public class TeamsAppDefinition
    {
        [JsonProperty(PropertyName = "teamsAppId")]
        public string TeamsAppId;

        [JsonProperty(PropertyName = "displayName")]
        public string DisplayName;
    }
}
