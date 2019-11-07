using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;


namespace Microsoft.Office.Interop.TeamsAuto
{
    public class GroupHelper
    {
        //https://docs.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0&tabs=http
        public static async Task<string> GetDemoGroupId(HttpClient graphHttpClient)
        {
            //Read the graph document to update the api path
            string apiPath = "Please update the api path";

            HttpResponseMessage response = await graphHttpClient.GetAsync(apiPath);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new FocusException($"List groups graph call failed: {responseMsg}");
            }

            GraphDataSet<Identity> dataSet = JsonConvert.DeserializeObject<GraphDataSet<Identity>>(responseMsg);
            return dataSet.Value.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.Mail) && Settings.DemoGroupMail.IndexOf(x.Mail, StringComparison.CurrentCultureIgnoreCase) >= 0).Id;
        }

        public static async Task<IEnumerable<Identity>> GetGroupUsersAsync(string groupId, HttpClient graphHttpClient)
        {
            HttpResponseMessage response = await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/groups/{groupId}/members");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new FocusException($"List group users graph call failed: {responseMsg}");
            }

            GraphDataSet<Identity> dataSet = JsonConvert.DeserializeObject<GraphDataSet<Identity>>(responseMsg);
            return dataSet.Value.Where(x => !string.IsNullOrEmpty(x.Mail));
        }

        public static async Task AddUserIntoGroupAsync(string groupId, string userMail, HttpClient graphHttpClient)
        {
            HttpResponseMessage getUserResponse = await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/users/{userMail}");
            string userId = JsonConvert.DeserializeObject<Identity>(await getUserResponse.Content.ReadAsStringAsync()).Id;
            HttpContent content = new StringContent(JsonConvert.SerializeObject(new Dictionary<string, object>() { { "@odata.id", $"https://graph.microsoft.com/v1.0/users/{userId}" } }), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await graphHttpClient.PostAsync($"{Settings.GraphBaseUri}/groups/{groupId}/members/$ref", content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.NoContent)
            {
                throw new FocusException($"Add user into group graph call failed: {responseMsg}");
            }
        }
    }
}
