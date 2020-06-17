// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Graph;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Mvc;


namespace Focus.Controllers
{
    public class WorkFlowController : Controller
    {
        public struct TeamsContext
        {
            public string teamsGroupID;
            public string teamsChannelID;
            public string teamsChannelName;
        }

        private static GraphServiceClient graphClient;
        private static TeamsContext m_Context = new TeamsContext();

        // GET: WorkFlow
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult InitiateAuth()
        {
            return View();
        }

        public ActionResult EndAuth()
        {
            return View();
        }
        public ActionResult SilentEndAuth()
        {
            return View();
        }

        public ActionResult SilentInitiateAuth()
        {
            return View();
        }

        // https://ourcodeworld.com/articles/read/322/how-to-convert-a-
        //
        // -image-into-a-image-file-and-upload-it-with-an-asynchronous-form-using-jquery
        [HttpPost]
        public void Initialize(string accessToken, string teamsGroupID, string teamsChannelID, string teamsChannelName)
        {
            graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                await Task.FromResult<object>(null);
            }));

            m_Context.teamsChannelID = teamsChannelID;
            m_Context.teamsChannelName = teamsChannelName;
            m_Context.teamsGroupID = teamsGroupID;

            User me = graphClient.Me.Request().GetAsync().GetAwaiter().GetResult();
            string logInUserName = me.DisplayName + graphClient.SubscribedSkus.Request().GetAsync().GetAwaiter().GetResult().First().SkuPartNumber;

        }

        [HttpPost]
        public ActionResult UploadPhoto(string imageData)
        {
            string data = imageData;
            Match match = Regex.Match(data, @"data:image/(?<type>.+?),(?<data>.+)");
            if (!match.Success)
            {
                Response.StatusCode = (int)HttpStatusCode.BadRequest;
                return Content("field imageData format invalid");
            }
            string base64Data = match.Groups["data"]?.Value;
            var binData = Convert.FromBase64String(base64Data);
            string folderPath = @"/" + m_Context.teamsChannelName + @"/" + Uri.EscapeUriString("Focus" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + ".jpg");

            using (var stream = new MemoryStream(binData))
            {
                DriveItem uploadedItem = graphClient.Groups[m_Context.teamsGroupID].Drive.Root.ItemWithPath(folderPath).Content.Request().PutAsync<DriveItem>(stream)
                    .GetAwaiter().GetResult();
            }
            return new HttpStatusCodeResult(HttpStatusCode.OK);
        }
    }
}