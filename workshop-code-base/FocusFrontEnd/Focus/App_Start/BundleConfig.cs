// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Web;
using System.Web.Optimization;

namespace Focus
{
    public class BundleConfig
    {
        // For more information on bundling, visit https://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/site.css"));

            bundles.Add(new ScriptBundle("~/bundles/Prod").Include(
                                  "~/Scripts/config.js",
                                  "~/Scripts/FlowCalls.js",
                                  "~/Scripts/Microlearning.js",
                                  "~/Scripts/teamsapp.js",
                                  "~/Scripts/Planner/planner.js",
                                  "~/Scripts/Cognitive/*.js"));
            bundles.Add(new ScriptBundle("~/bundles/Debug").Include(
                                 "~/Scripts/Cognitive/Test/*.js"));
        }
    }
}
