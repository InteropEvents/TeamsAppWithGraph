using System.Configuration;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class Settings
    {
        private const string AppTenantId = "AppTenantId";
        private const string AppClientId = "AppClientId";
        private const string AppClientSecret = "AppClientSecret";
        private const string GroupMail = "DemoGroupMailAddress";

        public const string GraphBaseUri = "https://graph.microsoft.com/v1.0";

        public const string GraphBetaUri = "https://graph.microsoft.com/beta";

        public static string TenantId => ConfigurationManager.AppSettings[AppTenantId];

        public static string ClientId => ConfigurationManager.AppSettings[AppClientId];

        public static string ClientSecret => ConfigurationManager.AppSettings[AppClientSecret];

        public static string DemoGroupMail => ConfigurationManager.AppSettings[GroupMail];
    }
}
