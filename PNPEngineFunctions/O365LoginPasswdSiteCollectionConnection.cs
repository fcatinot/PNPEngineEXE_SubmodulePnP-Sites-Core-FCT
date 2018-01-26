using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using Microsoft.Azure.ActiveDirectory.GraphClient;

namespace PNPEngineFunctions
{
    public class O365LoginPasswdSiteCollectionConnection
    {
        public Uri SiteUri { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public bool UseAppAuthentication { get; set; }

        public O365LoginPasswdSiteCollectionConnection(ConnectionSettings settings)
        {
            SiteUri = new Uri(settings.SiteUrl);
            Login = settings.Login;
            Password = settings.Password;
            UseAppAuthentication = settings.UseAppAuthentication;
        }

        public ClientContext Connect()
        {
            if (UseAppAuthentication == true)
            {
                AuthenticationManager authenticationManager = new AuthenticationManager();
                string realm = TokenHelper.GetRealmFromTargetUrl(new Uri(SiteUri.ToString()));
                return authenticationManager.GetAppOnlyAuthenticatedContext(SiteUri.ToString(), realm, Login, Password);
            }
            else
            {
                ClientContext context = new ClientContext(SiteUri.ToString());
                var securePassword = new SecureString();
                foreach (char c in Password)
                {
                    securePassword.AppendChar(c);
                }
                SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(Login, securePassword);
                context.Credentials = onlineCredentials;
                context.AuthenticationMode = ClientAuthenticationMode.Default;
                return context;
            }
        }
    }
}
