using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNPEngineFunctions
{
    public static class ContainerFolder
    {
        public static string SetContainerFolder(Web sourceWeb, ConnectionSettings _connectionSettings)
        {
            var pubWeb = PublishingWeb.GetPublishingWeb(sourceWeb.Context, sourceWeb);
            sourceWeb.Context.Load(pubWeb);
            sourceWeb.Context.ExecuteQueryRetry();
            string _containerFolder = "SiteAssets/tempazureHuman_" + _connectionSettings.Login.ToString().Split('@')[0];
            return _containerFolder;
        }
    }
}
