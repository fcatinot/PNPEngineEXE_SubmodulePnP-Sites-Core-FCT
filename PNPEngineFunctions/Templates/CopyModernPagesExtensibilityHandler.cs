using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using OfficeDevPnP.Core;

namespace PNPEngineFunctions.Templates
{
    public class CopyModernPagesExtensibilityHandler : IProvisioningExtensibilityHandler
    {
        private Web webExtract;
        private ProvisioningTemplate extractTemplate;

        public ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template,
         ProvisioningTemplateCreationInformation creationInformation,
         PnPMonitoredScope scope, string configurationData)
        {
            webExtract = ctx.Web;
            ctx.Load(webExtract);
            ctx.ExecuteQueryRetry();
            //Save Template
            extractTemplate = template;

            string libraryName = "SitePages";
            //Load list
            foreach (ListInstance templateList in extractTemplate.Lists)
            {
                if (templateList.Url.ToString().Contains(libraryName))
                {
                    //Load all ModernPage to copy
                    List sourceList = webExtract.GetListByUrl(templateList.Url.ToString());
                    ListItemCollection modernPagesToMigrate = sourceList.GetItems(CamlQuery.CreateAllItemsQuery());
                    webExtract.Context.Load(modernPagesToMigrate, pages => pages.Include(page => page.DisplayName));
                    webExtract.Context.ExecuteQueryRetry();
                    var fieldColl = sourceList.Fields;
                    webExtract.Context.Load(fieldColl);
                    webExtract.Context.ExecuteQuery();

                    //Load items
                    foreach (ListItem page in modernPagesToMigrate)
                    {
                        webExtract.Context.Load(page);
                        webExtract.Context.ExecuteQueryRetry();
                        webExtract.Context.Load(page.File);
                        webExtract.Context.ExecuteQueryRetry();

                        // Ignore the Home Page
                        webExtract.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);
                        var homePageUrl = webExtract.RootFolder.WelcomePage;

                        string pageUrl =page.File.ServerRelativeUrl.Substring(page.File.ServerRelativeUrl.IndexOf(libraryName));

                        if (homePageUrl != pageUrl && page["ContentTypeId"].ToString().StartsWith(BuiltInContentTypeId.ModernArticlePage)) {
                            // Extract Modern Page
                            new ClientSidePageContentsHelper().ExtractClientSidePage(webExtract, extractTemplate, creationInformation, scope, pageUrl, page.File.Name, false);
                        }
                    }
                }
            }
            return extractTemplate;
        }

        public IEnumerable<TokenDefinition> GetTokens(ClientContext ctx,
         ProvisioningTemplate template, string configurationData)
        {
            return new List<TokenDefinition>();
        }


        public void Provision(ClientContext ctx, ProvisioningTemplate template,
         ProvisioningTemplateApplyingInformation applyingInformation,
         TokenParser tokenParser, PnPMonitoredScope scope, string configurationData)
        {

        }

    }
}