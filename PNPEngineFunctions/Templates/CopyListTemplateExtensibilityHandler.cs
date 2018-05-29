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
    public class CopyListTemplateExtensibilityHandler : IProvisioningExtensibilityHandler
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

            using (scope)
            {
                // Check if this is not a noscript site as we're not allowed to update some properties
                bool isNoScriptSite = webExtract.IsNoScriptSite();

                webExtract.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);

                var serverRelativeUrl = webExtract.ServerRelativeUrl;


                // For each list in the site
                var lists = webExtract.Lists;

                webExtract.Context.Load(lists,
                    lc => lc.IncludeWithDefaultProperties(
                        l => l.ContentTypes,
                        l => l.Views,
                        l => l.EnableModeration,
                        l => l.ForceCheckout,
                        l => l.BaseTemplate,
                        l => l.OnQuickLaunch,
                        l => l.RootFolder.ServerRelativeUrl,
                        l => l.UserCustomActions,
#if !SP2013
                        l => l.MajorVersionLimit,
                        l => l.MajorWithMinorVersionsLimit,
#endif
                        l => l.DraftVersionVisibility,
                        l => l.DefaultDisplayFormUrl,
                        l => l.DefaultEditFormUrl,
                        l => l.ImageUrl,
                        l => l.DefaultNewFormUrl,
                        l => l.Direction,
                        l => l.IrmExpire,
                        l => l.IrmReject,
                        l => l.IrmEnabled,
                        l => l.IsApplicationList,
                        l => l.ValidationFormula,
                        l => l.ValidationMessage,
                        l => l.DocumentTemplateUrl,
                        l => l.NoCrawl,
#if !ONPREMISES
                        l => l.ListExperienceOptions,
                        l => l.ReadSecurity,
#endif
                        l => l.Fields.IncludeWithDefaultProperties(
                            f => f.Id,
                            f => f.Title,
                            f => f.Hidden,
                            f => f.InternalName,
                            f => f.DefaultValue,
                            f => f.Required)));

                webExtract.Context.ExecuteQueryRetry();

                var allLists = new List<List>();

                foreach (var list in lists)
                {
                    allLists.Add(list);
                }
                // Let's see if there are workflow subscriptions
                Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription[] workflowSubscriptions = null;
                try
                {
                    workflowSubscriptions = webExtract.GetWorkflowSubscriptions();
                }
                catch (ServerException)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    // Swallow the exception
                }

                // Retrieve all not hidden lists and the Workflow History Lists, just in case there are active workflow subscriptions
                var listsToProcess = lists.AsEnumerable().Where(l => (l.Hidden == false || ((workflowSubscriptions != null && workflowSubscriptions.Length > 0) && l.BaseTemplate == 140))).ToArray();
                var listCount = 0;
                foreach (var siteList in listsToProcess)
                {
                    listCount++;
                    ListInstance baseTemplateList = null;
                    if (creationInformation.BaseTemplate != null)
                    {
                        // Check if we need to skip this list...if so let's do it before we gather all the other information for this list...improves performance
                        var index = creationInformation.BaseTemplate.Lists.FindIndex(f => f.Url.Equals(siteList.RootFolder.ServerRelativeUrl.Substring(serverRelativeUrl.Length + 1)) &&
                                                                                   f.TemplateType.Equals(siteList.BaseTemplate));
                        if (index != -1)
                        {
                            baseTemplateList = creationInformation.BaseTemplate.Lists[index];
                        }
                    }

                    var contentTypeFields = new List<FieldRef>();
                    var list = new ListInstance
                    {
                        Description = siteList.Description,
                        EnableVersioning = siteList.EnableVersioning,
                        TemplateType = siteList.BaseTemplate,
                        Title = siteList.Title,
                        Hidden = siteList.Hidden,
                        EnableFolderCreation = siteList.EnableFolderCreation,
                        DocumentTemplate = Tokenize(siteList.DocumentTemplateUrl, webExtract.Url),
                        ContentTypesEnabled = siteList.ContentTypesEnabled,
                        Url = siteList.RootFolder.ServerRelativeUrl.Substring(serverRelativeUrl.Length).TrimStart('/'),
                        TemplateFeatureID = siteList.TemplateFeatureId,
                        EnableAttachments = siteList.EnableAttachments,
                        OnQuickLaunch = siteList.OnQuickLaunch,
                        DefaultDisplayFormUrl = Tokenize(siteList.DefaultDisplayFormUrl, webExtract.Url),
                        DefaultEditFormUrl = Tokenize(siteList.DefaultEditFormUrl, webExtract.Url),
                        DefaultNewFormUrl = Tokenize(siteList.DefaultNewFormUrl, webExtract.Url),
                        Direction = siteList.Direction.ToLower() == "none" ? ListReadingDirection.None : siteList.Direction.ToLower() == "rtl" ? ListReadingDirection.RTL : ListReadingDirection.LTR,
                        ImageUrl = Tokenize(siteList.ImageUrl, webExtract.Url),
                        IrmExpire = siteList.IrmExpire,
                        IrmReject = siteList.IrmReject,
                        IsApplicationList = siteList.IsApplicationList,
                        ValidationFormula = siteList.ValidationFormula,
                        ValidationMessage = siteList.ValidationMessage,
                        EnableModeration = siteList.EnableModeration,
                        NoCrawl = siteList.NoCrawl,
                        MaxVersionLimit =
                            siteList.IsPropertyAvailable("MajorVersionLimit") ? siteList.MajorVersionLimit : 0,
                        EnableMinorVersions = siteList.EnableMinorVersions,
                        MinorVersionLimit =
                            siteList.IsPropertyAvailable("MajorWithMinorVersionsLimit")
                                ? siteList.MajorWithMinorVersionsLimit
                                : 0,
                        ForceCheckout = siteList.IsPropertyAvailable("ForceCheckout") ?
                            siteList.ForceCheckout : false,
                        DraftVersionVisibility = siteList.IsPropertyAvailable("DraftVersionVisibility") ? (int)siteList.DraftVersionVisibility : 0,
                    };

                    if (siteList.BaseTemplate != (int)ListTemplateType.PictureLibrary)
                    {
                        siteList.EnsureProperties(l => l.InformationRightsManagementSettings);
                    }

                    if (baseTemplateList != null)
                    {
                        if (baseTemplateList.Url.Equals(list.Url) && template.Lists.Where(l => l.Url.ToString() == list.Url.ToString()).ToList().Count == 0)
                        {
                            template.Lists.Add(list);
                        }
                    }
                }

            }
            //Save Template
            extractTemplate = template;
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

        protected string Tokenize(string url, string webUrl, Web web = null)
        {
            String result = null;

            if (string.IsNullOrEmpty(url))
            {
                // nothing to tokenize...
                result = String.Empty;
            }
            else
            {
                // Decode URL
                url = Uri.UnescapeDataString(url);
                // Try with theme catalog
                if (url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    var subsite = false;
                    if (web != null)
                    {
                        subsite = web.IsSubSite();
                    }
                    if (subsite)
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/theme", "{sitecollection}/_catalogs/theme");
                    }
                    else
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/theme", "{themecatalog}");
                    }
                }

                // Try with master page catalog
                if (url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    var subsite = false;
                    if (web != null)
                    {
                        subsite = web.IsSubSite();
                    }
                    if (subsite)
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/masterpage", "{sitecollection}/_catalogs/masterpage");
                    }
                    else
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/masterpage", "{masterpagecatalog}");
                    }
                }

                // Try with site URL
                if (result != null)
                {
                    url = result;
                }
                Uri uri;
                if (Uri.TryCreate(webUrl, UriKind.Absolute, out uri))
                {
                    string webUrlPathAndQuery = System.Web.HttpUtility.UrlDecode(uri.PathAndQuery);
                    // Don't do additional replacement when masterpagecatalog and themecatalog (see #675)
                    if (url.IndexOf(webUrlPathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1 && (url.IndexOf("{masterpagecatalog}") == -1) && (url.IndexOf("{themecatalog}") == -1))
                    {
                        result = (uri.PathAndQuery.Equals("/") && url.StartsWith(uri.PathAndQuery))
                            ? "{site}" + url // we need this for DocumentTemplate attribute of pnp:ListInstance also on a root site ("/") without managed path
                            : url.Replace(webUrlPathAndQuery, "{site}");
                    }
                }

                // Default action
                if (String.IsNullOrEmpty(result))
                {
                    result = url;
                }
            }

            return (result);
        }

    }
}