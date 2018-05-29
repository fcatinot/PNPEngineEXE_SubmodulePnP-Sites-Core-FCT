using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PNPEngineFunctions.Templates
{
    class CopyWikiPagesExtensibilityHandler : IProvisioningExtensibilityHandler
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
                    ListItemCollection pagesToMigrate = sourceList.GetItems(CamlQuery.CreateAllItemsQuery());
                    webExtract.Context.Load(pagesToMigrate, pages => pages.Include(page => page.DisplayName));
                    webExtract.Context.ExecuteQueryRetry();
                    var fieldColl = sourceList.Fields;
                    webExtract.Context.Load(fieldColl);
                    webExtract.Context.ExecuteQuery();

                    //Load items
                    foreach (ListItem page in pagesToMigrate)
                    {
                        webExtract.Context.Load(page);
                        webExtract.Context.ExecuteQueryRetry();
                        webExtract.Context.Load(page.File);
                        webExtract.Context.ExecuteQueryRetry();

                        // Ignore the Home Page
                        webExtract.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);
                        var homePageUrl = webExtract.RootFolder.WelcomePage;
                        homePageUrl = webExtract.ServerRelativeUrl + '/' + homePageUrl;

                        string pageListUrl = page.File.ServerRelativeUrl.Substring(page.File.ServerRelativeUrl.IndexOf(libraryName));
                        string pageUrl = webExtract.ServerRelativeUrl + '/' + pageListUrl;

                        if (homePageUrl != pageUrl && page["ContentTypeId"].ToString().StartsWith(BuiltInContentTypeId.WikiDocument))
                        {
                            var file = webExtract.GetFileByServerRelativeUrl(pageUrl);
                            try
                            {
                                var listItem = file.EnsureProperty(f => f.ListItemAllFields);
                                if (listItem != null)
                                {
                                    if (listItem.FieldValues.ContainsKey("WikiField") && listItem.FieldValues["WikiField"] != null)
                                    {
                                        // Wiki page
                                        var fullUri = new Uri(UrlUtility.Combine(webExtract.Url, pageListUrl));

                                        var homeFile = webExtract.GetFileByServerRelativeUrl(pageUrl);

                                        var limitedWPManager = homeFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

                                        webExtract.Context.Load(limitedWPManager);

                                        var wikiPage = new Page()
                                        {
                                            Layout = WikiPageLayout.Custom,
                                            Overwrite = true,
                                            Url = Tokenize(fullUri.PathAndQuery, webExtract.Url),
                                        };
                                        var pageContents = listItem.FieldValues["WikiField"].ToString();

                                        Regex regexClientIds = new Regex(@"id=\""div_(?<ControlId>(\w|\-)+)");
                                        if (regexClientIds.IsMatch(pageContents))
                                        {
                                            foreach (Match webPartMatch in regexClientIds.Matches(pageContents))
                                            {
                                                String serverSideControlId = webPartMatch.Groups["ControlId"].Value;

                                                try
                                                {
                                                    var serverSideControlIdToSearchFor =
                                                        $"g_{serverSideControlId.Replace("-", "_")}";

                                                    var webPart = limitedWPManager.WebParts.GetByControlId(serverSideControlIdToSearchFor);
                                                    webExtract.Context.Load(webPart,
                                                        wp => wp.Id,
                                                        wp => wp.WebPart.Title,
                                                        wp => wp.WebPart.ZoneIndex
                                                        );
                                                    webExtract.Context.ExecuteQueryRetry();

                                                    var webPartxml = TokenizeWebPartXml(webExtract, webExtract.GetWebPartXml(webPart.Id, pageUrl));

                                                    wikiPage.WebParts.Add(new OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart()
                                                    {
                                                        Title = webPart.WebPart.Title,
                                                        Contents = webPartxml,
                                                        Order = (uint)webPart.WebPart.ZoneIndex,
                                                        Row = 1, // By default we will create a onecolumn layout, add the webpart to it, and later replace the wikifield on the page to position the webparts correctly.
                                                        Column = 1 // By default we will create a onecolumn layout, add the webpart to it, and later replace the wikifield on the page to position the webparts correctly.
                                                    });

                                                    pageContents = Regex.Replace(pageContents, serverSideControlId, $"{{webpartid:{webPart.WebPart.Title}}}", RegexOptions.IgnoreCase);
                                                }
                                                catch (ServerException)
                                                {
                                                    scope.LogWarning("Found a WebPart ID which is not available on the server-side. ID: {0}", serverSideControlId);
                                                }
                                            }
                                        }

                                        wikiPage.Fields.Add("WikiField", pageContents);
                                        template.Pages.Add(wikiPage);
                                    }
                                    else
                                    {
                                    }
                                }
                            }
                            catch (ServerException ex)
                            {

                                //ignore this error. The default page is not a page but a list view.
                                if (ex.ServerErrorCode != -2146232832 && ex.HResult != -2146233088)
                                {
                                    throw;
                                }
                                else
                                {
                                    if (ex.HResult != -2146233088)
                                    {
                                        if (webExtract.Context.HasMinimalServerLibraryVersion(Constants.MINIMUMZONEIDREQUIREDSERVERVERSION) || creationInformation.SkipVersionCheck)
                                        {
                                        }
                                        else
                                        {
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return extractTemplate;
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
                if (Uri.TryCreate(webUrl, UriKind.Absolute, out Uri uri))
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

        private string TokenizeWebPartXml(Web web, string xml)
        {
            var lists = web.Lists;
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id);
            web.Context.Load(lists, ls => ls.Include(l => l.Id, l => l.Title));
            web.Context.ExecuteQueryRetry();

            foreach (var list in lists)
            {
                xml = Regex.Replace(xml, list.Id.ToString(), $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
            }

            //some webparts already contains the site URL using ~sitecollection token (i.e: CQWP)
            xml = Regex.Replace(xml, "\"~sitecollection/(.)*\"", "\"{site}\"", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, "'~sitecollection/(.)*'", "'{site}'", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, ">~sitecollection/(.)*<", ">{site}<", RegexOptions.IgnoreCase);

            xml = Regex.Replace(xml, web.Id.ToString(), "{siteid}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, "(\"" + web.ServerRelativeUrl + ")(?!&)", "\"{site}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, "'" + web.ServerRelativeUrl, "'{site}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, ">" + web.ServerRelativeUrl, ">{site}", RegexOptions.IgnoreCase);
            return xml;
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
