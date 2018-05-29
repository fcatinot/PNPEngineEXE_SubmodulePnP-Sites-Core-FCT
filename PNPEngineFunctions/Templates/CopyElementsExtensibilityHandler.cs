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

namespace PNPEngineFunctions.Templates
{
    public class CopyElementsExtensibilityHandler : IProvisioningExtensibilityHandler
    {
        private Web webProvision;
        private Web webExtract;
        private string configurationXml;
        private ProvisioningTemplate extractTemplate;
        private ProvisioningTemplate provisionTemplate;
        string FolderConnecter;

        public ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template,
         ProvisioningTemplateCreationInformation creationInformation,
         PnPMonitoredScope scope, string configurationData)
        {
            //initialise Template
            extractTemplate = template;
            //Ignore security list
            extractTemplate.Lists.Where(w => w.Security != null).ToList().ForEach(x => x.Security = new ObjectSecurity());
            //traitement
            webExtract = ctx.Web;
            ctx.Load(webExtract);
            ctx.ExecuteQueryRetry();
            //SetFolderConnector
            FolderConnecter = template.Connector.Parameters["Container"].ToString().Split('/').LastOrDefault();
            //Copy Element
            CopyContentToTemplate();
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
            //webProvision = ctx.Web;
            //provisionTemplate = template;

            //List<string> Files = provisionTemplate.Connector.GetFiles();
            //for (int i = 0; i < Files.Count; i++)
            //{
            //    provisionTemplate.Connector.DeleteFile(Files[i]);
            //}
            //configurationXml = configurationData;
        }

        private void CopyContentToTemplate()
        {
            var excludedUrls = new string[] { "Style Library", "Lists/PublishedFeed", "Reports List", "WorkflowTasks", "FormServerTemplates" };
            //copie des fichiers
            foreach (ListInstance templateList in extractTemplate.Lists)
            {
                if (!excludedUrls.Contains(templateList.Url)) {
                    if (new[] { 101, 109 }.Contains(templateList.TemplateType))
                    {
                        CopyFilesToTemplate(webExtract, templateList.Url);
                    }
                    //Ajouter templateType si nécessaire et vérifié
                    if (new[] { 100, 102, 103, 104, 105, 106, 107, 108, 150, 170, 171, 200, 201, 202 }.Contains(templateList.TemplateType))
                    {
                        CopyItemsToTemplate(webExtract, templateList.Url);
                    }
                }
            }
        }

        private void CopyFilesToTemplate(Web sourceWeb, string folderWebRelativeUrl)
        {
            List sourceList = sourceWeb.GetListByUrl(folderWebRelativeUrl);
            Microsoft.SharePoint.Client.Folder sourceFolder = sourceWeb.GetFolderByServerRelativeUrl(sourceWeb.ServerRelativeUrl.TrimEnd('/') + folderWebRelativeUrl);
            CopyFolderContentToTemplate(sourceWeb, sourceList, sourceFolder, "");
        }

        private void CopyItemsToTemplate(Web sourceWeb, string listWebRelativeUrl)
        {
            //load all items to copy
            List sourceList = sourceWeb.GetListByUrl(listWebRelativeUrl);
            ListItemCollection itemsToMigrate = sourceList.GetItems(CamlQuery.CreateAllItemsQuery());
            sourceWeb.Context.Load(itemsToMigrate);
            sourceWeb.Context.ExecuteQueryRetry();
            CopyItemContentToTemplate(sourceWeb, sourceList, itemsToMigrate, listWebRelativeUrl);
        }

        private void CopyFolderContentToTemplate(Web sourceWeb, List sourceList, Microsoft.SharePoint.Client.Folder sourceFolder, string directories)
        {
            sourceFolder.EnsureProperty(f => f.Name);
            if (sourceFolder.Name.ToString() != FolderConnecter && sourceFolder.Name.ToString() != "Forms" && !sourceFolder.Name.StartsWith("_"))
            {
                Microsoft.SharePoint.Client.FileCollection sourceFiles = sourceFolder.Files;
                sourceFolder.Context.Load(sourceFiles);
                sourceFolder.Context.ExecuteQueryRetry();

                if (String.IsNullOrEmpty(directories))
                {
                    directories = sourceFolder.Name;
                }
                else
                {
                    directories = directories + "/" + sourceFolder.Name;
                }

                foreach (var file in sourceFiles)
                {
                    var fileStream = file.OpenBinaryStream();
                    sourceFolder.Context.Load(file.ListItemAllFields);
                    sourceFolder.Context.ExecuteQuery();

                    OfficeDevPnP.Core.Framework.Provisioning.Model.File targetFile =
                    new OfficeDevPnP.Core.Framework.Provisioning.Model.File
                    {
                        Folder = directories,
                        Src = file.Name,
                        Overwrite = true,
                        Level = (OfficeDevPnP.Core.Framework.Provisioning.Model.FileLevel)file.Level
                    };
                    populateProperties(sourceWeb, file.ListItemAllFields, sourceList, targetFile.Properties);
                    extractTemplate.Files.Add(targetFile);

                    try
                    {
                        extractTemplate.Connector.SaveFileStream(file.Name, fileStream.Value);
                    }
                    catch (Exception ex)
                    {
                        //The file name might contain encoded characters that prevent upload. Decode it and try again.
                        var fileName = file.Name.Replace("&", "");
                        extractTemplate.Connector.SaveFileStream(file.Name, fileStream.Value);
                    }
                }

                Microsoft.SharePoint.Client.FolderCollection sourceSubfolders = sourceFolder.Folders;
                sourceFolder.Context.Load(sourceSubfolders);
                sourceFolder.Context.ExecuteQueryRetry();

                foreach (var sourceSubFolder in sourceSubfolders)
                {
                    CopyFolderContentToTemplate(sourceWeb, sourceList, sourceSubFolder, directories);
                }
            }
        }

        private void CopyItemContentToTemplate(Web sourceWeb, Microsoft.SharePoint.Client.List sourceList, Microsoft.SharePoint.Client.ListItemCollection itemsToMigrate, string listWebRelativeUrl)
        {
            foreach (ListItem item in itemsToMigrate)
            {
                if (item.FileSystemObjectType != FileSystemObjectType.Folder)
                {
                    DataRow itemData =  new DataRow{};
                    populateProperties(sourceWeb, item, sourceList, itemData.Values);
                    if (itemData.Values.Count > 0)
                    {
                        extractTemplate.Lists.Where(w => w.Url == listWebRelativeUrl).FirstOrDefault().DataRows.Add(itemData);
                    }
                }
                else
                {
                    OfficeDevPnP.Core.Framework.Provisioning.Model.Folder folder =
                    new OfficeDevPnP.Core.Framework.Provisioning.Model.Folder
                    {
                        Name = item.FieldValues["Title"].ToString()
                    };
                    extractTemplate.Lists.Where(w => w.Url == listWebRelativeUrl).FirstOrDefault().Folders.Add(folder);
                }

            }
        }

        private static void populateProperties(Web sourceWeb, ListItem item, List sourceList, Dictionary<string, string> target)
        {
            sourceWeb.Context.Load(sourceList.Fields);
            sourceWeb.Context.ExecuteQueryRetry();
            foreach (Microsoft.SharePoint.Client.Field field in sourceList.Fields)
            {
                if (!field.ReadOnlyField && field.InternalName != "Attachments" && field.InternalName != "MetaInfo" && field.InternalName != "ContentType" && field.InternalName != "FileLeafRef")
                {
                    try
                    {      
                        string value = returnString(sourceWeb, item, field.InternalName);
                        if (!String.IsNullOrEmpty(value))
                        {
                            target.Add(field.InternalName, value);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        private static string returnString(Web sourceWeb, ListItem item, string internalName)
        {
            if (item[internalName] == null) { return ""; }
            //if (internalName == "FileLeafRef") { internalName = "Title"; }
            if (item[internalName] is FieldUserValue)
            {
                return ((FieldUserValue)item[internalName]).LookupValue;
            }
            if (item[internalName] is FieldUserValue[])
            {
                var lstuser = "";
                foreach (var user in (FieldUserValue[])item[internalName])
                    lstuser += user.LookupValue + ",";
                char[] charsToTrim = { ',' };
                return lstuser.Trim(charsToTrim);
            }
            if (item[internalName] is FieldLookupValue)
            {
                return ((FieldLookupValue)item[internalName]).LookupId.ToString();
            }
            if (item[internalName] is FieldLookupValue[])
            {
                var flv = (FieldLookupValue[])item[internalName];
                string multilookup = "";
                foreach (var lookupValue in flv)
                {
                    multilookup += lookupValue.LookupId.ToString();
                }
                char[] charsToTrim = { ',' };
                multilookup.Trim(charsToTrim);
                return multilookup;
            }
            if (item[internalName] is TaxonomyFieldValueCollection)
            {
                var lsttaxo = "";
                foreach (var taxo in ((TaxonomyFieldValueCollection)item[internalName]).ToList().Select(t => t.Label).ToArray())
                    lsttaxo += taxo + ";";
                return lsttaxo;
            }
            if (item[internalName] is TaxonomyFieldValue)
            {
                return ((TaxonomyFieldValue)item[internalName]).Label;
            }
            if (item[internalName] is FieldUrlValue)
            {
                return ((FieldUrlValue)item[internalName]).Url + "," + ((FieldUrlValue)item[internalName]).Description;
            }
            if (item[internalName] is DateTime)
            {
                DateTime date = (DateTime)item[internalName];
                return date.ToLocalTime().ToString();
            }
            if (item[internalName] is String)
            {
                return item[internalName].ToString();
            }
            if (item[internalName] is Boolean)
            {
                return item[internalName].ToString();
            }
            if (item[internalName] is Decimal)
            {
                return item[internalName].ToString().Replace(",", ".");
            }
            return "";
        }


        private bool ListShouldBeExported(ListInstance l)
        {
            var excludedUrls = new string[] { "Style Library", "Lists/PublishedFeed", "Reports List", "WorkflowTasks", "FormServerTemplates" };
            if (excludedUrls.Contains(l.Url) )
            {
                return false;
            }
            return true;
        }

    }
}