using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using PNPEngineFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNPEngineExport.Templates
{
    public static class TemplatesExtractor
    {
        public static void Apply(Web web, string templatePath, string templateFileName)
        {
            try
            {
                Handlers DefaultHandlers = Handlers.All;

                //Connector in Sharepoint Site
                //string _containerFolder = PNPEngineFunctions.ContainerFolder.SetContainerFolder(web, connectionSettings);
                //var connector = new SharePointConnector(web.Context, web.Url, _containerFolder);
                var connector = new FileSystemConnector(templatePath, string.Empty);

                ExtensibilityHandler CopyWikiPagesExtensibilityHandler = new ExtensibilityHandler
                {
                    Type = "PNPEngineFunctions.Templates.CopyWikiPagesExtensibilityHandler",
                    Assembly = "PNPEngineFunctions",
                    Enabled = true,
                    Configuration = ""
                };

                ExtensibilityHandler CopyElementsExtensibilityHandler = new ExtensibilityHandler
                {
                    Type = "PNPEngineFunctions.Templates.CopyElementsExtensibilityHandler",
                    Assembly = "PNPEngineFunctions",
                    Enabled = true,
                    Configuration = ""
                };

                ExtensibilityHandler CopyModernPagesExtensibilityHandler = new ExtensibilityHandler
                {
                    Type = "PNPEngineFunctions.Templates.CopyModernPagesExtensibilityHandler",
                    Assembly = "PNPEngineFunctions",
                    Enabled = true,
                    Configuration = ""
                };

                ExtensibilityHandler ConfigureXmlToImportExportExtensibilityHandler = new ExtensibilityHandler
                {
                    Type = "PNPEngineFunctions.Templates.ConfigureXmlToImportExportExtensibilityHandler",
                    Assembly = "PNPEngineFunctions",
                    Enabled = true,
                    Configuration = ""
                };

                ExtensibilityHandler CopyListTemplateExtensibilityHandler = new ExtensibilityHandler
                {
                    Type = "PNPEngineFunctions.Templates.CopyListTemplateExtensibilityHandler",
                    Assembly = "PNPEngineFunctions",
                    Enabled = true,
                    Configuration = ""
                };

                // Get source template
                var sourceCreationInformation = new ProvisioningTemplateCreationInformation(web)
                {
                    HandlersToProcess = DefaultHandlers,
                    FileConnector = connector,
                    ExtensibilityHandlers = new List<ExtensibilityHandler> { CopyWikiPagesExtensibilityHandler, CopyElementsExtensibilityHandler, CopyModernPagesExtensibilityHandler, ConfigureXmlToImportExportExtensibilityHandler, CopyListTemplateExtensibilityHandler },
                    MessagesDelegate = delegate (string message, ProvisioningMessageType messageType)
                    {
                        LogWriter.Current.WriteLine(string.Format("{0}:{1}", messageType.ToString(), message));
                    },
                    ProgressDelegate = delegate (string message, int progress, int total)
                    {
                        LogWriter.Current.WriteLine(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
                    }
                };

                // Get template from existing site
                ProvisioningTemplate template = web.GetProvisioningTemplate(sourceCreationInformation);

                // Save template using XML provider
                templatePath = templatePath.ToLower();
                XMLFileSystemTemplateProvider provider = new XMLFileSystemTemplateProvider(@""+templatePath, "");
                provider.SaveAs(template, templateFileName);
               
            }
            catch (Exception ex)
            {
                LogWriter.Current.WriteLine(string.Format("Error extracting template {0} : {1} - {2}", templateFileName, ex.Message, ex.StackTrace));
            }
        }
    }
}
