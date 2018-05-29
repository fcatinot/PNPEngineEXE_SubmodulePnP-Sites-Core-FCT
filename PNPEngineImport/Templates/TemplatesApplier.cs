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

namespace PNPEngineImport.Templates
{
    public static class TemplatesApplier
    {
        public static void Apply(Web web, string templatePath, List<string> templateNames)
        {
            foreach (string templateName in templateNames)
            {
                Apply(web, templatePath, templateName);
            }
        }

        public static void Apply(Web web, string templatePath, string templateName)
        {
            try
            {
                LogWriter.Current.WriteLine("Applying " + templateName);
                // Connector to file system. Use current .exe folder as root, don't specify a subfolder
                var connector = new FileSystemConnector(templatePath, string.Empty);
                // Provider to get template
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider();
                provider.Connector = connector;
                var result = provider.GetTemplate(templatePath + "\\" + templateName);

                // Connector needs to be specified once more in ProvisioningTemplate because he is not copied from the provider on template creation
                result.Connector = connector;

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

                //Applying template
                ProvisioningTemplateApplyingInformation applyingInformation = new ProvisioningTemplateApplyingInformation()
                {
                    ExtensibilityHandlers = new List<ExtensibilityHandler> { CopyWikiPagesExtensibilityHandler, CopyElementsExtensibilityHandler, CopyModernPagesExtensibilityHandler, ConfigureXmlToImportExportExtensibilityHandler, CopyListTemplateExtensibilityHandler },
                    ClearNavigation = false,
                    IgnoreDuplicateDataRowErrors = false,
                    OverwriteSystemPropertyBagValues = false,
                    HandlersToProcess = Handlers.All,
                    MessagesDelegate = delegate (string message, ProvisioningMessageType messageType)
                    {
                        LogWriter.Current.WriteLine(string.Format("{0}:{1}", messageType.ToString(), message));
                    },
                    ProgressDelegate = delegate (string message, int progress, int total)
                    {
                        LogWriter.Current.WriteLine(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
                    }
                };

                web.ApplyProvisioningTemplate(result, applyingInformation);
            }
            catch (Exception ex)
            {
                LogWriter.Current.WriteLine(string.Format("Error applying template {0} : {1} - {2}", templateName, ex.Message, ex.StackTrace));
            }
        }
    }
}
