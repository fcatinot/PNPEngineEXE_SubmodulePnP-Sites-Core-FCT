using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
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
                var connector = new FileSystemConnector(LocalFilePaths.LocalPath, string.Empty);
                // Provider to get template
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider();
                provider.Connector = connector;
                var result = provider.GetTemplate(templatePath + "\\" + templateName);

                // Connector needs to be specified once more in ProvisioningTemplate because he is not copied from the provider on template creation
                result.Connector = connector;
                var applyingInformation = new ProvisioningTemplateApplyingInformation();
                applyingInformation.ProgressDelegate = (message, step, total) =>
                {
                    LogWriter.Current.WriteLine(string.Format("{0}/{1} Provisioning {2}", step, total, message));
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
