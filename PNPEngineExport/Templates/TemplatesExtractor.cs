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
                // Get source template
                var sourceCreationInformation = new ProvisioningTemplateCreationInformation(web)
                {
                    HandlersToProcess = DefaultHandlers,
                    
                };
                sourceCreationInformation.ProgressDelegate = (message, step, total) =>
                {
                    LogWriter.Current.WriteLine(string.Format("{0}/{1} Extracting {2}", step, total, message));
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
