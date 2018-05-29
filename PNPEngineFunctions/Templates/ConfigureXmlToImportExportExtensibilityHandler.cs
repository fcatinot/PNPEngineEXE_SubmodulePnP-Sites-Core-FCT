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
    public class ConfigureXmlToImportExportExtensibilityHandler : IProvisioningExtensibilityHandler
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
            //toutes les vues = tous les évenements
            //template.Lists.Where(l => l.TemplateType.ToString() == "106").ToList().ForEach(sel => sel.Views.ToList().Where(v => v.SchemaXml.ToString().Contains("<DateRangesOverlap>")).ToList().ForEach(vw => vw.SchemaXml = vw.SchemaXml = sel.Views[0].SchemaXml.ToString()));
            //remove old navigation
            template.Navigation.GlobalNavigation.StructuralNavigation.RemoveExistingNodes = true;
            template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes = true;
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

    }
}