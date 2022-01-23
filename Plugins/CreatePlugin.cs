using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;

namespace Plugin
{
    public class CreatePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    if (!entity.Attributes.Keys.Contains("new_comment_field") && (bool)entity.Attributes["new_hubspot_integration"] == false)
                    {
                        Entity hubspotEntity = new Entity("new_hubspot_sync_log");
                        hubspotEntity["new_date_and_time_of_sync"] = DateTime.Now;

                        if (context.OutputParameters.Contains("id"))
                        {
                            Guid regardingobjectid = new Guid(context.OutputParameters["id"].ToString());
                            string regardingobjectidType = "contact";

                            hubspotEntity["new_parent_contactid"] = new EntityReference(regardingobjectidType, regardingobjectid);
                        }

                        //tracingService.Trace("FollowupPlugin: Creating the task activity.");
                        service.Create(hubspotEntity);
                        tracingService.Trace("FollowupPlugin: Hubspot record created.");
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                    throw ex;
                }
            }
        }
    }
}
