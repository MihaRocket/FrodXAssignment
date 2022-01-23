using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Plugin
{
    public class UpdatePlugin : IPlugin
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


                    string fetchXML = $@"  
                       <fetch mapping='logical'>  
                         <entity name='new_hubspot_sync_log'>   
                            <attribute name='new_parent_contactid'/>
                            <attribute name='new_date_and_time_of_sync'/>
                            <filter>
                                <condition attribute='new_parent_contactid' operator='eq' value='{entity.Id}' />
                            </filter>
                         </entity>   
                       </fetch> ";

                    EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXML));

                    if (result.Entities.Count != 0)
                    {
                        foreach (var resultEntity in result.Entities)
                        {
                            resultEntity["new_date_and_time_of_sync"] = DateTime.Now;

                            service.Update(resultEntity);

                            tracingService.Trace("UpdatePlugin: Hubspot record updated.");
                        }
                    }
                    else
                    {
                        Entity hubspotEntity = new Entity("new_hubspot_sync_log");
                        hubspotEntity["new_date_and_time_of_sync"] = DateTime.Now;

                        Guid regardingobjectid = entity.Id;
                        string regardingobjectidType = "contact";

                        hubspotEntity["new_parent_contactid"] = new EntityReference(regardingobjectidType, regardingobjectid);

                        service.Create(hubspotEntity);

                        tracingService.Trace("UpdatePlugin: Hubspot record created.");
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
