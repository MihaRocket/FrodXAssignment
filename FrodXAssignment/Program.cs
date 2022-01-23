using Microsoft.Xrm.Tooling.Connector;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using CrmEarlyBound;
using System;

namespace CRM
{
    class Programs
    {
        static void Main(string[] args)
        {
            string url = "https://org9210cc95.crm4.dynamics.com/";
            string userName = "MihaNabergoj@FrodX254.onmicrosoft.com";
            string password = "Geslo123";

            string conn =
                $@" Url = {url};
                    AuthType = OAuth;
                    UserName = {userName};
                    Password = {password};
                    AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;
                    RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;
                    LoginPrompt=Auto;
                    RequireNewInstance = True";

            CallCRM(conn);
        }

        private static void CallCRM(string conn)
        {
            using (var svc = new CrmServiceClient(conn))
            {
                CreateAttributes(svc);

                CreateHubspotEntity(svc);

                CreaterFieldForHubspot(svc);

                CreateRelationship(svc);
            }
        }

        static void CreateAttributes(CrmServiceClient svc)
        {
            var addedAttributes = new List<AttributeMetadata>();

            StringAttributeMetadata stringAttribute = new StringAttributeMetadata
            {
                SchemaName = "new_comment_field",
                DisplayName = new Label("Comment", 1033),
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                Description = new Label("Comment", 1033),
                MaxLength = 100
            };

            addedAttributes.Add(stringAttribute);

            BooleanAttributeMetadata booleanAttribute = new BooleanAttributeMetadata
            {
                SchemaName = "new_hubspot_integration",
                DisplayName = new Label("For Hubspot integration", 1033),
                Description = new Label("Hubspot integration", 1033),
                OptionSet = new BooleanOptionSetMetadata(
                            new OptionMetadata(new Label("True", 1033), 1),
                            new OptionMetadata(new Label("False", 1033), 0)
                            )
            };

            addedAttributes.Add(booleanAttribute);


            foreach (var attribute in addedAttributes)
            {
                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                {
                    EntityName = "contact",
                    Attribute = attribute
                };

                try
                {
                    svc.Execute(createAttributeRequest);
                }
                catch (Exception ex){ }
            }
        }
        static void CreateHubspotEntity(CrmServiceClient svc)
        {
            CreateEntityRequest createrequest = new CreateEntityRequest
            {
                //Define the entity
                Entity = new EntityMetadata
                {
                    SchemaName = "new_hubspot_sync_log",
                    DisplayName = new Label("Hubspot sync log", 1033),
                    DisplayCollectionName = new Label("Hubspot sync log", 1033),
                    Description = new Label("Hubspot sync log", 1033),
                    OwnershipType = OwnershipTypes.UserOwned,
                    IsActivity = false,

                },

                PrimaryAttribute = new StringAttributeMetadata
                {
                    SchemaName = "new_hubspotsynclog",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 100,
                    FormatName = StringFormatName.Text,
                    DisplayName = new Label("Hubspot sync log", 1033),
                    Description = new Label("The primary attribute for the Hubspot sync log entity.", 1033)
                }
            };

            try
            {
                svc.Execute(createrequest);
            }
            catch (Exception ex) { }
        }
        static void CreaterFieldForHubspot(CrmServiceClient svc)
        {
            DateTimeAttributeMetadata dateTimeAttributeMetadata = new DateTimeAttributeMetadata
            {
                SchemaName = "new_date_and_time_of_sync",
                DisplayName = new Label("Date and time of sync", 1033),
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                Description = new Label("Date and time of sync", 1033),
                Format = DateTimeFormat.DateAndTime
            };

            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
            {
                EntityName = "new_hubspot_sync_log",
                Attribute = dateTimeAttributeMetadata
            };

            try
            {
                svc.Execute(createAttributeRequest);
            }
            catch (Exception ex) { }
        }
        static void CreateRelationship(CrmServiceClient svc)
        {
            CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
            {
                OneToManyRelationship = new OneToManyRelationshipMetadata
                {
                    ReferencedEntity = "contact",
                    ReferencingEntity = "new_hubspot_sync_log",
                    SchemaName = "new_hubspot_ontomany_relationship",
                    AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                    {
                        Behavior = AssociatedMenuBehavior.UseLabel,
                        Group = AssociatedMenuGroup.Details,
                        Label = new Label("Contact", 1033),
                        Order = 10000
                    },
                    CascadeConfiguration = new CascadeConfiguration
                    {
                        Assign = CascadeType.NoCascade,
                        Delete = CascadeType.RemoveLink,
                        Merge = CascadeType.NoCascade,
                        Reparent = CascadeType.NoCascade,
                        Share = CascadeType.NoCascade,
                        Unshare = CascadeType.NoCascade
                    }
                },
                Lookup = new LookupAttributeMetadata
                {
                    SchemaName = "new_parent_contactid",
                    DisplayName = new Label("Contact Lookup", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("Sample Lookup", 1033)
                }
            };

            try
            {
                svc.Execute(createOneToManyRelationshipRequest);
            }
            catch (Exception ex) { }
        }
        static void CreateSavedQuery(CrmServiceClient svc)
        {
            string layoutXml = @"<grid name='resultset' object='2' jump='name' select='1' preview='1' icon='1'>
                                    <row name='CONTACT_INFORMATION' id='CONTACT_INFORMATION'>
                                        <cell name='new_comment_field' width='150' />
                                        <cell name='new_hubspot_integration' width='150' />
                                    </row>
                                </grid>";

            string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='contact'>
                                        <order attribute='new_string' descending='false' />
                                        <attribute name='new_string' />
                                        <attribute name='contactid' /> 
                                    </entity>
                                </fetch>";

            SavedQuery sq = new SavedQuery
            {
                Name = "Test public view",
                Description = "Created in Visual studio",
                ReturnedTypeCode = "contact",
                FetchXml = fetchXml,
                LayoutXml = layoutXml,
                QueryType = 0
            };

            svc.Create(sq);
        }
    }
}
