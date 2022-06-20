using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Utilities.Helpers
{

    public static class CRMHelper
    {
        public static Dictionary<string, string> GetConfigurationValues(IOrganizationService service, string[] configKeys)
        {
            try
            {
                var query = new QueryExpression("moci_systemconfig")
                {
                    ColumnSet = new ColumnSet("moci_name", "moci_value")
                };

                query.Criteria.AddCondition("moci_name", ConditionOperator.In, configKeys);

                EntityCollection result = service.RetrieveMultiple(query);


                if (result.Entities.Count != 0)
                {
                    Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();

                    foreach (var record in result.Entities)
                    {
                        var key = record.GetAttributeValue<string>("moci_name");

                        var value = record.GetAttributeValue<string>("moci_value");

                        keyValuePairs.Add(key, value);
                    }

                    return keyValuePairs;
                }
            }
            catch (Exception ex)
            {
            }

            return null;
        }

        public static EntityCollection DefaultGetByMultipleAttributes(IOrganizationService service, string entityName, Dictionary<string, object> paramters, IEnumerable<LinkedEntity> linkedEntities)
        {
            try
            {
                var query = new QueryExpression(entityName)
                {
                    ColumnSet = new ColumnSet(true)
                };

                var dictionaryKeys = paramters.Keys;

                foreach (var paramter in paramters)
                {
                    query.Criteria.Conditions.Add(new ConditionExpression(paramter.Key, ConditionOperator.Equal, paramter.Value));
                }

                foreach (var item in linkedEntities)
                {
                    var bu = query.AddLink(item.LinkToEntityName, item.LinkFromAttributeName, item.LinkToAttributeName, item.JoinOperator);

                    bu.EntityAlias = item.EntityAlias;

                    bu.LinkCriteria = item.Filters;

                    bu.Columns = new ColumnSet(true);
                }

                return service.RetrieveMultiple(query);
            }
            catch (Exception ex)
            {
            }

            return new EntityCollection();
        }

        public static EntityCollection DefaultGetByFetchXML(IOrganizationService service, string fetchXML)
        {
            try
            {
                return service.RetrieveMultiple(new FetchExpression(fetchXML));
            }
            catch (Exception ex)
            {
            }

            return new EntityCollection();
        }


        public static EntityCollection GetEntitiesBy(this IOrganizationService pService, string pEntityName, string[] pSearchColumn, object[] pSearchValue,
                            string pOrderby = "", OrderType ptype = OrderType.Descending, List<string> pAllCoulmn = null, bool pUsenonlook = false)
        {
            if (pOrderby == null) throw new ArgumentNullException(nameof(pOrderby));
            QueryExpression query = new QueryExpression(pEntityName);

            query.ColumnSet = pAllCoulmn == null ? new ColumnSet(true) : new ColumnSet(pAllCoulmn.ToArray());
            if (pOrderby != String.Empty)
                query.AddOrder(pOrderby, ptype);

            if (pUsenonlook)
                query.NoLock = true;

            for (int i = 0; i < pSearchColumn.Length; i++)
            {
                if (pSearchValue[i] is object[])
                {
                    query.Criteria.AddCondition(new ConditionExpression(pSearchColumn[i], ConditionOperator.In, (object[])pSearchValue[i]));
                }
                else
                {
                    query.Criteria.AddCondition(new ConditionExpression(pSearchColumn[i], ConditionOperator.Equal, pSearchValue[i]));
                }
            }

            return pService.RetrieveMultiple(query);

        }

        public static Entity GetOneEntityBy(this IOrganizationService pService, string pEntityName, string[] pSearchColumn, object[] pSearchValue
         , List<string> pAllCoulmn = null, bool pUsenonlook = false, string pOrderby = "", OrderType ptype = OrderType.Descending)
        {
            QueryExpression query = new QueryExpression(pEntityName);

            query.ColumnSet = pAllCoulmn == null ? new ColumnSet(true) : new ColumnSet(pAllCoulmn.ToArray());
            if (pOrderby != String.Empty)
                query.AddOrder(pOrderby, ptype);


            for (int i = 0; i < pSearchColumn.Length; i++)
            {
                if (pSearchValue[i] is object[])
                {
                    query.Criteria.AddCondition(new ConditionExpression(pSearchColumn[i], ConditionOperator.In, (object[])pSearchValue[i]));
                }
                else
                {
                    query.Criteria.AddCondition(new ConditionExpression(pSearchColumn[i], ConditionOperator.Equal, pSearchValue[i]));
                }
            }

            if (pUsenonlook)
            {
                query.NoLock = true;
            }

            EntityCollection entities = pService.RetrieveMultiple(query);

            if (entities.Entities.Count > 0)
            {
                return entities.Entities[0];
            }

            return null;

        }

        public static EntityCollection GetEntitiesBy(this IOrganizationService pService, string pEntityName, string pSearchColumn, object pSearchValue)
        {
            return GetEntitiesBy(pService, pEntityName, new string[] { pSearchColumn }, new object[] { pSearchValue }, String.Empty, OrderType.Descending, null, false);

        }

        public static EntityCollection GetEntitiesBy(this IOrganizationService pService, string pEntityName, string[] pSearchColumn, object[] pSearchValue)
        {
            return GetEntitiesBy(pService, pEntityName, pSearchColumn, pSearchValue, String.Empty, OrderType.Descending, null, false);

        }

        public static EntityCollection GetEntitiesBy(this IOrganizationService pService, string pEntityName, string pSearchColumn, object pSearchValue,
            string pOrderby, OrderType ptype = OrderType.Descending)
        {
            return GetEntitiesBy(pService, pEntityName, new string[] { pSearchColumn }, new object[] { pSearchValue }, pOrderby, ptype, null, false);

        }


        public static Entity GetOneEntityBy(this IOrganizationService pService, string pEntityName, string pSearchColumn, object pSearchValue)
        {
            return GetOneEntityBy(pService, pEntityName, new string[] { pSearchColumn }, new object[] { pSearchValue });
        }

        public static Entity GetOneEntityBy(this IOrganizationService pService, string pEntityName, string pSearchColumn, object pSearchValue, List<string> pAllCoulmn, bool pUsenonlook = false)
        {
            return GetOneEntityBy(pService, pEntityName, new string[] { pSearchColumn }, new object[] { pSearchValue }, pAllCoulmn, pUsenonlook);
        }
        public static EntityReference GetEntityLookup(IOrganizationService pService, string pEntityName, string pSearchAttributeName, object pSearchAttributeValue)
        {
            EntityReference refLookup = null;
            try
            {
                string strEntityIdColumnName = pEntityName + "id";
                Entity resultEntity = null;
                QueryExpression query = new QueryExpression { EntityName = pEntityName, ColumnSet = new ColumnSet(strEntityIdColumnName) };
                query.Criteria = new FilterExpression();
                query.Criteria.FilterOperator = LogicalOperator.And;

                ConditionExpression condExp = new ConditionExpression();
                condExp.AttributeName = pSearchAttributeName;
                condExp.Operator = ConditionOperator.Equal;
                condExp.Values.Add(pSearchAttributeValue);
                query.Criteria.Conditions.Add(condExp);

                var results = pService.RetrieveMultiple(query);
                if (results.Entities.Count > 0)
                {
                    resultEntity = results.Entities[0];

                    if (resultEntity != null && resultEntity.Contains(strEntityIdColumnName))
                    {
                        refLookup = new EntityReference(pEntityName, resultEntity.Id);
                    }
                }
                return refLookup;
            }
            catch (Exception ex)
            {
                throw new Exception("Lookup Fetch : Lookup Entity Name: " + pEntityName + " Exception :" + ex.Message.ToString());
            }
        }
        public static object GetColumnValueFromEntity(this IOrganizationService pService, string pEntityName, string columnName, string pSearchColumn, object pSearchValue)
        {
            QueryExpression query = new QueryExpression(pEntityName);
            query.ColumnSet = new ColumnSet(new string[] { columnName });
            query.Criteria.AddCondition(new ConditionExpression(pSearchColumn, ConditionOperator.Equal, pSearchValue));
            EntityCollection entities = pService.RetrieveMultiple(query);
            if (entities.Entities.Count > 0)
            {
                return entities.Entities[0][columnName];
            }
            return null;
        }

        public static bool AssociateManyToManyEntityRecords(this IOrganizationService pService, EntityReference pMoniker1,
            EntityReference pMoniker2, string strEntityRelationshipName)
        {
            try
            {
                AssociateRequest request = new AssociateRequest
                {
                    Target = new EntityReference(pMoniker1.LogicalName, pMoniker1.Id),
                    RelatedEntities = new EntityReferenceCollection
                {
                    new EntityReference(pMoniker2.LogicalName, pMoniker2.Id)
                },
                    Relationship = new Relationship(strEntityRelationshipName)
                };

                // Execute the request.
                pService.Execute(request);


                return true;
            }


            catch (Exception)
            {
                return false;
            }

        }

        public static ExecuteMultipleResponse BulkAction<T>(this IOrganizationService pService, List<Entity> entities)
            where T : OrganizationRequest
        {
            ExecuteMultipleResponse multipleResponse = null;
            try
            {

                // Create an ExecuteMultipleRequest object.
                var multipleRequest = new ExecuteMultipleRequest()
                {
                    // Assign settings that define execution behavior: continue on error, return responses. 
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };


                // Add a CreateRequest for each entity to the request collection.
                foreach (var entity in entities)
                {


                    switch (typeof(T).Name.ToLower())
                    {

                        case "createrequest":
                            CreateRequest createRequest = new CreateRequest { Target = entity };
                            multipleRequest.Requests.Add(createRequest);
                            break;
                        case "updaterequest":
                            UpdateRequest updateRequest = new UpdateRequest { Target = entity };
                            multipleRequest.Requests.Add(updateRequest);
                            break;
                        case "deleterequest":
                            DeleteRequest deleteRequest = new DeleteRequest { Target = new EntityReference(entity.LogicalName, entity.Id) };
                            multipleRequest.Requests.Add(deleteRequest);
                            break;


                    }


                }

                // Execute all the requests in the request collection using a single web method call.
                multipleResponse = (ExecuteMultipleResponse)pService.Execute(multipleRequest);
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message.ToString());
            }

            return multipleResponse;
        }

        public static void BulkCreate(this IOrganizationService pService, List<Entity> entities)
        {
            try
            {

                // Create an ExecuteMultipleRequest object.
                BulkAction<CreateRequest>(pService, entities);
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message.ToString());
            }
        }

        /// <summary>
        /// Call this method for bulk update
        /// </summary>
        /// <param name="pService">Org pService</param>
        /// <param name="entities">Collection of entities to Update</param>
        public static void BulkUpdate(this IOrganizationService pService, List<Entity> entities)
        {
            try
            {
                // Create an ExecuteMultipleRequest object.
                BulkAction<UpdateRequest>(pService, entities);

            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message.ToString());
            }

        }

        /// <summary>
        /// Call this method for bulk delete
        /// </summary>
        /// <param name="pService">Org pService</param>
        /// <param name="entityReferences">Collection of EntityReferences to Delete</param>
        public static void BulkDelete(this IOrganizationService pService, List<EntityReference> entityReferences)
        {
            try
            {





                List<Entity> entities = new List<Entity>();
                entityReferences.ForEach(c => entities.Add(new Entity() { LogicalName = c.LogicalName, Id = c.Id }));
                BulkAction<DeleteRequest>(pService, entities);
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message.ToString());
            }
        }
        /// <summary>
        /// Attach a note to the referenced Entity.
        /// </summary>
        /// <param name="pService"></param>
        /// <param name="Regarding">EntityReference of the Entity that the note will be attached to.</param>
        /// <param name="Subject">Subject line of the note</param>
        /// <param name="Text">Text of the note</param>
        public static void AttachNote(this IOrganizationService pService, EntityReference Regarding, string Subject, string Text)
        {
            Entity note = new Entity("annotation");
            note.Attributes.Add("subject", Subject);
            note.Attributes.Add("notetext", Text);
            note.Attributes.Add("objectid", Regarding);
            pService.Create(note);
        }


        /// <summary>
        /// Attach a note to the referenced Entity.
        /// </summary>
        /// <param name="pService"></param>
        /// <param name="EntityLogicalName">Logicl name of the Entity.</param>
        /// <param name="Id">GUID of the Entity</param>
        /// <param name="Subject">Subject line of the note.</param>
        /// <param name="Text">Text of the note.</param>
        public static void AttachNote(this IOrganizationService pService, string EntityLogicalName, Guid Id, string Subject, string Text)
        {
            EntityReference regarding = new EntityReference(EntityLogicalName, Id);
            pService.AttachNote(regarding, Subject, Text);
        }
        public static string GetCustomConfigValues(this IOrganizationService pService, string pConfigName)
        {
            string configValue = null;
            try
            {
                Entity configDetails = GetOneEntityBy(pService, "moci_systemconfig", "moci_name", pConfigName);
                if (configDetails != null)
                {
                    if (configDetails.Contains("moci_value"))
                        configValue = configDetails.Attributes["moci_value"].ToString();
                    else
                        throw new InvalidPluginExecutionException("Please configure the values for Configuration Name: " + pConfigName);
                }
                else
                {
                    throw new InvalidPluginExecutionException("Please configure the values for Configuration Name: " + pConfigName);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
            return configValue;
        }

        /// <summary>
        /// Returns a reference to the CRM Currency record identified by the specified currency code.
        /// </summary>
        /// <param name="pService"></param>
        /// <param name="currencyCode"></param>
        /// <returns></returns>
        public static EntityReference GetCurrencyByCode(this IOrganizationService pService, string currencyCode)
        {
            EntityReference rv = null;

            QueryByAttribute qry = new QueryByAttribute("transactioncurrency");
            qry.ColumnSet = new ColumnSet(new string[] { "transactioncurrencyid" });
            qry.AddAttributeValue("isocurrencycode", currencyCode);
            qry.AddAttributeValue("statecode", 0);

            EntityCollection results = pService.RetrieveMultiple(qry);
            if (results.Entities.Count == 1)
            {
                rv = results.Entities[0].ToEntityReference();
            }

            return rv;
        }

        /// <summary>
        /// Returns a reference to the CRM Currency record identified by USD currency code.
        /// </summary>
        /// <param name="pService"></param>
        /// <returns></returns>
        public static EntityReference GetCurrencyUSD(this IOrganizationService pService)
        {
            return GetCurrencyByCode(pService, "USD");
        }


        /// <summary>
        /// Returns an entity based on the target with additional column attributres merged in from the database if a 
        /// the attribues do not aleady exist in the target entity. pColumns can be specified as a ColumnSet object, a 
        /// string array, or a deliminated string. If pColumns are not provided, then all available columns 
        /// will be included in the merged record.
        /// </summary>
        /// <param name="service"></param>
        /// <param name="pService"></param>
        /// <param name="pTarget"></param>
        /// <param name="pColumns"></param>
        /// <returns></returns>
        public static Entity GetMergedRecord(this IOrganizationService pService, Entity pTarget, ColumnSet pColumns)
        {
            //create a new entity record and merge the target data into the new record
            Entity result = new Entity(pTarget.LogicalName);
            result.Id = pTarget.Id;
            result.MergeWith(pTarget);

            //merge the specified columns from the database into the new record.
            result.MergeWith(pService.GetEntity(pTarget.ToEntityReference(), pColumns));

            return result;
        }

        public static Entity GetEntity(this IOrganizationService pService, EntityReference pReference, ColumnSet pColumns)
        {
            return pService.Retrieve(pReference.LogicalName, pReference.Id, pColumns);
        }

        public static Entity GetMergedRecord(this IOrganizationService pService, Entity pTarget, string[] pColumns)
        {
            return GetMergedRecord(pService, pTarget, new ColumnSet(pColumns));
        }

        public static Entity GetMergedRecord(this IOrganizationService pService, Entity pTarget, string pColumns)
        {
            string[] cols = pColumns.Split(new char[] { ',', ';' });
            return GetMergedRecord(pService, pTarget, cols);
        }

        public static Entity GetMergedRecord(this IOrganizationService pService, Entity pTarget)
        {
            return GetMergedRecord(pService, pTarget, new ColumnSet(true));
        }

        /// <summary>
        /// Grants a specific user Read Access to a specific Entity record.
        /// </summary>
        /// <param name="pService"></param>
        /// <param name="pRecord">EntityReference of the record.</param>
        /// <param name="pUserId">GUID of user that is getting access.</param>
        public static void GrantUserReadAccess(this IOrganizationService pService, EntityReference pRecord, Guid pUserId)
        {

            PrincipalAccess principalAccess = new PrincipalAccess();
            principalAccess.Principal = new EntityReference("systemuser", pUserId);
            principalAccess.AccessMask = AccessRights.ReadAccess;

            GrantAccessRequest grantAccessRequest = new GrantAccessRequest();
            grantAccessRequest.Target = pRecord;
            grantAccessRequest.PrincipalAccess = principalAccess;

            GrantAccessResponse grantAccessResponse = (GrantAccessResponse)pService.Execute(grantAccessRequest);
        }

        /// <summary>
        /// Evaluates users role assignments and returns true if the user is assigned one or more of the role names provided.
        /// </summary>
        /// <param name="OrganizationService"></param>
        /// <param name="UserId">GUID of the user.</param>
        /// <param name="RoleNames">Name of the roles supplied in a string array.</param>
        /// <returns></returns>
        public static bool IsAssignedRole(this IOrganizationService pService, Guid UserId, string[] RoleNames)
        {
            //qry all roles linked to the user and return the name of the role
            QueryExpression qry = new QueryExpression()
            {
                EntityName = "role",
                ColumnSet = new ColumnSet(new string[] { "name" }),

                LinkEntities = {
                    new LinkEntity
                    {
                        JoinOperator = JoinOperator.Inner,
                        LinkFromEntityName = "role",
                        LinkFromAttributeName = "roleid",
                        LinkToEntityName = "systemuserroles",
                        LinkToAttributeName = "roleid",

                        LinkCriteria = new FilterExpression
                        {
                            FilterOperator = LogicalOperator.And,
                            Conditions =
                            {
                                new ConditionExpression("systemuserid",ConditionOperator.Equal,UserId)
                            }
                        }
                    }
                }
            };

            //create a filter to retrieve one or more role names
            FilterExpression roleFilter = new FilterExpression
            {
                IsQuickFindFilter = false,
                FilterOperator = LogicalOperator.Or,
            };

            foreach (string name in RoleNames)
            {
                roleFilter.Conditions.Add(new ConditionExpression("name", ConditionOperator.Equal, name));
            }

            //set the qry filter equal to the role filter
            qry.Criteria = roleFilter;

            EntityCollection result = pService.RetrieveMultiple(qry);

            return (result.Entities.Count >= 1);
        }


        /// <summary>
        /// Evaluates users role assignments and returns true if the user is assigned one or more of the role names provided.
        /// </summary>
        /// <param name="OrganizationService"></param>
        /// <param name="pUserId">GUID of the user.</param>
        /// <param name="pRoleName">Name of the role</param>
        /// <returns></returns>
        public static bool IsAssignedRole(this IOrganizationService pService, Guid pUserId, string pRoleName)
        {
            //convert the single role name to an array and call the overloaded function.
            string[] roleNames = new string[] { pRoleName };
            return pService.IsAssignedRole(pUserId, roleNames);
        }


        /// <summary>
        /// Evaluates a users membership in a team and returns true if the user is a team member.
        /// </summary>
        /// <param name="OrganizationService"></param>
        /// <param name="pSystemUser">EntityReference of the user.</param>
        /// <param name="pTeam">EntityReference of the team.</param>
        /// <returns></returns>
        public static bool IsTeamMember(this IOrganizationService pService, EntityReference pSystemUser, EntityReference pTeam)
        {
            if (pSystemUser == null || pSystemUser.LogicalName != "systemuser")
            {
                throw new Exception("SystemUser parameter must represent a CRM SystemUser record.");
            }

            if (pTeam == null || pTeam.LogicalName != "team")
            {
                throw new Exception("Team parameter must represent a team.");
            }

            QueryByAttribute qry = new QueryByAttribute("teammembership");
            qry.AddAttributeValue("teamid", pTeam.Id);
            qry.AddAttributeValue("systemuserid", pSystemUser.Id);
            qry.ColumnSet = new ColumnSet(new string[] { "teammembershipid" });

            EntityCollection results = pService.RetrieveMultiple(qry);
            if (results.Entities.Count > 0)
            {
                return true;
            }

            return false;
        }


        /// <summary>
        /// Sets the state of a record to the State Code and Status Code specified.
        /// </summary>
        /// <param name="OrganizationService"></param>
        /// <param name="Record">EntityReference of the record being modified.</param>
        /// <param name="pStateCode">StateCode value as an integer</param>
        /// <param name="pStatusCode">StatusCode value as an integer</param>
        public static void SetState(this IOrganizationService pService, EntityReference Record, int pStateCode, int pStatusCode)
        {
            SetStateRequest setStateReq = new SetStateRequest();
            setStateReq.EntityMoniker = Record;
            setStateReq.State = new Microsoft.Xrm.Sdk.OptionSetValue(pStateCode);
            setStateReq.Status = new Microsoft.Xrm.Sdk.OptionSetValue(pStatusCode);
            pService.Execute(setStateReq);
        }

        /// <summary>
        /// DeactivateRecord
        /// </summary>
        /// <param name="organizationService"></param>
        /// <param name="entityName"></param>
        /// <param name="recordId"></param>
        //Deactivate a record
        public static void DeactivateRecord(this IOrganizationService pService, string pEntityName, Guid pRecordId)
        {
            var cols = new ColumnSet(new[] { "statecode", "statuscode" });

            //Check if it is Active or not
            var entity = pService.Retrieve(pEntityName, pRecordId, cols);

            if (entity != null && entity.GetAttributeValue<Microsoft.Xrm.Sdk.OptionSetValue>("statecode").Value == 0)
            {
                //StateCode = 1 and StatusCode = 2 for deactivating Account or Contact
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = pRecordId,
                        LogicalName = pEntityName,
                    },
                    State = new Microsoft.Xrm.Sdk.OptionSetValue(1),
                    Status = new Microsoft.Xrm.Sdk.OptionSetValue(2)
                };
                pService.Execute(setStateRequest);
            }
        }

        /// <summary>
        /// ActivateRecord
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="recordId"></param>
        /// <param name="organizationService"></param>
        //Activate a record
        public static void ActivateRecord(this IOrganizationService pService, string pEntityName, Guid pRecordId)
        {
            var cols = new ColumnSet(new[] { "statecode", "statuscode" });

            //Check if it is Inactive or not
            var entity = pService.Retrieve(pEntityName, pRecordId, cols);

            if (entity != null && entity.GetAttributeValue<Microsoft.Xrm.Sdk.OptionSetValue>("statecode").Value == 1)
            {
                //StateCode = 0 and StatusCode = 1 for activating Account or Contact
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = pRecordId,
                        LogicalName = pEntityName,
                    },
                    State = new Microsoft.Xrm.Sdk.OptionSetValue(0),
                    Status = new Microsoft.Xrm.Sdk.OptionSetValue(1)
                };
                pService.Execute(setStateRequest);
            }
        }

        /// <summary>
        /// AssignRecord to user
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="recordId"></param>
        /// <param name="organizationService"></param>
        //AssignRecord a record
        public static void AssignRecord(this IOrganizationService pService, EntityReference pAssignee, EntityReference pRecord)
        {

            AssignRequest assignRequest = new AssignRequest
            {
                Assignee = new EntityReference(pAssignee.LogicalName, pAssignee.Id),
                Target = new EntityReference(pRecord.LogicalName, pRecord.Id)
            };
            pService.Execute(assignRequest);

        }


        public static string GetOptionsetText(string entityName, string attributeName, int OptionSetValue, IOrganizationService service)
        {

            string AttributeName = attributeName;
            string EntityLogicalName = entityName;
            string optionsetLabel = string.Empty;
            RetrieveEntityRequest retrieveDetails = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = EntityLogicalName
            };
            RetrieveEntityResponse retrieveEntityResponseObj = (RetrieveEntityResponse)service.Execute(retrieveDetails);
            EntityMetadata metadata = retrieveEntityResponseObj.EntityMetadata;
            PicklistAttributeMetadata picklistMetadata = metadata.Attributes.FirstOrDefault(attribute => String.Equals
                (attribute.LogicalName, attributeName, StringComparison.OrdinalIgnoreCase)) as PicklistAttributeMetadata;
            Microsoft.Xrm.Sdk.Metadata.OptionSetMetadata options = picklistMetadata.OptionSet;
            IList<OptionMetadata> OptionsList = (from o in options.Options
                                                 where o.Value.Value == OptionSetValue
                                                 select o).ToList();
            if (OptionsList != null && OptionsList.Count > 0)
                optionsetLabel = (OptionsList.First()).Label.UserLocalizedLabel.Label;
            return optionsetLabel;
        }

        public static void SetBusinessProcessFlowToNextStage(IOrganizationService pService, string pEntityName, Guid pEntityId, string pProcessName, string pStageName)
        {
            try
            {
                Guid NextactiveStageID = Guid.Empty;
                RetrieveProcessInstancesRequest processInstanceRequest = new RetrieveProcessInstancesRequest
                {
                    EntityId = pEntityId,
                    EntityLogicalName = pEntityName
                };
                RetrieveProcessInstancesResponse processInstanceResponse = (RetrieveProcessInstancesResponse)pService.Execute(processInstanceRequest);

                // Declare variables to store values returned in response
                Entity activeProcessInstance = processInstanceResponse.Processes.Entities[0];

                // First record is the active process instance
                Guid activeProcessInstanceID = activeProcessInstance.Id;

                // Id of the active process instance, which will be used later to retrieve the active path of the process instance
                // Retrieve the active stage ID of in the active process instance
                Guid activeStageID = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());

                // Retrieve the process stages in the active path of the current process instance
                RetrieveActivePathRequest pathReq = new RetrieveActivePathRequest { ProcessInstanceId = activeProcessInstanceID };
                RetrieveActivePathResponse pathResp = (RetrieveActivePathResponse)pService.Execute(pathReq);

                int activeStagePosition = -1;
                for (int i = 0; i < pathResp.ProcessStages.Entities.Count; i++)
                {
                    // Retrieve the active stage name and active stage position based on the pStageName for the process instance
                    if (pathResp.ProcessStages.Entities[i].Attributes["stagename"].ToString() == pStageName)
                    {
                        activeStagePosition = i;

                        break;
                    }
                }
                if (activeStagePosition > -1 && activeStagePosition < (pathResp.ProcessStages.Entities.Count - 1))
                {
                    // Retrieve the stage ID of the next stage that you want to set as active
                    NextactiveStageID = (Guid)pathResp.ProcessStages.Entities[activeStagePosition + 1].Attributes["processstageid"];

                    // Retrieve the process instance record to update its active stage
                    ColumnSet cols1 = new ColumnSet();
                    cols1.AddColumn("activestageid");

                    Entity retrievedProcessInstance = pService.Retrieve(pProcessName, activeProcessInstanceID, cols1);//
                                                                                                                      // Set the next stage as the active stage
                    retrievedProcessInstance["activestageid"] = new EntityReference("processstage", NextactiveStageID);
                    pService.Update(retrievedProcessInstance);
                }
                else if (activeStagePosition == (pathResp.ProcessStages.Entities.Count - 1)) // If Last Stage
                {
                    // Retrieve the process instance record to update its active stage
                    ColumnSet cols1 = new ColumnSet();
                    cols1.AddColumn("activestageid");
                    // change the status of BPF to Finish ( last stage of BPF)
                    Entity retrievedProcessInstance = pService.Retrieve(pProcessName, activeProcessInstanceID, cols1);//
                    retrievedProcessInstance["statecode"] = new OptionSetValue(1);
                    retrievedProcessInstance["statuscode"] = new OptionSetValue(2);                                                                              // Set the next stage as the active stage

                    pService.Update(retrievedProcessInstance);
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }

        }

        public static void SetBusinessProcessFlowStage(IOrganizationService pService, string pEntityName, Guid pEntityId, string pProcessName, string pStageName)
        {
            try
            {
                string[] workflowColumns = { "name", "statecode" };
                object[] workflowConditions = { pProcessName, 1 }; //WorkflowState.Activated
                Entity entityWorkflow = CRMHelper.GetOneEntityBy(pService, "workflow", workflowColumns, workflowConditions);
                if (entityWorkflow == null)
                {
                    throw new InvalidPluginExecutionException(string.Format("Worflow not found with name {0}", pProcessName));
                }
                else
                {
                    if (entityWorkflow.Attributes.Contains("workflowid") && entityWorkflow.Attributes["workflowid"] != null)
                    {
                        Guid workflowid = (Guid)entityWorkflow.Attributes["workflowid"];
                        string[] processStageColumns = { "stagename", "processid" };
                        object[] processStageConditions = { pStageName.ToUpper(), workflowid };
                        Entity entityStage = CRMHelper.GetOneEntityBy(pService, "processstage", processStageColumns, processStageConditions);
                        if (entityStage == null)
                        {
                            throw new InvalidPluginExecutionException(string.Format("Stage not found with name {0}", pStageName));
                        }
                        else
                        {
                            //Guid _processOpp1Id = Guid.Empty;
                            //string _procInstanceLogicalName = string.Empty;
                            //Guid _activeStageId = Guid.Empty;
                            //string _activeStageName;
                            //int _activeStagePosition = 0;


                            //RetrieveProcessInstancesRequest procOpp1Req = new RetrieveProcessInstancesRequest
                            //{
                            //    EntityId = pEntityId,
                            //    EntityLogicalName = pEntityName
                            //};
                            //RetrieveProcessInstancesResponse procOpp1Resp = (RetrieveProcessInstancesResponse)pService.Execute(procOpp1Req);

                            //// Declare variables to store values returned in response
                            //Entity activeProcessInstance = null;

                            //if (procOpp1Resp.Processes.Entities.Count > 0)
                            //{
                            //    activeProcessInstance = procOpp1Resp.Processes.Entities[0]; // First record is the active process instance
                            //    _processOpp1Id = activeProcessInstance.Id; // Id of the active process instance, which will be used                                                           
                            //    _procInstanceLogicalName = activeProcessInstance["name"].ToString().Replace(" ", string.Empty).ToLower();
                            //}
                            //else
                            //{


                            //}

                            //_activeStageId = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());

                            //// Retrieve the process stages in the active path of the current process instance
                            //RetrieveActivePathRequest pathReq = new RetrieveActivePathRequest
                            //{
                            //    ProcessInstanceId = _processOpp1Id
                            //};
                            //RetrieveActivePathResponse pathResp = (RetrieveActivePathResponse)pService.Execute(pathReq);

                            //for (int i = 0; i < pathResp.ProcessStages.Entities.Count; i++)
                            //{
                            //    if (pathResp.ProcessStages.Entities[i].Attributes["processstageid"].ToString() == _activeStageId.ToString())
                            //    {
                            //        _activeStageName = pathResp.ProcessStages.Entities[i].Attributes["stagename"].ToString();
                            //        _activeStagePosition = i;
                            //    }
                            //}

                            //_activeStageId = (Guid)pathResp.ProcessStages.Entities[_activeStagePosition + 1].Attributes["processstageid"];

                            //// Retrieve the process instance record to update its active stage
                            //ColumnSet cols1 = new ColumnSet();
                            //cols1.AddColumn("activestageid");
                            //Entity retrievedProcessInstance = pService.Retrieve("msdyn_bpf_d3d97bac8c294105840e99e37a9d1c39", _processOpp1Id, cols1);

                            //// Update the active stage to the next stage
                            //retrievedProcessInstance["activestageid"] = new EntityReference("processstage", _activeStageId);
                            //pService.Update(retrievedProcessInstance);


                            //SetProcessRequest request = new SetProcessRequest();
                            //request.Target = new EntityReference(pEntityName, pEntityId);


                            if (entityStage.Attributes.Contains("processstageid"))
                            {


                                Entity updateEntity = new Entity(pEntityName);
                                updateEntity.Id = pEntityId;
                                updateEntity["stageid"] = (Guid)entityStage.Attributes["processstageid"];
                                // updateEntity["processid"] = workflowid;
                                pService.Update(updateEntity);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //throw new InvalidPluginExecutionException(ex.Message.ToString());
            }

        }

    }


    /// <summary>
    /// Extensions for the IOrganizationService class to provide a set of common functions for accessing CRM MetaData.
    /// </summary>
    public static class IOganizationServiceMetaDataExtension
    {
        /// <summary>
        /// Returns a list of strings that represent each metadata attribute
        /// for the specified entity schema name that can be used in a create operation. 
        /// </summary>
        /// <param name="service"></param>
        /// <param name="EntityName"></param>
        /// <returns></returns>
        public static List<string> GetMetaDataCreateFieldNames(this IOrganizationService orgService, string entityLogicalName)
        {
            List<string> fieldNameList = new List<string>();

            Microsoft.Xrm.Sdk.Metadata.AttributeMetadata[] attributeMetadataArray = RetrieveAttributeMetaData(orgService, entityLogicalName);

            foreach (Microsoft.Xrm.Sdk.Metadata.AttributeMetadata item in attributeMetadataArray)
            {
                if ((bool)item.IsValidForCreate)
                {
                    fieldNameList.Add(item.LogicalName);
                }
            }

            return fieldNameList;
        }

        /// <summary>
        /// Returns a list of strings that represent each metadata attribute
        /// for the specified entity schema name that can be used in an update operation. 
        /// </summary>
        /// <param name="service"></param>
        /// <param name="EntityName"></param>
        /// <returns></returns>
        public static List<string> GetMetaDataUpdateFieldNames(this IOrganizationService orgService, string entityLogicalName)
        {
            List<string> fieldNameList = new List<string>();

            Microsoft.Xrm.Sdk.Metadata.AttributeMetadata[] attributeMetadataArray = RetrieveAttributeMetaData(orgService, entityLogicalName);

            foreach (Microsoft.Xrm.Sdk.Metadata.AttributeMetadata item in attributeMetadataArray)
            {
                if ((bool)item.IsValidForUpdate)
                {
                    fieldNameList.Add(item.LogicalName);
                }
            }

            return fieldNameList;
        }


        /// <summary>
        /// Returns the text associated with specified optionset value for the identified entity and attribute.
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="entityLogicalName">Schema name of the entity that contains the optionset attribute.</param>
        /// <param name="attributeName">Schema name of the attribute.</param>
        /// <param name="selectedValue">Numeric value of the optionset.</param>
        /// <returns></returns>
        public static string GetOptionSetLabel(this IOrganizationService orgService, string entityLogicalName, string attributeName, int OptionSetValue)
        {
            string returnLabel = string.Empty;

            OptionMetadataCollection optionsSetLabels = null;

            optionsSetLabels = RetrieveOptionSetMetaDataCollection(orgService, entityLogicalName, attributeName);

            foreach (OptionMetadata optionMetdaData in optionsSetLabels)
            {
                if (optionMetdaData.Value == OptionSetValue)
                {
                    returnLabel = optionMetdaData.Label.UserLocalizedLabel.Label;
                    break;
                }
            }

            return returnLabel;
        }

        /// <summary>
        /// Returns an array of AttributeMetadata for the specified entity name.
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="entityLogicalName"></param>
        /// <returns></returns>
        public static Microsoft.Xrm.Sdk.Metadata.AttributeMetadata[] RetrieveAttributeMetaData(this IOrganizationService orgService, string entityLogicalName)
        {
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest();
            RetrieveEntityResponse retrieveEntityResponse = new RetrieveEntityResponse();

            retrieveEntityRequest.LogicalName = entityLogicalName;
            retrieveEntityRequest.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes;

            retrieveEntityResponse = (RetrieveEntityResponse)orgService.Execute(retrieveEntityRequest);

            return retrieveEntityResponse.EntityMetadata.Attributes;
        }

        /// <summary>
        /// Returns the OptionMetadataCollection for the specified entity and attribute.
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="entityLogicalName"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        public static OptionMetadataCollection RetrieveOptionSetMetaDataCollection(this IOrganizationService orgService, string entityLogicalName, string attributeName)
        {
            OptionMetadataCollection returnOptionsCollection = null;
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest();
            RetrieveEntityResponse retrieveEntityResponse = new RetrieveEntityResponse();

            retrieveEntityRequest.LogicalName = entityLogicalName;
            retrieveEntityRequest.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes;

            retrieveEntityResponse = (RetrieveEntityResponse)orgService.Execute(retrieveEntityRequest);

            Microsoft.Xrm.Sdk.Metadata.AttributeMetadata[] attributeMetadataArray = RetrieveAttributeMetaData(orgService, entityLogicalName);

            foreach (Microsoft.Xrm.Sdk.Metadata.AttributeMetadata attributeMetadata in attributeMetadataArray)
            {
                if (attributeMetadata.AttributeType == AttributeTypeCode.Picklist &&
                    attributeMetadata.LogicalName == attributeName)
                {
                    returnOptionsCollection = ((PicklistAttributeMetadata)attributeMetadata).OptionSet.Options;
                    break;
                }
            }

            return returnOptionsCollection;
        }



    }

    /// <summary>
    /// Extensions for the Xrm Entity class to provide a set of common functions for manipulating Entity records.
    /// </summary>
    public static class EntityExtension
    {
        /// <summary>
        /// Adds the specified attribute, or updates the value if the attribute exists.
        /// </summary>
        /// <param name="pTarget"></param>
        /// <param name="AttributeName"></param>
        /// <param name="Value"></param>
        public static void AddUpdateAttribute(this Entity pTarget, string AttributeName, object Value)
        {
            if (pTarget.Attributes.Contains(AttributeName))
            {
                pTarget.Attributes[AttributeName] = Value;
            }
            else
            {
                pTarget.Attributes.Add(AttributeName, Value);
            }
        }


        /// <summary>
        /// Adds an attribute to the target entity by copying a specified attribute from a source entity. If 
        /// the target already contains the enitity then the target attribute will be updated.
        /// </summary>
        /// <param name="pTarget"></param>
        /// <param name="pTargetAttributeName"></param>
        /// <param name="pSource"></param>
        /// <param name="SourceAttributeName"></param>
        public static void AddUpdateAttributeFromSource(this Entity pTarget, string pTargetAttributeName, Entity pSource, string pSourceAttributeName)
        {
            if (pSource.Contains(pSourceAttributeName))
            {
                object value = pSource[pSourceAttributeName];
                pTarget.AddUpdateAttribute(pTargetAttributeName, value);
            }
        }


        /// <summary>
        /// Retrieves an attribute from the entity with of the specified type and returns a 
        /// default value if the attribute does not exist.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="pTarget"></param>
        /// <param name="Key"></param>
        /// <param name="DefaultValue"></param>
        /// <returns></returns>
        public static T GetValue<T>(this Entity pTarget, string Key, T DefaultValue)
        {
            if (!pTarget.Contains(Key))
            {
                return DefaultValue;
            }
            else if (pTarget[Key] == null)
            {
                return DefaultValue;
            }
            else
            {
                return pTarget.GetAttributeValue<T>(Key);
            }
        }


        /// <summary>
        /// Merge the entity with another entity to get a more complete list of attributes. If 
        /// the current entity and the source entity both have a value for a given attribute 
        /// then then attribute of the current entity will be preserved.
        /// </summary>
        /// <param name="copyTo"></param>
        /// <param name="copyFrom"></param>
        public static void MergeWith(this Entity copyTo, Entity copyFrom)
        {
            copyFrom.Attributes.ToList().ForEach(a =>
            {
                // if it already exists then dont copy
                if (!copyTo.Attributes.ContainsKey(a.Key))
                {
                    copyTo.Attributes.Add(a.Key, a.Value);
                }
            });
        }


        /// <summary>
        /// Removes an attribute from the entity attribute collection if that attribute exists.
        /// </summary>
        /// <param name="pTarget"></param>
        /// <param name="AttributeName"></param>
        public static void RemoveAttribute(this Entity pTarget, string AttributeName)
        {
            if (pTarget.Contains(AttributeName))
            {
                pTarget.Attributes.Remove(AttributeName);
            }
        }

        // MSM 3-1-2016
        public static bool IsFieldExist(this Entity pTarget, IOrganizationService pIOrganizationService, String fieldName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest
            {
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                LogicalName = pTarget.LogicalName
            };
            RetrieveEntityResponse response
              = (RetrieveEntityResponse)pIOrganizationService.Execute(request);
            return response.EntityMetadata.Attributes.FirstOrDefault(element
              => element.LogicalName == fieldName) != null;
        }




    }

    public class LinkedEntity
    {
        public string EntityAlias { get; set; }

        public string LinkToEntityName { get; set; }

        public string LinkFromAttributeName { get; set; }

        public string LinkToAttributeName { get; set; }

        public JoinOperator JoinOperator { get; set; }

        public FilterExpression Filters { get; set; } = new FilterExpression();
    }
}
