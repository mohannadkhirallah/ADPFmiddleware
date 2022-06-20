using ADPF.Utilities.Helpers;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Utilities.Manager
{
    public static class CRMBusinessManager
    {


        public static int RetrieveUserUiLanguageCode(IOrganizationService service, Guid userId)
        {
            var userSettingsQuery = new QueryExpression("usersettings");
            userSettingsQuery.ColumnSet.AddColumns("uilanguageid", "systemuserid");
            userSettingsQuery.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
            var userSettings = service.RetrieveMultiple(userSettingsQuery);
            if (userSettings.Entities.Count > 0)
            {
                return (int)userSettings.Entities[0]["uilanguageid"];
            }
            return 0;
        }

       

        public static bool IsTeamMemberByCode(IOrganizationService service, string Code, Guid UserID)
        {
            QueryExpression query = new QueryExpression("team");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition(new ConditionExpression("smc_teamcode", ConditionOperator.Equal, Code));
            LinkEntity link = query.AddLink("teammembership", "teamid", "teamid");
            link.LinkCriteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, UserID));

            try
            {
                var results = service.RetrieveMultiple(query);
                if (results.Entities.Count > 0) { return true; } else { return false; }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public static EntityCollection GetUserTeams(IOrganizationService service, Guid userID)
        {
            try
            {
                QueryExpression query = new QueryExpression("team");
                query.ColumnSet = new ColumnSet(true);
                LinkEntity link = query.AddLink("teammembership", "teamid", "teamid");
                link.LinkCriteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, userID));
                var results = service.RetrieveMultiple(query);
                return results;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }

        public static Entity GetTeamsDetails(IOrganizationService service, string teamCode)
        {
            try
            {
                Entity resultEntity = null;
                QueryExpression query = new QueryExpression("team");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition(new ConditionExpression("smc_teamcode", ConditionOperator.Equal, teamCode));
                var results = service.RetrieveMultiple(query);
                if (results.Entities.Count > 0)
                {
                    resultEntity = results.Entities[0];
                }
                return resultEntity;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }

        public static bool CreateTaks(IOrganizationService service, string pSubject, string pDescription, DateTime pScheduledstart, DateTime pScheduledend,
            EntityReference pRegrding)
        {
            bool result = false;
            try
            {
                Entity followup = new Entity();
                followup.LogicalName = "task";
                followup.Attributes = new AttributeCollection();
                followup.Attributes.Add("subject", pSubject);
                followup.Attributes.Add("description", pDescription);
                followup.Attributes.Add("scheduledstart", pScheduledstart);
                followup.Attributes.Add("scheduledend", pScheduledend);
                string regardingobjectidType = pRegrding.LogicalName;

                followup["regardingobjectid"] = pRegrding;
                // Refer to the account in the task activity.

                service.Create(followup);
                result = true;
            }
            catch (Exception)
            {

                //    throw;
            }

            return result;

        }

        public static void AddNotes(IOrganizationService service, EntityReference entityRef, string subject, string noteText, ITracingService tracingService)
        {
            try
            {
                Entity addingNotes = new Entity("annotation");
                addingNotes["subject"] = subject;
                addingNotes["notetext"] = noteText;
                addingNotes["objectid"] = entityRef;
                addingNotes["objecttypecode"] = entityRef.LogicalName;
                service.Create(addingNotes);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Attach Node: " + ex.Message.ToString());
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }

        }
        public static void ShareRecordToTeam(IOrganizationService service, string entityName, Guid sharingRecordId, EntityReference objUserOrTeam)
        {

            try
            {
                GrantAccessRequest grantRequest = new GrantAccessRequest()
                {
                    Target = new EntityReference(entityName, sharingRecordId),

                    PrincipalAccess = new PrincipalAccess()
                    {
                        //Principal = new EntityReference(objTeam.LogicalName, objTeam.Id),
                        Principal = objUserOrTeam,
                        AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess | AccessRights.ShareAccess
                    }
                };

                // Execute the request.

                GrantAccessResponse granted = (GrantAccessResponse)service.Execute(grantRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }
        public static void UnShareRecord(IOrganizationService service, EntityReference pRecordToUnshare, EntityReference pUserOrTeam)
        {
            ModifyAccessRequest modif = new ModifyAccessRequest(); ;
            modif.Target = new EntityReference(pRecordToUnshare.LogicalName, pRecordToUnshare.Id);

            PrincipalAccess principal = new PrincipalAccess();
            principal.Principal = new EntityReference(pUserOrTeam.LogicalName, pUserOrTeam.Id);
            principal.AccessMask = AccessRights.None;
            modif.PrincipalAccess = principal;

            try
            {
                ModifyAccessResponse modif_response = (ModifyAccessResponse)service.Execute(modif);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }
        public static void ShareRecordToTeamWithSpecificAccessMask(IOrganizationService service, string pEntityName, Guid pSharingRecordId, PrincipalAccess pPrincipalAccess)
        {

            try
            {

                GrantAccessRequest grantRequest = new GrantAccessRequest()
                {
                    Target = new EntityReference(pEntityName, pSharingRecordId),

                    PrincipalAccess = new PrincipalAccess()
                    {
                        //Principal = new EntityReference(objTeam.LogicalName, objTeam.Id),
                        Principal = pPrincipalAccess.Principal,

                        AccessMask = pPrincipalAccess.AccessMask
                    }
                };

                // Execute the request.

                GrantAccessResponse granted = (GrantAccessResponse)service.Execute(grantRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }

        public static void SetBusinessProcessFlowStage(IOrganizationService service, string pEntityName, Guid pEntityId, string pProcessName, string pStageName)
        {
            try
            {
                string[] workflowColumns = { "name", "statecode" };
                object[] workflowConditions = { pProcessName, 1 }; //WorkflowState.Activated
                Entity entityWorkflow = CRMHelper.GetOneEntityBy(service, "workflow", workflowColumns, workflowConditions);
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
                        Entity entityStage = CRMHelper.GetOneEntityBy(service, "processstage", processStageColumns, processStageConditions);
                        if (entityStage == null)
                        {
                            throw new InvalidPluginExecutionException(string.Format("Stage not found with name {0}", pStageName));
                        }
                        else
                        {
                            if (entityStage.Attributes.Contains("processstageid"))
                            {
                                Entity updateEntity = new Entity(pEntityName);
                                updateEntity.Id = pEntityId;
                                updateEntity["stageid"] = (Guid)entityStage.Attributes["processstageid"];
                                updateEntity["processid"] = workflowid;
                                service.Update(updateEntity);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }

        }

        public static void MakeOpportunityWon(IOrganizationService service, EntityReference pOpportunityId)//Staus Reason Won
        {
            try
            {
                WinOpportunityRequest req = new WinOpportunityRequest();
                Entity opportunityClose = new Entity("opportunityclose");
                opportunityClose.Attributes.Add("opportunityid", pOpportunityId);
                opportunityClose.Attributes.Add("subject", "Won the Opportunity!");
                req.OpportunityClose = opportunityClose;
                req.RequestName = "WinOpportunity";
                Microsoft.Xrm.Sdk.OptionSetValue optionSetStatusReason = new Microsoft.Xrm.Sdk.OptionSetValue();
                optionSetStatusReason.Value = 3;// Staus Reason Won
                req.Status = optionSetStatusReason;
                WinOpportunityResponse resp = (WinOpportunityResponse)service.Execute(req);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }

        public static void MakeOpportunityLost(IOrganizationService service, EntityReference pOpportunityId, int pStatusReason)
        {
            try
            {
                LoseOpportunityRequest req = new LoseOpportunityRequest();
                Entity opportunityClose = new Entity("opportunityclose");
                opportunityClose.Attributes.Add("opportunityid", pOpportunityId);
                opportunityClose.Attributes.Add("subject", "Lost the Opportunity!");
                req.OpportunityClose = opportunityClose;
                Microsoft.Xrm.Sdk.OptionSetValue optionsetStatusReason = new Microsoft.Xrm.Sdk.OptionSetValue();
                optionsetStatusReason.Value = pStatusReason;
                req.Status = optionsetStatusReason;
                LoseOpportunityResponse resp = (LoseOpportunityResponse)service.Execute(req);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }

        public static void MakeQuoteClose(IOrganizationService service, EntityReference pQuoteId, int pStatusReason)
        {
            try
            {
                CloseQuoteRequest req = new CloseQuoteRequest();
                Entity quoteClose = new Entity("quoteclose");
                quoteClose.Attributes.Add("quoteid", pQuoteId);
                quoteClose.Attributes.Add("subject", "Close Quote");
                req.QuoteClose = quoteClose;
                req.RequestName = "CloseQuote";
                Microsoft.Xrm.Sdk.OptionSetValue optionsetStatusReason = new Microsoft.Xrm.Sdk.OptionSetValue();
                optionsetStatusReason.Value = pStatusReason;
                req.Status = optionsetStatusReason;
                CloseQuoteResponse resp = (CloseQuoteResponse)service.Execute(req);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }

        public static void MakeQuoteWon(IOrganizationService service, EntityReference pQuoteId)
        {
            try
            {
                WinQuoteRequest reqWinQuote = new WinQuoteRequest();
                Entity quoteClose = new Entity("quoteclose");
                quoteClose.Attributes.Add("quoteid", pQuoteId);
                quoteClose.Attributes.Add("subject", "Won the Quote");
                reqWinQuote.QuoteClose = quoteClose;
                reqWinQuote.RequestName = "WinQuote";
                Microsoft.Xrm.Sdk.OptionSetValue optionsetStatusReason = new Microsoft.Xrm.Sdk.OptionSetValue();
                optionsetStatusReason.Value = 4;//Quote Status Reason - Won
                reqWinQuote.Status = optionsetStatusReason;
                WinQuoteResponse resWinQuote = (WinQuoteResponse)service.Execute(reqWinQuote);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }


        }


    }
}

