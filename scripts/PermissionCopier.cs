using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class PermissionCopier
    {
        private readonly IOrganizationService _service;

        public PermissionCopier(IOrganizationService service)
        {
            _service = service;
        }

        public async Task CopyBusinessUnit(Entity sourceUser, Entity targetUser)
        {
            if (!sourceUser.Contains("businessunitid"))
            {
                DynamicsCrmUtility.LogMessage("Source user does not have a Business Unit assigned.", "WARNING");
                return;
            }

            var sourceBuId = ((EntityReference)sourceUser["businessunitid"]).Id;
            await _service.UpdateAsync(new Entity("systemuser")
            {
                Id = targetUser.Id,
                ["businessunitid"] = new EntityReference("businessunit", sourceBuId)
            });
            DynamicsCrmUtility.LogMessage("Business Unit copied successfully.");
        }

        public async Task CopyTeams(Entity sourceUser, Entity targetUser)
        {
            var sourceTeams = await GetUserTeamsAsync(sourceUser.Id);
            var targetTeams = await GetUserTeamsAsync(targetUser.Id);

            foreach (var team in sourceTeams.Entities)
            {
                if (!targetTeams.Entities.Any(t => t.Id == team.Id))
                {
                    var addMembersRequest = new AddMembersTeamRequest
                    {
                        TeamId = team.Id,
                        MemberIds = new[] { targetUser.Id }
                    };

                    try
                    {
                        await Task.Run(() => _service.Execute(addMembersRequest));
                        DynamicsCrmUtility.LogMessage($"User added to team '{team.GetAttributeValue<string>("name")}'.", "SUCCESS");
                    }
                    catch (Exception ex)
                    {
                        DynamicsCrmUtility.LogMessage($"Error adding user to team '{team.GetAttributeValue<string>("name")}': {ex.Message}", "ERROR");
                    }
                }
            }

            DynamicsCrmUtility.LogMessage("\nTeam memberships copied successfully.\n");
        }

       public async Task CopyRoles(Entity sourceUser, Entity targetUser)
        {
            var sourceRoles = await GetUserRolesAsync(sourceUser.Id);
            var targetRoles = await GetUserRolesAsync(targetUser.Id);
            var targetUserBusinessUnit = ((EntityReference)targetUser["businessunitid"]).Id;

            foreach (var sourceRole in sourceRoles.Entities)
            {
                // Find an equivalent role in the target user's Business Unit
                var equivalentRole = await FindEquivalentRoleInBusinessUnitAsync(sourceRole, targetUserBusinessUnit);

                if (equivalentRole == null)
                {
                    DynamicsCrmUtility.LogMessage($"No equivalent role found for '{sourceRole.GetAttributeValue<string>("name")}' in the target user's Business Unit.", "WARNING");
                    continue;
                }

                if (!targetRoles.Entities.Any(r => r.Id == equivalentRole.Id))
                {
                    try
                    {
                        var request = new AssociateRequest
                        {
                            Target = new EntityReference("systemuser", targetUser.Id),
                            RelatedEntities = new EntityReferenceCollection
                    {
                        new EntityReference("role", equivalentRole.Id)
                    },
                            Relationship = new Relationship("systemuserroles_association")
                        };

                        await Task.Run(() => _service.Execute(request));
                        DynamicsCrmUtility.LogMessage($"Role '{equivalentRole.GetAttributeValue<string>("name")}' assigned to user.", "SUCCESS");
                    }
                    catch (Exception ex)
                    {
                        DynamicsCrmUtility.LogMessage($"Error assigning role '{equivalentRole.GetAttributeValue<string>("name")}' to user: {ex.Message}", "ERROR");
                    }
                }
            }

            DynamicsCrmUtility.LogMessage("Roles copying process completed.");
        }

        private async Task<Entity> FindEquivalentRoleInBusinessUnitAsync(Entity sourceRole, Guid targetBusinessUnitId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                        new ConditionExpression("name", ConditionOperator.Equal, sourceRole.GetAttributeValue<string>("name")),
                        new ConditionExpression("businessunitid", ConditionOperator.Equal,targetBusinessUnitId)
                    }
                }
            };
            var result = await _service.RetrieveMultipleAsync(query);

            return result.Entities.FirstOrDefault();
        }

        public async Task<EntityCollection> GetUserTeamsAsync(Guid userId)
        {
            // First, retrieve the user's Business Unit
            var userQuery = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("businessunitid"),
                Criteria = new FilterExpression
                {
                    Conditions =
            {
                new ConditionExpression("systemuserid", ConditionOperator.Equal, userId),
                new ConditionExpression("isdisabled", ConditionOperator.Equal, false)
            }
                }
            };
            var userResult = await _service.RetrieveMultipleAsync(userQuery);
            if (userResult.Entities.Count == 0)
            {
                throw new Exception($"User with ID {userId} not found or is disabled.");
            }
            var userBusinessUnitId = ((EntityReference)userResult.Entities[0]["businessunitid"]).Id;

            // Retrieve the Business Unit name
            var buQuery = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet("name"),
                Criteria = new FilterExpression
                {
                    Conditions =
            {
                new ConditionExpression("businessunitid", ConditionOperator.Equal, userBusinessUnitId)
            }
                }
            };
            var buResult = await _service.RetrieveMultipleAsync(buQuery);
            var businessUnitName = buResult.Entities[0]["name"].ToString();

            // Now, retrieve the user's teams, including the Business Unit ID for each team
            var teamQuery = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("name", "businessunitid"),
                LinkEntities =
        {
            new LinkEntity
            {
                LinkFromEntityName = "team",
                LinkToEntityName = "teammembership",
                LinkFromAttributeName = "teamid",
                LinkToAttributeName = "teamid",
                LinkCriteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                    }
                }
            }
        }
            };
            var result = await _service.RetrieveMultipleAsync(teamQuery);

            // Filter out only the team that exactly matches the user's Business Unit name
            var filteredTeams = result.Entities.Where(e => e["name"].ToString() != businessUnitName).ToList();
            return new EntityCollection(filteredTeams);
        }

        public async Task<EntityCollection> GetUserRolesAsync(Guid userId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("name"),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "role",
                        LinkToEntityName = "systemuserroles",
                        LinkFromAttributeName = "roleid",
                        LinkToAttributeName = "roleid",
                        JoinOperator = JoinOperator.Inner,
                        LinkEntities =
                        {
                            new LinkEntity
                            {
                                LinkFromEntityName = "systemuserroles",
                                LinkToEntityName = "systemuser",
                                LinkFromAttributeName = "systemuserid",
                                LinkToAttributeName = "systemuserid",
                                LinkCriteria = new FilterExpression
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression("systemuserid", ConditionOperator.Equal, userId),
                                        new ConditionExpression("isdisabled", ConditionOperator.Equal, false)
                                    }
                                }
                            }
                        }
                    }
                }
            };

            return await _service.RetrieveMultipleAsync(query);
        }
    }
}