using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Threading;

namespace RitmsHub.Scripts
{
    public abstract class BaseTeam
    {
        public string TeamName { get; protected set; }
        public string BusinessUnitName { get; protected set; }
        public string AdministratorName { get; protected set; }
        public string[] TeamRoles { get; protected set; }

        public abstract void SetTeamProperties(TransformedTeamData teamData);
    }

    public class ProprietaryTeam : BaseTeam
    {
        public override void SetTeamProperties(TransformedTeamData teamData)
        {
            TeamName = teamData.EquipaEDPR;
            BusinessUnitName = teamData.Bu;
            AdministratorName = CodesAndRoles.AdministratorNameEU;
            TeamRoles = CodesAndRoles.ProprietaryTeamRoles;
        }
    }

    public class StandardTeam : BaseTeam
    {
        public override void SetTeamProperties(TransformedTeamData teamData)
        {
            TeamName = teamData.EquipaContrata;
            BusinessUnitName = teamData.Bu;
            AdministratorName = CodesAndRoles.AdministratorNameEU;
            TeamRoles = CodesAndRoles.TeamRolesEU;
        }
    }

    public class TeamManager
    {
        private readonly CrmServiceClient _crmServiceClient;

        public TeamManager(string connectionString)
        {
            _crmServiceClient = DynamicsCrmUtility.CreateCrmServiceClient();
        }

        public async Task<TeamOperationResult> CreateOrUpdateTeamAsync(BaseTeam team, bool isProprietaryTeam, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            string teamType = isProprietaryTeam ? "Proprietary" : "Standard";
            Console.WriteLine($"\nStarting {teamType} team creation/update process for team: '{team.TeamName.Trim()}'");

            var existingTeam = await GetExistingTeamAsync(team.TeamName.Trim(), cancellationToken);
            if (existingTeam != null)
            {
                bool isCorrectType = (isProprietaryTeam && existingTeam.GetAttributeValue<OptionSetValue>("teamtype").Value == 0) ||
                                     (!isProprietaryTeam && existingTeam.GetAttributeValue<OptionSetValue>("teamtype").Value == 0);

                if (!isCorrectType)
                {
                    await DeleteTeamAsync(existingTeam.Id, cancellationToken);
                    return await CreateNewTeamAsync(team, isProprietaryTeam, cancellationToken);
                }
                else
                {
                    return await UpdateTeamIfNeededAsync(existingTeam, team, isProprietaryTeam, cancellationToken);
                }
            }
            else
            {
                return await CreateNewTeamAsync(team, isProprietaryTeam, cancellationToken);
            }
        }

        private async Task<Entity> GetExistingTeamAsync(string teamName, CancellationToken cancellationToken)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression("name", ConditionOperator.Equal, teamName)
                }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query), cancellationToken);
            return result.Entities.Count > 0 ? result.Entities[0] : null;
        }

        private async Task DeleteTeamAsync(Guid teamId, CancellationToken cancellationToken)
        {
            await Task.Run(() => _crmServiceClient.Delete("team", teamId), cancellationToken);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"Deleted existing team with incorrect type. Team ID: {teamId}");
            Console.ResetColor();
        }

        private async Task<TeamOperationResult> UpdateTeamIfNeededAsync(Entity existingTeam, BaseTeam team, bool isProprietaryTeam, CancellationToken cancellationToken)
        {
            bool updated = false;
            var updateEntity = new Entity("team") { Id = existingTeam.Id };

            var businessUnitId = await DynamicsCrmUtility.GetBusinessUnitIdAsync(_crmServiceClient, team.BusinessUnitName, cancellationToken);
            if (existingTeam.GetAttributeValue<EntityReference>("businessunitid").Id != businessUnitId)
            {
                updateEntity["businessunitid"] = new EntityReference("businessunit", businessUnitId);
                updated = true;
                Console.WriteLine($"Updating businessunitid for team '{team.TeamName.Trim()}' to '{businessUnitId}'");
            }

            var administratorId = await GetUserIdAsync(team.AdministratorName, cancellationToken);
            if (existingTeam.GetAttributeValue<EntityReference>("administratorid").Id != administratorId)
            {
                updateEntity["administratorid"] = new EntityReference("systemuser", administratorId);
                updated = true;
                Console.WriteLine($"Updating administratorid for team '{team.TeamName.Trim()}' to '{administratorId}'");
            }

            // Update the team entity if needed
            if (updated)
            {
                await Task.Run(() => _crmServiceClient.Update(updateEntity), cancellationToken);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Team '{team.TeamName.Trim()}' updated successfully.");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Team '{team.TeamName.Trim()}' properties are up to date. No changes needed.");
                Console.ResetColor();
            }

            // Check and update roles for proprietary teams
            if (isProprietaryTeam)
            {
                updated |= await UpdateTeamRolesIfNeededAsync(existingTeam.Id, team.TeamRoles, businessUnitId, cancellationToken);
                await UpdateBusinessUnitWithProprietaryTeamAsync(businessUnitId, existingTeam.Id, team.TeamName.Trim(), cancellationToken);
            }
            else
            {
                // Check and update roles for standard teams
                updated |= await UpdateTeamRolesIfNeededAsync(existingTeam.Id, team.TeamRoles, businessUnitId, cancellationToken);
            }

            return new TeamOperationResult { TeamName = team.TeamName, Exists = true, WasUpdated = updated };
        }


        private async Task<TeamOperationResult> CreateNewTeamAsync(BaseTeam team, bool isProprietaryTeam, CancellationToken cancellationToken)
        {
            var businessUnitId = await DynamicsCrmUtility.GetBusinessUnitIdAsync(_crmServiceClient, team.BusinessUnitName, cancellationToken);
            var administratorId = await GetUserIdAsync(team.AdministratorName, cancellationToken);

            var teamEntity = new Entity("team")
            {
                ["name"] = team.TeamName.Trim(),
                ["businessunitid"] = new EntityReference("businessunit", businessUnitId),
                ["administratorid"] = new EntityReference("systemuser", administratorId),
                ["teamtype"] = new OptionSetValue(0) // Always set to 0 for Owner team type
            };

            var newTeamId = await Task.Run(() => _crmServiceClient.Create(teamEntity), cancellationToken);
            Console.WriteLine($"New team created successfully. Team ID: {newTeamId}", "SUCCESS");
            Console.WriteLine($"Administrator '{team.AdministratorName}' assigned to team.", "SUCCESS");

            // Assign roles to proprietary and standard teams alike
            if (team.TeamRoles != null && team.TeamRoles.Length > 0)
            {
                await AssignRolesToTeamAsync(newTeamId, businessUnitId, team.TeamRoles, cancellationToken);
            }

            // Additional logic for proprietary teams
            if (isProprietaryTeam)
            {
                await UpdateBusinessUnitWithProprietaryTeamAsync(businessUnitId, newTeamId, team.TeamName.Trim(), cancellationToken);
            }

            return new TeamOperationResult { TeamName = team.TeamName, Exists = true, WasUpdated = false };
        }


        private async Task UpdateBusinessUnitWithProprietaryTeamAsync(Guid businessUnitId, Guid proprietaryTeamId, string proprietaryTeamName, CancellationToken cancellationToken)
        {
            try
            {
                var businessUnitUpdate = new Entity("businessunit", businessUnitId);
                businessUnitUpdate["atos_equipopropietarioid"] = new EntityReference("team", proprietaryTeamId);
                businessUnitUpdate["atos_equipopropietarioidname"] = proprietaryTeamName;

                await Task.Run(() => _crmServiceClient.Update(businessUnitUpdate));
                Console.WriteLine($"Business Unit updated with Proprietary Team information. BU ID: {businessUnitId}, Team ID: {proprietaryTeamId}", "SUCCESS");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating Business Unit with Proprietary Team: {ex.Message}", "ERROR");
            }
        }

        private async Task<bool> UpdateTeamRolesIfNeededAsync(Guid teamId, string[] desiredRoles, Guid businessUnitId, CancellationToken cancellationToken)
        {
            var currentRoles = await GetTeamRolesAsync(teamId);
            var desiredRoleSet = new HashSet<string>(desiredRoles);
            var currentRoleSet = new HashSet<string>(currentRoles.Select(r => r.GetAttributeValue<string>("name")));

            var rolesToAdd = desiredRoleSet.Except(currentRoleSet);
            var rolesToRemove = currentRoleSet.Except(desiredRoleSet);

            bool updated = false;

            foreach (var roleName in rolesToAdd)
            {
                var roleInfo = await GetRoleInfoAsync(roleName, businessUnitId);
                if (roleInfo.HasValue)
                {
                    await AssignRoleToTeamAsync(teamId, roleInfo.Value.roleId, roleInfo.Value.roleName);
                    updated = true;
                }
            }

            foreach (var roleName in rolesToRemove)
            {
                var roleToRemove = currentRoles.FirstOrDefault(r => r.GetAttributeValue<string>("name") == roleName);
                if (roleToRemove != null)
                {
                    await RemoveRoleFromTeamAsync(teamId, roleToRemove.Id, roleName);
                    updated = true;
                }
            }

            return updated;
        }

        private async Task AssignRolesToTeamAsync(Guid teamId, Guid businessUnitId, string[] roleNames, CancellationToken cancellationToken)
        {
            foreach (var roleName in roleNames)
            {
                var roleInfo = await GetRoleInfoAsync(roleName, businessUnitId);
                if (roleInfo.HasValue)
                {
                    await AssignRoleToTeamAsync(teamId, roleInfo.Value.roleId, roleInfo.Value.roleName);
                }
            }
        }

        private async Task<(Guid roleId, string roleName)?> GetRoleInfoAsync(string roleName, Guid businessUnitId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions = {
                        new ConditionExpression("name", ConditionOperator.Equal, roleName),
                        new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId)
                    }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query));
            if (result.Entities.Count == 0)
            {
                Console.WriteLine($"Role not found: {roleName}", "WARNING");
                return null;
            }

            var role = result.Entities[0];
            return (role.Id, role.GetAttributeValue<string>("name"));
        }

        private async Task AssignRoleToTeamAsync(Guid teamId, Guid roleId, string roleName)
        {
            try
            {
                var query = new QueryExpression("teamroles")
                {
                    ColumnSet = new ColumnSet("teamroleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions = {
                            new ConditionExpression("teamid", ConditionOperator.Equal, teamId),
                            new ConditionExpression("roleid", ConditionOperator.Equal, roleId)
                        }
                    }
                };

                var existingAssignment = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query));
                if (existingAssignment.Entities.Count > 0)
                {
                    Console.WriteLine($"Role '{roleName}' is already assigned to the team.", "INFO");
                    return;
                }

                var request = new AssociateRequest
                {
                    Target = new EntityReference("team", teamId),
                    RelatedEntities = new EntityReferenceCollection
                    {
                        new EntityReference("role", roleId)
                    },
                    Relationship = new Relationship("teamroles_association")
                };

                await Task.Run(() => _crmServiceClient.Execute(request));
                Console.WriteLine($"Role '{roleName}' assigned to team.", "SUCCESS");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error assigning role '{roleName}' to team: {ex.Message}", "ERROR");
            }
        }

        private async Task RemoveRoleFromTeamAsync(Guid teamId, Guid roleId, string roleName)
        {
            try
            {
                var request = new DisassociateRequest
                {
                    Target = new EntityReference("team", teamId),
                    RelatedEntities = new EntityReferenceCollection
                    {
                        new EntityReference("role", roleId)
                    },
                    Relationship = new Relationship("teamroles_association")
                };

                await Task.Run(() => _crmServiceClient.Execute(request));
                Console.WriteLine($"Role '{roleName}' removed from team.", "SUCCESS");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing role '{roleName}' from team: {ex.Message}", "ERROR");
            }
        }

        private async Task<Guid> GetUserIdAsync(string fullName, CancellationToken cancellationToken)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("fullname", ConditionOperator.Equal, fullName) }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query), cancellationToken);
            if (result.Entities.Count == 0)
            {
                throw new Exception($"User not found: {fullName}");
            }

            return result.Entities[0].Id;
        }

        private async Task<List<Entity>> GetTeamRolesAsync(Guid teamId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "role",
                        LinkToEntityName = "teamroles",
                        LinkFromAttributeName = "roleid",
                        LinkToAttributeName = "roleid",
                        JoinOperator = JoinOperator.Inner,
                        LinkCriteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("teamid", ConditionOperator.Equal, teamId)
                            }
                        }
                    }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query));
            return result.Entities.ToList();
        }
    }

    public class CreateTeam
    {

        // sequential team creation, slowler but reduces the chances of lockout of other users
        public static async Task<List<TeamOperationResult>> RunAsync(List<TransformedTeamData> transformedTeams, TeamType teamType, CancellationToken cancellationToken)
        {
            var results = new List<TeamOperationResult>();
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                var teamManager = new TeamManager(connectionString);

                foreach (var team in transformedTeams)
                {
                    // Check for cancellation before each iteration
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        BaseTeam baseTeam = teamType == TeamType.Proprietary
                            ? new ProprietaryTeam()
                            : new StandardTeam();

                        baseTeam.SetTeamProperties(team);

                        var result = await teamManager.CreateOrUpdateTeamAsync(baseTeam, teamType == TeamType.Proprietary, cancellationToken);
                        result.BuName = team.Bu;
                        results.Add(result);

                        Console.WriteLine($"{teamType} Team '{result.TeamName}' {(result.Exists ? (result.WasUpdated ? "updated" : "already exists") : "created")}.");
                    }
                    catch (OperationCanceledException)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Operation cancelled while processing {teamType} Team for '{team.Bu}'.");
                        Console.ResetColor();
                        throw; // Re-throw to stop further processing
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error processing {teamType} Team for '{team.Bu}': {ex.Message}");
                        Console.ResetColor();
                    }

                    // Optional: Add a small delay between operations
                    await Task.Delay(500, cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"{teamType} Team creation process was cancelled.");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error in {teamType} Team creation/update process: {ex.Message}");
                Console.ResetColor();
                return results;
            }
            return results;
        }

    }

    public enum TeamType
    {
        Standard,
        Proprietary
    }

    public class TeamOperationResult
    {
        public string TeamName { get; set; }
        public bool Exists { get; set; }
        public bool WasUpdated { get; set; }
        public string BuName { get; set; } 
    }
}