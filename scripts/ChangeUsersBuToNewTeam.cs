using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class ChangeUsersBuToNewTeam
    {
        private IOrganizationService _service;
        private UserRetriever _userRetriever;

        public async Task Run(List<BuUserDomains> buUserDomainsList, List<TransformedTeamData> transformedTeams)
        {
            try
            {
                await ConnectToDynamicsAsync();

                var preprocessedData = PreprocessData(buUserDomainsList);

                foreach (var item in preprocessedData)
                {
                    foreach (var userDomain in item.UserDomains)
                    {
                        await ProcessUserAsync(userDomain, item.OriginalNewCreatedPark, item.ProcessedNewBU, transformedTeams);
                    }
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Process completed successfully.");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                LogError($"An error occurred in the main process: {ex.Message}", ex);
            }
        }

        private async Task ConnectToDynamicsAsync()
        {
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine($"Attempting to connect with: {connectionString}");
                Console.ResetColor();

                var serviceClient = new CrmServiceClient(connectionString);
                if (serviceClient is null || !serviceClient.IsReady)
                {
                    throw new Exception($"Failed to connect. Error: {(serviceClient?.LastCrmError ?? "Unknown error")}");

                }

                _service = serviceClient;
                _userRetriever = new UserRetriever(_service);
                Console.ForegroundColor = ConsoleColor.Green;
                //Console.WriteLine($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                LogError("Failed to connect to Dynamics CRM", ex);
                throw;
            }
        }

        private List<PreprocessedBuData> PreprocessData(List<BuUserDomains> buUserDomainsList)
        {
            return buUserDomainsList.Select(item => new PreprocessedBuData
            {
                OriginalNewCreatedPark = item.NewCreatedPark,
                ProcessedNewBU = ProcessNewCreatedPark(item.NewCreatedPark),
                UserDomains = item.UserDomains
            }).ToList();
        }

        private string ProcessNewCreatedPark(string newCreatedPark)
        {
            return Regex.Replace(newCreatedPark, @"^Equipo contrata\s*", "", RegexOptions.IgnoreCase).Trim();
        }

        private async Task ProcessUserAsync(string userDomain, string originalNewCreatedPark, string processedNewBU, List<TransformedTeamData> transformedTeams)
        {
            try
            {
                var users = await _userRetriever.FindUsersAsync(userDomain);
                if (users.Count == 0)
                {
                    LogWarning($"User with domain {userDomain} not found.");
                    return;
                }
                var user = users[0]; // Take the first user if multiple are found
                string fullName = user.GetAttributeValue<string>("yomifullname");
                var currentBu = user.GetAttributeValue<EntityReference>("businessunitid");
                Console.WriteLine($"Processing user: {fullName}, Current BU: {currentBu.Name}");
                if (!string.Equals(currentBu.Name, processedNewBU, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"User {fullName} current BU does not match processed BU. Skipping.");
                    return;
                }
                var matchingTeam = FindMatchingTeam(originalNewCreatedPark, transformedTeams);
                if (!matchingTeam.HasValue)
                {
                    LogWarning($"No matching team found for {originalNewCreatedPark}. Skipping user {fullName}.");
                    return;
                }
                if (!UserNameContainsContractor(fullName, matchingTeam.Value.Contractor))
                {
                    Console.WriteLine($"User {fullName} name does not contain contractor {matchingTeam.Value.Contractor}. Skipping.");
                    return;
                }
                string newBuName = DeriveNewBuName(originalNewCreatedPark);
                await ChangeBuAndReapplyRolesAsync(user, newBuName, originalNewCreatedPark);
            }
            catch (Exception ex)
            {
                LogError($"Error processing user {userDomain}", ex);
            }
        }

        private TransformedTeamData? FindMatchingTeam(string originalNewCreatedPark, List<TransformedTeamData> transformedTeams)
        {
            return transformedTeams.FirstOrDefault(t =>
                string.Equals(t.EquipaContrataContrata, originalNewCreatedPark, StringComparison.OrdinalIgnoreCase));
        }

        private bool UserNameContainsContractor(string userName, string contractor)
        {
            return contractor.Split(' ').Any(word =>
                userName.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private string DeriveNewBuName(string originalNewCreatedPark)
        {
            return Regex.Replace(originalNewCreatedPark, @"^Equipo contrata\s*", "", RegexOptions.IgnoreCase).Trim();
        }

        private async Task ChangeBuAndReapplyRolesAsync(Entity user, string newBuName, string originalNewCreatedPark)
        {
            var currentRoles = await GetUserRolesAsync(user.Id);
            var newBu = await GetBusinessUnitByNameAsync(newBuName);

            if (newBu == null)
            {
                LogWarning($"Business Unit {newBuName} not found. Skipping BU change for user {user.GetAttributeValue<string>("yomifullname")}.");
                return;
            }

            // Change the user's BU
            user["businessunitid"] = new EntityReference("businessunit", newBu.Id);
            await _service.UpdateAsync(user);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Changed BU for user {user.GetAttributeValue<string>("yomifullname")} to {newBuName}.");
            Console.ResetColor();

            // Reapply roles
            await ReapplyRolesToUserAsync(user.Id, newBu.Id, currentRoles);

            // Remove the Contrata Contrata team
            await RemoveContrataContrataTeamAsync(user, originalNewCreatedPark);
        }

        private async Task<List<Entity>> GetUserRolesAsync(Guid userId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                LinkEntities = {
                    new LinkEntity
                    {
                        LinkFromEntityName = "role",
                        LinkToEntityName = "systemuserroles",
                        LinkFromAttributeName = "roleid",
                        LinkToAttributeName = "roleid",
                        LinkCriteria = new FilterExpression
                        {
                            Conditions = {
                                new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                            }
                        }
                    }
                }
            };

            var result = await _service.RetrieveMultipleAsync(query);
            return result.Entities.ToList();
        }

        private async Task ReapplyRolesToUserAsync(Guid userId, Guid newBuId, List<Entity> roles)
        {
            foreach (var role in roles)
            {
                try
                {
                    var newRole = await FindEquivalentRoleInBusinessUnitAsync(role, newBuId);
                    if (newRole != null)
                    {
                        await AssignRoleToUserAsync(userId, newRole.Id);
                        Console.WriteLine($"Reapplied role {newRole.GetAttributeValue<string>("name")} to user.");
                    }
                    else
                    {
                        LogWarning($"Equivalent role for {role.GetAttributeValue<string>("name")} not found in new BU.");
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Error reapplying role {role.GetAttributeValue<string>("name")}", ex);
                }
            }
        }


        // ----------------------------------------------------------
        // remove extra contrata contra team from users that have in their name the subcontractor of the team being created
        // for example remove  Equipo contrata 0-PT-CES-01 Contrata

        private async Task RemoveContrataContrataTeamAsync(Entity user, string newCreatedPark)
        {
            // Derive the contrataContrataTeam name
            string contrataContrataTeam = DeriveContrataContrataTeam(newCreatedPark);

            // Retrieve the user's teams
            var userTeams = await GetUserTeamsAsync(user.Id);

            // Find the matching team (case-insensitive)
            var teamToRemove = userTeams.FirstOrDefault(t =>
                string.Equals(t.GetAttributeValue<string>("name"), contrataContrataTeam, StringComparison.OrdinalIgnoreCase));

            if (teamToRemove != null)
            {
                // Remove the team from the user
                await RemoveTeamFromUserAsync(user.Id, teamToRemove.Id);
                Console.WriteLine($"Removed team '{contrataContrataTeam}' from user {user.GetAttributeValue<string>("yomifullname")}.");
            }
        }

        private string DeriveContrataContrataTeam(string newCreatedPark)
        {
            var match = Regex.Match(newCreatedPark, @"(Equipo contrata.*?Contrata)");
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }

        private async Task<List<Entity>> GetUserTeamsAsync(Guid userId)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid", "name"),
                LinkEntities = {
            new LinkEntity
            {
                LinkFromEntityName = "team",
                LinkToEntityName = "teammembership",
                LinkFromAttributeName = "teamid",
                LinkToAttributeName = "teamid",
                LinkCriteria = new FilterExpression
                {
    Conditions = {
        new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
    }
}
            }
        }
            };

            var result = await _service.RetrieveMultipleAsync(query);
            return result.Entities.ToList();
        }

        private async Task RemoveTeamFromUserAsync(Guid userId, Guid teamId)
        {
            var request = new DisassociateRequest
            {
                Target = new EntityReference("systemuser", userId),
                RelatedEntities = new EntityReferenceCollection
        {
            new EntityReference("team", teamId)
        },
                Relationship = new Relationship("teammembership_association")
            };

            await Task.Run(() => _service.Execute(request));
        }


        // ----------------------------------------------------------

        private async Task<Entity> FindEquivalentRoleInBusinessUnitAsync(Entity sourceRole, Guid targetBusinessUnitId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions = {
                        new ConditionExpression("name", ConditionOperator.Equal, sourceRole.GetAttributeValue<string>("name")),
                        new ConditionExpression("businessunitid", ConditionOperator.Equal, targetBusinessUnitId)
                    }
                }
            };

            var result = await _service.RetrieveMultipleAsync(query);
            return result.Entities.FirstOrDefault();
        }

        private async Task AssignRoleToUserAsync(Guid userId, Guid roleId)
        {
            var request = new AssociateRequest
            {
                Target = new EntityReference("systemuser", userId),
                RelatedEntities = new EntityReferenceCollection
                {
                    new EntityReference("role", roleId)
                },
                Relationship = new Relationship("systemuserroles_association")
            };

            await Task.Run(() => _service.Execute(request));
        }

        private async Task<Entity> GetBusinessUnitByNameAsync(string buName)
        {
            var query = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet("businessunitid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions = {
                        new ConditionExpression("name", ConditionOperator.Equal, buName)
                    }
                }
            };

            var result = await _service.RetrieveMultipleAsync(query);
            return result.Entities.FirstOrDefault();
        }

        private void LogError(string message, Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"ERROR: {message}");
            Console.WriteLine($"Exception: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            Console.ResetColor();
        }

        private void LogWarning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"WARNING: {message}");
            Console.ResetColor();
            
        }
    }

    public class PreprocessedBuData
    {
        public string OriginalNewCreatedPark { get; set; }
        public string ProcessedNewBU { get; set; }
        public List<string> UserDomains { get; set; }
    }
}