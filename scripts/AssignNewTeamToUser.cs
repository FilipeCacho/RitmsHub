using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class ProcessedUser
    {
        public string UserDomain { get; set; }
        public string AssignedPark { get; set; }
    }

    public class AssignNewTeamToUser
    {
        private List<BuUserDomains> buUserDomainsList;
        private Dictionary<string, string> parkToEquipoContrata;
        private IOrganizationService service;

        public AssignNewTeamToUser(List<BuUserDomains> buUserDomainsList)
        {
            this.buUserDomainsList = buUserDomainsList;
            this.parkToEquipoContrata = new Dictionary<string, string>();
        }

        public async Task<List<ProcessedUser>> ProcessUsersAsync()
        {
            try
            {
                ConfigureSecurityProtocol();
                await ConnectToCrmAsync();
                CreateEquipoContrataMapping();
                return await AssignUsersToParksAsync();

            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred in ProcessUsersAsync: {ex.Message}", "ERROR");
                throw;
            }


        }

        private void ConfigureSecurityProtocol()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }

        private async Task ConnectToCrmAsync()
        {
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                DynamicsCrmUtility.LogMessage($"Attempting to connect with: {connectionString}");

                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();
                this.service = serviceClient;
                //DynamicsCrmUtility.LogMessage($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred while connecting to CRM: {ex.Message}", "ERROR");
                throw;
            }
        }

        private void CreateEquipoContrataMapping()
        {
            foreach (var buUserDomain in buUserDomainsList)
            {
                string equipoContrata = AdjustCreatedTeamName(buUserDomain.NewCreatedPark);
                parkToEquipoContrata[buUserDomain.NewCreatedPark] = equipoContrata;
            }
        }

        private string AdjustCreatedTeamName(string parkName)
        {
            string[] words = parkName.Trim().Split(' ');
            int lastValidIndex = words.Length - 1;

            for (int i = words.Length - 1; i >= 0; i--)
            {
                if (Regex.IsMatch(words[i], @"^[a-zA-Z]+-[a-zA-Z]+-[a-zA-Z]+$"))
                {
                    lastValidIndex = i;
                    break;
                }
            }

            return string.Join(" ", words.Take(lastValidIndex + 1)).Trim();
        }

        private async Task<List<ProcessedUser>> AssignUsersToParksAsync()
        {
            List<ProcessedUser> processedUsers = new List<ProcessedUser>();
            foreach (var buUserDomain in buUserDomainsList)
            {
                string equipaContrata = parkToEquipoContrata[buUserDomain.NewCreatedPark];
                foreach (var userDomain in buUserDomain.UserDomains)
                {
                    bool assigned = await AssignUserToTeamAsync(userDomain, equipaContrata);
                    if (assigned)
                    {
                        processedUsers.Add(new ProcessedUser
                        {
                            UserDomain = userDomain,
                            AssignedPark = buUserDomain.NewCreatedPark
                        });
                        DynamicsCrmUtility.LogMessage($"User {userDomain} successfully processed for park {buUserDomain.NewCreatedPark}");
                    }
                    else
                    {
                        DynamicsCrmUtility.LogMessage($"User {userDomain} skipped for park {buUserDomain.NewCreatedPark}. Check previous warnings for details.", "WARNING");
                    }
                }
            }

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Press any key to continue...");
            Console.ResetColor();
            Console.ReadKey();

            return processedUsers;
        }

        public static string FindContrataContrataTeam(string input)
        {
            int lastContrataIndex = input.LastIndexOf("Contrata", StringComparison.Ordinal);

            if (lastContrataIndex == -1)
            {
                return input;
            }

            int lastSpaceIndex = input.LastIndexOf(" ", lastContrataIndex, StringComparison.Ordinal);

            if (lastSpaceIndex == -1)
            {
                return input;
            }

            return input.Substring(0, lastContrataIndex + "Contrata".Length);
        }

        private async Task<bool> AssignUserToTeamAsync(string userDomain, string equipoContrata)
        {
            string contrataContrataTeam = FindContrataContrataTeam(equipoContrata);

            try
            {
                // Retrieve the user
                var userQuery = new QueryExpression("systemuser")
                {
                    ColumnSet = new ColumnSet("systemuserid"),
                    Criteria = new FilterExpression()
                };
                userQuery.Criteria.AddCondition("domainname", ConditionOperator.Equal, userDomain);

                EntityCollection userResults = await Task.Run(() => service.RetrieveMultiple(userQuery));
                if (userResults.Entities.Count == 0)
                {
                    DynamicsCrmUtility.LogMessage($"User not found: {userDomain}", "WARNING");
                    return false;
                }

                Guid userId = userResults.Entities[0].Id;

                // Check if the user is already a member of the Contrata Contrata team
                var contrataTeamQuery = new QueryExpression("team")
                {
                    ColumnSet = new ColumnSet("teamid"),
                    Criteria = new FilterExpression()
                };
                contrataTeamQuery.Criteria.AddCondition("name", ConditionOperator.Equal, contrataContrataTeam);

                EntityCollection contrataTeamResults = await Task.Run(() => service.RetrieveMultiple(contrataTeamQuery));
                if (contrataTeamResults.Entities.Count == 0)
                {
                    DynamicsCrmUtility.LogMessage($"Contrata Contrata team not found: {contrataContrataTeam}", "WARNING");
                    return false;
                }

                Guid contrataTeamId = contrataTeamResults.Entities[0].Id;

                var membershipQuery = new QueryExpression("teammembership")
                {
                    Criteria = new FilterExpression()
                };
                membershipQuery.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
                membershipQuery.Criteria.AddCondition("teamid", ConditionOperator.Equal, contrataTeamId);

                EntityCollection membershipResults = await Task.Run(() => service.RetrieveMultiple(membershipQuery));

                if (membershipResults.Entities.Count == 0)
                {
                    DynamicsCrmUtility.LogMessage($"User {userDomain} is not a member of the required Contrata Contrata team {contrataContrataTeam}", "WARNING");
                    return false;
                }

                // User is a member of the Contrata Contrata team, now assign them to the new team
                var newTeamQuery = new QueryExpression("team")
                {
                    ColumnSet = new ColumnSet("teamid"),
                    Criteria = new FilterExpression()
                };
                newTeamQuery.Criteria.AddCondition("name", ConditionOperator.Equal, equipoContrata);

                EntityCollection newTeamResults = await Task.Run(() => service.RetrieveMultiple(newTeamQuery));
                if (newTeamResults.Entities.Count == 0)
                {
                    DynamicsCrmUtility.LogMessage($"New team not found: {equipoContrata}", "WARNING");
                    return false;
                }

                Guid newTeamId = newTeamResults.Entities[0].Id;

                // Check if user is already a member of the new team
                var newMembershipQuery = new QueryExpression("teammembership")
                {
                    Criteria = new FilterExpression()
                };
                newMembershipQuery.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
                newMembershipQuery.Criteria.AddCondition("teamid", ConditionOperator.Equal, newTeamId);

                EntityCollection newMembershipResults = await Task.Run(() => service.RetrieveMultiple(newMembershipQuery));

                if (newMembershipResults.Entities.Count == 0)
                {
                    // User is not a member of the new team, so add them
                    try
                    {
                        await Task.Run(() => service.Associate(
                            "team",
                            newTeamId,
                            new Relationship("teammembership_association"),
                            new EntityReferenceCollection() { new EntityReference("systemuser", userId) }
                        ));
                        DynamicsCrmUtility.LogMessage($"User {userDomain} successfully assigned to new team {equipoContrata}");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        DynamicsCrmUtility.LogMessage($"Failed to assign user {userDomain} to new team {equipoContrata}: {ex.Message}", "ERROR");
                        return false;
                    }
                }
                else
                {
                    DynamicsCrmUtility.LogMessage($"User {userDomain} is already a member of the new team {equipoContrata}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred while assigning user to team: {ex.Message}", "ERROR");
                return false;
            }
        }
    }
}