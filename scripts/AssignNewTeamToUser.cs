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
                Console.WriteLine($"An error occurred in ProcessUsersAsync: {ex.Message}", "ERROR");
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
                //Console.WriteLine($"Attempting to connect with: {connectionString}");

                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();
                this.service = serviceClient;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while connecting to CRM: {ex.Message}", "ERROR");
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
                        Console.WriteLine($"User {userDomain} successfully processed for park {buUserDomain.NewCreatedPark}");
                    }
                    else
                    {
                        Console.WriteLine($"User {userDomain} could not be processed for park {buUserDomain.NewCreatedPark}. Check previous errors for details.", "WARNING");
                    }
                }
            }

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Press any key to continue...");
            Console.ResetColor();
            Console.ReadKey();

            return processedUsers;
        }

        //find equipo dados maestros
        public static string FindParkMasterDataTeam(string input)
        {
            string lowercaseInput = input.ToLower();
            int firstContrataIndex = lowercaseInput.IndexOf("contrata");
            if (firstContrataIndex == -1)
            {
                return input.Trim();
            }

            int secondContrataIndex = lowercaseInput.IndexOf("contrata", firstContrataIndex + 1);
            if (secondContrataIndex == -1)
            {
                return input.Trim();
            }

            return input.Substring(0, secondContrataIndex).Trim();
        }

        private async Task<bool> AssignUserToTeamAsync(string userDomain, string equipoContrata)
        {
            string contrataContrataTeam = FindParkMasterDataTeam(equipoContrata);

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
                    Console.WriteLine($"User not found: {userDomain}", "WARNING");
                    return false;
                }

                Guid userId = userResults.Entities[0].Id;

                // Check if the Contrata Contrata team exists
                var contrataTeamQuery = new QueryExpression("team")
                {
                    ColumnSet = new ColumnSet("teamid"),
                    Criteria = new FilterExpression()
                };
                contrataTeamQuery.Criteria.AddCondition("name", ConditionOperator.Equal, contrataContrataTeam);

                EntityCollection contrataTeamResults = await Task.Run(() => service.RetrieveMultiple(contrataTeamQuery));
                if (contrataTeamResults.Entities.Count == 0)
                {
                    Console.WriteLine($"Contrata Contrata team not found: {contrataContrataTeam}", "WARNING");
                    return false;
                }

                Guid contrataTeamId = contrataTeamResults.Entities[0].Id;

                // Check if the user is a member of the Contrata Contrata team
                var membershipQuery = new QueryExpression("teammembership")
                {
                    Criteria = new FilterExpression()
                };
                membershipQuery.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
                membershipQuery.Criteria.AddCondition("teamid", ConditionOperator.Equal, contrataTeamId);

                EntityCollection membershipResults = await Task.Run(() => service.RetrieveMultiple(membershipQuery));

                bool contrataTeamAssigned = false;

                if (membershipResults.Entities.Count == 0)
                {
                    // User is not a member of the Contrata Contrata team, so add them
                    try
                    {
                        await Task.Run(() => service.Associate(
                            "team",
                            contrataTeamId,
                            new Relationship("teammembership_association"),
                            new EntityReferenceCollection() { new EntityReference("systemuser", userId) }
                        ));
                        Console.WriteLine($"User {userDomain} successfully assigned to Contrata Contrata team {contrataContrataTeam}");
                        contrataTeamAssigned = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to assign user {userDomain} to Contrata Contrata team {contrataContrataTeam}: {ex.Message}", "ERROR");
                        return false;
                    }
                }
                else
                {
                    contrataTeamAssigned = true;
                }

                // Only proceed with new team assignment if Contrata Contrata team assignment was successful
                if (contrataTeamAssigned)
                {
                    // Now assign the user to the new team
                    var newTeamQuery = new QueryExpression("team")
                    {
                        ColumnSet = new ColumnSet("teamid"),
                        Criteria = new FilterExpression()
                    };
                    newTeamQuery.Criteria.AddCondition("name", ConditionOperator.Equal, equipoContrata);

                    EntityCollection newTeamResults = await Task.Run(() => service.RetrieveMultiple(newTeamQuery));
                    if (newTeamResults.Entities.Count == 0)
                    {
                        Console.WriteLine($"New team not found: {equipoContrata}", "WARNING");
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
                            Console.WriteLine($"User {userDomain} successfully assigned to new team {equipoContrata}");
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to assign user {userDomain} to new team {equipoContrata}: {ex.Message}", "ERROR");
                            return false;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"User {userDomain} is already a member of the new team {equipoContrata}");
                        return true;
                    }
                }
                else
                {
                    Console.WriteLine($"User {userDomain} could not be assigned to new team {equipoContrata} because Contrata Contrata team assignment failed", "WARNING");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while assigning user to team: {ex.Message}", "ERROR");
                return false;
            }
        }
    }
}