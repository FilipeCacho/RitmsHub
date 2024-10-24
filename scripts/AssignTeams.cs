using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace RitmsHub.Scripts
{
    public class AssignTeams
    {
        private IOrganizationService _service;
        private UserRetriever _userRetriever;

        public async Task ProcessAssignTeamsAsync()
        {
            //list to store disabled users to make sure they are not repeated when processing teams
            List<string> disabledUser = new List<string>();

            try
            {
                await ConnectToCrmAsync();
                List<AssignTeamData> assignTeamDataList = ExcelReader.ReadAssignTeamsData();

                if (!ValidateExcelData(assignTeamDataList))
                {
                    Console.ForegroundColor = ConsoleColor.Blue;

                    Console.WriteLine("Press any key to return to the menu.");
                    Console.ReadKey();
                    return;
                }

                foreach (var data in assignTeamDataList)
                {
                    await ProcessUserAsync(data.Username, data.TeamName, disabledUser);
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\nTeam assignment process completed. Press any key to return to the menu.");
                Console.ResetColor();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Press any key to return to the menu.");
                Console.ResetColor();
                Console.ReadKey();
            }
        }

        private bool ValidateExcelData(List<AssignTeamData> data)
        {
            bool isValid = true;
            for (int i = 0; i < data.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(data[i].Username) || string.IsNullOrWhiteSpace(data[i].TeamName))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Warning: Row {i + 2} has missing data. Username: '{data[i].Username}', Team Name: '{data[i].TeamName}'");
                    Console.ResetColor();
                    isValid = false;
                }
            }
            return isValid;
        }

        private Task ConnectToCrmAsync()
        {
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                //DynamicsCrmUtility.LogMessage($"Attempting to connect with: {connectionString}");

                var serviceClient = new CrmServiceClient(connectionString);

                if (serviceClient is null || !serviceClient.IsReady)
                {
                    throw new Exception($"Failed to connect. Error: {serviceClient?.LastCrmError ?? "Unknown error"}");
                }

                _service = serviceClient;
                _userRetriever = new UserRetriever(_service);
                //DynamicsCrmUtility.LogMessage($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred while connecting to CRM: {ex.Message}", "ERROR");
                throw;
            }

            return Task.CompletedTask;
        }

        private async Task ProcessUserAsync(string userIdentifier, string teamName, List<string> disabledUser)
        {
            // if a user is disabled the code just exists, 
            if (disabledUser.Contains(userIdentifier))
            {
                return;
            }

            var users = await _userRetriever.FindUsersAsync(userIdentifier);
            if (users.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"User {userIdentifier} not found.");
                Console.ResetColor();
                disabledUser.Add(userIdentifier); // adds user to disabled list to avoid processing it again if it's disabled
                return;
            }

            var user = users[0]; // Take the first user if multiple are found

            if (user.GetAttributeValue<bool>("isdisabled"))
            {
                disabledUser.Add(userIdentifier); // adds user to disabled list to avoid processing it again if it's disabled
                return;
            }

            string username = user.GetAttributeValue<string>("domainname").Split('@')[0];
            Console.WriteLine($"User {username} (active) - assigning team:");
            await EnsureUserHasTeam(user, teamName);
        }


        private async Task EnsureUserHasTeam(Entity user, string teamName)
        {
            var currentTeams = await GetUserTeamsAsync(user.Id);
            var currentTeamNames = currentTeams.Entities.Select(t => t.GetAttributeValue<string>("name")).ToList();

            if (currentTeamNames.Contains(teamName))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{teamName} (already assigned, skipped)");
                Console.ResetColor();
                return;
            }

            var team = await GetTeamByNameAsync(teamName);
            if (team != null)
            {
                var addMembersRequest = new AddMembersTeamRequest
                {
                    TeamId = team.Id,
                    MemberIds = new[] { user.Id }
                };

                await Task.Run(() => _service.Execute(addMembersRequest));
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{teamName} (assigned)");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{teamName} (team not found, skipped)");
                Console.ResetColor();
            }
        }

        private async Task<EntityCollection> GetUserTeamsAsync(Guid userId)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("name"),
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

            return await Task.Run(() => _service.RetrieveMultiple(query));
        }

        private async Task<Entity> GetTeamByNameAsync(string teamName)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("name", ConditionOperator.Equal, teamName)
                    }
                }
            };

            var result = await Task.Run(() => _service.RetrieveMultiple(query));
            return result.Entities.FirstOrDefault();
        }
    }

    public class AssignTeamData
    {
        public string Username { get; set; }
        public string TeamName { get; set; }
    }
}