using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class UserRetriever
    {
        private readonly IOrganizationService _service;

        // Define a static field to represent the go back to main menu
        public static readonly Entity Exit = new Entity("exit");

        public UserRetriever(IOrganizationService service)
        {
            _service = service;
        }




        public async Task<Entity> PromptAndRetrieveUser(string prompt)
        {
            while (true)
            {
                Console.Write(prompt);
                var input = Console.ReadLine();

                //if the users inputs 0 it sends entity type back with a exit to indicate the code must go back to main
                if (input == "0")
                {
                    return Exit; 
                }



                var users = await FindUsersAsync(input);

                if (users.Count == 0)
                {
                    Console.WriteLine("No users found. Please try again.");
                    Console.WriteLine("If the user is not found it might be disabled");
                    continue;
                }

                if (users.Count == 1)
                {
                    return users[0];
                }

                while (true)
                {
                    Console.WriteLine("\nMultiple users found:");
                    for (int i = 0; i < users.Count; i++)
                    {
                        string domainName = users[i].Contains("domainname") ? users[i]["domainname"].ToString() : "N/A";
                        string username = domainName != "N/A" ? domainName.Split('@')[0] : "N/A";
                        string fullName = users[i].Contains("fullname") ? users[i]["fullname"].ToString() : "N/A";
                        Console.WriteLine($"({i + 1}) {fullName} (Username: {username})");
                    }

                    Console.Write($"\nSelect one of the users (1-{users.Count}), or press 0 to go back to the previous search: ");
                    if (int.TryParse(Console.ReadLine(), out int selection))
                    {
                        if (selection == 0)
                        {
                            // Go back to the previous search
                            

                            break;
                        }
                        else if (selection >= 1 && selection <= users.Count)
                        {
                            return users[selection - 1];
                        }
                    }

                    Console.WriteLine("\nInvalid selection. Please try again. Press any key to retry");

                    Console.ReadKey();
                    Console.Clear();
                }
            }
        }

        public async Task<List<Entity>> FindUsersAsync(string input)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("fullname", "businessunitid", "domainname", "internalemailaddress", "windowsliveid"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };

            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);

            var orFilter = new FilterExpression(LogicalOperator.Or);
            orFilter.AddCondition("domainname", ConditionOperator.BeginsWith, input);
            orFilter.AddCondition("internalemailaddress", ConditionOperator.BeginsWith, input);
            orFilter.AddCondition("windowsliveid", ConditionOperator.BeginsWith, input);
            orFilter.AddCondition("fullname", ConditionOperator.Like, $"%{input}%");
            orFilter.AddCondition("yomifullname", ConditionOperator.Like, $"%{input}%");

            query.Criteria.AddFilter(orFilter);

            var result = await _service.RetrieveMultipleAsync(query);

            // Perform additional case-insensitive filtering in memory
            var filteredResults = result.Entities.Where(e =>
                (e.Contains("domainname") && e["domainname"].ToString().IndexOf(input, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (e.Contains("internalemailaddress") && e["internalemailaddress"].ToString().IndexOf(input, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (e.Contains("windowsliveid") && e["windowsliveid"].ToString().IndexOf(input, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (e.Contains("fullname") && e["fullname"].ToString().IndexOf(input, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (e.Contains("yomifullname") && e["yomifullname"].ToString().IndexOf(input, StringComparison.OrdinalIgnoreCase) >= 0)
            ).ToList();

            return filteredResults;
        }

        public async Task<Entity> RetrieveUserAsync(Guid userId)
        {
            return await Task.Run(() => _service.Retrieve("systemuser", userId, new ColumnSet(true)));
        }
    }


}

