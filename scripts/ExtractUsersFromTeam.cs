using ClosedXML.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public static class ExtractUsersFromTeam
    {
        public static async Task<List<TransformedTeamData>> FormatTeamData()
        {
            List<TransformedTeamData> dynamicTeams = new List<TransformedTeamData>();
            List<TeamRow> validTeams = ExcelReader.ValidateTeamsToCreate();

            if (validTeams.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nNo valid teams found to process.");
                Console.ResetColor();
                return null;
            }

            foreach (var team in validTeams)
            {
                string bu = FormatBusinessUnitName(team);

                // Remove the contractor code so that we can find the BU and the respective contrata team that hold the users
                // The contrata contrata team is also affected because its name is generated from the BU
                string buWithoutContractor = bu.Substring(0, bu.LastIndexOf(' '));

                dynamicTeams.Add(new TransformedTeamData
                {
                    Bu = buWithoutContractor,
                    EquipaContrataContrata = $"Equipo contrata {buWithoutContractor}".Trim(),
                    FileName = team.ColumnA,
                    // Add this new property to store the full BU name including the contractor code
                    FullBuName = bu
                });
            }

            DisplayTransformedTeamData(dynamicTeams);

            if (!GetUserConfirmation())
            {
                return null;
            }

            // Validate existence of BUs and Teams before proceeding
            bool allExist = await ValidateBusAndTeamsExistAsync(dynamicTeams);
            if (!allExist)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nProcess stopped due to missing Business Units or Teams. Please check the datacenter XLS file and correct the information.");
                Console.ResetColor();
                return null;
            }

            return dynamicTeams;
        }

        private static async Task<bool> ValidateBusAndTeamsExistAsync(List<TransformedTeamData> teams)
        {
            ConfigureSecurityProtocol();
            string connectionString = DynamicsCrmUtility.CreateConnectionString();
            var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();

            bool allExist = true;

            foreach (var team in teams)
            {
                bool buExists = await CheckBusinessUnitExistsAsync(serviceClient, team.Bu);
                bool teamExists = await CheckTeamExistsAsync(serviceClient, team.EquipaContrataContrata);

                if (!buExists)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"\nBusiness Unit '{team.Bu}' not found in Dynamics.");
                    Console.ResetColor();
                    allExist = false;
                }

                if (!teamExists)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"\nTeam '{team.EquipaContrataContrata}' not found in Dynamics.");
                    Console.ResetColor();
                    allExist = false;
                }
            }

            if ( !allExist ) {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nPress any key to go back to the main menu");
                Console.ResetColor();
                Console.ReadKey();
            }

            return allExist;
        }

        private static async Task<bool> CheckBusinessUnitExistsAsync(IOrganizationService service, string businessUnitName)
        {
            var query = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet("businessunitid"),
                Criteria = new FilterExpression()
            };
            query.Criteria.AddCondition("name", ConditionOperator.Equal, businessUnitName);

            EntityCollection results = await service.RetrieveMultipleAsync(query);
            return results.Entities.Count > 0;
        }

        private static async Task<bool> CheckTeamExistsAsync(IOrganizationService service, string teamName)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid"),
                Criteria = new FilterExpression()
            };
            query.Criteria.AddCondition("name", ConditionOperator.Equal, teamName);

            EntityCollection results = await service.RetrieveMultipleAsync(query);
            return results.Entities.Count > 0;
        }

        public static async Task<List<BuUserDomains>> CreateExcel(List<TransformedTeamData> transformedBus)
        {
            List<BuUserDomains> result = new List<BuUserDomains>();

            try
            {
                ConfigureSecurityProtocol();
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();

                string excelFolderPath = CreateExcelFolder();

                using (var cts = new CancellationTokenSource())
                {
                    foreach (var team in transformedBus)
                    {
                        var allUsers = await RetrieveAllUsersAsync(serviceClient, team, cts.Token);
                        var buUserDomains = CreateBuUserDomains(allUsers, team);
                        result.Add(buUserDomains);

                        await CreateExcelFileAsync(allUsers, team.FileName, excelFolderPath);
                    }
                }
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nAll excel files created, press any key to continue");
                Console.ResetColor();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}", "ERROR");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}", "ERROR");
            }

            return result;
        }

        // Helper methods

        private static string FormatBusinessUnitName(TeamRow team)
        {
            string bu = team.ColumnA;
            if (team.ColumnC != "ZP1")
            {
                bu += $" {team.ColumnC}";
            }
            return $"{bu} Contrata {team.ColumnB}";
        }

        private static void DisplayTransformedTeamData(List<TransformedTeamData> dynamicTeams)
        {
            foreach (var team in dynamicTeams)
            {
                Console.WriteLine("\nTransformed Team Data:");
                Console.WriteLine($"bu: {team.Bu}");
                Console.WriteLine($"Equipa contrata Contrata: {team.EquipaContrataContrata}");
                Console.WriteLine($"Users will be stored in the file: {team.FileName}.xls\n");
            }
        }

        private static bool GetUserConfirmation()
        {
            string input;
            do
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Do you want to extract the users from these BUs and their corresponding general Team (the Contrata contrata Team)?");
                Console.WriteLine("Each BU-Team pair will have its own excel file inside the 'generated excels' folder in your Downloads folder\n");
                Console.WriteLine("The users from each BU and it's general team will be processed, duplicates will be removed and only active users will be included\n");
                Console.ResetColor();
                Console.Write("Enter your choice (y/n): ");
                input = Console.ReadLine().ToLower();

                if (input == "y")
                {
                    return true;
                }
                else if (input == "n")
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine("You chose No. Returning to the previous menu.");
                    Console.ResetColor();
                    return false;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Invalid input. Please try again.");
                    Console.ResetColor();
                    Console.WriteLine();
                }
            } while (input != "y" && input != "n");

            return false;
        }

        private static void ConfigureSecurityProtocol()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }

        private static string CreateExcelFolder()
        {
            string downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string excelFolderPath = Path.Combine(downloadsFolder, "generated excels");
            Directory.CreateDirectory(excelFolderPath);
            return excelFolderPath;
        }

        private static async Task<List<UserData>> RetrieveAllUsersAsync(CrmServiceClient serviceClient, TransformedTeamData team, CancellationToken cancellationToken)
        {
            List<UserData> usersFromBu = await RetrieveUsersFromBuAsync(serviceClient, team.Bu, cancellationToken);
            List<UserData> usersFromTeam = await RetrieveUsersFromTeamAsync(serviceClient, team.EquipaContrataContrata, cancellationToken);

            return usersFromBu.Concat(usersFromTeam)
                              .GroupBy(u => u.YomiFullName)
                              .Select(g => g.First())
                              .ToList();
        }

        private static BuUserDomains CreateBuUserDomains(List<UserData> allUsers, TransformedTeamData team)
        {
            return new BuUserDomains
            {
                NewCreatedPark = $"Equipo contrata {team.FullBuName}".Trim(),
                UserDomains = allUsers.Select(u => u.DomainName).ToList()
            };
        }

        public static async Task CreateExcelFileAsync(List<UserData> users, string fileName, string folderPath)
        {
            string filePath = Path.Combine(folderPath, $"{fileName}.xlsx");

            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Users");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "Yomi Full Name";
                    worksheet.Cell(1, 2).Value = "Domain Name";
                    worksheet.Cell(1, 3).Value = "Business Unit";
                    worksheet.Cell(1, 4).Value = "Team";

                    // Add data
                    for (int i = 0; i < users.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = users[i].YomiFullName;
                        worksheet.Cell(i + 2, 2).Value = users[i].DomainName;
                        worksheet.Cell(i + 2, 3).Value = users[i].BusinessUnit;
                        worksheet.Cell(i + 2, 4).Value = users[i].Team;
                    }

                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();

                    // Save the workbook
                    workbook.SaveAs(filePath);
                }
            });
            Console.WriteLine("\n");
            Console.WriteLine($"Excel file created successfully at {filePath}");
        }

        private static async Task<List<UserData>> RetrieveUsersFromBuAsync(IOrganizationService service, string businessUnitName, CancellationToken cancellationToken)
        {
            var businessUnitId = await DynamicsCrmUtility.GetBusinessUnitIdAsync(service, businessUnitName, cancellationToken);

            if (businessUnitId == Guid.Empty)
            {
                Console.WriteLine($"Business Unit '{businessUnitName}' not found.", "WARNING");
                return new List<UserData>();
            }

            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("yomifullname", "domainname", "isdisabled"),
                Criteria = new FilterExpression()
            };

            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, businessUnitId);
            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);

            EntityCollection results = await service.RetrieveMultipleAsync(query);

            return results.Entities.Select(e => new UserData
            {
                YomiFullName = e.GetAttributeValue<string>("yomifullname"),
                DomainName = e.GetAttributeValue<string>("domainname"),
                BusinessUnit = businessUnitName,
                Team = ""
            }).ToList();
        }

        private static async Task<List<UserData>> RetrieveUsersFromTeamAsync(IOrganizationService service, string teamName, CancellationToken cancellationToken)
        {
            var teamQuery = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid"),
                Criteria = new FilterExpression()
            };
            teamQuery.Criteria.AddCondition("name", ConditionOperator.Equal, teamName);

            EntityCollection teamResults = await service.RetrieveMultipleAsync(teamQuery);

            if (teamResults.Entities.Count == 0)
            {
                Console.WriteLine($"Team '{teamName}' not found.", "WARNING");
                return new List<UserData>();
            }

            Guid teamId = teamResults.Entities[0].Id;

            var userQuery = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("yomifullname", "domainname", "isdisabled", "businessunitid"),
                Criteria = new FilterExpression()
            };

            userQuery.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);

            var teamLink = userQuery.AddLink("teammembership", "systemuserid", "systemuserid");
            teamLink.LinkCriteria.AddCondition("teamid", ConditionOperator.Equal, teamId);

            var buLink = userQuery.AddLink("businessunit", "businessunitid", "businessunitid");
            buLink.Columns = new ColumnSet("name");
            buLink.EntityAlias = "bu";

            EntityCollection userResults = await service.RetrieveMultipleAsync(userQuery);

            return userResults.Entities.Select(e => new UserData
            {
                YomiFullName = e.GetAttributeValue<string>("yomifullname"),
                DomainName = e.GetAttributeValue<string>("domainname"),
                BusinessUnit = e.GetAttributeValue<AliasedValue>("bu.name")?.Value?.ToString() ?? "Unknown",
                Team = teamName
            }).ToList();
        }
    }
}