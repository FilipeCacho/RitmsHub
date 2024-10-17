using RitmsHub.scripts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Discovery;

namespace RitmsHub.Scripts
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                if (!ExcelTemplateManager.CheckAndCreateExcelFile())
                {
                    Console.WriteLine("Failed to initialize the Excel file. The program will now exit.");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    return;
                }

                var menuHandler = new MenuHandler();
                await menuHandler.RunMenuAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }
    }

    class MenuHandler
    {

        private List<BuUserDomains> buUserDomainsList;
        private CancellationTokenSource _cts;

        private static readonly Dictionary<int, string> EnvironmentNames = new Dictionary<int, string>
        {
            {1, "DEV" },
            {2, "PRE" },
            {3, "PRD" }
        };


        public async Task RunMenuAsync()
        {
            while (true)
            {
                DisplayMenu();
                string choice = Console.ReadLine();

                if (!ExcelTemplateManager.ExcelFileExists() && choice != "14" && choice != "0")
                {
                    Console.WriteLine("Excel file not found. Please extract the template first (Option 14).");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    continue;
                }

                switch (choice)
                {
                    case "0":
                        Console.WriteLine("Exiting program...");
                        return;
                    case "1":
                        Console.Clear();
                        await CreateTeamProcessWithCancellation();

                        break;
                    case "2":
                        Console.Clear();
                        await ExtractUsersFromTeams();
                        break;
                    case "3":
                        Console.Clear();
                        await AssignTeamToUsers();
                        break;
                    case "4":
                        Console.Clear();
                        await NotificationsAndWorkorderViewCreatorAsync();
                        break;
                    case "5":
                        Console.Clear();
                        await ChangeUsersBuToNewTeamAsync();
                        break;
                    case "6":
                        Console.Clear();
                        var userNormalizer = new UserNormalizer();
                        List<UserNormalizationResult> results = await userNormalizer.Run();

                        if (results != null && results.Count > 0)
                        {
                            // Process SAP normalization only if users were normalized
                            UserSAPNormalizer.ProcessUsers(results);

                            // Run workflow for normalized users
                            await RunNewUserWorkFlow.ExecuteWorkflowForUsersAsync(results);
                        }
                        break;
                    case "7":
                        Console.Clear();
                        await ListUserWorkflows();
                        break;
                    case "8":
                        Console.Clear();
                        await UserCopier();
                        break;

                    case "9":
                        Console.Clear();
                        await HoldUserRoles();
                        break;

                    case "10":
                        Console.Clear();
                        await AssignTeamsFromExcel();
                        break;

                    case "11":
                        Console.Clear();
                        await ViewUserInfo();
                        break;

                    case "12":
                        Console.Clear();
                        await Releaselocks();
                        break;

                    case "13":
                        ChangeEnvironment();
                        break;

                    case "14":
                        ExtractExcelTemplate();
                        break;


                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        Console.WriteLine("\nPress any key to return to the menu...");
                        Console.ReadKey();
                        break;
                }

                if (choice != "14" && choice != "0" && !ExcelTemplateManager.ValidateExcelFile())
                {
                    Console.WriteLine("Excel file validation failed. Please check the file contents.");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
            }
        }

        private void DisplayMenu()
        {
            Console.Clear();

                //color set in the excel template manager, for some unknown reason only works there
                Console.WriteLine($"CRM Operations Menu (connected to {EnvironmentNames[ExcelReader.CurrentEnvironment]}):\n");
            

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Create Teams flow");
            Console.WriteLine("-----------------------------------------------------------------------------------------------------");
            Console.ResetColor();
            Console.WriteLine("1.  Create/Update Team Process (BU, Contrata team, EDPR Team) (1-5 is only for EU teams only)");
            Console.WriteLine("2.  Extract users from BU and it's contrata Contrata team");
            Console.WriteLine("3.  Give newly created team to extracted users");
            Console.WriteLine("4.  Create views for workorders and notifications (ordens de trabajo e avisos)");
            Console.WriteLine("5.  Change BU of users that have in their name the Contractor (Puesto de trabajo)");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("-----------------------------------------------------------------------------------------------------");


            Console.WriteLine("\nOther flows");
            Console.WriteLine("-----------------------------------------------------------------------------------------------------");
            Console.ResetColor();
            Console.WriteLine("6.  Normalize Users");
            Console.WriteLine("7.  Find user workflows");
            Console.WriteLine("8.  Copy BU, Teams and Roles from one user to the other");
            Console.WriteLine("9.  Hold user roles while BU is replaced");
            Console.WriteLine("10. Assign Teams to Users in the 'Assign Teams' worksheet");
            Console.WriteLine("11. View info about 1 or 2 users");
            Console.WriteLine("12. Release any pending system locks");
            Console.WriteLine("13. Change connection to DEV, PRE or PRD");
            Console.WriteLine("14. Extract Excel template");
            Console.WriteLine("0.  Exit");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("-----------------------------------------------------------------------------------------------------");

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\nEditing the .xls file data while the console is open can lead to cached data being used, do it while the project is not running \n");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Closing the console windows or press CTRL + C while the code is making requests to the DB can cause table locks, don't do it");
            Console.ResetColor();

            Console.WriteLine("\nEnter your choice (0-13): ");
        }

        //sequential team creation
        private async Task CreateTeamProcessWithCancellation()
        {
            using (_cts = new CancellationTokenSource())
            {
                Console.WriteLine("Press 'Q' at any time to cancel the operation.");

                // Start a task to listen for the 'Q' key
                var cancellationTask = Task.Run(() =>
                {
                    while (!_cts.Token.IsCancellationRequested)
                    {
                        if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
                        {
                            _cts.Cancel();
                            Console.WriteLine("Cancellation requested. Please wait...");
                        }
                    }
                });

                try
                {
                    List<TransformedTeamData> transformedTeams = FormatBUandTeams.FormatTeamData();
                    if (transformedTeams != null && transformedTeams.Count > 0)
                    {
                        // Create or verify Business Units
                        var buResults = await CreateBu.RunAsync(transformedTeams, _cts.Token);

                        // Process teams
                        var standardTeamResults = await CreateTeam.RunAsync(transformedTeams, TeamType.Standard, _cts.Token);
                        var proprietaryTeamResults = await CreateTeam.RunAsync(transformedTeams, TeamType.Proprietary, _cts.Token);



                        // Display results
                        DisplayResults(buResults, standardTeamResults, proprietaryTeamResults);
                    }
                    else
                    {
                        Console.WriteLine("No teams to create were found, press any key to continue");
                    }
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Operation was cancelled by the user.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
                finally
                {
                    // Cancel the token to stop the cancellation task
                    if (!_cts.IsCancellationRequested)
                    {
                        _cts.Cancel();
                    }

                    // Wait for the cancellation task to complete
                    await cancellationTask;
                }
            }


        }

        private void DisplayResults(List<BuCreationResult> buResults, List<TeamOperationResult> standardTeamResults, List<TeamOperationResult> proprietaryTeamResults)
        {
            Console.WriteLine("\nProcess Results:");
            Console.WriteLine("------------------");

            foreach (var buResult in buResults)
            {
                Console.WriteLine($"BU: {buResult.BuName} - {(buResult.Exists ? "Exists/Created" : "Failed to create")}");

                var standardTeam = standardTeamResults.FirstOrDefault(tr => tr.BuName == buResult.BuName);
                var proprietaryTeam = proprietaryTeamResults.FirstOrDefault(tr => tr.BuName == buResult.BuName);

                if (standardTeam != null)
                    Console.WriteLine($"  Standard Team: {(standardTeam.Exists ? (standardTeam.WasUpdated ? "Updated" : "Already Exists") : "Created")}");
                if (proprietaryTeam != null)
                    Console.WriteLine($"  Proprietary Team: {(proprietaryTeam.Exists ? (proprietaryTeam.WasUpdated ? "Updated" : "Already Exists") : "Created")}");

                Console.WriteLine();
            }

            Console.WriteLine("Press any key 2 times (except 'q' to continue)...");
            Console.ReadKey();
        }

        // New classes to hold result information
        public class TeamCreationResult
        {
            public string TeamName { get; set; }
            public bool Exists { get; set; }
            public bool WasUpdated { get; set; }
            public TeamType TeamType { get; set; }
            public string BuName { get; set; }
            public bool Success { get; set; }
        }

        private async Task ExtractUsersFromTeams()
        {
            List<TransformedTeamData> transformedBus = await ExtractUsersFromTeam.FormatTeamData();



            if (transformedBus != null)
            {
                buUserDomainsList = await ExtractUsersFromTeam.CreateExcel(transformedBus);
                PrintBuUserDomains();
            }
        }

        private async Task AssignTeamToUsers()
        {
            if (buUserDomainsList != null)
            {
                var processor = new AssignNewTeamToUser(buUserDomainsList);
                List<ProcessedUser> processedUsers = await processor.ProcessUsersAsync();
            }
            else
            {
                Console.WriteLine("You must run Option 2 first");
                Console.WriteLine("Press any key to return to the main menu");
                Console.ReadKey();
            }


        }

        private async Task NotificationsAndWorkorderViewCreatorAsync()
        {
            var teamDataList = FormatBUandTeams.FormatTeamData();
            if (teamDataList != null)
            {
                Console.Clear();

                //create workorder view
                var viewWorkOrderCreator = new WorkOrderViewCreator(teamDataList);
                await viewWorkOrderCreator.CreateAndSaveWorkOrderViewAsync();

                Console.Clear();

                //create notifications view
                var viewNotificationsCreator = new NotificationsViewCreator(teamDataList);
                await viewNotificationsCreator.CreateAndSaveNotificationsViewAsync();
            }
        }

        private async Task ChangeUsersBuToNewTeamAsync()
        {
            if (buUserDomainsList != null)
            {
                List<TransformedTeamData> transformedTeams = FormatBUandTeams.FormatTeamData();

                if (transformedTeams != null && transformedTeams.Count > 0)
                {
                    var changeUsersBuManager = new ChangeUsersBuToNewTeam();
                    await changeUsersBuManager.Run(buUserDomainsList, transformedTeams);
                }
                else
                {
                    Console.WriteLine("No transformed teams data available. Please run option 1 first.");
                }
            }
            else
            {
                Console.WriteLine("You must run Option 2 first to extract users from teams.");
                Console.WriteLine("Press any key to return to the main menu");
                Console.ReadKey();
            }
        }

        private async Task ListUserWorkflows()
        {
            var workflowLister = new WorkflowLister();
            await workflowLister.ListUserWorkflowsAsync("EX121935@edpr.com");
            await workflowLister.DiagnoseWorkflowRetrievalAsync();
        }

        private async Task UserCopier()
        {
            var copier = new UserPermissionCopier();
            await copier.Run();
        }

        private void PrintBuUserDomains()
        {
            Console.Clear();
            Console.WriteLine("Processed Users (only active and distinct) from each BU that will need to receive the new team\n");
            foreach (var buUserDomains in buUserDomainsList)
            {
                Console.WriteLine($"{buUserDomains.NewCreatedPark.Replace("Equipo contrata", "").Trim()}:");

                foreach (var userDomain in buUserDomains.UserDomains)
                {
                    Console.WriteLine(userDomain);
                }
                Console.WriteLine(); // Empty line between business units
            }
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }

        private async Task HoldUserRoles()
        {
            var userRoleManager = new UserRoleManager();
            await userRoleManager.Run();
        }

        private async Task AssignTeamsFromExcel()
        {
            try
            {
                var assignTeams = new AssignTeams();
                await assignTeams.ProcessAssignTeamsAsync();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.ResetColor();
                Console.WriteLine("Press any key to return to the menu.");
                Console.ReadKey();
            }
        }

        private async Task ViewUserInfo()
        {
            var userInfoRetriever = new UserBasicInfoRetriever();
            await userInfoRetriever.RetrieveAndCompareUserInfoAsync();
        }

        private async Task Releaselocks()
        {
            string connectionString = DynamicsCrmUtility.CreateConnectionString();
            var lockManager = new CrmLockManager(connectionString);
            await lockManager.ReleaseLocks();
            Console.WriteLine("Press any key to return to the menu.");
            Console.ReadKey();

        }

        private void ChangeEnvironment()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Select environment:");
                Console.WriteLine("1. DEV");
                Console.WriteLine("2. PRE");
                Console.WriteLine("3. PRD");
                Console.WriteLine("0. Go back to main menu");
                Console.Write("\nEnter your choice (0-3): ");

                string input = Console.ReadLine();
                if (input == "0")
                {
                    return;
                }

                if (int.TryParse(input, out int choice) && choice >= 1 && choice <= 3)
                {
                    ExcelReader.CurrentEnvironment = choice;
                    DynamicsCrmUtility.ResetConnection();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"\nEnvironment changed to {EnvironmentNames[choice]}.");

                    // Test the connection to verify the change
                    try
                    {
                        var client = DynamicsCrmUtility.CreateCrmServiceClient();
                        if (client.IsReady)
                        {
                            var orgUrl = client.ConnectedOrgPublishedEndpoints[EndpointType.WebApplication];
                            Console.WriteLine($"Successfully connected to: {orgUrl}");
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Failed to connect: Client is not ready.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Failed to connect: {ex.Message}");
                    }

                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine("\nPress any key to return to the main menu...");
                    Console.ReadKey();
                    Console.Clear();
                    return;
                }

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nInvalid choice. Please try again.");
                Console.WriteLine("Press any key to continue...");
                Console.ResetColor();

                Console.ReadKey();
            }
        }

        private void ExtractExcelTemplate()
        {
            try
            {
                ExcelTemplateManager.ExtractExcelTemplate();
                Console.WriteLine("Excel template extracted successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting Excel template: {ex.Message}");
            }
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

    }
}