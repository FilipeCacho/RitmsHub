using Microsoft.Xrm.Sdk;
using System;
using System.Net;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserPermissionCopier
    {
        private IOrganizationService _service;
        private UserRetriever _userRetriever;
        private PermissionCopier _permissionCopier;

        public async Task Run()
        {
            try
            {
                await ConnectToCrmAsync();
                Console.Clear();
                Console.WriteLine("This code can copy from one user (the source user) its BU's, Roles or Teams to other user (the target user)");
                Console.WriteLine("Roles or Teams that already exist in the target user will be skipped");
                Console.WriteLine("(Or press 0 to return to main menu)");

                var sourceUser = await _userRetriever.PromptAndRetrieveUser("\nEnter the source user's name or EX: ");
                if (sourceUser == UserRetriever.Exit) return;
                if (sourceUser == null || !await IsUserInitializedAsync(sourceUser))
                {
                    Console.WriteLine("Invalid source user. Press any key to exit.");
                    Console.ReadKey();
                    return;
                }

                var targetUser = await _userRetriever.PromptAndRetrieveUser("\nEnter the target user's name or ID: ");
                if (targetUser == UserRetriever.Exit) return;
                if (targetUser == null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Target user not found. Press any key to exit.");
                    Console.ResetColor();
                    Console.ReadKey();
                    return;
                }

                Console.Clear();
                DisplayUserInfo(sourceUser, targetUser);

                bool copyBU = await PromptForPermissionCopy("Business Unit", sourceUser, targetUser);
                bool copyTeams = await PromptForPermissionCopy("Teams", sourceUser, targetUser);
                bool copyRoles = await PromptForPermissionCopy("Roles", sourceUser, targetUser);

                if (copyBU)
                {
                    await _permissionCopier.CopyBusinessUnit(sourceUser, targetUser);
                    targetUser = await _userRetriever.RetrieveUserAsync(targetUser.Id);
                }

                if (copyTeams)
                {
                    await _permissionCopier.CopyTeams(sourceUser, targetUser);
                }

                if (copyRoles)
                {
                    await _permissionCopier.CopyRoles(sourceUser, targetUser);
                }

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nAll operations completed");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nPress any key to continue");
                Console.ResetColor();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred: {ex.Message}", "ERROR");
                DynamicsCrmUtility.LogMessage($"Stack Trace: {ex.StackTrace}", "ERROR");
            }
        }

        private async Task ConnectToCrmAsync()
        {
            try
            {
                ConfigureSecurityProtocol();
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                DynamicsCrmUtility.LogMessage($"Attempting to connect with: {connectionString}");

                var serviceClient = await Task.Run(() => DynamicsCrmUtility.CreateCrmServiceClient());


                if (serviceClient is null || !serviceClient.IsReady)
                {
                    throw new Exception($"Failed to connect. Error: {(serviceClient?.LastCrmError ?? "Unknown error")}");
                }

                this._service = serviceClient;
                this._userRetriever = new UserRetriever(_service);
                this._permissionCopier = new PermissionCopier(_service);
                //DynamicsCrmUtility.LogMessage($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred while connecting to CRM: {ex.Message}", "ERROR");
                throw;
            }
        }

        private static void ConfigureSecurityProtocol()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }
    }
}