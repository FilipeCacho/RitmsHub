using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class UserRoleManager
    {
        private IOrganizationService _service;
        private UserRetriever _userRetriever;
        private PermissionCopier _permissionCopier;
        private List<string> _savedRoleNames;
        private Guid _savedUserId;

        public async Task Run()
        {
            try
            {
                await ConnectToCrmAsync();
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nThis code will retrieve and manage user roles.");
                Console.ResetColor();

                var user = await _userRetriever.PromptAndRetrieveUser("\nEnter the user's name or username (or 0 to exit): ");

                if (user == UserRetriever.Exit)
                {
                    return;
                }

                if (user == null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("User not found. Press any key to exit.");
                    Console.ResetColor();
                    Console.ReadKey();
                    return;
                }

                _savedUserId = user.Id;
                await DisplayAndSaveUserRoles(user);

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nPress any key to continue...");
                Console.ReadKey();

                Console.Clear();
                Console.WriteLine("Options:");
                Console.WriteLine("1. Reapply saved roles");
                Console.WriteLine("2. Exit");
                Console.ResetColor();
                Console.Write("\nEnter your choice (1 or 2): ");

                string choice = Console.ReadLine();
                if (choice == "1")
                {
                    await ReapplySavedRoles();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                Console.ResetColor();
            }

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("\nPress any key to continue");
            Console.ResetColor();
            Console.ReadKey();
        }

        private Task ConnectToCrmAsync()
        {
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                //Console.WriteLine($"Attempting to connect with: {connectionString}");

                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();

                if (serviceClient is null || !serviceClient.IsReady)
                {
                    throw new Exception($"Failed to connect. Error: {serviceClient?.LastCrmError ?? "Unknown error"}");
                }

                _service = serviceClient;
                _userRetriever = new UserRetriever(_service);
                _permissionCopier = new PermissionCopier(_service);
                //Console.WriteLine($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while connecting to CRM: {ex.Message}", "ERROR");
                throw;
            }

            return Task.CompletedTask;
        }

        private async Task DisplayAndSaveUserRoles(Entity user)
        {
            var roles = await _permissionCopier.GetUserRolesAsync(user.Id);

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("\nUser Roles:");
            Console.ResetColor();
            foreach (var role in roles.Entities.OrderByDescending(r => r.GetAttributeValue<string>("name")))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"- {role.GetAttributeValue<string>("name")}");
                Console.ResetColor();
            }

            _savedRoleNames = roles.Entities.Select(r => r.GetAttributeValue<string>("name")).ToList();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\nSaved {_savedRoleNames.Count} role names for later use.");
            Console.ResetColor();
        }

        private async Task ReapplySavedRoles()
        {
            if (_savedRoleNames == null || _savedRoleNames.Count == 0)
            {
                Console.WriteLine("No role names were saved. Cannot reapply.");
                return;
            }

            Console.WriteLine($"\nReapplying {_savedRoleNames.Count} saved roles...\n");

            // Fetch the user's current information
            Entity currentUser = await Task.Run(() => _service.Retrieve("systemuser", _savedUserId, new ColumnSet("businessunitid")));
            var userBusinessUnitId = ((EntityReference)currentUser["businessunitid"]).Id;

            var currentRoles = await GetCurrentUserRoles(_savedUserId);

            foreach (var roleName in _savedRoleNames)
            {
                var equivalentRole = await FindRoleInBusinessUnitAsync(roleName, userBusinessUnitId);

                if (equivalentRole == null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"No equivalent role found for '{roleName}' in the user's current Business Unit. Skipping.");
                    Console.ResetColor();
                    continue;
                }

                if (currentRoles.Any(r => r.Id == equivalentRole.Id))
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.WriteLine($"Role '{equivalentRole.GetAttributeValue<string>("name")}' is already assigned. Skipping.");
                    Console.ResetColor();
                    continue;
                }

                try
                {
                    var request = new AssociateRequest
                    {
                        Target = new EntityReference("systemuser", _savedUserId),
                        RelatedEntities = new EntityReferenceCollection
                {
                    new EntityReference("role", equivalentRole.Id)
                },
                        Relationship = new Relationship("systemuserroles_association")
                    };

                    await Task.Run(() => _service.Execute(request));
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Role '{equivalentRole.GetAttributeValue<string>("name")}' reapplied successfully.");
                    Console.ResetColor();
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error reapplying role '{roleName}': {ex.Message}");
                    Console.ResetColor();
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nRole reapplication process completed.");
            Console.ResetColor();
        }

        private async Task<Entity> FindRoleInBusinessUnitAsync(string roleName, Guid businessUnitId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("name", ConditionOperator.Equal, roleName),
                        new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId)
                    }
                }
            };
            var result = await _service.RetrieveMultipleAsync(query);

            return result.Entities.FirstOrDefault();
        }

        private async Task<List<Entity>> GetCurrentUserRoles(Guid userId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
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
                                        new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var result = await _service.RetrieveMultipleAsync(query);
            return result.Entities.ToList();
        }
    }
}