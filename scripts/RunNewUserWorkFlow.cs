using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using System.Linq;
using System.ServiceModel;

namespace RitmsHub.Scripts
{
    public class RunNewUserWorkFlow
    {
        public static async Task ExecuteWorkflowForUsersAsync(List<UserNormalizationResult> users)
        {
            try
            {
                ConfigureSecurityProtocol();

                string connectionString = DynamicsCrmUtility.CreateConnectionString();

                using var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();

                if (!serviceClient.IsReady)
                {
                    Console.WriteLine($"Failed to connect. Error: {serviceClient.LastCrmError}");
                    Console.WriteLine($"Detailed error: {serviceClient.LastCrmException?.ToString()}");
                    return;
                }

                //Console.WriteLine($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}\n");

                // Retrieve the workflow
                var workflow = await RetrieveWorkflowAsync(serviceClient, "Usuario-Proceso para crear un recurso desde el usuario");

                if (workflow == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Workflow not found or not in a published state.");
                    Console.WriteLine("Press any key to return to the main menu");
                    Console.ResetColor();
                    Console.ReadKey();
                    return;
                }

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine($"\nFound workflow: {workflow.GetAttributeValue<string>("name")}");
                Console.ResetColor();

                foreach (var user in users)
                {
                    await ProcessUserWorkflowAsync(serviceClient, workflow, user);
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                Console.ResetColor();

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nPress any key to return to the main menu");
                Console.ResetColor();
                Console.ReadKey();
            }
        }

        private static async Task ProcessUserWorkflowAsync(CrmServiceClient serviceClient, Entity workflow, UserNormalizationResult user)
        {
            var systemUser = await RetrieveSystemUserAsync(user.Username, serviceClient);
            if (systemUser == null)
            {
                Console.WriteLine($"System user not found for {user.Username}. Skipping processing.");

                Console.WriteLine("Press any key to return to the main menu");
                Console.ReadKey();

                return;
            }

            var mailbox = await RetrieveUserMailboxAsync(systemUser.Id, serviceClient);
            string mailboxName = mailbox?.GetAttributeValue<string>("name");

            if (string.IsNullOrEmpty(mailboxName))
            {
                Console.WriteLine($"Not found for user: {user.Username}. Skipping processing.");

                Console.WriteLine("Press any key to return to the main menu");
                Console.ReadKey();

                return;
            }

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Attempting to execute workflow for user: {user.Username}, {mailboxName}");
            Console.ResetColor();

            var (success, errorMessage) = await ExecuteWorkflowAsync(serviceClient, workflow.Id, systemUser.Id);

            if (success)
            {
                Console.WriteLine($"Workflow executed successfully for user: {user.Name}, Username: {user.Username}");

                Console.WriteLine("\nPress any key to return to the main menu");
                Console.ReadKey();
            }
            else if (errorMessage.Contains("There is already another resource record associated to this user"))
            {
                Console.WriteLine($"User: {user.Username} is already registered.");

                Console.WriteLine("\nPress any key to return to the main menu");
                Console.ReadKey();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Workflow execution failed. Error: {errorMessage}");
                Console.WriteLine("This is ok means the user is already is already activated");

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nPress any key to return to the main menu");
                Console.ResetColor();

                Console.ReadKey();
            }
        }

        private static async Task<Entity> RetrieveWorkflowAsync(IOrganizationService service, string workflowName)
        {
            var query = new QueryExpression("workflow")
            {
                ColumnSet = new ColumnSet("workflowid", "name", "statecode", "statuscode"),
                Criteria = new FilterExpression()
            };

            query.Criteria.AddCondition("name", ConditionOperator.Equal, workflowName);
            query.Criteria.AddCondition("type", ConditionOperator.Equal, 1); // Definition
            query.Criteria.AddCondition("category", ConditionOperator.Equal, 0); // Process
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 1); // Active

            EntityCollection results = await Task.Run(() => service.RetrieveMultiple(query));

            return results.Entities.FirstOrDefault();
        }

        private static async Task<(bool Success, string ErrorMessage)> ExecuteWorkflowAsync(IOrganizationService service, Guid workflowId, Guid userId)
        {
            try
            {
                var executeWorkflowRequest = new OrganizationRequest("ExecuteWorkflow")
                {
                    ["WorkflowId"] = workflowId,
                    ["EntityId"] = userId
                };

                var result = await Task.Run(() =>
                {
                    try
                    {
                        service.Execute(executeWorkflowRequest);
                        return (true, (string)null);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        return (false, ex.Message);
                    }
                });

                return result;
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private static async Task<Entity> RetrieveSystemUserAsync(string username, CrmServiceClient serviceClient)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid", "domainname", "isdisabled"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("domainname", ConditionOperator.Like, $"%{username}%")
                    }
                }
            };

            EntityCollection result = await Task.Run(() => serviceClient.RetrieveMultiple(query));

            if (result.Entities.Count > 1)
            {
                Console.WriteLine($"Multiple users found for username: {username}. Selecting the non-disabled user.");
                return result.Entities.FirstOrDefault(e => e.GetAttributeValue<bool>("isdisabled") == false);
            }

            return result.Entities.FirstOrDefault();
        }

        private static async Task<Entity> RetrieveUserMailboxAsync(Guid userId, CrmServiceClient serviceClient)
        {
            var query = new QueryExpression("mailbox")
            {
                ColumnSet = new ColumnSet("mailboxid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("regardingobjectid", ConditionOperator.Equal, userId)
                    }
                }
            };

            EntityCollection result = await Task.Run(() => serviceClient.RetrieveMultiple(query));

            return result.Entities.FirstOrDefault();
        }

        private static void ConfigureSecurityProtocol()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }
    }

    public class UserNormalizationResult
    {
        public string Username { get; set; }
        public string Name { get; set; }
        public bool IsInternal { get; set; }
    }
}