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
    public class UserNormalizationResult
    {
        public string Username { get; set; }
        public string Name { get; set; }
        public bool IsInternal { get; set; }
    }

    public class RunNewUserWorkFlow
    {
        private static void ConfigureSecurityProtocol()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }

        public static async Task ExecuteWorkflowForUsersAsync(List<UserNormalizationResult> users)
        {
            try
            {
                ConfigureSecurityProtocol();
                DynamicsCrmUtility.ResetConnection();
                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();

                if (serviceClient == null || !serviceClient.IsReady)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"CRM service client is not ready. {serviceClient?.LastCrmError}");
                    Console.ResetColor();
                    return;
                }

                // Query for the workflow
                var query = new QueryExpression("workflow")
                {
                    ColumnSet = new ColumnSet("workflowid", "name", "statecode", "statuscode"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };

                query.Criteria.AddCondition("name", ConditionOperator.Equal, CodesAndRoles.NewUserWorkflow);
                query.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                query.Criteria.AddCondition("category", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 1);

                var workflowResults = serviceClient.RetrieveMultiple(query);
                if (workflowResults?.Entities == null || !workflowResults.Entities.Any())
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Workflow not found or not in a published state.");
                    Console.ResetColor();
                    return;
                }

                var workflow = workflowResults.Entities.First();

                foreach (var user in users)
                {
                    await ProcessUserWorkflowAsync(serviceClient, workflow, user);
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.ResetColor();
            }
        }

        private static async Task ProcessUserWorkflowAsync(CrmServiceClient serviceClient, Entity workflow, UserNormalizationResult user)
        {
            // First get the system user to get their ID
            var systemUser = await RetrieveSystemUserAsync(serviceClient, user.Username);
            if (systemUser == null)
            {
                Console.WriteLine($"System user not found for {user.Username}");
                return;
            }

            try
            {
                var executeWorkflowRequest = new OrganizationRequest("ExecuteWorkflow")
                {
                    ["WorkflowId"] = workflow.Id,
                    ["EntityId"] = systemUser.Id
                };

                // Execute the workflow and handle the exception directly
                await Task.Run(() =>
                {
                    try
                    {
                        serviceClient.Execute(executeWorkflowRequest);
                        Console.WriteLine($"Workflow executed successfully for user: {user.Name}, Username: {user.Username}");
                    }
                    catch (FaultException<OrganizationServiceFault> ex) when (
                        ex.Message.Contains("Ya hay otro registro de recurso asociado a este usuario") ||
                        ex.Message.Contains("There is already another resource record associated to this user"))
                    {
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.WriteLine("\nRunning workflow to activate user");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"User {user.Username} is already activated (Resource record exists)");
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("\nPress any key to continue");
                        Console.ResetColor();
                        Console.ReadLine();
                    }
                    catch (Exception ex)
                    {
                        throw; // Rethrow other exceptions to be caught by outer catch block
                    }
                });
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error executing workflow for {user.Username}: {ex.Message}");
                Console.ResetColor();
            }
        }

        private static async Task<Entity> RetrieveSystemUserAsync(CrmServiceClient serviceClient, string username)
        {
            try
            {
                var query = new QueryExpression("systemuser")
                {
                    ColumnSet = new ColumnSet("systemuserid", "domainname"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("domainname", ConditionOperator.Like, $"%{username}%")
                        }
                    }
                };

                var result = serviceClient.RetrieveMultiple(query);
                return result.Entities.FirstOrDefault();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving system user: {ex.Message}");
                return null;
            }
        }
    }
}