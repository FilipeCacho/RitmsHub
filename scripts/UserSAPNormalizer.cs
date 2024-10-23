using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace RitmsHub.Scripts
{
    public static class UserSAPNormalizer
    {
        public static void ProcessUsers(List<UserNormalizationResult> results)
        {
            if (results.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nProcessing SAP Credentials for Users");
                Console.ResetColor();

                using (var serviceClient = CreateServiceClient())
                {
                    foreach (var result in results)
                    {
                        ProcessUser(result, serviceClient);
                    }
                }
            }
        }

        private static CrmServiceClient CreateServiceClient()
        {
            string connectionString = DynamicsCrmUtility.CreateConnectionString();
            var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();
            if (serviceClient is null || !serviceClient.IsReady)
            {
                throw new Exception($"Failed to connect. Error: {(serviceClient?.LastCrmError ?? "Unknown error")}");
            }
            return serviceClient;
        }

        private static void ProcessUser(UserNormalizationResult result, CrmServiceClient serviceClient)
        {
            if (!result.IsInternal)
            {
                Console.WriteLine($"User {result.Username} is not internal. Skipping process.");
                return;
            }

            var systemUser = RetrieveSystemUser(result.Username, serviceClient);
            if (systemUser == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"System user not found for {result.Username}. Skipping processing.");
                Console.ResetColor();
                return;
            }

            if (systemUser.GetAttributeValue<bool>("isdisabled"))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"User {result.Username} is disabled. Skipping processing.");
                Console.ResetColor();
                return;
            }

            var atos_usuariosapid = systemUser.GetAttributeValue<EntityReference>("atos_usuariosapid");
            if (atos_usuariosapid != null)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"User {result.Username} already has a SAP user assigned to it. Skipping processing.");
                Console.ResetColor();
                return;
            }

            var existingAtosUsuarios = RetrieveAtosUsuarios(result.Username, serviceClient);
            if (existingAtosUsuarios != null)
            {
                LinkExistingAtosUsuarios(systemUser, existingAtosUsuarios, serviceClient);
            }
            else
            {
                CreateAndLinkAtosUsuarios(result, systemUser, serviceClient);
            }
        }

        private static Entity RetrieveSystemUser(string username, CrmServiceClient serviceClient)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid", "isdisabled", "atos_usuariosapid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("domainname", ConditionOperator.Like, username + "%")
                    }
                }
            };
            var result = serviceClient.RetrieveMultiple(query);
            return result.Entities.FirstOrDefault();
        }

        private static Entity RetrieveAtosUsuarios(string username, CrmServiceClient serviceClient)
        {
            var query = new QueryExpression("atos_usuarios")
            {
                ColumnSet = new ColumnSet("atos_usuariosid", "atos_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("atos_codigo", ConditionOperator.Equal, username)
                    }
                }
            };
            var result = serviceClient.RetrieveMultiple(query);
            return result.Entities.FirstOrDefault();
        }

        private static void LinkExistingAtosUsuarios(Entity systemUser, Entity atosUsuarios, CrmServiceClient serviceClient)
        {
            var systemUserUpdate = new Entity("systemuser")
            {
                Id = systemUser.Id,
                ["atos_usuariosapid"] = new EntityReference("atos_usuarios", atosUsuarios.Id)
            };

            serviceClient.Update(systemUserUpdate);
            Console.WriteLine($"Sucefully updated user's SAP credentials");
        }

        private static void CreateAndLinkAtosUsuarios(UserNormalizationResult result, Entity systemUser, CrmServiceClient serviceClient)
        {
            var nameParts = result.Name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var firstName = nameParts[0];
            var lastName = string.Join(" ", nameParts.Skip(1));

            var atosUsuarios = new Entity("atos_usuarios")
            {
                ["atos_apellidos"] = lastName,
                ["atos_codigo"] = result.Username,
                ["atos_name"] = $"{result.Username}: {firstName} {lastName}",
                ["atos_nombre"] = firstName,
                ["atos_fechafin"] = new DateTime(9999, 12, 31)
            };

            try
            {
                var newRecordId = serviceClient.Create(atosUsuarios);
                Console.WriteLine($"Created new atos_usuarios record with ID: {newRecordId}");

                var systemUserUpdate = new Entity("systemuser")
                {
                    Id = systemUser.Id,
                    ["atos_usuariosapid"] = new EntityReference("atos_usuarios", newRecordId)
                };

                serviceClient.Update(systemUserUpdate);
                Console.WriteLine($"Sucefully updated user's SAP credentials");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating or linking atos_usuarios for {result.Username}: {ex.Message}");
                // If creation fails due to duplicate, try to retrieve and link the existing record
                var existingAtosUsuarios = RetrieveAtosUsuarios(result.Username, serviceClient);
                if (existingAtosUsuarios != null)
                {
                    LinkExistingAtosUsuarios(systemUser, existingAtosUsuarios, serviceClient);
                }
            }
        }
    }
}