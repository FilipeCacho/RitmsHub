using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserNormalizer
    {
        private IOrganizationService _service;
        private UserRetriever _userRetriever;
        private PermissionCopier _permissionCopier;

        public async Task<List<UserNormalizationResult>> Run()
        {
            var results = new List<UserNormalizationResult>();
            try
            {
                await ConnectToCrmAsync();

                Console.Clear();

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("This code will normalize a user's permissions based on their region (EU or NA)");
                Console.ResetColor();

                var user = await _userRetriever.PromptAndRetrieveUser("\nEnter the user's name or username (or 0 to exit): ");

                if (user == UserRetriever.Exit || user == null)
                {
                    return results; // Return empty list if no user is selected
                }

                await DisplayUserInfo(user);

                if (ConfirmUserNormalization(user))
                {
                    string regionChoice = GetRegionChoice();

                    if (regionChoice == "0")
                    {
                        return results; // Return empty list if user chooses to go back
                    }

                    await NormalizeUser(user, regionChoice);
                    await GiveResco(user, regionChoice);

                    string username = user.Contains("domainname") ? user["domainname"].ToString().Split('@')[0] : "";
                    bool isInternal = IsInternalUser(username);

                    results.Add(new UserNormalizationResult
                    {
                        Username = username,
                        IsInternal = isInternal,
                        Name = GetUserName(user, isInternal)
                    });
                }
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"An error occurred: {ex.Message}", "ERROR");
                DynamicsCrmUtility.LogMessage($"Stack Trace: {ex.StackTrace}", "ERROR");
            }

            return results;
        }

        private async Task ConnectToCrmAsync()
        {
            try
            {
                ConfigureSecurityProtocol();
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                //DynamicsCrmUtility.LogMessage($"Attempting to connect with: {connectionString}");

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

        private async Task DisplayUserInfo(Entity user)
        {
            string username = user.Contains("domainname") ? user["domainname"].ToString().Split('@')[0] : "N/A";
            string fullName = user.Contains("fullname") ? user["fullname"].ToString() : "N/A";
            string businessUnit = user.Contains("businessunitid") ? ((EntityReference)user["businessunitid"]).Name : "N/A";
            string regionName = await GetUserRegionName(user.Id);

            Console.Clear();
            Console.WriteLine($"User: {fullName}");
            Console.WriteLine($"ID: {username}");
            Console.WriteLine($"Business Unit: {businessUnit}");
            Console.WriteLine($"Region Name: {regionName}");
        }

        private async Task<string> GetUserRegionName(Guid userId)
        {
            try
            {
                var query = new QueryExpression("systemuser")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                        }
                    }
                };

                var result = await Task.Run(() => _service.RetrieveMultiple(query));

                if (result.Entities.Count > 0)
                {
                    var user = result.Entities[0];

                    if (user.Contains("atos_regionname"))
                    {
                        return user["atos_regionname"].ToString();
                    }

                    if (user.Contains("atos_region"))
                    {
                        var optionSetValue = (OptionSetValue)user["atos_region"];
                        return GetOptionSetLabel("systemuser", "atos_region", optionSetValue.Value);
                    }

                    return "Region not set";
                }
                else
                {
                    return "User not found";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving region name: {ex.Message}");
                return "Error";
            }
        }

        private string GetOptionSetLabel(string entityName, string attributeName, int value)
        {
            var request = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };

            var response = (RetrieveAttributeResponse)_service.Execute(request);
            var metadata = (EnumAttributeMetadata)response.AttributeMetadata;

            var option = metadata.OptionSet.Options.FirstOrDefault(o => o.Value == value);
            return option?.Label.UserLocalizedLabel.Label ?? $"Unknown ({value})";
        }

        private bool ConfirmUserNormalization(Entity user)
        {
            if (user.Contains("businessunitid") && ((EntityReference)user["businessunitid"]).Name.Equals("edpr", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("\nWarning: The user's business unit is set to 'edpr'.");
                Console.WriteLine("This means this user with this BU and if it has roles can see all parks");
                Console.WriteLine("Are you sure you want to continue? (y/n)");

                return Console.ReadLine().Trim().ToLower() == "y";
            }
            return true;
        }

        private string GetRegionChoice()
        {
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nSelect region for normalization:");
                Console.ResetColor();
                Console.WriteLine("1. EU");
                Console.WriteLine("2. NA");
                Console.WriteLine("0. Go back to the main menu");
                Console.Write("\nEnter your region choice (1, 2, or 0): ");
                string choice = Console.ReadLine();

                if (choice == "1" || choice == "2" || choice == "0")
                {
                    return choice;
                }

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nInvalid choice. Please try again.");
                Console.ResetColor();
            }
        }

        

        

        private bool IsInternalUser(string username)
        {
            return username.StartsWith("e", StringComparison.OrdinalIgnoreCase) &&
                   username.Length > 1 &&
                   char.IsDigit(username[1]) &&
                   !username.Substring(1, 1).Equals("x", StringComparison.OrdinalIgnoreCase);
        }

        private string GetUserName(Entity user, bool isInternal)
        {
            return isInternal
                ? (user.Contains("fullname") ? user["fullname"].ToString() : "")
                : (user.Contains("firstname") ? user["firstname"].ToString() : "");
        }

        

        private async Task<List<Entity>> RetrieveUserRolesAsync(Guid userId, Guid businessUnitId)
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("name", "roleid"),
                Criteria = new FilterExpression()
            };

            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, businessUnitId);

            var linkEntity = new LinkEntity("role", "systemuserroles", "roleid", "roleid", JoinOperator.Inner);
            linkEntity.LinkCriteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
            query.LinkEntities.Add(linkEntity);

            EntityCollection results = await Task.Run(() => _service.RetrieveMultiple(query));

            return results.Entities.ToList();
        }

        

        private async Task UpdateUserRegion(Entity user, string[] region)
        {
            Entity userUpdate = new Entity("systemuser");
            userUpdate.Id = user.Id;
            userUpdate["atos_regionname"] = region[0];
            userUpdate["atos_region"] = new OptionSetValue(int.Parse(region[1]));
            await Task.Run(() => _service.Update(userUpdate));

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\nUpdated user region to: {region[0]}");
            Console.ResetColor();
        }

        

        
    }
}

public class UserNormalizationResult
{
    public string Username { get; set; }
    public string Name { get; set; }
    public bool IsInternal { get; set; }
}