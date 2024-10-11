//uses new login methods

using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using RitmsHub.Scripts;

namespace RitmsHub.Scripts
{
    public class WorkflowLister
    {
        private readonly CrmServiceClient _crmServiceClient;
        private readonly string _outputFileName;

        public WorkflowLister()
        {
            string downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            _outputFileName = Path.Combine(downloadsFolder, "workflow_diagnosis.txt");

            //Extract dyn env and login info from the excel
            var (url, username, password, appId, redirectUri, authType, loginPrompt, requireNewInstance) = ExcelReader.ReadLoginValues();

            // Construct the connection string
            string connectionString = $"AuthType={authType};Url={url};Username={username};Password={password};AppId={appId};RedirectUri={redirectUri};LoginPrompt={loginPrompt};RequireNewInstance={requireNewInstance};TokenCacheStorePath=c:\\MyTokenCache;Prompt=login";

            _crmServiceClient = new CrmServiceClient(connectionString);
            if (!_crmServiceClient.IsReady)
            {
                throw new Exception("Failed to connect to Dynamics CRM");
            }
        }

        public async Task ListUserWorkflowsAsync(string username)
        {
            using (StreamWriter writer = new StreamWriter(_outputFileName, true))
            {
                try
                {
                    await writer.WriteLineAsync($"\nWorkflows for user: {username}");
                    await writer.WriteLineAsync("----------------------------------------");

                    var userId = await GetUserIdAsync(username);
                    if (userId == Guid.Empty)
                    {
                        await writer.WriteLineAsync($"User not found: {username}");
                        return;
                    }

                    var query = new QueryExpression("workflow")
                    {
                        ColumnSet = new ColumnSet("name", "category", "type", "primaryentity", "statecode"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("statecode", ConditionOperator.Equal, 1), // Active workflows only
                            }
                        }
                    };

                    var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query));

                    await writer.WriteLineAsync($"Total workflows found: {result.Entities.Count}");
                    await writer.WriteLineAsync();

                    foreach (var entity in result.Entities)
                    {
                        await writer.WriteLineAsync($"Name: {entity.GetAttributeValue<string>("name")}");
                        await writer.WriteLineAsync($"Category: {entity.GetAttributeValue<OptionSetValue>("category")?.Value}");
                        await writer.WriteLineAsync($"Type: {entity.GetAttributeValue<OptionSetValue>("type")?.Value}");
                        await writer.WriteLineAsync($"Primary Entity: {entity.GetAttributeValue<string>("primaryentity")}");
                        await writer.WriteLineAsync();
                    }
                }
                catch (Exception ex)
                {
                    await writer.WriteLineAsync($"An error occurred: {ex.Message}");
                    await writer.WriteLineAsync($"Stack Trace: {ex.StackTrace}");
                }
            }
            Console.WriteLine($"Workflow information has been written to: {_outputFileName}");
        }

        public async Task DiagnoseWorkflowRetrievalAsync()
        {
            using (StreamWriter writer = new StreamWriter(_outputFileName, true))
            {
                try
                {
                    await writer.WriteLineAsync("\nWorkflow Diagnosis Report");
                    await writer.WriteLineAsync("========================");

                    await writer.WriteLineAsync($"CRM Connection Status: {(_crmServiceClient.IsReady ? "Ready" : "Not Ready")}");

                    var query = new QueryExpression("workflow")
                    {
                        ColumnSet = new ColumnSet(true)
                    };

                    var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query));

                    await writer.WriteLineAsync($"Total workflows found (including inactive): {result.Entities.Count}");
                    await writer.WriteLineAsync();

                    foreach (var entity in result.Entities)
                    {
                        await writer.WriteLineAsync($"Name: {entity.GetAttributeValue<string>("name")}");
                        await writer.WriteLineAsync($"State: {entity.GetAttributeValue<OptionSetValue>("statecode")?.Value}");
                        await writer.WriteLineAsync($"Category: {entity.GetAttributeValue<OptionSetValue>("category")?.Value}");
                        await writer.WriteLineAsync($"Type: {entity.GetAttributeValue<OptionSetValue>("type")?.Value}");
                        await writer.WriteLineAsync($"Primary Entity: {entity.GetAttributeValue<string>("primaryentity")}");
                        await writer.WriteLineAsync($"Owner: {entity.GetAttributeValue<EntityReference>("ownerid")?.Name}");
                        await writer.WriteLineAsync();
                    }

                    var currentUserId = _crmServiceClient.GetMyCrmUserId();
                    await writer.WriteLineAsync($"Current User ID: {currentUserId}");

                    var userRolesQuery = new QueryExpression("systemuserroles")
                    {
                        ColumnSet = new ColumnSet("roleid"),
                        Criteria = new FilterExpression
                        {
                            Conditions = { new ConditionExpression("systemuserid", ConditionOperator.Equal, currentUserId) }
                        }
                    };

                    var userRolesResult = await Task.Run(() => _crmServiceClient.RetrieveMultiple(userRolesQuery));
                    await writer.WriteLineAsync($"User has {userRolesResult.Entities.Count} roles assigned.");

                    foreach (var userRole in userRolesResult.Entities)
                    {
                        var roleId = userRole.GetAttributeValue<EntityReference>("roleid").Id;
                        await writer.WriteLineAsync($"Role ID: {roleId}");

                        var privilegesQuery = new QueryExpression("roleprivileges")
                        {
                            ColumnSet = new ColumnSet("privilegedepthmask", "privilegeid"),
                            Criteria = new FilterExpression
                            {
                                Conditions = { new ConditionExpression("roleid", ConditionOperator.Equal, roleId) }
                            }
                        };

                        var privilegesResult = await Task.Run(() => _crmServiceClient.RetrieveMultiple(privilegesQuery));
                        await writer.WriteLineAsync($"  This role has {privilegesResult.Entities.Count} privileges.");
                    }
                }
                catch (Exception ex)
                {
                    await writer.WriteLineAsync($"An error occurred during diagnosis: {ex.Message}");
                    await writer.WriteLineAsync($"Stack Trace: {ex.StackTrace}");
                }
            }
            Console.WriteLine($"Diagnosis information has been written to: {_outputFileName}");
        }

        private async Task<Guid> GetUserIdAsync(string username)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("domainname", ConditionOperator.Equal, username) }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query));
            if (result.Entities.Count > 0)
            {
                return result.Entities[0].Id;
            }
            return Guid.Empty;
        }
    }
}