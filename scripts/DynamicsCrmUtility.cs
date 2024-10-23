using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Threading;

namespace RitmsHub.Scripts
{
    public static class DynamicsCrmUtility
    {
        private static string cachedConnectionString;
        private static CrmServiceClient cachedServiceClient;

        public static void ResetConnection()
        {
            cachedConnectionString = null;
            cachedServiceClient?.Dispose();
            cachedServiceClient = null;
        }

        public static string CreateConnectionString()
        {
            var (url, username, password, appId, redirectUri, authType, loginPrompt, requireNewInstance) = ExcelReader.ReadLoginValues();
            return $"AuthType={authType};Url={url};Username={username};Password={password};AppId={appId};RedirectUri={redirectUri};LoginPrompt={loginPrompt};RequireNewInstance={requireNewInstance};TokenCacheStorePath=c:\\MyTokenCache;Prompt=login";
        }

        public static CrmServiceClient CreateCrmServiceClient()
        {
            if (cachedServiceClient == null || !cachedServiceClient.IsReady)
            {
                string connectionString = CreateConnectionString();
                cachedServiceClient = new CrmServiceClient(connectionString);

                // Check for specific connection errors
                if (!cachedServiceClient.IsReady)
                {
                    string errorMessage;
                    if (cachedServiceClient.LastCrmError.Contains("401") ||
                        cachedServiceClient.LastCrmError.Contains("authentication failed"))
                    {
                        errorMessage = "Authentication failed. Please check your username and password.";
                    }
                    else if (cachedServiceClient.LastCrmError.Contains("Unable to Login to Dynamics CRM"))
                    {
                        errorMessage = "Unable to connect to Dynamics CRM. Please check your connection details and try again.";
                    }
                    else
                    {
                        errorMessage = "An unexpected error occurred while connecting to Dynamics CRM. Please try again or contact support.";
                    }

                    // Display error message using LogMessage
                    Console.ForegroundColor = ConsoleColor.Red;
                    LogMessage(errorMessage, "ERROR");
                    LogMessage($"Detailed Error: {cachedServiceClient.LastCrmError}", "ERROR");

                    throw new InvalidOperationException(errorMessage);
                }
            }
            return cachedServiceClient;
        }

        public static void LogMessage(string message, string type = "INFO")
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Console.Write($"{timestamp} [{type}] ");

            var highlights = new Dictionary<string, ConsoleColor>
            {
                { "Starting business unit creation/update", ConsoleColor.Cyan },
                { "team creation/update process for team", ConsoleColor.Cyan },
                { "New team created successfully", ConsoleColor.Green },
                { "Error", ConsoleColor.Red },
                { "Role not found\r\n", ConsoleColor.Red },
                { "Creating new business unit", ConsoleColor.Blue },
                { "New business unit created successfully", ConsoleColor.Green },
                { "No Planner Group found for Code", ConsoleColor.Red },
                { "Multiple Planner Groups found for Code", ConsoleColor.Red },
                { "No Work Center found for Contractor Code", ConsoleColor.Red },
                { "Multiple Work Centers found for Contractor Code", ConsoleColor.Red },
                { "assigned to user", ConsoleColor.Green },
                { "(\"Source user does not have a Business Unit assigned.", ConsoleColor.Yellow },
                { "process for team", ConsoleColor.DarkGreen }
            };

            int position = 0;
            while (position < message.Length)
            {
                bool highlighted = false;
                foreach (var highlight in highlights)
                {
                    if (message.IndexOf(highlight.Key, position, StringComparison.OrdinalIgnoreCase) == position)
                    {
                        Console.ForegroundColor = highlight.Value;
                        Console.Write(message.Substring(position, highlight.Key.Length));
                        Console.ResetColor();
                        position += highlight.Key.Length;
                        highlighted = true;
                        break;
                    }
                }

                if (!highlighted)
                {
                    Console.Write(message[position]);
                    position++;
                }
            }

            Console.WriteLine(); // New line at the end of the message
        }

        public static async Task<Guid> GetBusinessUnitIdAsync(IOrganizationService service, string businessUnitName, CancellationToken cancellationToken)
        {
            var query = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet("businessunitid"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("name", ConditionOperator.Equal, businessUnitName) }
                }
            };

            var result = await Task.Run(() => service.RetrieveMultiple(query), cancellationToken);

            if (result.Entities.Count == 0)
            {
                throw new Exception($"Business Unit not found: {businessUnitName}");
            }

            return result.Entities[0].Id;
        }
    }
}