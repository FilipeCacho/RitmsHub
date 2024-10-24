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
                    Console.WriteLine(errorMessage, "ERROR");
                    Console.WriteLine($"Detailed Error: {cachedServiceClient.LastCrmError}", "ERROR");
                    Console.ResetColor();

                    throw new InvalidOperationException(errorMessage);
                }
            }
            return cachedServiceClient;
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