using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class CrmLockManager
    {
        private readonly CrmServiceClient _crmServiceClient;

        public CrmLockManager(string connectionString)
        {
            _crmServiceClient = new CrmServiceClient(connectionString);
        }

        public async Task ReleaseLocks()
        {
            try
            {
                if (_crmServiceClient == null || !_crmServiceClient.IsReady)
                {
                    Console.WriteLine("CRM connection is not ready. Unable to release locks.");
                    return;
                }

                // Attempt to release any potential locks
                await Task.Run(() =>
                {
                    // Release locks on businessunit entity
                    ReleaseLockOnEntity("businessunit");

                    // Release locks on team entity
                    ReleaseLockOnEntity("team");

                });

                Console.WriteLine("Attempted to release all potential locks.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while attempting to release locks: {ex.Message}");
            }
            finally
            {
                // Ensure the CRM connection is closed
                _crmServiceClient?.Dispose();
            }
        }

        private void ReleaseLockOnEntity(string entityName)
        {
            try
            {
                // attempts to update a non-existent record, which should release any locks
                var query = $@"
                    UPDATE {entityName} 
                    SET modifiedon = GETDATE() 
                    WHERE 1 = 0
                ";

                _crmServiceClient.Execute(new OrganizationRequest("ExecuteNonQuery")
                {
                    ["Query"] = query
                });

                Console.WriteLine($"Attempted to release locks on {entityName} entity.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error releasing locks on {entityName}: {ex.Message}");
            }
        }
    }
}