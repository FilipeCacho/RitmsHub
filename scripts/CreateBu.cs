using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class CreateBu
    {
        public static async Task<List<BuCreationResult>> RunAsync(List<TransformedTeamData> transformedTeams, CancellationToken cancellationToken)
        {
            var results = new List<BuCreationResult>();
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                var buManager = new DynamicsCrmBusinessUnitManager(connectionString);

                foreach (var team in transformedTeams)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        var (exists, wasUpdated) = await buManager.CreateOrUpdateBusinessUnitAsync(team, cancellationToken);
                        results.Add(new BuCreationResult { BuName = team.Bu, Exists = exists, WasUpdated = wasUpdated });

                        Console.WriteLine($"Business Unit '{team.Bu}' {(exists ? (wasUpdated ? "updated" : "already exists") : "created")}.");
                    }
                    catch (OperationCanceledException)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Operation cancelled while processing Business Unit '{team.Bu}'.");
                        Console.ResetColor();
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error processing Business Unit '{team.Bu}': {ex.Message}");
                        Console.ResetColor();
                        results.Add(new BuCreationResult { BuName = team.Bu, Exists = false, WasUpdated = false });
                    }

                    await Task.Delay(500, cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Business Unit creation process was cancelled.");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error in Business Unit creation process: {ex.Message}");
                Console.ResetColor();
            }

            return results;
        }
    }

    public class DynamicsCrmBusinessUnitManager
    {
        private readonly CrmServiceClient _crmServiceClient;

        public DynamicsCrmBusinessUnitManager(string connectionString)
        {
            _crmServiceClient = DynamicsCrmUtility.CreateCrmServiceClient();
        }

        public async Task<(bool, bool)> CreateOrUpdateBusinessUnitAsync(TransformedTeamData team, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            DynamicsCrmUtility.LogMessage($"Starting business unit creation/update process for: '{team.Bu.Trim()}'");

            var existingBu = await GetExistingBusinessUnit(team.Bu.Trim(), cancellationToken);
            if (existingBu != null)
            {
                return await UpdateBusinessUnitIfNeeded(existingBu, team, cancellationToken);
            }
            else
            {
                var created = await CreateNewBusinessUnitAsync(team, cancellationToken);
                return (created, false);
            }
        }

        private async Task<Entity> GetExistingBusinessUnit(string businessUnitName, CancellationToken cancellationToken)
        {
            var query = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("name", ConditionOperator.Equal, businessUnitName)
                    }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query), cancellationToken);
            return result.Entities.Count > 0 ? result.Entities[0] : null;
        }

        private async Task<(bool, bool)> UpdateBusinessUnitIfNeeded(Entity existingBu, TransformedTeamData team, CancellationToken cancellationToken)
        {
            bool updated = false;
            var updateEntity = new Entity("businessunit") { Id = existingBu.Id };

            if (existingBu.GetAttributeValue<string>("atos_codigout") != team.Bu)
            {
                updateEntity["atos_codigout"] = team.Bu;
                updated = true;
                Console.WriteLine($"Updating atos_codigout for BU '{team.Bu.Trim()}' from '{existingBu.GetAttributeValue<string>("atos_codigout")}' to '{team.Bu}'");
            }

            var parentBuId = await DynamicsCrmUtility.GetBusinessUnitIdAsync(_crmServiceClient, team.PrimaryCompany, cancellationToken);
            if (existingBu.GetAttributeValue<EntityReference>("parentbusinessunitid").Id != parentBuId)
            {
                updateEntity["parentbusinessunitid"] = new EntityReference("businessunit", parentBuId);
                updated = true;
                Console.WriteLine($"Updating parentbusinessunitid for BU '{team.Bu.Trim()}' to '{parentBuId}'");
            }

            var plannerGroupId = await GetPlannerGroupIdAsync(team.PlannerGroup, team.PlannerCenterName, cancellationToken);
            if (plannerGroupId.HasValue && existingBu.GetAttributeValue<EntityReference>("atos_grupoplanificadorid")?.Id != plannerGroupId.Value)
            {
                updateEntity["atos_grupoplanificadorid"] = new EntityReference("atos_grupoplanificador", plannerGroupId.Value);
                updated = true;
                Console.WriteLine($"Updating atos_grupoplanificadorid for BU '{team.Bu.Trim()}' to '{plannerGroupId.Value}'");
            }

            var workCenterId = await GetWorkCenterIdAsync(team.ContractorCode, team.PlannerCenterName, cancellationToken);
            if (workCenterId.HasValue && existingBu.GetAttributeValue<EntityReference>("atos_puestodetrabajoid")?.Id != workCenterId.Value)
            {
                updateEntity["atos_puestodetrabajoid"] = new EntityReference("atos_puestodetrabajo", workCenterId.Value);
                updated = true;
                Console.WriteLine($"Updating atos_puestodetrabajoid for BU '{team.Bu.Trim()}' to '{workCenterId.Value}'");
            }

            if (updated)
            {
                await Task.Run(() => _crmServiceClient.Update(updateEntity), cancellationToken);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Business Unit '{team.Bu.Trim()}' updated successfully.");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Business Unit '{team.Bu.Trim()}' is up to date. No changes needed.");
                Console.ResetColor();
            }

            return (true, updated);
        }

        private async Task<bool> CreateNewBusinessUnitAsync(TransformedTeamData team, CancellationToken cancellationToken)
        {
            DynamicsCrmUtility.LogMessage($"Creating new business unit: '{team.Bu.Trim()}'");

            var parentBuId = await DynamicsCrmUtility.GetBusinessUnitIdAsync(_crmServiceClient, team.PrimaryCompany, cancellationToken);
            var plannerGroupId = await GetPlannerGroupIdAsync(team.PlannerGroup, team.PlannerCenterName, cancellationToken);
            var workCenterId = await GetWorkCenterIdAsync(team.ContractorCode, team.PlannerCenterName, cancellationToken);

            if (!plannerGroupId.HasValue || !workCenterId.HasValue)
            {
                throw new Exception("Creation halted due to missing Planner Group or Work Center data");
            }

            var buEntity = new Entity("businessunit")
            {
                ["name"] = team.Bu.Trim(),
                ["atos_codigout"] = team.Bu,
                ["parentbusinessunitid"] = new EntityReference("businessunit", parentBuId),
                ["atos_grupoplanificadorid"] = new EntityReference("atos_grupoplanificador", plannerGroupId.Value),
                ["atos_puestodetrabajoid"] = new EntityReference("atos_puestodetrabajo", workCenterId.Value)
            };

            var newBuId = await Task.Run(() => _crmServiceClient.Create(buEntity), cancellationToken);
            DynamicsCrmUtility.LogMessage($"New business unit created successfully. BU ID: {newBuId}", "SUCCESS");

            return true;
        }

        private async Task<Guid?> GetPlannerGroupIdAsync(string plannerGroupCode, string planningCenterName, CancellationToken cancellationToken)
        {
            var query = new QueryExpression("atos_grupoplanificador")
            {
                ColumnSet = new ColumnSet("atos_grupoplanificadorid", "atos_codigo", "atos_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("atos_codigo", ConditionOperator.Equal, plannerGroupCode)
                    }
                },
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "atos_grupoplanificador",
                        LinkToEntityName = "atos_centrodeplanificacion",
                        LinkFromAttributeName = "atos_centroplanificacionid",
                        LinkToAttributeName = "atos_centrodeplanificacionid",
                        JoinOperator = JoinOperator.Inner,
                        Columns = new ColumnSet("atos_name"),
                        EntityAlias = "pc",
                        LinkCriteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("atos_name", ConditionOperator.Like, $"%{planningCenterName}%")
                            }
                        }
                    }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query), cancellationToken);

            if (result.Entities.Count == 0)
            {
                DynamicsCrmUtility.LogMessage($"No Planner Group found for Code: {plannerGroupCode}, Planning Center: {planningCenterName}", "WARNING");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("BU creation halted. No matching Planner Group found.");
                Console.ResetColor();
                return null;
            }
            else if (result.Entities.Count > 1)
            {
                DynamicsCrmUtility.LogMessage($"Multiple Planner Groups found for Code: {plannerGroupCode}, Planning Center: {planningCenterName}", "WARNING");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("BU creation halted. Multiple matching Planner Groups found.");
                Console.ResetColor();
                return null;
            }

            return result.Entities[0].Id;
        }

        private async Task<Guid?> GetWorkCenterIdAsync(string contractorCode, string planningCenterName, CancellationToken cancellationToken)
        {
            var query = new QueryExpression("atos_puestodetrabajo")
            {
                ColumnSet = new ColumnSet("atos_puestodetrabajoid", "atos_codigopuestodetrabajo", "atos_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("atos_codigopuestodetrabajo", ConditionOperator.Equal, contractorCode)
                    }
                },
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "atos_puestodetrabajo",
                        LinkToEntityName = "atos_centrodeemplazamiento",
                        LinkFromAttributeName = "atos_centroemplazamientoid",
                        LinkToAttributeName = "atos_centrodeemplazamientoid",
                        JoinOperator = JoinOperator.Inner,
                        Columns = new ColumnSet("atos_name"),
                        EntityAlias = "ce",
                        LinkCriteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("atos_name", ConditionOperator.Like, $"%{planningCenterName}%")
                            }
                        }
                    }
                }
            };

            var result = await Task.Run(() => _crmServiceClient.RetrieveMultiple(query), cancellationToken);

            if (result.Entities.Count == 0)
            {
                DynamicsCrmUtility.LogMessage($"No Work Center found for Contractor Code: {contractorCode}, Planning Center: {planningCenterName}", "WARNING");

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("BU creation halted. No matching Work Center found.");
                Console.ResetColor();
                return null;
            }
            else if (result.Entities.Count > 1)
            {
                DynamicsCrmUtility.LogMessage($"Multiple Work Centers found for Contractor Code: {contractorCode}, Planning Center: {planningCenterName}", "WARNING");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("BU creation halted. Multiple matching Work Centers found.");
                Console.ResetColor();
                return null;
            }

            return result.Entities[0].Id;
        }
    }

    public class BuCreationResult
    {
        public string BuName { get; set; }
        public bool Exists { get; set; }
        public bool WasUpdated { get; set; }
    }
}