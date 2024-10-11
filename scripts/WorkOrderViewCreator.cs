using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using RitmsHub.Scripts;
using System.Text.RegularExpressions;

namespace RitmsHub.scripts
{
    public class WorkOrderViewCreator
    {
        private readonly List<TransformedTeamData> teamDataList;
        private IOrganizationService service;

        public WorkOrderViewCreator(List<TransformedTeamData> teamDataList)
        {
            this.teamDataList = teamDataList;
        }

        private string ExtractBuCode(string input)
        {
            // Regular expression to match the pattern x-xx-xx-xx
            var regex = new Regex(@"^\d-[A-Z]{2}-[A-Z]{3}-\d{2}");
            var match = regex.Match(input);

            return match.Success ? match.Value : input;
        }

        public async Task CreateAndSaveWorkOrderViewAsync()
        {
            string fetchXml = BuildWorkOrderQuery();
            DynamicsCrmUtility.LogMessage("Generated FetchXML Query:");
            DynamicsCrmUtility.LogMessage(fetchXml);

            Console.Write("\nDo you want to save this query as a new personal Workorder view? (y/n): ");
            if (Console.ReadLine().Trim().ToLower() == "y")
            {
                Console.Write("Enter a name for the new personal view: ");
                string viewName = Console.ReadLine().Trim();

                await ConnectToCrmAsync();
                await SavePersonalViewAsync(fetchXml, viewName);
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
            }
            else
            {
                DynamicsCrmUtility.LogMessage("Personal view creation cancelled.");
            }
        }

        private string BuildWorkOrderQuery()
        {
            XDocument doc = new XDocument(
                new XElement("fetch",
                    new XElement("entity", new XAttribute("name", "msdyn_workorder"),
                        new XElement("attribute", new XAttribute("name", "statecode")),
                        new XElement("attribute", new XAttribute("name", "msdyn_name")),
                        new XElement("attribute", new XAttribute("name", "msdyn_serviceaccount")),
                        new XElement("attribute", new XAttribute("name", "atos_codigoordentrabajosap")),
                        new XElement("attribute", new XAttribute("name", "atos_titulo")),
                        new XElement("attribute", new XAttribute("name", "statuscode")),
                        new XElement("attribute", new XAttribute("name", "atos_puestotrabajoprincipalid")),
                        new XElement("attribute", new XAttribute("name", "atos_grupoplanificadorid")),
                        new XElement("attribute", new XAttribute("name", "atos_estadousuario")),
                        new XElement("attribute", new XAttribute("name", "atos_fechadecreacinensap")),
                        new XElement("attribute", new XAttribute("name", "createdon")),
                        new XElement("attribute", new XAttribute("name", "atos_inicextr")),
                        new XElement("attribute", new XAttribute("name", "atos_fechainicioprogramado")),
                        new XElement("attribute", new XAttribute("name", "atos_finextr")),
                        new XElement("attribute", new XAttribute("name", "atos_fechafinprogramado")),
                        new XElement("order", new XAttribute("attribute", "createdon"),
                        new XAttribute("descending", "true")),
                        new XElement("attribute", new XAttribute("name", "msdyn_workorderid")),
                        new XElement("filter", new XAttribute("type", "and"),
                            new XElement("condition", new XAttribute("attribute", "statuscode"),
                            new XAttribute("operator", "ne"), new XAttribute("value", "300000005"))
                        )
                    )
                )
            );

            var mainFilter = doc.Descendants("filter").First();

            // Create separate filters for owneridname, atos_grupoplanificadoridname, and atos_puestotrabajoprincipalidname
            var ownerFilter = new XElement("filter", new XAttribute("type", "or"));
            var plannerGroupFilter = new XElement("filter", new XAttribute("type", "or"));
            var contractorFilter = new XElement("filter", new XAttribute("type", "or"));

            // Add owneridname conditions
            var truncatedUniqueContractorCodes = teamDataList
                .Select(t => t.ContractorCode)
                .Select(code => code.Length >= 4 ? code.Substring(0, 4) : code)
                .Distinct();

            foreach (var code in truncatedUniqueContractorCodes)
            {
                ownerFilter.Add(new XElement("condition",
                    new XAttribute("attribute", "owneridname"),
                    new XAttribute("operator", "not-like"),
                    new XAttribute("value", $"%{code}%")));
            }

            // Add atos_grupoplanificadoridname conditions
            var plannerGroups = teamDataList.Select(t => t.PlannerGroup).Distinct();
            foreach (var group in plannerGroups)
            {
                plannerGroupFilter.Add(new XElement("condition",
                    new XAttribute("attribute", "atos_grupoplanificadoridname"),
                    new XAttribute("operator", "like"),
                    new XAttribute("value", $"%{group}%")));
            }

            // Add atos_puestotrabajoprincipalidname conditions
            var contractors = teamDataList.SelectMany(t => t.Contractor.Split(' ')).Distinct();
            foreach (var contractor in contractors)
            {
                contractorFilter.Add(new XElement("condition",
                    new XAttribute("attribute", "atos_puestotrabajoprincipalidname"),
                    new XAttribute("operator", "like"),
                    new XAttribute("value", $"%{contractor}%")));
            }

            mainFilter.Add(ownerFilter);
            mainFilter.Add(plannerGroupFilter);
            mainFilter.Add(contractorFilter);

            // Add msdyn_serviceaccountname conditions in a separate filter
            var buFilter = new XElement("filter", new XAttribute("type", "or"));
            foreach (var bu in teamDataList.Select(t => t.Bu).Distinct())
            {
                string extractedBU = ExtractBuCode(bu);

                buFilter.Add(new XElement("condition",
                    new XAttribute("attribute", "msdyn_serviceaccountname"),
                    new XAttribute("operator", "like"),
                    new XAttribute("value", $"%{extractedBU}%")));
            }
            mainFilter.Add(buFilter);

            return doc.ToString();
        }

        private async Task ConnectToCrmAsync()
        {
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                DynamicsCrmUtility.LogMessage($"Attempting to connect with: {connectionString}");

                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();
                this.service = serviceClient;

                //DynamicsCrmUtility.LogMessage($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"Failed to connect to Dynamics CRM: {ex.Message}", "ERROR");
                throw;
            }
        }

        private async Task SavePersonalViewAsync(string fetchXml, string viewName)
        {
            var userQuery = new Entity("userquery");
            userQuery["returnedtypecode"] = "msdyn_workorder";
            userQuery["name"] = viewName;
            userQuery["fetchxml"] = fetchXml;
            userQuery["layoutxml"] = CreateLayoutXml();
            userQuery["querytype"] = 0;

            try
            {
                Guid viewId = await Task.Run(() => service.Create(userQuery));
                DynamicsCrmUtility.LogMessage($"\n\nPersonal view '{viewName}' created successfully with ID: {viewId}");
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"Error creating personal view: {ex.Message}", "ERROR");
            }
        }

        private string CreateLayoutXml()
        {
            return @"
            <grid name='resultset' object='10010' jump='name' select='1' icon='1' preview='1'>
              <row name='result' id='msdyn_workorderid'>
                <cell name='msdyn_name' width='300' />
                <cell name='msdyn_serviceaccount' width='150' />
                <cell name='atos_codigoordentrabajosap' width='100' />
                <cell name='atos_titulo' width='200' />
                <cell name='statuscode' width='100' />
                <cell name='atos_puestotrabajoprincipalid' width='150' />
                <cell name='atos_grupoplanificadorid' width='150' />
                <cell name='atos_estadousuario' width='100' />
                <cell name='createdon' width='125' />
              </row>
            </grid>";
        }
    }
}