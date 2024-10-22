using Microsoft.Xrm.Sdk;
using RitmsHub.Scripts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RitmsHub.scripts
{
    public class NotificationsViewCreator
    {
        private readonly List<TransformedTeamData> teamDataList;
        private IOrganizationService service;

        public NotificationsViewCreator(List<TransformedTeamData> teamDataList)
        {
            this.teamDataList = teamDataList;
        }

        private string ExtractBuCode(string input)
        {
            var regex = new Regex(@"^\d-[A-Z]{2}-[A-Z]{3}-\d{2}");
            var match = regex.Match(input);
            return match.Success ? match.Value : input;
        }

        public async Task CreateAndSaveNotificationsViewAsync()
        {
            string fetchXml = BuildNotificationsQuery();

            Console.Clear();
            
            DynamicsCrmUtility.LogMessage("Generated FetchXML Query for Notifications:");
            DynamicsCrmUtility.LogMessage(fetchXml);

            Console.Write("\nDo you want to save this query as a new personal Notifications view? (y/n): ");
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

        private string BuildNotificationsQuery()
        {
            XDocument doc = new XDocument(
                new XElement("fetch",
                    new XElement("entity",
                        new XAttribute("name", "atos_aviso"),
                        CreateAttributeElements(),
                        CreateOrderElements(),
                        new XElement("filter",
                            new XAttribute("type", "and"),
                            CreateConditionElement("atos_indicadorborrado", "ne", "1"),
                            CreateOwnerFilter(),
                            CreatePlannerGroupFilter(),
                            CreateContractorFilter(),
                            CreateBuFilter()
                        )
                    )
                )
            );

            return doc.ToString(SaveOptions.None);
        }

        private IEnumerable<XElement> CreateAttributeElements()
        {
            string[] attributes = new string[]
            {
                "statecode", "atos_name", "atos_ubicaciontecnicaid", "statuscode",
                "atos_fechainicioaveria", "atos_fechanotificacion", "atos_prioridadid",
                "atos_esprincipal", "atos_fechafinaveria", "atos_ordendetrabajoid",
                "atos_equipoid", "atos_clasedeavisoid", "atos_esplantilla",
                "atos_codigosap", "atos_descripcioncorta", "atos_avisoid"
            };

            return attributes.Select(attr => new XElement("attribute", new XAttribute("name", attr)));
        }

        private IEnumerable<XElement> CreateOrderElements()
        {
            yield return new XElement("order",
                new XAttribute("attribute", "atos_name"),
                new XAttribute("descending", "true"));
            yield return new XElement("order",
                new XAttribute("attribute", "atos_clasedeavisoid"),
                new XAttribute("descending", "false"));
        }

        private XElement CreateConditionElement(string attribute, string @operator, string value)
        {
            return new XElement("condition",
                new XAttribute("attribute", attribute),
                new XAttribute("operator", @operator),
                new XAttribute("value", value));
        }

        private XElement CreateOwnerFilter()
        {
            var contractorCodes = teamDataList
                .Select(t => t.ContractorCode)
                .Select(code => code.Length >= 4 ? code.Substring(0, 4) : code)
                .Distinct();

            return new XElement("filter",
                new XAttribute("type", "or"),
                contractorCodes.Select(code =>
                    CreateConditionElement("owneridname", "not-like", $"%{code}%")
                )
            );
        }

        private XElement CreatePlannerGroupFilter()
        {
            var plannerGroups = teamDataList.Select(t => t.PlannerGroup).Distinct();
            return new XElement("filter",
                new XAttribute("type", "or"),
                plannerGroups.Select(group =>
                    CreateConditionElement("atos_grupoplanificadoridname", "like", $"%{group}%")
                )
            );
        }

        private XElement CreateContractorFilter()
        {
            var contractors = teamDataList.SelectMany(t => t.Contractor.Split(' ')).Distinct();
            return new XElement("filter",
                new XAttribute("type", "or"),
                contractors.Select(contractor =>
                    CreateConditionElement("atos_puestotrabajoprincipalidname", "like", $"%{contractor}%")
                )
            );
        }

        private XElement CreateBuFilter()
        {
            return new XElement("filter",
                new XAttribute("type", "or"),
                teamDataList.Select(t => t.Bu)
                    .Distinct()
                    .Select(bu => CreateConditionElement("atos_ubicaciontecnicaidname", "like", $"%{ExtractBuCode(bu)}%"))
            );
        }

        private Task ConnectToCrmAsync()
        {
            try
            {
                string connectionString = DynamicsCrmUtility.CreateConnectionString();
                //DynamicsCrmUtility.LogMessage($"Attempting to connect with: {connectionString}");

                var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();
                service = serviceClient;

                //DynamicsCrmUtility.LogMessage($"Connected successfully to {serviceClient.ConnectedOrgUniqueName}");
            }
            catch (Exception ex)
            {
                DynamicsCrmUtility.LogMessage($"Failed to connect to Dynamics CRM: {ex.Message}", "ERROR");
                throw;
            }

            return Task.CompletedTask;
        }

        private async Task SavePersonalViewAsync(string fetchXml, string viewName)
        {
            var userQuery = new Entity("userquery");
            userQuery["returnedtypecode"] = "atos_aviso";
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
            <grid name='resultset' object='10010' jump='atos_name' select='1' icon='1' preview='1'>
              <row name='result' id='atos_avisoid'>
                <cell name='atos_name' width='300' />
                <cell name='atos_ubicaciontecnicaid' width='150' />
                <cell name='statuscode' width='100' />
                <cell name='atos_fechainicioaveria' width='125' />
                <cell name='atos_fechanotificacion' width='125' />
                <cell name='atos_prioridadid' width='100' />
                <cell name='atos_esprincipal' width='100' />
                <cell name='atos_ordendetrabajoid' width='150' />
                <cell name='atos_equipoid' width='150' />
                <cell name='atos_clasedeavisoid' width='150' />
              </row>
            </grid>";
        }
    }
}