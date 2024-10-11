using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserNormalizer
    {
        private async Task EnsureUserHasRoles(Entity user, string[] rolesToEnsure)
        {
            var userBusinessUnitId = ((EntityReference)user["businessunitid"]).Id;

            var currentRoles = await RetrieveUserRolesAsync(user.Id, userBusinessUnitId);
            var currentRoleNames = currentRoles.Select(r => r.GetAttributeValue<string>("name")).ToHashSet();

            var roleQuery = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("businessunitid", ConditionOperator.Equal, userBusinessUnitId),
                        new ConditionExpression("name", ConditionOperator.In, rolesToEnsure)
                    }
                }
            };

            var availableRoles = await Task.Run(() => _service.RetrieveMultiple(roleQuery));

            foreach (var role in availableRoles.Entities)
            {
                string roleName = role.GetAttributeValue<string>("name");
                if (!currentRoleNames.Contains(roleName))
                {
                    try
                    {
                        var request = new AssociateRequest
                        {
                            Target = new EntityReference("systemuser", user.Id),
                            RelatedEntities = new EntityReferenceCollection
                            {
                                new EntityReference("role", role.Id)
                            },
                            Relationship = new Relationship("systemuserroles_association")
                        };

                        await Task.Run(() => _service.Execute(request));
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Added role: {roleName}");
                        Console.ResetColor();
                        
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error adding role {roleName}: {ex.Message}");
                        Console.ResetColor();
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Role already assigned: {roleName}");
                    Console.ResetColor();
                }
            }

            var missingRoles = rolesToEnsure.Except(availableRoles.Entities.Select(r => r.GetAttributeValue<string>("name")));
            foreach (var roleName in missingRoles)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Role not found in user's business unit: {roleName}");
                Console.ResetColor();
            }
        }

        private async Task EnsureUserHasTeams(Entity user, string[] teamsToEnsure)
        {
            var currentTeams = await _permissionCopier.GetUserTeamsAsync(user.Id);
            var currentTeamNames = currentTeams.Entities.Select(t => t.GetAttributeValue<string>("name")).ToList();

            var teamsToAdd = teamsToEnsure.Except(currentTeamNames).ToList();

            foreach (var teamName in teamsToAdd)
            {
                var teamQuery = new QueryExpression("team")
                {
                    ColumnSet = new ColumnSet("teamid"),
                    Criteria = new FilterExpression
                    {
                        Conditions = { new ConditionExpression("name", ConditionOperator.Equal, teamName) }
                    }
                };
                var teams = await Task.Run(() => _service.RetrieveMultiple(teamQuery));
                if (teams.Entities.Count > 0)
                {
                    var teamId = teams.Entities[0].Id;
                    var addMembersRequest = new AddMembersTeamRequest
                    {
                        TeamId = teamId,
                        MemberIds = new[] { user.Id }
                    };

                    await Task.Run(() => _service.Execute(addMembersRequest));
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Added team: {teamName}");
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Team not found: {teamName}");
                    Console.ResetColor();
                }
            }
        }


    }
}