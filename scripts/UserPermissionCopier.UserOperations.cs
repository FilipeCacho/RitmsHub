using Microsoft.Xrm.Sdk;
using System;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserPermissionCopier
    {
        private void DisplayUserInfo(Entity sourceUser, Entity targetUser)
        {
            string GetUserDisplayName(Entity user)
            {
                string fullName = user.Contains("fullname") ? user["fullname"].ToString() : "N/A";
                string domainName = user.Contains("domainname") ? user["domainname"].ToString() : "N/A";
                string username = domainName != "N/A" ? domainName.Split('@')[0] : "N/A";
                return $"{fullName} (Username: {username})";
            }

            Console.WriteLine("Source User: " + GetUserDisplayName(sourceUser));
            Console.WriteLine("Target User: " + GetUserDisplayName(targetUser));
            Console.WriteLine();
        }

        private async Task<bool> IsUserInitializedAsync(Entity user)
        {
            int conditionsMet = 0;

            // Check Business Unit
            if (user.Contains("businessunitid") && ((EntityReference)user["businessunitid"]).Name == "edpr")
            {
                conditionsMet++;
            }

            // Check Roles
            var roles = await _permissionCopier.GetUserRolesAsync(user.Id);
            if (roles.Entities.Count == 0)
            {
                conditionsMet++;
            }

            // Check Teams
            var teams = await _permissionCopier.GetUserTeamsAsync(user.Id);
            if (teams.Entities.Count == 0)
            {
                conditionsMet++;
            }

            if (conditionsMet >= 2)
            {
                Console.WriteLine("This is not a valid user to take permissions from.");
                return false;
            }
            return true;
        }
    }
}