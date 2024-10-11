using Microsoft.Xrm.Sdk;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace RitmsHub.Scripts
{
    public partial class UserPermissionCopier
    {
        private async Task<bool> PromptForPermissionCopy(string permissionType, Entity sourceUser, Entity targetUser)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"\n{permissionType}:");
            Console.ResetColor();
            var sourceInfo = await GetPermissionInfo(permissionType, sourceUser);
            var targetInfo = await GetPermissionInfo(permissionType, targetUser);

            if (permissionType == "Business Unit")
            {
                Console.WriteLine($"\nSource user: {sourceInfo[0]}");
                Console.WriteLine($"Target user: {targetInfo[0]}");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\n(Replacing the current user BU deletes its current roles)");
                Console.ResetColor();
            }
            else
            {
                DisplayComparison(permissionType, sourceInfo, targetInfo);
            }

            Console.Write($"\nDo you want to copy the {permissionType.ToLower()} from the Source User to the Target User? (Y/N): ");
            return Console.ReadLine().Trim().ToUpper() == "Y";
        }

        private async Task<List<string>> GetPermissionInfo(string permissionType, Entity user)
        {
            switch (permissionType)
            {
                case "Business Unit":
                    return new List<string> { user.Contains("businessunitid") ? ((EntityReference)user["businessunitid"]).Name : "N/A" };
                case "Teams":
                    var teams = await _permissionCopier.GetUserTeamsAsync(user.Id);
                    return teams.Entities.Select(t => t.GetAttributeValue<string>("name")).ToList();
                case "Roles":
                    var roles = await _permissionCopier.GetUserRolesAsync(user.Id);
                    return roles.Entities.Select(r => r.GetAttributeValue<string>("name")).ToList();
                default:
                    return new List<string> { "Unknown" };
            }
        }

        private void DisplayComparison(string title, List<string> sourceItems, List<string> targetItems)
        {
            Console.WriteLine($"\n{title}:");
            Console.WriteLine("Source User".PadRight(50) + "Target User");
            Console.WriteLine(new string('-', 100));

            var commonItems = sourceItems.Intersect(targetItems).ToHashSet();

            int maxCount = Math.Max(sourceItems.Count, targetItems.Count);
            for (int i = 0; i < maxCount; i++)
            {
                string sourceItem = i < sourceItems.Count ? sourceItems[i] : "";
                string targetItem = i < targetItems.Count ? targetItems[i] : "";

                if (commonItems.Contains(sourceItem))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write(sourceItem.PadRight(50));
                    Console.ResetColor();
                }
                else
                {
                    Console.Write(sourceItem.PadRight(50));
                }

                if (commonItems.Contains(targetItem))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(targetItem);
                    Console.ResetColor();
                }
                else
                {
                    Console.WriteLine(targetItem);
                }
            }
        }
    }
}