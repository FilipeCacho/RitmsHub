using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public class UserBasicInfoRetriever
    {
        private IOrganizationService _service;
        private UserRetriever _userRetriever;
        private PermissionCopier _permissionCopier;

        public async Task RetrieveAndCompareUserInfoAsync()
        {
            try
            {
                await ConnectToCrmAsync();

                while (true)
                {
                    var user1 = await PromptAndRetrieveUserAsync("Enter the first user's name or username (or 0 to exit): ");
                    if (user1 == null) return;

                    var user1Info = await GetUserInfoAsync(user1);
                    Console.Clear();
                    DisplayUserInfo(user1Info, "User 1");

                    while (true)
                    {
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.Write("\nDo you want to compare with another user? (If yes, roles and or teams both have will be highlighted) (y/n):  \n");

                        Console.ResetColor();
                        var response = Console.ReadLine().ToLower();

                        if (response == "n") return;
                        if (response == "y") break;
                        Console.WriteLine("Invalid input. Please enter 'y' or 'n'.");
                    }

                    var user2 = await PromptAndRetrieveUserAsync("\nEnter the second user's name or username (or 0 to exit): ");
                    if (user2 == null) return;

                    var user2Info = await GetUserInfoAsync(user2);
                    DisplayUsersInfoSideBySide(user1Info, user2Info);

                    Console.WriteLine("\nPress any key to return to the main menu...");
                    Console.ReadKey();
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine("Stack Trace: " + ex.StackTrace);
            }
        }

        private async Task ConnectToCrmAsync()
        {
            string connectionString = DynamicsCrmUtility.CreateConnectionString();
            var serviceClient = DynamicsCrmUtility.CreateCrmServiceClient();

            if (!serviceClient.IsReady)
            {
                throw new Exception(string.Format("Failed to connect. Error: {0}", serviceClient.LastCrmError));
            }

            _service = serviceClient;
            _userRetriever = new UserRetriever(_service);
            _permissionCopier = new PermissionCopier(_service);

            //Console.WriteLine(string.Format("Connected successfully to {0}", serviceClient.ConnectedOrgUniqueName));
        }

        private async Task<Entity> PromptAndRetrieveUserAsync(string prompt)
        {
            while (true)
            {
                var user = await _userRetriever.PromptAndRetrieveUser(prompt);
                if (user == UserRetriever.Exit || user == null) return null;
                return user;
            }
        }

        private async Task<UserInfo> GetUserInfoAsync(Entity user)
        {
            var businessUnit = user.GetAttributeValue<EntityReference>("businessunitid").Name;
            var roles = await _permissionCopier.GetUserRolesAsync(user.Id);
            var teams = await _permissionCopier.GetUserTeamsAsync(user.Id);

            return new UserInfo
            {
                FullName = user.GetAttributeValue<string>("fullname"),
                Username = user.GetAttributeValue<string>("domainname").Split('@')[0],
                BusinessUnit = businessUnit,
                Roles = roles.Entities.Select(r => r.GetAttributeValue<string>("name")).OrderBy(r => r, new AlphanumericComparer()).ToList(),
                Teams = teams.Entities.Select(t => t.GetAttributeValue<string>("name")).OrderBy(t => t, new AlphanumericComparer()).ToList()
            };
        }

        private void DisplayUserInfo(UserInfo userInfo, string userLabel)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine(userLabel + " Information:");
            Console.ResetColor();
            Console.WriteLine("Name: " + userInfo.FullName);
            Console.WriteLine("Username: " + userInfo.Username);
            Console.WriteLine("Business Unit: " + userInfo.BusinessUnit);
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("\nRoles:");
            Console.ResetColor();
            foreach (var role in userInfo.Roles)
            {
                Console.WriteLine("  - " + role);
            }
            Console.WriteLine("\nTeams:");
            foreach (var team in userInfo.Teams)
            {
                Console.WriteLine("  - " + team);
            }
        }

        private void DisplayUsersInfoSideBySide(UserInfo user1, UserInfo user2)
        {
            Console.Clear();
            int maxLength = Math.Max(
                user1.Teams.Concat(user1.Roles).Max(s => s.Length),
                Math.Max(user1.FullName.Length, Math.Max(user1.Username.Length, user1.BusinessUnit.Length))
            );
            int padding = maxLength + 3;

            Console.WriteLine(string.Format("{0,-" + padding + "}{1,-" + padding + "}", "User 1", "User 2"));
            Console.WriteLine(new string('-', padding * 2));
            Console.WriteLine(string.Format("{0,-" + padding + "}{1,-" + padding + "}", user1.FullName, user2.FullName));
            Console.WriteLine(string.Format("{0,-" + padding + "}{1,-" + padding + "}", user1.Username, user2.Username));
            Console.WriteLine(string.Format("{0,-" + padding + "}{1,-" + padding + "}", user1.BusinessUnit, user2.BusinessUnit));

            Console.WriteLine(string.Format("\n{0,-" + padding + "}{1,-" + padding + "}", "Roles:", "Roles:"));
            DisplayLists(user1.Roles, user2.Roles, padding);

            Console.WriteLine(string.Format("\n{0,-" + padding + "}{1,-" + padding + "}", "Teams:", "Teams:"));
            DisplayLists(user1.Teams, user2.Teams, padding);
        }

        private void DisplayLists(List<string> list1, List<string> list2, int padding)
        {
            var commonItems = list1.Intersect(list2).ToHashSet();

            for (int i = 0; i < Math.Max(list1.Count, list2.Count); i++)
            {
                string item1 = i < list1.Count ? list1[i] : "";
                string item2 = i < list2.Count ? list2[i] : "";

                if (commonItems.Contains(item1))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write(string.Format("{0,-" + padding + "}", item1));
                    Console.ResetColor();
                }
                else
                {
                    Console.Write(string.Format("{0,-" + padding + "}", item1));
                }

                if (commonItems.Contains(item2))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(string.Format("{0,-" + padding + "}", item2));
                    Console.ResetColor();
                }
                else
                {
                    Console.WriteLine(string.Format("{0,-" + padding + "}", item2));
                }
            }
        }

        public class UserInfo
        {
            public string FullName { get; set; }
            public string Username { get; set; }
            public string BusinessUnit { get; set; }
            public List<string> Roles { get; set; }
            public List<string> Teams { get; set; }
        }

        public class AlphanumericComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (x == null) return y == null ? 0 : -1;
                if (y == null) return 1;

                int len1 = x.Length;
                int len2 = y.Length;
                int marker1 = 0;
                int marker2 = 0;

                while (marker1 < len1 && marker2 < len2)
                {
                    char ch1 = x[marker1];
                    char ch2 = y[marker2];

                    var text1 = new System.Text.StringBuilder();
                    var text2 = new System.Text.StringBuilder();

                    while (marker1 < len1 && !char.IsDigit(x[marker1]))
                        text1.Append(x[marker1++]);

                    while (marker2 < len2 && !char.IsDigit(y[marker2]))
                        text2.Append(y[marker2++]);

                    int result = string.Compare(text1.ToString(), text2.ToString(), StringComparison.OrdinalIgnoreCase);

                    if (result != 0)
                        return result;

                    text1.Clear();
                    text2.Clear();

                    while (marker1 < len1 && char.IsDigit(x[marker1]))
                        text1.Append(x[marker1++]);

                    while (marker2 < len2 && char.IsDigit(y[marker2]))
                        text2.Append(y[marker2++]);

                    if (text1.Length == 0 && text2.Length > 0) return -1;
                    if (text1.Length > 0 && text2.Length == 0) return 1;
                    if (text1.Length > 0 && text2.Length > 0)
                    {
                        int num1, num2;
                        if (int.TryParse(text1.ToString(), out num1) && int.TryParse(text2.ToString(), out num2))
                        {
                            if (num1 != num2)
                                return num1.CompareTo(num2);
                        }
                    }
                }

                return len1 - len2;
            }
        }
    }
}