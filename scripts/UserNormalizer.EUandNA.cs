using Microsoft.Xrm.Sdk;
using System;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserNormalizer
    {
        private async Task NormalizeUser(Entity user, string regionChoice)
        {
            switch (regionChoice)
            {
                case "1":
                    await NormalizeEUUser(user);
                    break;
                case "2":
                    await NormalizeNAUser(user);
                    break;
            }
        }

        private async Task NormalizeEUUser(Entity user)
        {
            string username = user.Contains("domainname") ? user["domainname"].ToString().Split('@')[0] : "";
            bool isInternal = IsInternalUser(username);

            string[] rolesToAdd = isInternal ? CodesAndRoles.EUDefaultRolesForInternalUsers : CodesAndRoles.EUDefaultRolesForExternalUsers;
            string[] teamsToAdd = isInternal ? CodesAndRoles.EUDefaultTeamsForInteralUsers : CodesAndRoles.EUDefaultTeamsForExternalUsers;

            await EnsureUserHasRoles(user, rolesToAdd);

            await EnsureUserHasTeams(user, teamsToAdd);

            await UpdateUserRegion(user, CodesAndRoles.EURegion);

            await AddPortugalSpainTeams(user);
        }

        private async Task NormalizeNAUser(Entity user)
        {
            string username = user.Contains("domainname") ? user["domainname"].ToString().Split('@')[0] : "";
            bool isInternal = IsInternalUser(username);

            if (isInternal)
            {
                await EnsureUserHasRoles(user, CodesAndRoles.NADefaultRolesForInternalUser);
                await UpdateUserRegion(user, CodesAndRoles.NARegion);
            }
            else
            {
                Console.WriteLine("User does not match NA internal pattern.");
            }
        }
    }
}
