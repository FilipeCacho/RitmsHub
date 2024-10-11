using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserNormalizer
    {

        private async Task AddPortugalSpainTeams(Entity user)
        {
            if (user.Contains("businessunitid"))
            {
                string buName = ((EntityReference)user["businessunitid"]).Name;
                if (buName.Contains("Portugal") || buName.Contains("Spain") ||
                    (buName.Contains("-") && (buName.Substring(buName.IndexOf('-') + 1, 2) == "PT" || buName.Substring(buName.IndexOf('-') + 1, 2) == "ES")))
                {
                    await EnsureUserHasTeams(user, CodesAndRoles.EUDefaultTeamForPortugueseAndSpanishUsers);
                }
            }
        }

    }
}