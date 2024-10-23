using Microsoft.Xrm.Sdk;
using System;
using System.Threading.Tasks;

namespace RitmsHub.Scripts
{
    public partial class UserNormalizer
    {
        private async Task GiveResco(Entity user, string regionChoice)
        {
            string askResco;
            do
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("\nDo you want to give RESCO access and Role to this user? (y, or n)");
                Console.ResetColor();

                askResco = Console.ReadLine();

                if (askResco == "y")
                {
                    if (regionChoice == "1") //eur
                    {
                        await EnsureUserHasTeams(user, CodesAndRoles.RescoTeamEU);
                        await EnsureUserHasRoles(user, CodesAndRoles.RescoRole);

                        Console.WriteLine("\nRESCO role and team were given to the user");


                    }

                    else if (regionChoice == "2") //na
                    {
                        await EnsureUserHasTeams(user, CodesAndRoles.RescoTeamNA);
                        await EnsureUserHasRoles(user, CodesAndRoles.RescoRole);

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("\nRESCO role and team were given to the user");
                        Console.ResetColor();

                    }
                }
                else if (askResco == "n")
                {
                    break;
                }


                else if (askResco != "y" || askResco != "n")
                {
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\nInvalid region choice. Please enter y, n");
                    Console.ResetColor();
                }
            } while (askResco != "y" && askResco != "n");
        }

    }
}