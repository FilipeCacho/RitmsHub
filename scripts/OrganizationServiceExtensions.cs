using System;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace RitmsHub.Scripts
{
    public static class OrganizationServiceExtensions
    {
        public static Task<EntityCollection> RetrieveMultipleAsync(this IOrganizationService service, QueryBase query)
        {
            return Task.Run(() => service.RetrieveMultiple(query));
        }

        public static Task<Guid> CreateAsync(this IOrganizationService service, Entity entity)
        {
            return Task.Run(() => service.Create(entity));
        }

        public static Task UpdateAsync(this IOrganizationService service, Entity entity)
        {
            return Task.Run(() => service.Update(entity));
        }
    }
}