using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilDuplication
{
    public static class Service
    {
        public static List<Entity> RetrieveAll(this IOrganizationService service, FetchExpression query)
        {
            var conversionRequest = new FetchXmlToQueryExpressionRequest
            {
                FetchXml = query.Query
            };

            var conversionResponse =
                (FetchXmlToQueryExpressionResponse)service.Execute(conversionRequest);

            return RetrieveAll(service, conversionResponse.Query);
        }

        public static List<Entity> RetrieveAll(this IOrganizationService service, QueryExpression query)
        {
            var result = new List<Entity>();

            var entities = service.RetrieveMultiple(query);
            result.AddRange(entities.Entities);

            var page = 2;
            while (entities.MoreRecords)
            {
                query.PageInfo = new PagingInfo
                {
                    PagingCookie = entities.PagingCookie,
                    PageNumber = page
                };

                entities = service.RetrieveMultiple(query);
                result.AddRange(entities.Entities);
                page++;
            }

            return result;
        }
    }
}
