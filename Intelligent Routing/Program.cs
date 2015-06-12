using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using System.ServiceModel.Description;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;
using XRM;

namespace Intelligent_Routing
{
    public class CredentialProvider
    {
        public static string CRMUser
        {
            get
            {
                return ConfigurationManager.AppSettings["CRMUser"];
            }
        }
        public static string CRMPwd
        {
            get
            {
                return ConfigurationManager.AppSettings["CRMPwd"];
            }
        }
        public static ClientCredentials CRMCreds
        {
            get
            {
                ClientCredentials cre = new ClientCredentials();
                cre.UserName.UserName = CredentialProvider.CRMUser;
                cre.UserName.Password = CredentialProvider.CRMPwd;
                return cre;
            }
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Route();
        }
        static void Route()
        {
            Uri serviceUri = new Uri(ConfigurationManager.AppSettings["CRMUri"]);
            OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, CredentialProvider.CRMCreds, null);
            proxy.EnableProxyTypes();
            IOrganizationService service = proxy;

            using (XrmServiceContext orgContext = new XrmServiceContext(service))
            {
                Guid CurrentUserId = ((WhoAmIResponse)proxy.Execute(new WhoAmIRequest())).UserId;
                Guid destinationQueueId = new Guid();
                Guid incidentId = new Guid();

                IEnumerable<Incident> queryCases = from a in orgContext.IncidentSet
                                                   where a.Title == "Defective item delivered"
                                                   select a;

                foreach(var queryCase in queryCases)
                {
                    incidentId = queryCase.Id;
                }

                RetrieveUserQueuesRequest retrieveUserQueuesRequest = new RetrieveUserQueuesRequest
                {
                    UserId = CurrentUserId,
                    IncludePublic = false
                };

                RetrieveUserQueuesResponse retrieveUserQueuesResponse = (RetrieveUserQueuesResponse)proxy.Execute(retrieveUserQueuesRequest);
                EntityCollection queues = (EntityCollection)retrieveUserQueuesResponse.EntityCollection;

                foreach (Entity entity in queues.Entities)
                {
                    Queue queue = (Queue)entity;

                    if (queue.Name == "Churn Risk")
                    {
                        destinationQueueId = queue.Id;
                    }
                }

                AddToQueueRequest routeRequest = new AddToQueueRequest
                {
                    Target = new EntityReference(Incident.EntityLogicalName, incidentId),
                    DestinationQueueId = destinationQueueId
                };

                proxy.Execute(routeRequest);
            }
        }
    }
}