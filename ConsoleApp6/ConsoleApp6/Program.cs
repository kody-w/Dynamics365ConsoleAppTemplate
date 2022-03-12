using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using System.Xml;
using Microsoft.Xrm.Sdk.Messages;

namespace ConsoleApp6
{
    class Program
    {
        static void Main(string[] args)
        {
            IOrganizationService organizationService = null;

            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = ""; //Email Address Here
                clientCredentials.UserName.Password = ""; // Password Here

                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources
                // Copy and Paste Organization Service Endpoint Address URL
                organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri("https://.api.crm.dynamics.com/XRMServices/2011/Organization.svc"), //fill in url details
                 null, clientCredentials, null);

                if (organizationService != null)
                {
                    Guid userid = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Established Successfully...");

                        //GenerateEmailForMobile(organizationService, new Guid());

                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught - " + ex.Message);
            }
            Console.WriteLine("End of Code");
            Console.ReadKey();
        }

        public static List<Entity> ExecutePagingFetch(string fetchXML, IOrganizationService orgService)
        {
            try
            {
                bool hasMoreRecords = false;
                int fetchCount = 5000;
                int pageNumber = 1;
                string pagingCookie = null;

                List<Entity> records = new List<Entity>();

                do
                {
                    string xml = CreateXml(fetchXML, pagingCookie, pageNumber, fetchCount);

                    var retMulRes = orgService.RetrieveMultiple(new FetchExpression(xml));

                    records.AddRange(retMulRes.Entities);

                    if (retMulRes.MoreRecords)
                    {
                        pagingCookie = retMulRes.PagingCookie;
                        hasMoreRecords = true;
                        pageNumber++;
                    }
                    else
                    {
                        hasMoreRecords = false;
                    }

                } while (hasMoreRecords);

                return records;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static string CreateXml(string xml, string cookie, int page, int count)
        {
            try
            {
                StringReader stringReader = new StringReader(xml);
                XmlTextReader reader = new XmlTextReader(stringReader);

                // Load document
                XmlDocument doc = new XmlDocument();
                doc.Load(reader);

                return CreateXml(doc, cookie, page, count);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            try
            {
                XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

                if (cookie != null)
                {
                    XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                    pagingAttr.Value = cookie;
                    attrs.Append(pagingAttr);
                }

                XmlAttribute pageAttr = doc.CreateAttribute("page");
                pageAttr.Value = System.Convert.ToString(page);
                attrs.Append(pageAttr);

                XmlAttribute countAttr = doc.CreateAttribute("count");
                countAttr.Value = System.Convert.ToString(count);
                attrs.Append(countAttr);

                StringBuilder sb = new StringBuilder(1024);
                StringWriter stringWriter = new StringWriter(sb);

                XmlTextWriter writer = new XmlTextWriter(stringWriter);
                doc.WriteTo(writer);
                writer.Close();

                return sb.ToString();
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static Entity GenerateEmailForMobile(IOrganizationService service, Guid id)
        {
            var fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='tm_emailformobile'>
                                    <attribute name='activityid' />
                                    <attribute name='subject' />
                                    <attribute name='to' />
                                    <attribute name='from' />
                                    <attribute name='description' />
                                    <attribute name='cc' />
                                    <attribute name='bcc' />
                                    <attribute name='regardingobjectid' />
                                    <order attribute='subject' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='activityid' operator='eq' value='" + id + @"' />
                                    </filter>
                                  </entity>
                                </fetch>";

            return service.RetrieveMultiple(new FetchExpression(fetchXML)).Entities.FirstOrDefault();
        }
    }
}
