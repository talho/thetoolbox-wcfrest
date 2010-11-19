using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Text;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Security;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.EnterpriseServices;
using System.Management.Automation;
using System.Xml;
using System.Xml.Serialization;
using System.ServiceModel.Channels;
using ToolBoxUtility;

namespace TALHO
{
    // Start the service and browse to http://<machine_name>:<port>/ExchSvc/help to view the service's generated help page
    // NOTE: By default, a new instance of the service is created for each call; change the InstanceContextMode to Single if you want
    // a single instance of the service to process all calls.	
    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single, AddressFilterMode = AddressFilterMode.Prefix)]
    // NOTE: If the service is renamed, remember to update the global.asax.cs file
    public class DstrSvc
    {
        // CreateDistributionGroup()
        // desc: Method creates a new distribution group
        // params: XML post from the client server
        //         string alias    - User login name
        // method: Web Post HTTP/XML
        // return: DistributionGroup Object
        [OperationContract]
        [WebInvoke(UriTemplate = "/", Method = "POST", BodyStyle = WebMessageBodyStyle.Bare, RequestFormat = WebMessageFormat.Xml, ResponseFormat = WebMessageFormat.Xml)]
        public Message Create(Stream body)
        {
            StreamReader rd = new StreamReader(body);
            string bodyXml = rd.ReadToEnd();
            DistributionGroupCreationParams parms = XmlSerializationHelper.Deserialize<DistributionGroupCreationParams>(bodyXml);

            //begin read in xml           
            string result = DistributionRepo.CreateDistributionGroup(parms.Name, parms.OrganizationalUnit, parms.AuthEnabled);
            DistributionGroup group = XmlSerializationHelper.Deserialize<DistributionGroup>(result);
            if (group.error != "")
            {
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
            }

            return MessageBuilder.CreateResponseMessage(group);
        }

        // Get()
        // desc: Find and return Distribution group
        // params: string identity - Name of Distribution List
        // method: Web Get
        // return: Serialized XML representation of Distribution object
        [WebGet(UriTemplate = "{identity}")]
        public Message Get(string identity)
        {
                identity = HttpUtility.UrlDecode(identity);
                string result = DistributionRepo.GetDistributionGroup(identity, 0, 0, "");
            try
            {
                if (result == null || result.IndexOf("Error") != -1)
                {
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
                    return null;
                }

                DistributionGroupsShorter shorty = XmlSerializationHelper.Deserialize<DistributionGroupsShorter>(result);

                if (shorty.groups.Count < 1)
                    return null;
                else
                    return MessageBuilder.CreateResponseMessage(shorty.groups[0]);
            }
            catch (Exception e)
            {
                string message = "";                
                while (e != null)
                {
                    message += e.Message;
                    e = e.InnerException;
                }
                return MessageBuilder.CreateResponseMessage(result + message);
            }
        }
        
        [WebGet(UriTemplate = "?page={current_page}&per_page={per_page}&ou={ou}")]
        public Message GetAll(int current_page = 0, int per_page = 0, string ou = "")
        {
            try
            {
                //current_page = current_page < 1 ? 1 : current_page;
                //per_page = per_page < 1 ? 10 : per_page;
                DistributionGroupsShorter shorty = new DistributionGroupsShorter();
                
                string result = DistributionRepo.GetDistributionGroup("", current_page, per_page, ou);

                shorty = XmlSerializationHelper.Deserialize<DistributionGroupsShorter>(result);
                  
                return MessageBuilder.CreateResponseMessage(shorty);
            }
            catch (Exception e)
            {
                string message = "";
                while (e != null)
                {
                    message += e.Message;
                    e = e.InnerException;
                }
                XmlDocument doc = new XmlDocument();
                XmlElement elem = doc.CreateElement("error");
                elem.InnerText = message;
                return MessageBuilder.CreateResponseMessage(doc);
            }
        }

        [WebInvoke(UriTemplate = "{identity}/update", Method = "POST")]
        public Message Update(string identity, Stream body)
        {
            identity = identity.Replace("+", " ");
            StreamReader rd = new StreamReader(body);
            string bodyXml = rd.ReadToEnd().Replace("DstrSvc", "distribution-group");

            try
            {
                DistributionGroup updated = DistributionRepo.UpdateDistributionGroup(XmlSerializationHelper.Deserialize<DistributionGroup>(bodyXml));
  
                return MessageBuilder.CreateResponseMessage(updated);
            }
            catch (Exception e)
            {
                if (e.Message == "Contact Conflict")
                {
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.Conflict;
                    return MessageBuilder.CreateResponseMessage(e.Message);
                }
                else
                {
                    string message = "";
                    while (e != null)
                    {
                        message += e.Message;
                        e = e.InnerException;
                    }
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.InternalServerError;                
                    return MessageBuilder.CreateResponseMessage(message);
                }
            }
        }

        [WebInvoke(UriTemplate = "{identity}/delete", Method = "POST")]
        public Message Delete(string identity, Stream body)
        {
            identity = identity.Replace("+", " ");
            StreamReader rd = new StreamReader(body);
            string bodyXml = rd.ReadToEnd().Replace("DstrSvc", "distribution-group");
            
            try
            {
                DistributionRepo.DeleteDistributionGroup(XmlSerializationHelper.Deserialize<DistributionGroup>(bodyXml));
                return MessageBuilder.CreateResponseMessage(true);
            }
            catch (Exception e)
            {
                string message = "";
                while (e != null)
                {
                    message += e.Message;
                    e = e.InnerException;
                } 
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.InternalServerError;
                return MessageBuilder.CreateResponseMessage(message);
            }
        }
    }
}
