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
        public DistributionGroup Create()
        {
            //begin read in xml
            /*
            StringWriter sw = new StringWriter();
            XmlTextWriter xtw = new XmlTextWriter(sw);
            OperationContext.Current.RequestContext.RequestMessage.WriteMessage(xtw);
            xtw.Flush();
            xtw.Close();
            XmlTextReader xtr = new XmlTextReader(new StringReader(sw.ToString()));
            string s = xtr.ReadElementString("Binary");
            string s1 = System.Text.ASCIIEncoding.ASCII.GetString(System.Convert.FromBase64String(s));
            System.Xml.XmlDocument xml_doc = new System.Xml.XmlDocument();
            xml_doc.LoadXml(@s1);

            //read in attributes from xml_doc, store them in Dictionary variable attributes
            string group_name = xml_doc.SelectSingleNode("/distribution-group/group-name") == null ? "" : xml_doc.SelectSingleNode("/distribution-group/group-name").InnerText;
            string ou = xml_doc.SelectSingleNode("/distribution-group/ou") == null ? "" : xml_doc.SelectSingleNode("/distribution-group/ou").InnerText;

            string result = DistributionGroup.CreateDistributionGroup(group_name, ou);
            if (result.IndexOf("Error") != -1)
            {
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
            }
             * */
            XmlSerializer serializer = new XmlSerializer(typeof(DistributionGroup));
            StringReader textReader  = new StringReader("");
            DistributionGroup group        = (DistributionGroup)serializer.Deserialize(textReader);
            textReader.Close();
            return group;
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
                string result = DistributionRepo.GetDistributionGroup(identity, 0, 0);
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



        [WebGet(UriTemplate = "?page={current_page}&per_page={per_page}", ResponseFormat=WebMessageFormat.Json)]
        public Message GetAll(int current_page = 1, int per_page = 10)
        {
            try
            {
                current_page = current_page < 1 ? 1 : current_page;
                per_page = per_page < 1 ? 10 : per_page;
                DistributionGroupsShorter shorty = new DistributionGroupsShorter();
                
                string result = DistributionRepo.GetDistributionGroup("", current_page, per_page);

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

        // CreateMailContact
        // desc: Method creates a new mail contact
        // params: string name  - Name of mail contact
        //         string email - Email of mail contact
        //         string ou    - Name of Organizational Unit to create mail contact
        // method: Web Post HTTP/XML
        // return: void, method sets WebOperationContact status code to OK or Not Found
        /*
        [OperationContract]
        [WebInvoke(UriTemplate = "/{identity}/create_mail_contact", Method = "POST", BodyStyle = WebMessageBodyStyle.Bare, RequestFormat = WebMessageFormat.Xml)]
        public void CreateMailContact()
        {
            //begin read in xml
            StringWriter sw = new StringWriter();
            XmlTextWriter xtw = new XmlTextWriter(sw);
            OperationContext.Current.RequestContext.RequestMessage.WriteMessage(xtw);
            xtw.Flush();
            xtw.Close();
            XmlTextReader xtr = new XmlTextReader(new StringReader(sw.ToString()));
            string s = xtr.ReadElementString("Binary");
            string s1 = System.Text.ASCIIEncoding.ASCII.GetString(System.Convert.FromBase64String(s));
            System.Xml.XmlDocument xml_doc = new System.Xml.XmlDocument();
            xml_doc.LoadXml(@s1);

            //read in attributes from xml_doc, store them in Dictionary variable attributes
            string name  = xml_doc.SelectSingleNode("/distribution-group/name") == null ? "" : xml_doc.SelectSingleNode("/distribution-group/name").InnerText;
            string email = xml_doc.SelectSingleNode("/distribution-group/email") == null ? "" : xml_doc.SelectSingleNode("/distribution-group/email").InnerText;
            string ou = xml_doc.SelectSingleNode("/distribution-group/email") == null ? "" : xml_doc.SelectSingleNode("/distribution-group/email").InnerText;

            bool r = DistributionGroup.CreateMailContact(name, email, ou);
            if (r)
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.OK;
            else
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
        }*/

        //[WebGet(UriTemplate = "/GetAllTest?page={current_page}&per_page={per_page}")]
        //public DistributionGroups GetAllTest(int current_page, int per_page)
        //{
        //    DistributionGroupCollection groups = null;
        //    try
        //    {
        //        current_page = current_page < 1 ? 1 : current_page;
        //        per_page = per_page < 1 ? 10 : per_page;
        //        string result = DistributionRepo.GetDistributionGroup("", current_page, per_page);
        //        string[] seperator = new string[1];
        //        seperator[0] = "THEWORLDSLARGESTSEPERATOR";
        //        string[] results = result.Split(seperator, StringSplitOptions.RemoveEmptyEntries);
        //        result = results[0];

        //        //groups = new DistributionGroupCollection() { per_page = 1, total_entries = 1, current_page = 1 };
        //        //groups.Add(new DistributionGroup() { error = result });

        //        XmlSerializer serializer = new XmlSerializer(typeof(DistributionGroupCollection));
        //        StringReader textReader = new StringReader(result);
        //        groups = (DistributionGroupCollection)serializer.Deserialize(textReader);
        //        groups.current_page = current_page < 1 ? 1 : current_page;
        //        groups.per_page = per_page < 1 ? 10 : per_page;
        //        groups.total_entries = Int32.Parse(results[1]);
                
        //        textReader.Close();
        //    }
        //    catch (Exception e)
        //    {
        //        DistributionGroup group = new DistributionGroup();
        //        while (e != null)
        //        {
        //            group.error += e.Message;
        //            e = e.InnerException;
        //        }
        //        groups = new DistributionGroupCollection() { total_entries = 1, per_page = 1, current_page = 1};
        //        groups.Add(group);
        //    }

        //    return groups.toXML();
        //}
    }
}
