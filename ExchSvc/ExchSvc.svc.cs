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
using System.EnterpriseServices;
using System.Management.Automation;
using System.Xml;
using System.Xml.Serialization;



namespace TALHO
{
    // Start the service and browse to http://<machine_name>:<port>/ExchSvc/help to view the service's generated help page
    // NOTE: By default, a new instance of the service is created for each call; change the InstanceContextMode to Single if you want
    // a single instance of the service to process all calls.	
    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single, AddressFilterMode = AddressFilterMode.Prefix)]
    // NOTE: If the service is renamed, remember to update the global.asax.cs file
    public class ExchSvc
    {
        [WebGet(UriTemplate = "")]
        public ExchangeUsers GetCollection()
        {
            string result = ExchangeUser.GetUser("");
            XmlSerializer serializer = new XmlSerializer(typeof(ExchangeUsers));
            StringReader textReader = new StringReader(result);
            ExchangeUsers users = (ExchangeUsers)serializer.Deserialize(textReader);
            textReader.Close();
            return users;
        }

        [OperationContract]
        [WebInvoke(UriTemplate = "/", Method = "POST", BodyStyle = WebMessageBodyStyle.Bare, RequestFormat = WebMessageFormat.Xml, ResponseFormat = WebMessageFormat.Xml)]
        public ExchangeUser Create()
        {

            StringWriter sw   = new StringWriter();
            XmlTextWriter xtw = new XmlTextWriter(sw);
            OperationContext.Current.RequestContext.RequestMessage.WriteMessage(xtw);
            xtw.Flush();
            xtw.Close();
            XmlTextReader xtr = new XmlTextReader(new StringReader(sw.ToString()));
            string s          = xtr.ReadElementString("Binary");
            string s1         = System.Text.ASCIIEncoding.ASCII.GetString(System.Convert.FromBase64String(s));

            System.Xml.XmlDocument xml_doc = new System.Xml.XmlDocument();
            xml_doc.LoadXml(@s1);
            string alias = xml_doc.SelectSingleNode("/ExchSvc/alias").InnerText;
            string domain = xml_doc.SelectSingleNode("/ExchSvc/domain").InnerText;
            string upn = alias + "@" + domain;

            string result = ExchangeUser.EnableMailbox(upn, alias).ToString();
            XmlSerializer serializer = new XmlSerializer(typeof(ExchangeUser));
            StringReader textReader = new StringReader(result);
            ExchangeUser user = (ExchangeUser)serializer.Deserialize(textReader);
            user.upn = upn;
            user.alias = alias;
            textReader.Close();
            return user;

        }

        [WebGet(UriTemplate = "{alias}")]
        public ExchangeUser Get(string alias)
        {
            string result = ExchangeUser.GetUser(alias);
            XmlSerializer serializer = new XmlSerializer(typeof(ExchangeUser));
            StringReader textReader = new StringReader(result);
            ExchangeUser user = (ExchangeUser)serializer.Deserialize(textReader);
            textReader.Close();
            return user;
        }

        [WebInvoke(UriTemplate = "{id}", Method = "PUT")]
        public bool Update(string id)
        {
            // TODO: Update the given instance of SampleItem in the collection
            //throw new NotImplementedException();
            return true;
        }

        [OperationContract]
        [WebInvoke(UriTemplate = "{alias}/delete", Method = "POST")]
        public void Delete(string alias)
        {
            ExchangeUser result = Get(alias);
            if (result.upn.CompareTo("") == 0)
            {
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.Gone;
            }
            else
            {
                bool r = ExchangeUser.RemoveMailbox(alias);
                if (r)
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.OK;
                else
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.Gone;
            }
        }
        
    }
}
