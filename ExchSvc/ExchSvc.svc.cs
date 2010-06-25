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
            XmlTextReader xtr                     = new XmlTextReader(new StringReader(sw.ToString()));
            string s                              = xtr.ReadElementString("Binary");
            string s1                             = System.Text.ASCIIEncoding.ASCII.GetString(System.Convert.FromBase64String(s));
            Dictionary<string, string> attributes = new Dictionary<string,string>();              
            System.Xml.XmlDocument xml_doc        = new System.Xml.XmlDocument();
            xml_doc.LoadXml(@s1);

            attributes.Add("alias",  xml_doc.SelectSingleNode("/ExchSvc/alias").InnerText);
            attributes.Add("domain", xml_doc.SelectSingleNode("/ExchSvc/domain").InnerText);
            attributes.Add("upn", xml_doc.SelectSingleNode("/ExchSvc/userPrincipalName").InnerText);
            attributes.Add("cn", xml_doc.SelectSingleNode("/ExchSvc/cn").InnerText);
            attributes.Add("name", xml_doc.SelectSingleNode("/ExchSvc/name").InnerText);
            attributes.Add("displayName", xml_doc.SelectSingleNode("/ExchSvc/displayName").InnerText);
            attributes.Add("dn" , xml_doc.SelectSingleNode("/ExchSvc/distinguishedName").InnerText);
            attributes.Add("givenName", xml_doc.SelectSingleNode("/ExchSvc/givenName").InnerText);
            attributes.Add("samAccountName", xml_doc.SelectSingleNode("/ExchSvc/samAccountName").InnerText);
            attributes.Add("unicodePwd", xml_doc.SelectSingleNode("/ExchSvc/unicodePwd").InnerText);
            attributes.Add("sn", xml_doc.SelectSingleNode("/ExchSvc/sn").InnerText);
            attributes.Add("changePwd", xml_doc.SelectSingleNode("/ExchSvc/changePwd").InnerText);
            attributes.Add("isVPN", xml_doc.SelectSingleNode("/ExchSvc/isVPN").InnerText);
            attributes.Add("acctDisabled", xml_doc.SelectSingleNode("/ExchSvc/acctDisabled").InnerText);
            attributes.Add("pwdExpires", xml_doc.SelectSingleNode("/ExchSvc/pwdExpires").InnerText);

            string newResult = ExchangeUser.NewADUser(attributes);
            
            if (newResult.IndexOf("Error") != -1)
            {
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
            }
            //ExchangeUser userResult = Get(attributes["alias"]);
            if (attributes["isVPN"] == "1")
            {
                attributes["alias"]          = attributes["alias"] + "-vpn";
                attributes["upn"]            = attributes["upn"].Replace("@", "-vpn@");
                attributes["dn"]             = attributes["dn"].Replace("OU=TALHO", "OU=VPN");
                attributes["samAccountName"] = attributes["samAccountName"] + "-vpn";
                ExchangeUser.NewADUser(attributes);
            }
            string exchResult        = ExchangeUser.EnableMailbox(attributes["upn"], attributes["alias"]).ToString();
            XmlSerializer serializer = new XmlSerializer(typeof(ExchangeUser));
            StringReader textReader  = new StringReader(exchResult);
            ExchangeUser user        = (ExchangeUser)serializer.Deserialize(textReader);
            user.upn                 = attributes["upn"];
            user.alias               = attributes["alias"];
            textReader.Close();
            
            return user;

        }


        [WebGet(UriTemplate = "{alias}")]
        public ExchangeUser Get(string alias)
        {
            string result = ExchangeUser.GetUser(alias);
            if (result == null || result.IndexOf("Error") != -1)
            {
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
                return null;
            }
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
                WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
            }
            else
            {
                bool r = ExchangeUser.RemoveMailbox(alias);
                if (r)
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.OK;
                else
                    WebOperationContext.Current.OutgoingResponse.StatusCode = System.Net.HttpStatusCode.NotFound;
            }
        }
        
    }
}
