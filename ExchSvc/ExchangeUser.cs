using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace TALHO
{
    [DataContract(Name="exchange-user")]
    public class ExchangeUser
    {
        [DataMember(Name = "alias")]
        public string alias { get; set; }
        [DataMember(Name = "dn")]
        public string dn { get; set; }
        [DataMember(Name = "cn")]
        public string cn { get; set; }
        [DataMember(Name = "upn")]
        public string upn { get; set; }
        [DataMember(Name = "mailboxEnabled")]
        public bool mailboxEnabled { get; set; }
        [DataMember(Name = "ou")]
        public string ou { get; set; }
        [DataMember(Name = "login")]
        public string login { get; set; }
        [DataMember(Name = "email")]
        public string email { get; set; }
        [DataMember(Name = "has_vpn")]
        public bool has_vpn { get; set; }


        public static string GetUser(string thealias, int current_page, int per_page)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();

            string result;
            result = objManage.GetUser(thealias, current_page, per_page);

            objManage = null;
            return result;
        }

        public static string NewExchangeUser(Dictionary<string, string> attributes)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();

            string Results;
            Results = objManage.NewExchangeUser(attributes);

            objManage = null;
            return Results;
        }

        public static string EnableMailbox(string upn, string thealias)
        {

            // Create the COM+ component
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();

            string Results;
            Results = objManage.EnableMailbox(upn, thealias);

            objManage = null;
            return Results;
        }

        public static string NewADUser(Dictionary<string, string> attributes)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();

            string Results;
            Results = objManage.NewADUser(attributes);

            objManage = null;
            return Results;
        }

        public static bool RemoveMailbox(string identity)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            bool Results;
            Results = objManage.DeleteUser(identity);
            objManage = null;
            return Results;
        }

        
    }

    public class ExchangeUserCollection : List<ExchangeUser>
    {
        public int current_page { get; set; }
        public int per_page { get; set; }
        public int total_entries { get; set; }

        public ExchangeUsers toXML()
        {
            DataContractSerializer serializer = new DataContractSerializer(typeof(ExchangeUser));
            using (MemoryStream memoryStream = new MemoryStream())
            {
                foreach (ExchangeUser user in this) serializer.WriteObject(memoryStream, user);
                memoryStream.Position = 0;

                String xml = null;
                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    xml = reader.ReadToEnd();
                    reader.Close();
                }
                memoryStream.Close();
                return new ExchangeUsers(xml, total_entries, current_page, per_page);
            }
        }
    }

    [XmlRoot(ElementName = "exchange-users")]
    public class ExchangeUsers : IXmlSerializable
    {
        private int size;
        private int current_page;
        private int per_page;

        public ExchangeUsers(string content, int size, int current_page, int per_page)
        {
            this.Content = content;
            this.size = size;
            this.current_page = current_page;
            this.per_page = per_page;
        }
        public string Content { get; set; }

        #region IXmlSerializable Members

        public System.Xml.Schema.XmlSchema GetSchema()
        {
            return null;
        }

        public void ReadXml(XmlReader reader)
        {
            throw new NotImplementedException();
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.WriteAttributeString("type", "array");
            writer.WriteAttributeString("current_page", current_page.ToString());
            writer.WriteAttributeString("per_page", per_page.ToString());
            writer.WriteAttributeString("total_entries", size.ToString());
            writer.WriteRaw(this.Content);
        }
        #endregion
    }
}
