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
    // Class DistributionGroup
    [DataContract(Name = "distribution-group")]
    public class DistributionGroup
    {
        public DistributionGroup()
        {
            error = "";
        }

        [DataMember(Name = "name")]
        [XmlElement("name")]
        public string Name { get; set; }
        [DataMember(Name = "displayName")]
        [XmlElement("displayName")]
        public string displayName { get; set; }
        [DataMember(Name = "groupType")]
        [XmlElement("groupType")]
        public string groupType { get; set; }
        [DataMember(Name = "primarySmtpAddress")]
        [XmlElement("primarySmtpAddress")]
        public string primarySmtpAddress { get; set; }
        [DataMember(Name = "error")]
        [XmlElement("error")]
        public string error { get; set; }

        [DataMember(Name = "exchange-users")]
        [XmlArray("users")]
        [XmlArrayItem("ExchangeUser")]
        public List<ExchangeUser> users { get; set; }

        // GetDistributionGroup()
        // desc: Method calls PowerShellComponent command GetDistributionGroup, returns list of DistributionGroup/DistributionGroups
        // params: string identity  - Distribution group name, passed in by ExchSvc
        //         int current_page - Page to return
        //         int per_page     - Number of entries to return per page
        // method: public 
        // return: string, XML string of DistributionGroup object
        public static string GetDistributionGroup(string identity, int current_page, int per_page)
        {
            string result;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            result = objManage.GetDistributionGroup(identity, current_page, per_page);
            objManage = null;
            return result;
        }

        // CreateDistributionList()
        // desc: Method calls PowershellComponent command CreateDistributionList and creates a new distribution group under specified OU
        // params: string group_name - Name of Distribution Group to create
        //         string ou         - Name of Organizational Unit to create Organizational Group in
        // method: public
        // return: bool
        public static string CreateDistributionGroup(string group_name, string ou)
        {
            string Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results = objManage.CreateDistributionGroup(group_name, ou);
            objManage = null;
            return Results;
        }

        // AddToDistributionGroup()
        // desc: Method calls PowershellComponent command AddToDistributionGroup and adds a user to a distribution group
        // params: string group_name - Name of Distribution Group to add user to 
        //         string alias      - Name of user to add to group
        // method: public
        // return: bool
        public static bool AddToDistributionGroup(string group_name, string alias)
        {
            bool Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results = objManage.AddToDistributionGroup(group_name, alias);
            objManage = null;
            return Results;
        }

        // CreateMailContact()
        // desc: Method calls PowerShellComponent command CreateMailContact and creates a new mail contact under given OU
        // params: string name  - Name of mail contact
        //         string email - Email of mail contact
        //         string ou    - Organizational Unit to create mail contact
        // method: public
        // return: bool
        public static bool CreateMailContact(string name, string email, string ou)
        {
            bool Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results = objManage.CreateMailContact(name, email, ou);
            objManage = null;
            return Results;
        }
    }

    // Class DistributionGroupCollection, can take in a list of DistributionGroup, serialize them, then return an XML with a list of DistributionGroups
    public class DistributionGroupCollection : List<DistributionGroup>
    {
        public int current_page { get; set; }
        public int per_page { get; set; }
        public int total_entries { get; set; }

        // toXML()
        // desc: Method serializes a list of DistributionGroups, returns serialized XML
        // params: none
        // return: DistributionGroups, IXmlSerialized
        public DistributionGroups toXML()
        {
            DataContractSerializer serializer = new DataContractSerializer(typeof(DistributionGroup));
            using (MemoryStream memoryStream = new MemoryStream())
            {
                foreach (DistributionGroup user in this) serializer.WriteObject(memoryStream, user);
                memoryStream.Position = 0;

                String xml = null;
                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    xml = reader.ReadToEnd();
                    reader.Close();
                }
                memoryStream.Close();
                return new DistributionGroups(xml, total_entries, current_page, per_page);
            }
        }
    }

    // Class DistributionGroups, takes in an XML object of DistributionGroup
    [XmlRoot(ElementName = "distribution-groups")]
    public class DistributionGroups : IXmlSerializable
    {
        private int size;
        private int current_page;
        private int per_page;

        // ExchangeUsers()
        // desc: Constructor, set Content, size, current_page, and per_page
        // params: string content   - XML string 
        //         int size         - total number of entries
        //         int current_page - page to return
        //         int per_page     - number of entries to return per page
        // return: none
        public DistributionGroups(string content, int size, int current_page, int per_page)
        {
            this.Content = content;
            this.size = size;
            this.current_page = current_page;
            this.per_page = per_page;
        }

        public string Content { get; set; }

        //Overload IXmlSerializable
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

    [XmlRoot("distribution-groups")]
    public class DistributionGroupsShorter
    {
        [XmlAttribute("current_page")]
        public int CurrentPage { get; set; }
        [XmlAttribute("per_page")]
        public int PerPage { get; set; }
        [XmlAttribute("total_entries")]
        public int TotalEntries { get; set; }

        [XmlAttribute("type")]
        public string type = "array";

        [XmlElement("distribution-group")]
        public List<DistributionGroup> groups { get; set; }
    }
}