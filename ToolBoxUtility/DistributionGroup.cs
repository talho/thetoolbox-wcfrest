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

namespace ToolBoxUtility
{
    // Class DistributionGroup
    [DataContract(Name = "distribution-group")]
    [XmlRoot("distribution-group")]
    public class DistributionGroup
    {
        public DistributionGroup()
        {
            error = "";
        }

        [XmlElement("id")]
        public string Alias { get; set; }

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
        [XmlElement("ExchangeUsers")]
        public ExchangeUserMembers users { get; set; }

        [XmlElement("ou")]
        public string OrganizationalUnit { get; set; }
    }

    [XmlRoot("distribution-groups")]
    public class DistributionGroupsShorter
    {

        [XmlAttribute("current_page")]
        public string CurrentPageDisplay
        {
            get { return CurrentPage.HasValue && CurrentPage.Value > 0 ? CurrentPage.Value.ToString() : null; }
            set
            {
                int val;
                if (int.TryParse(value, out val))
                {
                    CurrentPage = val;
                }
                else
                {
                    CurrentPage = null;
                }
            }
        }

        [XmlIgnore]
        public int? CurrentPage { get; set; }
        
        [XmlAttribute("per_page")]
        public string PerPageDisplay
        {
            get { return PerPage.HasValue && PerPage.Value > 0 ? PerPage.Value.ToString() : null; }
            set
            {
                int val;
                if (int.TryParse(value, out val))
                {
                    PerPage = val;
                }
                else
                {
                    PerPage = null;
                }
            }
        }

        [XmlIgnore]
        public int? PerPage { get; set; }

        [XmlAttribute("total_entries")]
        public int TotalEntries { get; set; }

        [XmlAttribute("type")]
        public string type = "array";

        [XmlElement("distribution-group")]
        public List<DistributionGroup> groups { get; set; }
    }

    [XmlRoot("DstrSvc")]
    public class DistributionGroupCreationParams
    {
        [XmlElement("group-name")]
        public string Name { get; set; }

        [XmlElement("ou")]
        public string OrganizationalUnit { get; set; }
    }

    /*
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
    */
}