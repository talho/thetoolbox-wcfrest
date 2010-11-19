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
using ToolBoxUtility;

namespace TALHO
{
    // Class DistributionGroup
    public static class DistributionRepo
    {
        // GetDistributionGroup()
        // desc: Method calls PowerShellComponent command GetDistributionGroup, returns list of DistributionGroup/DistributionGroups
        // params: string identity  - Distribution group name, passed in by ExchSvc
        //         int current_page - Page to return
        //         int per_page     - Number of entries to return per page
        // method: public 
        // return: string, XML string of DistributionGroup object
        public static string GetDistributionGroup(string identity, int current_page, int per_page, string ou)
        {
            string result;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            result = objManage.GetDistributionGroup(identity, current_page, per_page, ou);
            objManage = null;
            return result;
        }

        // CreateDistributionList()
        // desc: Method calls PowershellComponent command CreateDistributionList and creates a new distribution group under specified OU
        // params: string group_name - Name of Distribution Group to create
        //         string ou         - Name of Organizational Unit to create Organizational Group in
        // method: public
        // return: bool
        public static string CreateDistributionGroup(string group_name, string ou, string auth_enabled)
        {
            string Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results = objManage.CreateDistributionGroup(group_name, ou, auth_enabled);
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

        public static DistributionGroup UpdateDistributionGroup(DistributionGroup group)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            string updatedGroupXml = objManage.UpdateDistributionGroup(XmlSerializationHelper.Serialize(group));
            return XmlSerializationHelper.Deserialize<DistributionGroup>(updatedGroupXml);
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

        public static void DeleteDistributionGroup(DistributionGroup group)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            objManage.DeleteDistributionGroup(XmlSerializationHelper.Serialize(group));
        }
    }

}