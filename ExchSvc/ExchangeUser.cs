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
    public class ExchangeUser
    {
        public string givenName { get; set; }
        public string sn { get; set;} 
        public string dn { get; set; }
        public string cn { get; set; }
        public string mailbox { get; set; }
        public string alias { get; set; }
        public string upn { get; set; }
        public bool mailboxEnabled { get; set; }

        public static string GetUser(string thealias)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();

            string result;
            result = objManage.GetUser(thealias);

            objManage = null;
            return result;
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

        public static bool RemoveMailbox(string identity)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            bool Results;
            Results = objManage.DeleteUser(identity);
            objManage = null;
            return Results;
        }
    }

    public class ExchangeUsers : List<ExchangeUser>
    {
       
    }
}
