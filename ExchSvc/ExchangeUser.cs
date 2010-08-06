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

// Scope TALHO
namespace TALHO
{
    // Class ExchangeUser
    public static class ExchangeRepo
    {
        // GetUser()
        // desc: Method calls PowerShellComponent command GetUser, returns list of user/users
        // params: string thealias  - User login name, passed in by ExchSvc
        //         int current_page - Page to return
        //         int per_page     - Number of entries to return per page
        // method: public 
        // return: string, XML string of ExchangeUser object
        public static string GetUser(string thealias, int current_page, int per_page)
        {
            string result;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            result                                           = objManage.GetUser(thealias, current_page, per_page);
            objManage                                        = null;
            return result;
        }

        // NewExchangeUser()
        // desc: Method calls PowerShellComponent command NewExchangeUser, creates new mailbox user on Exchange Server
        // params: Dictionary<string, string> attributes - Dictionary object, contains attributes for new user
        // method: public
        // return: string, XML string of ExchangeUser object
        public static string NewExchangeUser(Dictionary<string, string> attributes)
        {
            PowerShellComponent.ManagementCommands objManage = null;
            string Results                                   = null;
            try{
                objManage = new PowerShellComponent.ManagementCommands();
                Results   = objManage.NewExchangeUser(attributes);
            }catch (Exception e){
                return e.Message;
            }
            objManage = null;
            return Results;
        }

        // EnableMailbox()
        // desc: Method calls PowerShellComponent command EnableMailbox, enables user mailbox on Exchange Server
        // params: string upn      - User Principal Name
        //         string thealias - User login name
        // method: public
        // return: string, XML string of ExchangeUser object
        public static string EnableMailbox(string upn, string thealias)
        {
            string Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results                                          = objManage.EnableMailbox(upn, thealias);
            objManage                                        = null;
            return Results;
        }

        // NewADUser()
        // desc: Method calls PowerShellComponent command NewADUser, creates a user in ActiveDirectory only
        // params: Dictionary<string, string> attributes - Dictionary object, contains attributes for new user
        // method: public
        // return: string, XML string of ExchangeUser object
        public static string NewADUser(Dictionary<string, string> attributes)
        {
            PowerShellComponent.ManagementCommands objManage = null;
            string Results                                   = null;
            try{
                objManage = new PowerShellComponent.ManagementCommands();
                Results   = objManage.NewADUser(attributes);
            }catch (Exception e){
                return e.Message;
            }
            objManage = null;
            return Results;
        }

        // ChangePassword()
        // desc: Method calls PowerShellComponent command ChangePassword, changes user password with password given
        // params: Dictionary<string, string> attributes - Dictionary object, contains attributes to change user password
        // method: public
        // return: bool
        public static bool ChangePassword(string identity, string password)
        {
            bool Results;
            PowerShellComponent.ManagementCommands objManange = new PowerShellComponent.ManagementCommands();
            Results                                           = objManange.ChangePassword(identity, password);
            objManange                                        = null;
            return Results;
        }

        // RemoveMailBox()
        // desc: Method calls PowerShellComponent command DeleteUser, deletes user from Exchange and AD Server
        // params: string identity - User login name
        // method: public
        // return: bool
        public static bool RemoveMailbox(string identity)
        {
            bool Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results                                          = objManage.DeleteUser(identity);
            objManage                                        = null;
            return Results;
        }

        public static void DeleteMailContact(string alias)
        {
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            objManage.DeleteMailContact(alias);
        }

        internal static bool RemoveADUser(string identity)
        {
            bool Results;
            PowerShellComponent.ManagementCommands objManage = new PowerShellComponent.ManagementCommands();
            Results = objManage.DeleteADUser(identity);
            objManage = null;
            return Results;
        }
    }
}
