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
using iThink.Net;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
[System.Web.Script.Services.ScriptService]
public class AutoCompletion : WebService
{
    public AutoCompletion()
    {
    }

    [WebMethod]
    public string[] GetSupplierList(string prefixText, int count)
    {
        DataTable dtCompany;
        
        dtCompany = new DataTable();
        
        string sql;

        sql = "select ";
        
        sql = sql + " name from " + ConfigAppSettingsXyra.SchemaName + ".supplier where name like  " + Utility.QStr(prefixText+"%");
        sql = sql + " order by name ";
        sql = sql + " limit " + count.ToString();
        //string sql = "select description from unit_of_measure" ;
        dtCompany = Utility.CreateDataSet(sql).Tables[0];

        List<string> items = new List<string>(count);
        foreach (DataRow drCompany in dtCompany.Rows)
        {
            items.Add(drCompany["name"].ToString());
        }
        return items.ToArray();
    }

    [WebMethod]
    public string[] GetStreetNameList(string prefixText, int count)
    {
        DataTable dtTable;
        
        dtTable = new DataTable();
        
        string sql;

        sql = "select distinct ";
        //sql = sql + " top " + count.ToString();
        sql = sql + " street_name from " + ConfigAppSettingsXyra.SchemaName + ".supplier where street_name like  " + Utility.QStr(prefixText + "%");
        sql = sql + " order by street_name ";
        sql = sql + " limit " + count.ToString();
        //string sql = "select description from unit_of_measure" ;
        dtTable = Utility.CreateDataSet(sql).Tables[0];

        List<string> items = new List<string>(count);
        foreach (DataRow drRow in dtTable.Rows)
        {
            items.Add(drRow["street_name"].ToString());
        }
        return items.ToArray();
    }

    [WebMethod]
    public string[] GetBarangayList(string prefixText, int count)
    {
        DataTable dtTable;
        
        dtTable = new DataTable();
        
        string sql;

        sql = "select distinct ";
        //sql = sql + " top " + count.ToString();
        sql = sql + " barangay from " + ConfigAppSettingsXyra.SchemaName + ".supplier where barangay like  " + Utility.QStr(prefixText + "%");
        sql = sql + " order by barangay ";
        sql = sql + " limit " + count.ToString();
        //string sql = "select description from unit_of_measure" ;
        dtTable = Utility.CreateDataSet(sql).Tables[0];

        List<string> items = new List<string>(count);
        foreach (DataRow drRow in dtTable.Rows)
        {
            items.Add(drRow["barangay"].ToString());
        }
        return items.ToArray();
    }

    [WebMethod]
    public string[] GetCompoundList(string prefixText, int count)
    {
        DataTable dtTable;
        
        dtTable = new DataTable();
        
        string sql;

        sql = "select distinct ";
        //sql = sql + " top " + count.ToString();
        sql = sql + " compound from " + ConfigAppSettingsXyra.SchemaName + ".supplier where compound like  " + Utility.QStr(prefixText + "%");
        sql = sql + " order by compound ";
        sql = sql + " limit " + count.ToString();
        //string sql = "select description from unit_of_measure" ;
        dtTable = Utility.CreateDataSet(sql).Tables[0];

        List<string> items = new List<string>(count);
        foreach (DataRow drRow in dtTable.Rows)
        {
            items.Add(drRow["compound"].ToString());
        }
        return items.ToArray();
    }
}