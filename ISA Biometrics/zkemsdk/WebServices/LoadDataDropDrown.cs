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
public class LoadDataDropDown : WebService
{
    public LoadDataDropDown()
    {
    }

    [WebMethod]
    public System.Collections.IEnumerable LoadSupplierDropDown(Telerik.Web.UI.RadComboBoxContext context)
    {

        string query;
        //string[] parameter;
        DataTable dataSource;
        //DataRow dr;

        XmlParser xmlParser;

        xmlParser = new XmlParser();


        //query = xmlParser.GetSelectSQLQuery("dropdown", "Supplier", FileFolderPathsXyra.QUERY_PATH_SUPPLIER);
        query = "select ";
        query = query + "id, ";
        query = query + "name ";
        query = query + " from ";
        query = query + ConfigAppSettingsXyra.SchemaName + ".supplier ";
        query = query + " order by ";
        query = query + " name asc ";
        

        dataSource = Utility.CreateDataSet(query).Tables[0];

        //dr = dataSource.NewRow();
        //dr["name"] = " --- ";

        //dataSource.Rows.Add(dr);
        //dataSource.DefaultView.Sort = "name ASC";


        List<ComboBoxItemData> items = new List<ComboBoxItemData>();

        foreach (DataRow dataRow in dataSource.Rows)
        {
            ComboBoxItemData itemData = new ComboBoxItemData();
            itemData.Text = dataRow["name"].ToString();
            items.Add(itemData);
        }

        return items;
    }

    class ComboBoxItemData
    {
        private string text;

        public string Text
        {
            get { return text; }
            set { text = value; }
        }
    }
}