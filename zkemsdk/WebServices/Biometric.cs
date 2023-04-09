// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/resources/sharedsource/licensingbasics/sharedsourcelicenses.mspx.
// All other rights reserved.


using System;
using System.Collections.Generic;
using System.Web.Services;
using System.Data;
using System.Text;

using iThink.Net.Library;
using iThink.Net.DataObjects;
using iThink.Net.Components;
using iThink.Net.Tools;
using iThink.Net.DataClass.Structures;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
[System.Web.Script.Services.ScriptService]
public class Biometric : WebService
{
    zkemkeeper.CZKEMClass zConnection = new zkemkeeper.CZKEMClass();
    public Biometric()
    {
    }

    [WebMethod]
    public string HelloWorld()
    {
        return "Hello World";
    }

    [WebMethod]
    public int add(int a, int b)
    {
        return a + b;
    }

    [WebMethod]
    public DeviceLogs[] Logs(string strIpAddress,ref bool blnConnect,string strLastDownload)
    {
        int intPort = 4370;
        int intMachineNumber = 1;

        int dwEnrollNumber=0;
        int dwVerifyMode=0;
        int dwInOutMode=0;
        string strTime="";

        
        //DeviceLogs[] logs;
        DeviceLogs log;

        List<DeviceLogs> logs = new List<DeviceLogs>();
        if (zConnection.Connect_Net(strIpAddress, intPort))
        {

            blnConnect = (zConnection.ReadGeneralLogData(intMachineNumber));
            if (blnConnect)
            {
                while (zConnection.GetGeneralLogDataStr(intMachineNumber, ref dwEnrollNumber, ref dwVerifyMode,
                    ref dwInOutMode, ref strTime))
                {
                    if (Convert.ToDateTime(strTime) >= Convert.ToDateTime(strLastDownload))
                    {
                        log.enrollNo = Convert.ToString(dwEnrollNumber);
                        log.verifyMode = Convert.ToString(dwVerifyMode);
                        log.inOutMode = Convert.ToString(dwInOutMode);
                        log.inOutTime = Convert.ToDateTime(strTime);
                        logs.Add(log);
                    }
                    //count = count + 1;
                }
            }
        }
        else
        {
            blnConnect = false;            
        }

        zConnection.Disconnect();

        return logs.ToArray();
    }

    //[WebMethod]
    //public string[] GetStreetNameList(string prefixText, int count)
    //{
    //    DataTable dtTable;
    //    DBConnection dbcnn;

    //    dtTable = new DataTable();
    //    dbcnn = new DBConnection(true);

    //    string sql;

    //    sql = "select distinct ";
    //    //sql = sql + " top " + count.ToString();
    //    sql = sql + " street_name from supplier where street_name like  " + Utility.QStr(prefixText + "%");
    //    sql = sql + " order by street_name ";
    //    //string sql = "select description from unit_of_measure" ;
    //    dtTable = dbcnn.Execute(sql).Tables[0];

    //    List<string> items = new List<string>(count);
    //    foreach (DataRow drRow in dtTable.Rows)
    //    {
    //        items.Add(drRow["street_name"].ToString());
    //    }
    //    return items.ToArray();
    //}

    //[WebMethod]
    //public string[] GetBarangayList(string prefixText, int count)
    //{
    //    DataTable dtTable;
    //    DBConnection dbcnn;

    //    dtTable = new DataTable();
    //    dbcnn = new DBConnection(true);

    //    string sql;

    //    sql = "select distinct ";
    //    //sql = sql + " top " + count.ToString();
    //    sql = sql + " barangay from supplier where barangay like  " + Utility.QStr(prefixText + "%");
    //    sql = sql + " order by barangay ";
    //    //string sql = "select description from unit_of_measure" ;
    //    dtTable = dbcnn.Execute(sql).Tables[0];

    //    List<string> items = new List<string>(count);
    //    foreach (DataRow drRow in dtTable.Rows)
    //    {
    //        items.Add(drRow["barangay"].ToString());
    //    }
    //    return items.ToArray();
    //}

    //[WebMethod]
    //public string[] GetCompoundList(string prefixText, int count)
    //{
    //    DataTable dtTable;
    //    DBConnection dbcnn;

    //    dtTable = new DataTable();
    //    dbcnn = new DBConnection(true);

    //    string sql;

    //    sql = "select distinct ";
    //    //sql = sql + " top " + count.ToString();
    //    sql = sql + " compound from supplier where compound like  " + Utility.QStr(prefixText + "%");
    //    sql = sql + " order by compound ";
    //    //string sql = "select description from unit_of_measure" ;
    //    dtTable = dbcnn.Execute(sql).Tables[0];

    //    List<string> items = new List<string>(count);
    //    foreach (DataRow drRow in dtTable.Rows)
    //    {
    //        items.Add(drRow["compound"].ToString());
    //    }
    //    return items.ToArray();
    //}
}