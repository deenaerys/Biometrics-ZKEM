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
public class CascadingDropDown : WebService
{
    public CascadingDropDown()
    {
    }

    #region Region - Province - Town - Barangay

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetRegion(
      string knownCategoryValues,
      string category)
    {

        DataTable dtRegion;
        
        dtRegion = new DataTable();
        
        string sql;

        sql = "select ";
        sql = sql + " id, description ";
        sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".region ";
        sql = sql + " order by description ";
        dtRegion = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtRegion.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["description"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetProvince(
      string knownCategoryValues,
      string category)
    {

        DataTable dtProvince;
        
        string regionId;
        string categories;
        string filter = "";

        dtProvince = new DataTable();
        
        categories = knownCategoryValues;

        if (!Utility.IsNullOrEmpty(categories))
        {
            regionId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);
            filter = " where ";
            filter = filter + " region_id = " + Utility.QStr(regionId);
        }

        string sql;

        sql = "select ";
        sql = sql + " id, description ";
        sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".province ";
        sql = sql + filter;
        sql = sql + " order by description ";
        dtProvince = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtProvince.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["description"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetTown(
      string knownCategoryValues,
      string category)
    {

        DataTable dtTown;
        string provinceId;
        string categories;

        dtTown = new DataTable();
        categories = knownCategoryValues;

        provinceId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);

        string sql;

        sql = "select ";
        sql = sql + " id, description ";
        sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".town ";
        sql = sql + " where ";
        sql = sql + " province_id = " + Utility.QStr(provinceId);
        sql = sql + " order by description ";
        dtTown = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtTown.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["description"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetBarangay(
      string knownCategoryValues,
      string category)
    {

        DataTable dtBarangay;
        string townId;
        string categories;

        dtBarangay = new DataTable();
        categories = knownCategoryValues;

        townId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);

        string sql;

        sql = "select ";
        sql = sql + " id, description ";
        sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".barangay ";
        sql = sql + " where ";
        sql = sql + " town_id = " + Utility.QStr(townId);
        sql = sql + " order by description ";
        dtBarangay = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtBarangay.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["description"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    #endregion

    #region ProductLine - Product Type - Product Sub-Type - Product Model - Obsolete

    //[WebMethod]
    //public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductLine(
    //  string knownCategoryValues,
    //  string category)
    //{

    //    DataTable dtProductLine;

    //    dtProductLine = new DataTable();

    //    string sql;

    //    sql = "select ";
    //    sql = sql + " id, description ";
    //    sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".product ";
    //    sql = sql + " where left(internal_code,2) = " + Utility.QStr(GlobalXyraConstants.PRODUCT_PREFIX_TRADE_ITEMS);
    //    sql = sql + " and length(internal_code) = 4";
    //    sql = sql + " order by description ";
    //    dtProductLine = Utility.CreateDataSet(sql).Tables[0];


    //    List<AjaxControlToolkit.CascadingDropDownNameValue> values =
    //        new List<AjaxControlToolkit.CascadingDropDownNameValue>();

    //    foreach (DataRow dr in dtProductLine.Rows)
    //    {
    //        values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
    //          (string)dr["description"], dr["id"].ToString()));
    //    }

    //    return values.ToArray();
    //}

    //[WebMethod]
    //public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductType(
    //  string knownCategoryValues,
    //  string category)
    //{

    //    DataTable dtProductType;
    //    string productLineId;
    //    string categories;

    //    dtProductType = new DataTable();
    //    categories = knownCategoryValues;

    //    productLineId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);

    //    string sql;

    //    sql = "select distinct ";
    //    sql = sql + " camera_type_map.code, camera_type_map.description ";
    //    sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".camera_type_map ";
    //    sql = sql + " inner join " + ConfigAppSettingsXyra.SchemaName + ".camera_model ";
    //    sql = sql + "   on (camera_type_map.code = camera_model.camera_type_code) ";
    //    sql = sql + " where ";
    //    sql = sql + " camera_model.product_line_id = " + Utility.QStr(productLineId);
    //    sql = sql + " order by description ";
    //    dtProductType = Utility.CreateDataSet(sql).Tables[0];


    //    List<AjaxControlToolkit.CascadingDropDownNameValue> values =
    //        new List<AjaxControlToolkit.CascadingDropDownNameValue>();

    //    foreach (DataRow dr in dtProductType.Rows)
    //    {
    //        values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
    //          (string)dr["description"], dr["code"].ToString()));
    //    }

    //    return values.ToArray();
    //}

    //[WebMethod]
    //public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductSubType(
    //  string knownCategoryValues,
    //  string category)
    //{

    //    DataTable dtProductSubType;
    //    string productTypeId;
    //    string categories;

    //    dtProductSubType = new DataTable();
    //    categories = knownCategoryValues;

    //    productTypeId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);

    //    string sql;

    //    sql = "select distinct ";
    //    sql = sql + " camera_sub_type.id, camera_sub_type.description ";
    //    sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".camera_sub_type ";
    //    sql = sql + " inner join " + ConfigAppSettingsXyra.SchemaName + ".camera_model ";
    //    sql = sql + "   on (camera_sub_type.id = camera_model.camera_sub_type_id) ";
    //    sql = sql + " where ";
    //    sql = sql + " camera_model.camera_type_code = " + productTypeId;
    //    sql = sql + " order by description ";
    //    dtProductSubType = Utility.CreateDataSet(sql).Tables[0];


    //    List<AjaxControlToolkit.CascadingDropDownNameValue> values =
    //        new List<AjaxControlToolkit.CascadingDropDownNameValue>();

    //    foreach (DataRow dr in dtProductSubType.Rows)
    //    {
    //        values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
    //          (string)dr["description"], dr["id"].ToString()));
    //    }

    //    return values.ToArray();
    //}

    //[WebMethod]
    //public AjaxControlToolkit.CascadingDropDownNameValue[] GetModels(
    //  string knownCategoryValues,
    //  string category)
    //{

    //    DataTable dtModels;
    //    string productSubTypeId;
    //    string categories;

    //    dtModels = new DataTable();
    //    categories = knownCategoryValues;

    //    productSubTypeId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);

    //    string sql;

    //    sql = "select ";
    //    sql = sql + " id, description ";
    //    sql = sql + " from " + ConfigAppSettingsXyra.SchemaName + ".camera_sub_type ";
    //    //sql = sql + " where ";
    //    //sql = sql + " town_id = " + Utility.QStr(townId);
    //    //sql = sql + " order by description ";
    //    dtModels = Utility.CreateDataSet(sql).Tables[0];


    //    List<AjaxControlToolkit.CascadingDropDownNameValue> values =
    //        new List<AjaxControlToolkit.CascadingDropDownNameValue>();

    //    foreach (DataRow dr in dtModels.Rows)
    //    {
    //        values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
    //          (string)dr["description"], dr["id"].ToString()));
    //    }

    //    return values.ToArray();
    //}
    #endregion


    #region Product Category - Product Line - Product Model - Product Brand
    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductCategory(
      string knownCategoryValues,
      string category)
    {

        DataTable dtProductCategory;

        dtProductCategory = new DataTable();

        string sql;

        sql = "select ";
        sql = sql + " id, ";
        sql = sql + " description ";
        sql = sql + " from ";
        sql = sql + ConfigAppSettingsXyra.SchemaName + ".product ";
        sql = sql + " where ";
        sql = sql + " left(internal_code,2) = " + Utility.QStr(ConfigAppSettingsXyra.TradeItemInternalCode);
        sql = sql + " and length(internal_code) = " + ConfigAppSettingsXyra.TradeItemCategoryLength;
        sql = sql + " order by ";
        sql = sql + " description ";

        dtProductCategory = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtProductCategory.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["description"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductLine(
      string knownCategoryValues,
      string category)
    {

        DataTable dtProductLine;
        string categoryId;
        string categories;

        dtProductLine = new DataTable();
        categories = knownCategoryValues;

        categoryId = categories.Substring(categories.LastIndexOf(":") + 1, (categories.Length - categories.LastIndexOf(":")) - 2);

        string sql;

        sql = "select ";
        sql = sql + " id, ";
        sql = sql + " description ";
        sql = sql + " from ";
        sql = sql + ConfigAppSettingsXyra.SchemaName + ".product ";
        sql = sql + " where ";
        sql = sql + " category_id = " + Utility.QStr(categoryId);
        //sql = sql + " left(internal_code,2) = " + Utility.QStr(ConfigAppSettingsXyra.TradeItemInternalCode);
        //sql = sql + " and length(internal_code) = " + ConfigAppSettingsXyra.TradeItemProductLineLength;
        sql = sql + " order by ";
        sql = sql + " description ";


        dtProductLine = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtProductLine.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["description"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductModel(
      string knownCategoryValues,
      string category)
    {

        DataTable dtProductModel;
        string productLineId;
        string productLines;

        dtProductModel = new DataTable();
        productLines = knownCategoryValues;

        productLineId = productLines.Substring(productLines.LastIndexOf(":") + 1, (productLines.Length - productLines.LastIndexOf(":")) - 2);

        string sql;

        sql = "select ";
        sql = sql + " product_model.id, ";
        sql = sql + " product_model.code, ";
        sql = sql + " product_model.description, ";
        sql = sql + " brand.description brand_description, ";
        sql = sql + " unit_of_measure.description uom_description ";

        sql = sql + " from ";
        sql = sql + ConfigAppSettingsXyra.SchemaName + ".product_model ";
        sql = sql + " inner join " + ConfigAppSettingsXyra.SchemaName + ".brand ";
        sql = sql + "  on (product_model.brand_id = brand.id) ";
        sql = sql + " inner join " + ConfigAppSettingsXyra.SchemaName + ".unit_of_measure ";
        sql = sql + "  on (product_model.unit_of_measure_id = unit_of_measure.id) ";

        sql = sql + " where ";
        sql = sql + " product_model.product_line_id = " + Utility.QStr(productLineId);

        sql = sql + " order by ";
        sql = sql + " product_model.code asc ";

        dtProductModel = Utility.CreateDataSet(sql).Tables[0];

        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();

        foreach (DataRow dr in dtProductModel.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["code"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

    //[WebMethod]
    //public AjaxControlToolkit.CascadingDropDownNameValue[] GetProductBrand(
    //  string knownCategoryValues,
    //  string category)
    //{

    //    DataTable dtProductBrand;
    //    string productModelId;
    //    string productModels;

    //    dtProductBrand = new DataTable();
    //    productModels = knownCategoryValues;

    //    productModelId = productModels.Substring(productModels.LastIndexOf(":") + 1, (productModels.Length - productModels.LastIndexOf(":")) - 2);

    //    string sql;

    //    sql = "select ";
    //    sql = sql + " product.id, ";
    //    sql = sql + " description ";
    //    sql = sql + " from ";
    //    sql = sql + ConfigAppSettingsXyra.SchemaName + ".product ";
    //    sql = sql + " where ";
    //    sql = sql + " category_id = " + Utility.QStr(categoryId);
    //    //sql = sql + " left(internal_code,2) = " + Utility.QStr(ConfigAppSettingsXyra.TradeItemInternalCode);
    //    //sql = sql + " and length(internal_code) = " + ConfigAppSettingsXyra.TradeItemProductLineLength;
    //    sql = sql + " order by ";
    //    sql = sql + " description ";


    //    dtProductBrand = Utility.CreateDataSet(sql).Tables[0];


    //    List<AjaxControlToolkit.CascadingDropDownNameValue> values =
    //        new List<AjaxControlToolkit.CascadingDropDownNameValue>();

    //    foreach (DataRow dr in dtProductBrand.Rows)
    //    {
    //        values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
    //          (string)dr["description"], dr["id"].ToString()));
    //    }

    //    return values.ToArray();
    //}
    #endregion

    [WebMethod]
    public AjaxControlToolkit.CascadingDropDownNameValue[] LoadSupplierDropDown(
      string knownCategoryValues,
      string category)
    {

        DataTable dtSupplier;

        dtSupplier = new DataTable();

        string sql;

        sql = "select ";
        sql = sql + " id, ";
        sql = sql + " name ";
        sql = sql + " from ";
        sql = sql + ConfigAppSettingsXyra.SchemaName + ".supplier ";
        //sql = sql + " where ";
        sql = sql + " order by ";
        sql = sql + " name ";

        dtSupplier = Utility.CreateDataSet(sql).Tables[0];


        List<AjaxControlToolkit.CascadingDropDownNameValue> values =
            new List<AjaxControlToolkit.CascadingDropDownNameValue>();


        //values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
        //      "---", null));

        foreach (DataRow dr in dtSupplier.Rows)
        {
            values.Add(new AjaxControlToolkit.CascadingDropDownNameValue(
              (string)dr["name"], dr["id"].ToString()));
        }

        return values.ToArray();
    }

}