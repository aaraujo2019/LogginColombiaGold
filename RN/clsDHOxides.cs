using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;

public class clsDHOxides
{
    public string sOpcion;
    public string sHoleID;
    public double dFrom;
    public double dTo;
    public double? dHem;
    public string sGt;
    public string sJar;
    public string sLim;
    public string sCuO;
    public string smmnox_Perc;
    public double? dOther;
    public double? dOtherGr;
    public string sDist;
    public string sIntensity;
    public Int64 iSKDHOxides;

    public static string sStaticFrom;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();
    public string sOxides;
    public double? dRate;

    public DataTable getDHOxides()
    {
        try
        {

            DataSet dtDH_Oxides = new DataSet();
            SqlParameter[] arr = oData.GetParameters(2);
            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            dtDH_Oxides = oData.ExecuteDataset("usp_DH_Oxides_List", arr, CommandType.StoredProcedure);
            return dtDH_Oxides.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in DH_Oxides: " + eX.Message);
        }
    }

    public string DH_Oxides_Add()
    {
        try
        {

            object oRes;
            SqlParameter[] arr = oData.GetParameters(15);

            arr[0].ParameterName = "@Opcion";
            arr[0].Value = sOpcion;
            arr[1].ParameterName = "@HoleID";
            arr[1].Value = sHoleID;
            arr[2].ParameterName = "@From";
            arr[2].Value = dFrom;
            arr[3].ParameterName = "@To";
            arr[3].Value = dTo;
            arr[4].ParameterName = "@Oxides";
            if (this.sOxides == null)
                arr[4].Value = SqlString.Null;
            else
                arr[4].Value = this.sOxides;

            arr[5].ParameterName = "@Rate";
            if (!this.dRate.HasValue)
                arr[5].Value = SqlDouble.Null;
            else
                arr[5].Value = this.dRate;

            arr[6].ParameterName = "@SKDHOxides";
            arr[6].Value = iSKDHOxides;

            oRes = oData.ExecuteScalar("usp_DH_OxidesInd_Insert", arr, CommandType.StoredProcedure);
            return oRes.ToString();


        }
        catch (Exception eX)
        {
            throw new Exception("Save error Oxidation. " + eX.Message); ;
        }
    }

    public DataTable getRfIntensityOxides_List()
    {
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfOxidationIntensity_List", parameters, CommandType.StoredProcedure);
            return dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public DataTable getRfMineralOxides_List()
    {
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfMineralOxides_List", parameters, CommandType.StoredProcedure);
            return dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    public string DH_Oxides_Delete()
    {
        try
        {
            object oRes;
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@SKDHOxides";
            arr[0].Value = iSKDHOxides;

            oRes = oData.ExecuteScalar("usp_DH_Oxides_Delete", arr, CommandType.StoredProcedure);
            return oRes.ToString();
        }
        catch (Exception eX)
        {
            throw new Exception("Delete error Oxidation. " + eX.Message); ;
        }
    }

    public DataTable getDHOxidesFromToValid()
    {
        try
        {
            DataSet dtDHOxFromToValid = new DataSet();
            SqlParameter[] arr = oData.GetParameters(4);
            arr[0].ParameterName = "@From";
            arr[0].Value = dFrom;
            arr[1].ParameterName = "@To";
            arr[1].Value = dTo;
            arr[2].ParameterName = "@HoleID";
            arr[2].Value = sHoleID;
            arr[3].ParameterName = "@SKDHOxides";
            arr[3].Value = iSKDHOxides;
            dtDHOxFromToValid = oData.ExecuteDataset("usp_DH_Oxides_ListFromToValid", arr, CommandType.StoredProcedure);
            return dtDHOxFromToValid.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in dtDHOxFromToValid: " + eX.Message);
        }
    }


}

