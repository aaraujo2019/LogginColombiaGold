using System;
using System.Collections.Generic;
using System.Text;
using DataAccess;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;

public class clsDHSamples
{
    public string sOpcion;
    public string sHoleID;
    public string lFrom;
    public string lTo;
    public double iFrom;
    public double iTo;
    public string sSample;
    public string sSampleType;
    public string sDupDe;
    public string sLith;
    public long iDHSampID;
    public string sComments;
    public string sVeinLocation;
    public string sVein;
    public static string sConsLoggin;
    public static string sStaticFrom;
    private ManagerDA oData = new ManagerDA();
    public string sVnMod;

    public DataTable getDHSamplesFromToValid()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(4);
            parameters[0].ParameterName = "@From";
            parameters[0].Value = this.iFrom;
            parameters[1].ParameterName = "@To";
            parameters[1].Value = this.iTo;
            parameters[2].ParameterName = "@HoleID";
            parameters[2].Value = this.sHoleID;
            parameters[3].ParameterName = "@DHSamplesID";
            parameters[3].Value = this.iDHSampID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListFromToValid", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamplesFromToValid: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesValidFromToNext()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@HoleID";
            parameters[0].Value = this.sHoleID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListValidFromToNext", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamplesFromToValidFromToNext: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamples_Litho_ListValid()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DHSamples_Litho_ListValid", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamples_Litho_ListValid: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesValid()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(4);
            parameters[0].ParameterName = "@From";
            parameters[0].Value = this.iFrom;
            parameters[1].ParameterName = "@To";
            parameters[1].Value = this.iTo;
            parameters[2].ParameterName = "@HoleID";
            parameters[2].Value = this.sHoleID;
            parameters[3].ParameterName = "@DHSamplesID";
            parameters[3].Value = this.iDHSampID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListValid", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamplesFromToValid: " + ex.Message);
        }
        return result;
    }
    public string DHSamples_DeleteLoggin()
    {
        string result;
        try
        {
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@SKDHSamples";
            parameters[0].Value = this.iDHSampID;
            object obj = this.oData.ExecuteScalar("usp_DH_Samples_Delete", parameters, CommandType.StoredProcedure);
            result = obj.ToString();
        }
        catch (Exception ex)
        {
            throw new Exception("Delete error DHSample. " + ex.Message);
        }
        return result;
    }
    public string DHSamples_AddLoggin()
    {
        string result;
        try
        {
            SqlParameter[] parameters = this.oData.GetParameters(13);
            parameters[0].ParameterName = "@Opcion";
            parameters[0].Value = this.sOpcion;
            parameters[1].ParameterName = "@HoleID";
            parameters[1].Value = this.sHoleID;
            parameters[2].ParameterName = "@Sample";
            parameters[2].Value = this.sSample;
            parameters[3].ParameterName = "@From";
            parameters[3].Value = this.iFrom;
            parameters[4].ParameterName = "@To";
            parameters[4].Value = this.iTo;
            parameters[5].ParameterName = "@SampleType";
            parameters[5].Value = this.sSampleType;
            parameters[6].ParameterName = "@DupDe";
            parameters[6].Value = this.sDupDe;
            parameters[7].ParameterName = "@Lithology";
            if (this.sLith == null)
            {
                parameters[7].Value = SqlString.Null;
            }
            else
            {
                parameters[7].Value = this.sLith;
            }
            parameters[8].ParameterName = "@Comments";
            parameters[8].Value = this.sComments;
            parameters[9].ParameterName = "@DHSamplesID";
            parameters[9].Value = this.iDHSampID;
            parameters[10].ParameterName = "@VeinLocation";
            if (this.sVeinLocation == null)
            {
                parameters[10].Value = SqlString.Null;
            }
            else
            {
                parameters[10].Value = this.sVeinLocation;
            }
            parameters[11].ParameterName = "@Vein";
            if (this.sVein == null)
            {
                parameters[11].Value = SqlString.Null;
            }
            else
            {
                parameters[11].Value = this.sVein;
            }

            parameters[12].ParameterName = "@VnMod";
            if (sVnMod == null)
            {
                parameters[12].Value = SqlString.Null;
            }
            else
            {
                parameters[12].Value = sVnMod;
            }

            object obj = this.oData.ExecuteScalar("usp_DH_Samples_InsertLoggin", parameters, CommandType.StoredProcedure);
            result = obj.ToString();
        }
        catch (Exception ex)
        {
            throw new Exception("Save error DHSamples. " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesList()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(2);
            parameters[0].ParameterName = "@Opcion";
            parameters[0].Value = this.sOpcion;
            parameters[1].ParameterName = "@HoleId";
            parameters[1].Value = this.sHoleID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in getDHSamples: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesAll()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListAll", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamplesAll: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesDistinct()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListDistinct", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamplesAll: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesId()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@sampleId";
            parameters[0].Value = this.sSample;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListSampleId", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in getDHSamplesId: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamples_getRangeValid(string _sIni, string _sFin, string _Id)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(3);
            parameters[0].ParameterName = "@SampleIni";
            parameters[0].Value = _sIni;
            parameters[1].ParameterName = "@SampleFin";
            parameters[1].Value = _sFin;
            parameters[2].ParameterName = "@Id";
            parameters[2].Value = int.Parse(_Id);
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListRange", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamples_getRangeValid: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamples()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(2);
            parameters[0].ParameterName = "@Opcion";
            parameters[0].Value = this.sOpcion;
            parameters[1].ParameterName = "@HoleId";
            parameters[1].Value = this.sHoleID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_List_Login", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamples: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesBySamp(string _sample1, string _sample2)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(3);
            parameters[0].ParameterName = "@Sample1";
            parameters[0].Value = _sample1;
            parameters[1].ParameterName = "@Sample2";
            parameters[1].Value = _sample2;
            parameters[2].ParameterName = "@HoleID";
            parameters[2].Value = this.sHoleID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListSample", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in getDHSamplesBySamp: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesById(string _Id1, string _Id2)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(3);
            parameters[0].ParameterName = "@Id1";
            parameters[0].Value = _Id1;
            parameters[1].ParameterName = "@Id2";
            parameters[1].Value = _Id2;
            parameters[2].ParameterName = "@HoleID";
            parameters[2].Value = this.sHoleID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListDHID", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in getDHSamplesBySamp: " + ex.Message);
        }
        return result;
    }
    public DataTable getDHSamplesFromTo()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(3);
            parameters[0].ParameterName = "@From";
            parameters[0].Value = this.lFrom;
            parameters[1].ParameterName = "@To";
            parameters[1].Value = this.lTo;
            parameters[2].ParameterName = "@HoleID";
            parameters[2].Value = this.sHoleID;
            dataSet = this.oData.ExecuteDataset("usp_DH_Samples_ListFromTo", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in DHSamplesFromTo: " + ex.Message);
        }
        return result;
    }
}

