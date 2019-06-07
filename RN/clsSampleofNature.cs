using System;
using System.Collections.Generic;
using System.Text;
using DataAccess;
using System.Data;
using System.Data.SqlClient;


public class clsSampleofNature
{
    public int iID;
    public string sDescription;

    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public DataTable getSampleofNature()
    {
        try
        {
            DataSet dtSampleofNature = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Id";
            arr[0].Value = iID;
            dtSampleofNature = oData.ExecuteDataset("usp_Rfsampleofnature_List", arr, CommandType.StoredProcedure);
            return dtSampleofNature.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in SampleofNature: " + eX.Message);
        }
    }

        public DataTable getSampleofNatureAll()
    {
        try
        {
            DataSet dtSampleofNature = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtSampleofNature = oData.ExecuteDataset("usp_Rfsampleofnature_List_All", arr, CommandType.StoredProcedure);
            return dtSampleofNature.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in SampleofNature: " + eX.Message);
        }
    }
    
}

