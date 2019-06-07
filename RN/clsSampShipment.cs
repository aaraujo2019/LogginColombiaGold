using System;
using System.Collections.Generic;
using System.Text;
using DataAccess;
using System.Data;
using System.Data.SqlClient;


public class clsSampShipment
{
    public int iID;
    public string sCode;
    public string sDescription;
    public bool bStatus;
    public string sObservation;
    public long lCost;


    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public DataTable getSampShipment()
    {
        try
        {
            DataSet dtSampShipment = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Code";
            arr[0].Value = sCode;
            dtSampShipment = oData.ExecuteDataset("usp_RfTypeSampShipment_List", arr, CommandType.StoredProcedure);
            return dtSampShipment.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in SampShipment: " + eX.Message);
        }
    }

    public DataTable getSampShipmentAll()
    {
        try
        {
            DataSet dtSampShipment = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);
            dtSampShipment = oData.ExecuteDataset("usp_RfTypeSampShipment_List_All", arr, CommandType.StoredProcedure);
            return dtSampShipment.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in SampShipment: " + eX.Message);
        }
    }
}

