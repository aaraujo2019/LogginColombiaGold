using System;
using System.Collections.Generic;
using System.Text;
using DataAccess;
using System.Data;
using System.Data.SqlClient;


public class clsShipment
{
    public int iID;
    public string sCode;
    public string sDescription;
    public bool bStatus;
    public string sObservation;


    private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

    public DataTable getShipment()
    {
        try
        {
            DataSet dtShipment = new DataSet();
            SqlParameter[] arr = oData.GetParameters(1);
            arr[0].ParameterName = "@Code";
            arr[0].Value = sCode;
            dtShipment = oData.ExecuteDataset("usp_RfTypeShipment_List", arr, CommandType.StoredProcedure);
            return dtShipment.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in Shipment: " + eX.Message);
        }
    }

    public DataTable getShipmentAll()
    {
        try
        {
            DataSet dtShipment = new DataSet();
            SqlParameter[] arr = oData.GetParameters(0);

            dtShipment = oData.ExecuteDataset("usp_RfTypeShipment_List_All", arr, CommandType.StoredProcedure);
            return dtShipment.Tables[0];

        }
        catch (Exception eX)
        {
            throw new Exception("Error in Shipment: " + eX.Message);
        }
    }
}

