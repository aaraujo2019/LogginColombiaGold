using DataAccess;
using System;
using System.Data;
using System.Data.SqlClient;


public class clsRf
{
    private ManagerDA oData = new ManagerDA();
    public static string sUser;
    public static string sIdentification;
    public static string sIdGrupo;
    public static DataSet dsPermisos;
    public int iIdProject;
    public string sCodeLith;
    public DataTable getCollarsPlatf(string _sHoleId)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@HoleId";
            parameters[0].Value = _sHoleId;
            dataSet = this.oData.ExecuteDataset("usp_DH_Platform_ListReport", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getTarget(string _sCode)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@CODE";
            parameters[0].Value = _sCode;
            dataSet = this.oData.ExecuteDataset("usp_RfTarget_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getLocation(string _sCode)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@CODE";
            parameters[0].Value = _sCode;
            dataSet = this.oData.ExecuteDataset("usp_RfLocation_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getVersionProject()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@Id";
            parameters[0].Value = this.iIdProject;
            dataSet = this.oData.ExecuteDataset("[usp_getProject]", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfTextures_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@Code";
            parameters[0].Value = this.sCodeLith;
            dataSet = this.oData.ExecuteDataset("usp_RfTextures_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRFGsize_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@Code";
            parameters[0].Value = this.sCodeLith;
            dataSet = this.oData.ExecuteDataset("usp_RFGsize_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public string InsertTrans(string _sTABLE, string _TRANS, string _LOGINTRANS, string _DATOSTRANS)
    {
        string result;
        try
        {
            SqlParameter[] parameters = this.oData.GetParameters(4);
            parameters[0].ParameterName = "@sTABLE";
            parameters[0].Value = _sTABLE;
            parameters[1].ParameterName = "@TRANS";
            parameters[1].Value = _TRANS;
            parameters[2].ParameterName = "@LOGINTRANS";
            parameters[2].Value = _LOGINTRANS;
            parameters[3].ParameterName = "@DATOSTRANS";
            parameters[3].Value = _DATOSTRANS;
            object obj = this.oData.ExecuteScalar("[usp_TransactionsAdd]", parameters, CommandType.StoredProcedure);
            result = obj.ToString();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public string UpdatePass(string _sPassOld, string _sPass, string _sLogin)
    {
        string result;
        try
        {
            SqlParameter[] parameters = this.oData.GetParameters(3);
            parameters[0].ParameterName = "@PasswdOld";
            parameters[0].Value = _sPassOld;
            parameters[1].ParameterName = "@PasswdNew";
            parameters[1].Value = _sPass;
            parameters[2].ParameterName = "@LoginUser";
            parameters[2].Value = _sLogin;
            object obj = this.oData.ExecuteScalar("[usp_saveUserPasswd]", parameters, CommandType.StoredProcedure);
            result = obj.ToString();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getTransList(string _sUser)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@LOGINTRANS";
            parameters[0].Value = _sUser;
            dataSet = this.oData.ExecuteDataset("usp_TransactionsList", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfWorkerCred(string _sCod, string _sPass)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(2);
            parameters[0].ParameterName = "@Cod";
            parameters[0].Value = _sCod;
            parameters[1].ParameterName = "@Password";
            parameters[1].Value = _sPass;
            dataSet = this.oData.ExecuteDataset("usp_RfWorker_ListByCred", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getInvSamples()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_Inv_Samples_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfPrefixW_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfPrefixW_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfTypeStructure_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfTypeStructure_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfGSize_ListMin(string _sOpcion)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@Opcion";
            parameters[0].Value = _sOpcion;
            dataSet = this.oData.ExecuteDataset("usp_RFGsize_ListAll", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfFillStructure_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfFillStructure_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfStyleAlt_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfStyleAlt_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfMinerAlt_ListMin(string _sType)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@TypeAlt";
            parameters[0].Value = _sType;
            dataSet = this.oData.ExecuteDataset("usp_RfMinerAlt_ListTypeAlt", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfMinerAlt_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfMinerAlt_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfIntensityAlt_List(string _sProjectGC)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@Project";
            parameters[0].Value = _sProjectGC;
            dataSet = this.oData.ExecuteDataset("usp_RfIntensityAlt_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfTypeAlt_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfTypeAlt_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfMinerPercent_List(string _sProjectGC)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@Project";
            parameters[0].Value = _sProjectGC;
            dataSet = this.oData.ExecuteDataset("usp_RfMinPercent_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfMinerMinSt_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfMinerStyle_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfMinerMin_ListOxid()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfMinerMin_ListOxid", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfMinerMin_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfMinerMin_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfColour_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfColour_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfOxidation_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfOxidation_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfOxides_List()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfOxides_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getWeathering()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfWeathering_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfGeotechHardness()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfGeotechHardness_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfGeotechBreak()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfGeotechBreak_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfPrefix()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_Prefix_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfWorker()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfWorker_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfCodeLab()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfCodeLab_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfPreparationCode()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfPreparationCode_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfAnalysisCodeLab()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfAnalysisCodeLab_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getUsersPortal(string _sLogin)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@sLogin";
            parameters[0].Value = _sLogin;
            dataSet = this.oData.ExecuteDataset("usp_getUsersSubpartners_PORTAL", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfTypeSample()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DH_RfTypeSample_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataSet getDsRfLithology()
    {
        DataSet result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DH_RfLithology_List", parameters, CommandType.StoredProcedure);
            result = dataSet;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfLithology()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DH_RfLithology_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getRfLithologyDH()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_DH_RfLithology_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[1];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataSet getFormsByGrupoAll(string _IdGrupo)
    {
        DataSet result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@idGrupo";
            parameters[0].Value = _IdGrupo;
            dataSet = this.oData.ExecuteDataset("usp_getFormByGrupoAll_PORTAL", parameters, CommandType.StoredProcedure);
            result = dataSet;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getFormsByGrupo(string _sIdGrupo, string _sIDGrupo)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(2);
            parameters[0].ParameterName = "@idGrupo";
            parameters[0].Value = _sIdGrupo;
            parameters[1].ParameterName = "@ID_Project";
            parameters[1].Value = _sIDGrupo;
            dataSet = this.oData.ExecuteDataset("usp_getFormulariosByGrupo", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getPermisosFormsByGrupo(string _IdGrupo)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@idGrupo";
            parameters[0].Value = _IdGrupo;
            dataSet = this.oData.ExecuteDataset("usp_getPermisosFormByGrupo_PORTAL", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }
    public DataTable getUsuarios(string _IdUser)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@sUsuario";
            parameters[0].Value = _IdUser;
            dataSet = this.oData.ExecuteDataset("usp_getUsuarios_PORTAL", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }


    public DataTable getRfStage()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfStage_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }

    public DataTable getTypeInfill()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfTypeInfill_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }

    public DataTable getRfMineral(int opcion)
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(1);
            parameters[0].ParameterName = "@pOpcion";
            parameters[0].Value = opcion;
            dataSet = this.oData.ExecuteDataset("usp_RfMineral_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }

    public DataTable getRfTextureInfill()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = this.oData.GetParameters(0);
            dataSet = this.oData.ExecuteDataset("usp_RfTextureInfill_List", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        return result;
    }

    public DataTable getRfVeinsCodes()
    {
        DataTable result;
        try
        {
            DataSet dataSet = new DataSet();
            SqlParameter[] parameters = oData.GetParameters(0);
            dataSet = oData.ExecuteDataset("[dbo].[usp_getRfVeinsCodes]", parameters, CommandType.StoredProcedure);
            result = dataSet.Tables[0];
        }
        catch (Exception ex)
        {
            throw new Exception("Error in dtRfVeinsCodesAll: " + ex.Message);
        }
        return result;
    }
}

