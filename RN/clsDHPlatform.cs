using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

    public class clsDHPlatform
    {
        public string sPlatformID;
        public string sPlatformIDHold;
        public string sRfTorelocate;
        public string sRfStatusPlatform;
        public string sRfZone;
        public string sRfSurface;
        public string sRfPriority;
        public string sRfContractor;
        public string sRfSurveryor;
        public string sRfCompanyServices;
        public string sRfLandPermit;
        public string sRfLandPermitStatus;
        public string sRfPercentProgress;
        public string sRfVeins;
        public string sCommentsHistory;
        public string sUser;
        
        public string sOpcion;
        public string sPlatform;
        public string sSection;
        public string sEastPlanned;
        public string sNothPlanned;
        public string sElevationPlanned;
        public string sAzimuthPlanned;
        public string sInclinationPlanned;
        public string sLengthPlanned;
        public string sTorelocate;
        public string sStatus;
        public string sZone;
        public string sSurface;
        public string sPriorityPlan;
        public string sCommentsPlanned;
        public string sEdit;
        public string sRfTarguet;
        public string sTarguet;
        public string sCode;

        //Adicionado el dia 29/03/2012
        public string sDepth1;
        public string sDepth2;
        public string sDepth3;
        public string sBeta1;
        public string sBeta2;
        public string sBeta3;
        public string sOrientation1;
        public string sOrientation2;
        public string sOrientation3;

        public string sHoleID;
        public string sHolePlatform;
        //Drilling
        public string sEOH;
        public string sStartDate;
        public string sFinalDate;
        public string sRigUsed;
        public string sContractor;
        public string sRodLost;
        public string sCasing;

        //Drilling Progress
        public string sFrom;
        public string sTo;
        public string sComments;
        public string sID;
        public string sDate;
        public string sRfCoreDiameter;

        //Topo
        public string sEastGPS;
        public string sNothGPS;
        public string sElevationGPS;
        public string sSurveryor;
        public string sEastST;
        public string sNothST;
        public string sElevationST;
        public string sSurveryorST;
        public string sCompanyService;
        public string sLocation;
        public string sEastCS;
        public string sNorthCS;
        public string sElevationCS;
        public string sCommentsTopo;
        public string sDateGps;
        public string sDateSt;
        public string sDateCs;

        public string sIDProject;
        public string sModule;


        //Company Drill
        public string sIdDel;
        public string sTurn;
        public string sRig;
        public string sProject;
        public string sAzimuth;
        public string sSize;
        public string sIdDc;
        public string sIdDt;
        public string sSerial;
        public string sIdCc;
        public string sIdTs;
        public string sAmount;
        public string sPercentPay;
        public string sPercentPayAdmon;
        public string sIdBA;
        public string sResTimeCont;
        public string sResTimeComp;
        public string sTimeReportDrill;
        public string sTimeApprovedInter;
        //public string sIdDT;

        public string sRfRig;
        public string sRfTurn;
        public string sRegistro;


        public string sDateini;
        public string sDatefin;
        public string sEmpresa;
        public string sMaquina;


        public string sGroup;
        public string sSubGroup;

        public string sWaterColPoint;
        public string sDateDevelopment;
        public string sDateReview;
        public string sConclutions;
        public double? dEastWC;
        public double? dNorthWC;
        public double? dElevationWC;
        public string sCoordinateSystemWC;
        /*@East	numeric(18, 3)	,
        @North	numeric(18, 3)	,
        @Elevation	numeric(18, 3),
        @CoordinateSystem	varchar(20)*/



        public string sIDPH;
        public string sIDSG;
        public string sIDQ;
        public string sOpt;
        public string sIDG;
        public string sIDI;
        public string sIDA;
        public string sIDO;
        public string sIDH;

        //Funciones......

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();
        public string DelDHDrillingTime()
        {
            try
            {
                object oRes;
                //DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sIdDel;
                oRes = oData.ExecuteScalar("usp_DH_DrillingTimeDel", arr, CommandType.StoredProcedure);
                return oRes.ToString();

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public string DelDHlostTools()
        {
            try
            {
                object oRes;
                //DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sIdDel;
                oRes = oData.ExecuteScalar("usp_DH_DrillLostTools_Del", arr, CommandType.StoredProcedure);
                return oRes.ToString();

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public string DelDHBillAdditives()
        {
            try
            {
                object oRes;
                //DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sIdDel;
                oRes = oData.ExecuteScalar("usp_DH_DrillBillableAdditives_Del", arr, CommandType.StoredProcedure);
                return oRes.ToString();

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public string DelDHHoleInProgress()
        {
            try
            {
                object oRes;
                //DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                                oRes = oData.ExecuteScalar("usp_DH_DrillProgress_Delete", arr, CommandType.StoredProcedure);
                return oRes.ToString();

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public DataTable getDHHoleStructure()
        {
            try
            {

                DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IDProyect";
                arr[0].Value = sIDProject;
                dtDHHoleInProgress = oData.ExecuteDataset("usp_DH_HOLE_Structure", arr, CommandType.StoredProcedure);
                return dtDHHoleInProgress.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Structure: " + eX.Message);
            }
        }
        public DataTable getDHHoleInProgress()
        {
            try
            {

                DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHHoleInProgress = oData.ExecuteDataset("usp_DH_Platform_Hole_InProgress_HD", arr, CommandType.StoredProcedure);
                return dtDHHoleInProgress.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public DataTable getDHDrillProgressValidacionFT()
        {
            try
            {

                DataSet dtDHDrillProgressFT = new DataSet();
                SqlParameter[] arr = oData.GetParameters(3);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@From";
                arr[1].Value = sFrom;
                arr[2].ParameterName = "@To";
                arr[2].Value = sTo;
                dtDHDrillProgressFT = oData.ExecuteDataset("usp_DrillProgress_ValidaFT", arr, CommandType.StoredProcedure);
                return dtDHDrillProgressFT.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public DataTable getDHDrillProgress()
        {
            try
            {

                DataSet dtDHDrillProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHDrillProgress = oData.ExecuteDataset("usp_DH_Drill_Progress_List", arr, CommandType.StoredProcedure);
                return dtDHDrillProgress.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public DataTable getDHMailSend()
        {
            try
            {

                DataSet dtDHMailSend = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@IDProject";
                arr[0].Value = sIDProject;
                arr[1].ParameterName = "@Module";
                arr[1].Value = sModule;
                dtDHMailSend = oData.ExecuteDataset("usp_DH_Mail_Send", arr, CommandType.StoredProcedure);
                return dtDHMailSend.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Mail Send: " + eX.Message);
            }
        }
        public DataTable getDHPlatform()
        {
            try
            {

                DataSet dtDHPlatform = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@PlatformID";
                arr[0].Value = sPlatformID;
                dtDHPlatform = oData.ExecuteDataset("usp_DH_Platform_List", arr, CommandType.StoredProcedure);
                return dtDHPlatform.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Platform: " + eX.Message);
            }
        }
        public DataTable getDHPlatformHold()
        {
            try
            {

                DataSet dtDHPlatformHold = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatformIDHold;
                dtDHPlatformHold = oData.ExecuteDataset("usp_DH_Platform_Hold_List", arr, CommandType.StoredProcedure);
                return dtDHPlatformHold.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Platform: " + eX.Message);
            }
        }
        public DataTable getRfTorelocate()
        {
            try
            {
                DataSet dtRfTorelocate = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfTorelocate;
                dtRfTorelocate = oData.ExecuteDataset("usp_RfTorelocate_List", arr, CommandType.StoredProcedure);
                return dtRfTorelocate.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfLocation()
        {
            try
            {
                DataSet dtRfLocation = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfTorelocate;
                dtRfLocation = oData.ExecuteDataset("usp_RfLocation_List", arr, CommandType.StoredProcedure);
                return dtRfLocation.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfStatusPlatform()
        {
            try
            {
                DataSet dtRfStatusPlatform = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfStatusPlatform;
                dtRfStatusPlatform = oData.ExecuteDataset("usp_RfStatusPlatform_List", arr, CommandType.StoredProcedure);
                return dtRfStatusPlatform.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
       }
        public DataTable getRfZone()
        {
            try
            {
                DataSet dtRfZone = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfZone;
                dtRfZone = oData.ExecuteDataset("usp_RfZone_List", arr, CommandType.StoredProcedure);
                return dtRfZone.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfVeins()
        {
            try
            {
                DataSet dtRfVein = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfVeins;
                dtRfVein = oData.ExecuteDataset("usp_RfVetas_List", arr, CommandType.StoredProcedure);
                return dtRfVein.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfTarguet()
        {
            try
            {
                DataSet dtRfTarguet = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfTarguet;
                dtRfTarguet = oData.ExecuteDataset("usp_RfTarget_List", arr, CommandType.StoredProcedure);
                return dtRfTarguet.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfSurface()
        {
            try
            {
                DataSet dtRfSurface = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfZone;
                dtRfSurface = oData.ExecuteDataset("usp_RfSurface_List", arr, CommandType.StoredProcedure);
                return dtRfSurface.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfPriority()
        {
            try
            {
                DataSet dtRfPriority = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Code";
                arr[0].Value = sRfPriority;
                dtRfPriority = oData.ExecuteDataset("usp_RfPriority_List", arr, CommandType.StoredProcedure);
                return dtRfPriority.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRFEnvironmentGroup()
        {
            try
            {
                DataSet RFEnvironmentGroup = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sGroup;
                RFEnvironmentGroup = oData.ExecuteDataset("usp_RF_EnvironmentGroup_List", arr, CommandType.StoredProcedure);
                return RFEnvironmentGroup.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRFEnvironmentSubGroup()
        {
            try
            {
                DataSet RFEnvironmentSubGroup = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID_G";
                arr[0].Value = sGroup;
                RFEnvironmentSubGroup = oData.ExecuteDataset("usp_RF_EnvironmentSubGroup_List", arr, CommandType.StoredProcedure);
                return RFEnvironmentSubGroup.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getRFEnvironmentQuestion()
        {
            try
            {
                DataSet RFEnvironmentQuestion = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID_SG";
                arr[0].Value = sSubGroup;
                RFEnvironmentQuestion = oData.ExecuteDataset("usp_RF_EnvironmentQuestion_List", arr, CommandType.StoredProcedure);
                return RFEnvironmentQuestion.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfContractor()
        {
            try
            {
                DataSet dtRfContractor = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sRfContractor;
                dtRfContractor = oData.ExecuteDataset("usp_RfContractor_List", arr, CommandType.StoredProcedure);
                return dtRfContractor.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfSurveryor()
        {
            try
            {
                DataSet dtRfSurveryor = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@CODE";
                arr[0].Value = sRfSurveryor;
                dtRfSurveryor = oData.ExecuteDataset("usp_RfSurveryor_List", arr, CommandType.StoredProcedure);
                return dtRfSurveryor.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfCompanyServices()
        {
            try
            {
                DataSet dtRfCompanyServices = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@CODE";
                arr[0].Value = sRfCompanyServices;
                dtRfCompanyServices = oData.ExecuteDataset("usp_RfCompanyServices_List", arr, CommandType.StoredProcedure);
                return dtRfCompanyServices.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getHolePlatform()
        {
            try
            {
                DataSet dtHolePlatform = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sHolePlatform;
                dtHolePlatform = oData.ExecuteDataset("usp_DH_Collar_Platform_List", arr, CommandType.StoredProcedure);
                return dtHolePlatform.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfLandPermit()
        {
            try
            {
                DataSet dtRfLandPermit = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@CODE";
                arr[0].Value = sRfLandPermit;
                dtRfLandPermit = oData.ExecuteDataset("usp_RfLandPermit_List", arr, CommandType.StoredProcedure);
                return dtRfLandPermit.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfLandPermitStatus()
        {
            try
            {
                DataSet dtRfLandPermitStatus = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@CODE";
                arr[0].Value = sRfLandPermitStatus;
                dtRfLandPermitStatus = oData.ExecuteDataset("usp_RfLandPermitStatus_List", arr, CommandType.StoredProcedure);
                return dtRfLandPermitStatus.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfPercentProgress()
        {
            try
            {
                DataSet dtRfPercentProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@CODE";
                arr[0].Value = sRfPercentProgress;
                dtRfPercentProgress = oData.ExecuteDataset("usp_RfPercentProgress_List", arr, CommandType.StoredProcedure);
                return dtRfPercentProgress.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDHPlatformPlanned()
        {
            try
            {

                DataSet dtDHPlatform = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatformID;
                dtDHPlatform = oData.ExecuteDataset("usp_DH_Platform_Planned", arr, CommandType.StoredProcedure);
                return dtDHPlatform.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Query: " + eX.Message);
            }
        }
        public DataTable getDHPlatformPlannedHistory()
        {
            try
            {

                DataSet dtDHPlatform = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatformID;
                dtDHPlatform = oData.ExecuteDataset("usp_DH_Platform_Planned_History", arr, CommandType.StoredProcedure);
                return dtDHPlatform.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Query: " + eX.Message);
            }
        }
        public DataTable getDHDrillProgressReport()
        {
            try
            {

                DataSet dtDHDrillProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = "";
                dtDHDrillProgress = oData.ExecuteDataset("usp_DH_DrillProgress_Report", arr, CommandType.StoredProcedure);
                return dtDHDrillProgress.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Query: " + eX.Message);
            }
        }
        public DataTable getDHDrillProgressMaxTo()
        {
            try
            {

                DataSet dtDHDrillProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHDrillProgress = oData.ExecuteDataset("usp_DH_DrillProgress_MaxTo", arr, CommandType.StoredProcedure);
                return dtDHDrillProgress.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Query: " + eX.Message);
            }
        }
        public string DH_Platform_Drilling_Upd()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(10);


                string fechai = sStartDate;
                DateTime fechaSd = new DateTime();
                fechaSd = DateTime.Parse(fechai);

                string fechaf = sFinalDate;
                DateTime fechaEd = new DateTime();
                fechaEd = DateTime.Parse(fechaf);

                
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@EOH";
                if (sEOH == "")
                    arr[2].Value = System.Data.SqlTypes.SqlInt32.Null;
                else arr[2].Value = sEOH;
                //arr[2].Value = sEOH;
                
                arr[3].ParameterName = "@StartDate";
                arr[3].Value = fechaSd;
                arr[4].ParameterName = "@FinalDate";
                arr[4].Value = fechaEd;
                arr[5].ParameterName = "@RigUsed";
                arr[5].Value = sRigUsed;
                arr[6].ParameterName = "@Contractor";
                arr[6].Value = sContractor;
                arr[7].ParameterName = "@RodLost";
                arr[7].Value = sRodLost;
                arr[8].ParameterName = "@Casing";
                arr[8].Value = sCasing;
                arr[9].ParameterName = "@Edit";
                arr[9].Value = sEdit;


                oRes = oData.ExecuteScalar("usp_DH_Platform_Drilling_Update", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Drilling. " + eX.Message);
            }
        }
        public string DH_Platform_Topo_Upd()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(17);
                
                string fechaGps = sDateGps;
                DateTime fechaGps_N = new DateTime();
                fechaGps_N = DateTime.Parse(fechaGps);

                string fechaSt = sDateSt;
                DateTime fechaSt_N = new DateTime();
                fechaSt_N = DateTime.Parse(fechaSt);

                string fechaCs = sDateCs;
                DateTime fechaCs_N = new DateTime();
                fechaCs_N = DateTime.Parse(fechaCs);

                //if (sEOH == "")
                //    arr[2].Value = System.Data.SqlTypes.SqlInt32.Null;
                //else arr[2].Value = sEOH;

                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                arr[1].ParameterName = "@EastGPS";
                if (sEastGPS =="")
                    arr[1].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                    arr[1].Value = sEastGPS;

                arr[2].ParameterName = "@NorthGPS";
                if (sNothGPS=="")
                    arr[2].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[2].Value = sNothGPS;

                arr[3].ParameterName = "@ElevationGPS";
                if(sElevationGPS=="")
                    arr[3].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[3].Value = sElevationGPS;
                arr[4].ParameterName = "@Surveryor";
                arr[4].Value = sSurveryor;
                arr[5].ParameterName = "@DateGPS";
                arr[5].Value = fechaGps_N;

                arr[6].ParameterName = "@EastST";
                if (sEastST=="")
                    arr[6].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[6].Value = sEastST;

                arr[7].ParameterName = "@NorthST";
                if (sNothST=="")
                    arr[7].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[7].Value = sNothST;

                arr[8].ParameterName = "@ElevationST";
                if (sElevationST == "")
                    arr[8].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[8].Value = sElevationST;

                arr[9].ParameterName = "@SurveryorST";
                arr[9].Value = sSurveryorST;
                arr[10].ParameterName = "@DateST";
                arr[10].Value = fechaSt_N;
                arr[11].ParameterName = "@CompanyService";
                arr[11].Value = sCompanyService;
                //arr[12].ParameterName = "@Location";
                //arr[12].Value = sLocation;


                arr[12].ParameterName = "@EastCS";
                if (sEastCS == "")
                    arr[12].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[12].Value = sEastCS;

                arr[13].ParameterName = "@NorthCS";
                if (sNorthCS == "")
                    arr[13].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[13].Value = sNorthCS;

                arr[14].ParameterName = "@ElevationCS";
                if (sElevationCS == "")
                    arr[14].Value = System.Data.SqlTypes.SqlInt32.Null;
                else
                arr[14].Value = sElevationCS;

                arr[15].ParameterName = "@DateCS";
                arr[15].Value = fechaCs_N;
                arr[16].ParameterName = "@CommentsTopo";
                arr[16].Value = sCommentsTopo;

                oRes = oData.ExecuteScalar("usp_DH_Platform_Topo_Update", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Drilling. " + eX.Message);
            }
        }
        public string DH_Platform_Planned_Add()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(25);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@Platform";
                arr[1].Value = sPlatform;
                arr[2].ParameterName = "@Section";
                arr[2].Value = sSection;
                arr[3].ParameterName = "@EastPlanned";
                arr[3].Value = sEastPlanned;
                arr[4].ParameterName = "@NorthPlanned";
                arr[4].Value = sNothPlanned;
                arr[5].ParameterName = "@ElevationPlanned";
                arr[5].Value = sElevationPlanned;
                arr[6].ParameterName = "@AzimuthPlanned";
                arr[6].Value = sAzimuthPlanned;
                arr[7].ParameterName = "@InclinationPlanned";
                arr[7].Value = sInclinationPlanned;
                arr[8].ParameterName = "@LengthPlanned";
                arr[8].Value = sLengthPlanned;
                arr[9].ParameterName = "@Torelocate";
                arr[9].Value = sTorelocate;
                arr[10].ParameterName = "@Status";
                arr[10].Value = sStatus;
                arr[11].ParameterName = "@LocationPlanned";
                arr[11].Value = sZone;
                arr[12].ParameterName = "@Surface";
                arr[12].Value = sSurface;
                arr[13].ParameterName = "@PriorityPlan";
                arr[13].Value = sPriorityPlan;
                arr[14].ParameterName = "@CommentsPlanned";
                arr[14].Value = sCommentsPlanned;
                arr[15].ParameterName = "@Edit";
                arr[15].Value = sEdit;


                //Opciones adicionadas el día 29/03/2012
                arr[16].ParameterName = "@Depth1";
                if (sDepth1 == "")
                    arr[16].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[16].Value = sDepth1;

                arr[17].ParameterName = "@Depth2";
                if (sDepth2 == "")
                    arr[17].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[17].Value = sDepth2;

                arr[18].ParameterName = "@Depth3";
                if (sDepth3 == "")
                    arr[18].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[18].Value = sDepth3;
                
                arr[19].ParameterName = "@target1";
                arr[19].Value = sBeta1;
                arr[20].ParameterName = "@target2";
                arr[20].Value = sBeta2;
                arr[21].ParameterName = "@target3";
                arr[21].Value = sBeta3;

                arr[22].ParameterName = "@Orientation1";
                arr[22].Value = sOrientation1;
                arr[23].ParameterName = "@Orientation2";
                arr[23].Value = sOrientation2;
                arr[24].ParameterName = "@Orientation3";
                arr[24].Value = sOrientation3;
                
                //arr[7].ParameterName = "@row";
                //if (sRow == null)
                //    arr[7].Value = System.Data.SqlTypes.SqlString.Null;
                //else arr[7].Value = sRow;

                oRes = oData.ExecuteScalar("usp_DH_Platform_Form_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();


            }
            catch (Exception eX)
            {
                throw new Exception("Save error Platform. " + eX.Message); ;
            }
        }
        public string DH_Platform_History_Add()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(3);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                arr[1].ParameterName = "@Comments";
                arr[1].Value = sCommentsHistory;
                arr[2].ParameterName = "@User";
                arr[2].Value = sUser;
                oRes = oData.ExecuteScalar("usp_DH_Platform_History_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error History. " + eX.Message); ;
            }
        }

        public string DH_EnvironmentPollH_Insert()
        {
            try
            {
                object oRes;

                DateTime sDate_N = new DateTime();
                sDate_N = DateTime.Parse(sDateDevelopment);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDateReview);


                SqlParameter[] arr = oData.GetParameters(11);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@ID";
                arr[1].Value = sIDPH;
                arr[2].ParameterName = "@Platform";
                arr[2].Value = sPlatform;
                arr[3].ParameterName = "@WaterColPoint";
                arr[3].Value = sWaterColPoint;
                arr[4].ParameterName = "@DateDevelopment";
                arr[4].Value = sDate_N;
                arr[5].ParameterName = "@DateReview";
                arr[5].Value = sDate_F;
                arr[6].ParameterName = "@Conclusions";
                arr[6].Value = sConclutions;


                //arr[5].ParameterName = "@Stand";
                //if (iStand == null)
                //    arr[5].Value = System.Data.SqlTypes.SqlInt32.Null;
                //else arr[5].Value = iStand;

                arr[7].ParameterName = "@East";
                if (dEastWC == null)
                    arr[7].Value = System.Data.SqlTypes.SqlDouble.Null;
                else
                    arr[7].Value = dEastWC;

                arr[8].ParameterName = "@North";
                if (dNorthWC == null)
                    arr[8].Value = System.Data.SqlTypes.SqlDouble.Null;
                else
                    arr[8].Value = dNorthWC;

                arr[9].ParameterName = "@Elevation";
                if (dElevationWC == null)
                    arr[9].Value = System.Data.SqlTypes.SqlDouble.Null;
                else
                    arr[9].Value = dElevationWC;

                arr[10].ParameterName = "@CoordinateSystem";
                if (sCoordinateSystemWC == null)
                    arr[10].Value = System.Data.SqlTypes.SqlString.Null;
                else
                    arr[10].Value = sCoordinateSystemWC;


                oRes = oData.ExecuteScalar("usp_DH_EnvironmentPollH_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Poll Header. " + eX.Message); ;
            }
        }
        public string DH_EnvironmentPollC_Insert()
        {
            try
            {
                object oRes;

                //DateTime sDate_N = new DateTime();
                //sDate_N = DateTime.Parse(sDateDevelopment);

                //DateTime sDate_F = new DateTime();
                //sDate_F = DateTime.Parse(sDateReview);


                SqlParameter[] arr = oData.GetParameters(7);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@ID";
                arr[1].Value = sID;
                arr[2].ParameterName = "@IDPH";
                arr[2].Value = sIDPH;
                arr[3].ParameterName = "@IDSG";
                arr[3].Value = sIDSG;
                arr[4].ParameterName = "@IDQ";
                arr[4].Value = sIDQ;
                arr[5].ParameterName = "@Opt";
                arr[5].Value = sOpt;
                arr[6].ParameterName = "@Comments";
                arr[6].Value = sConclutions;
                oRes = oData.ExecuteScalar("usp_DH_EnvironmentPollC_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Poll Header. " + eX.Message); ;
            }
        }
        public string DH_Collar_Platform_Add()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@Platform";
                arr[1].Value = sPlatform;
                oRes = oData.ExecuteScalar("usp_DH_Collar_Platform_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Collar. " + eX.Message); ;
            }
        }
        public string DH_Drill_Progress_Add()
        {
            try
            {
                object oRes;
                DateTime sDate_N = new DateTime();
                sDate_N = DateTime.Parse(sDate);

                SqlParameter[] arr = oData.GetParameters(8);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@From";
                arr[2].Value = sFrom;
                arr[3].ParameterName = "@To";
                arr[3].Value = sTo;
                arr[4].ParameterName = "@Comments";
                arr[4].Value = sComments;
                arr[5].ParameterName = "@ID";
                arr[5].Value = sID;
                arr[6].ParameterName = "@Date";
                arr[6].Value = sDate_N;
                arr[7].ParameterName = "@CoreID";
                arr[7].Value = sRfCoreDiameter;
                oRes = oData.ExecuteScalar("usp_DH_Drill_Progress_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Drill Progress. " + eX.Message); ;
            }
        }

        // Company Drill

        public string DH_Company_Drill_Add()
        {
            try
            {
                object oRes;
                DateTime sDate_N = new DateTime();
                sDate_N = DateTime.Parse(sDate);

                SqlParameter[] arr = oData.GetParameters(7);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IdDc";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@Date";
                arr[2].Value = sDate_N;
                arr[3].ParameterName = "@Turn";
                arr[3].Value = sTurn;
                arr[4].ParameterName = "@Rig";
                arr[4].Value = sRig;
                arr[5].ParameterName = "@Project";
                arr[5].Value = sProject;
                arr[6].ParameterName = "@Comments";
                arr[6].Value = sComments;
                //arr[7].ParameterName = "@ID";
                //arr[7].Value = sComments;
                oRes = oData.ExecuteScalar("usp_DH_DrillCompany_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Company Drill. " + eX.Message); ;
            }
        }
        public DataTable getRfRig()
        {
            try
            {
                DataSet dtRfRig = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sRfRig;
                dtRfRig = oData.ExecuteDataset("usp_RfRig_List", arr, CommandType.StoredProcedure);
                return dtRfRig.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfCoreDiameter()
        {
            try
            {
                DataSet dtRfCore = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sRfCoreDiameter;
                dtRfCore = oData.ExecuteDataset("usp_RfCoreDiameter_List", arr, CommandType.StoredProcedure);
                return dtRfCore.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getRfTurn()
        {
            try
            {
                DataSet dtRfTurn = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sRfTurn;
                dtRfTurn = oData.ExecuteDataset("usp_RfTurn_List", arr, CommandType.StoredProcedure);
                return dtRfTurn.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRegistroListt()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sRegistro;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillCompany_List", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRegistroListtID()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sRegistro;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillCompany_List_ID", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string DH_DrillMeterTurn_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(9);
                arr[0].ParameterName = "@HoleId";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@Azimuth";
                arr[2].Value = sAzimuth;
                arr[3].ParameterName = "@Size";
                arr[3].Value = sSize;
                arr[4].ParameterName = "@From";
                arr[4].Value = sFrom;
                arr[5].ParameterName = "@To";
                arr[5].Value = sTo;
                arr[6].ParameterName = "@Comments";
                arr[6].Value = sComments;
                arr[7].ParameterName = "@Opcion";
                arr[7].Value = sOpcion;
                arr[8].ParameterName = "@ID";
                arr[8].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_DrillMeterTurn_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Drill Meter. " + eX.Message); ;
            }
        }
        public DataTable getDrillMeterTurn_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillMeterTurn_IdDc", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getRfDrillDownTCD()
        {
            try
            {

                DataSet dtRfDrillDown = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRfDrillDown = oData.ExecuteDataset("usp_RfDownTime_List", arr, CommandType.StoredProcedure);
                return dtRfDrillDown.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Drill Down T: " + eX.Message);
            }
        }
        public DataTable getRfChangeCrownCD()
        {
            try
            {

                DataSet dtRfChangeCrown = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRfChangeCrown = oData.ExecuteDataset("usp_RfChangeCrown_List", arr, CommandType.StoredProcedure);
                return dtRfChangeCrown.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Change Crown : " + eX.Message);
            }
        }
        public DataTable getRfTurnSuppliesCD()
        {
            try
            {

                DataSet dtRfTurnSupplies = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRfTurnSupplies = oData.ExecuteDataset("usp_RfTurnSupplies_List", arr, CommandType.StoredProcedure);
                return dtRfTurnSupplies.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Turn Supplies: " + eX.Message);
            }
        }
        public DataTable getRfBiabilityCD()
        {
            try
            {

                DataSet dtRfBiability = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRfBiability = oData.ExecuteDataset("usp_RfBiability_List", arr, CommandType.StoredProcedure);
                return dtRfBiability.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Biability : " + eX.Message);
            }
        }

        public DataTable getRfLostToolsCD()
        {
            try
            {

                DataSet dtRfLostTools = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRfLostTools = oData.ExecuteDataset("usp_RfLostTools_List", arr, CommandType.StoredProcedure);
                return dtRfLostTools.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Lost Tools : " + eX.Message);
            }
        }

        public DataTable getRfBillableAdditivesCD()
        {
            try
            {

                DataSet dtRfBillableAdditives = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRfBillableAdditives = oData.ExecuteDataset("usp_RfBillableAdditives_List", arr, CommandType.StoredProcedure);
                return dtRfBillableAdditives.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Billable Additives : " + eX.Message);
            }
        }

        public string DH_DrillDownTime_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(7);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IdDt";
                arr[2].Value = sIdDt;
                arr[3].ParameterName = "@From";
                arr[3].Value = sFrom;
                arr[4].ParameterName = "@To";
                arr[4].Value = sTo;
                arr[5].ParameterName = "@Comments";
                arr[5].Value = sComments;
                arr[6].ParameterName = "@ID";
                arr[6].Value = sID;
                oRes = oData.ExecuteScalar("usp_DrillDownTime_Add", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Down Time. " + eX.Message); ;
            }
        }

        public DataTable getDrillDownTime_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillDownTime_IdDc", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getDrillChangeCrown_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillChangeCrown_IdDc", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }



        public string DH_DrillChageCrown_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(8);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IdCc";
                arr[2].Value = sIdCc;
                arr[3].ParameterName = "@Serial";
                arr[3].Value = sSerial;
                arr[4].ParameterName = "@From";
                arr[4].Value = sFrom;
                arr[5].ParameterName = "@To";
                arr[5].Value = sTo;
                arr[6].ParameterName = "@Comments";
                arr[6].Value = sComments;
                arr[7].ParameterName = "@ID";
                arr[7].Value = sID;
                oRes = oData.ExecuteScalar("usp_DrillChangeCrown_Add", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Change Crown. " + eX.Message); ;
            }
        }

        public string DH_DrillTurnSupplies_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(6);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IdTs";
                arr[2].Value = sIdTs;
                arr[3].ParameterName = "@Amount";
                arr[3].Value = sAmount;
                arr[4].ParameterName = "@Comments";
                arr[4].Value = sComments;
                arr[5].ParameterName = "@ID";
                arr[5].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_Drill_TurnSupplies_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Change Crown. " + eX.Message); ;
            }
        }
        public DataTable getTurnSupplies_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_Drill_TurnSupplies_ID", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string DH_BiabilityOfTimeCom_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IDBiTiCom";
                arr[2].Value = sIdTs;
                arr[3].ParameterName = "@ID";
                arr[3].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_Drill_BiabilityOfTimeCom_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Biability of Time. " + eX.Message); ;
            }
        }
        public DataTable getDH_BiabilityOfTimeCom_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_Drill_BiabilityOfTimeCom_ID", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string DH_BiabilityOfTimeCon_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IDBiTiCon";
                arr[2].Value = sIdTs;
                arr[3].ParameterName = "@ID";
                arr[3].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_Drill_BiabilityOfTimeCon_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Biability of Time. " + eX.Message); ;
            }
        }
        public DataTable getDH_BiabilityOfTimeCon_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_Drill_BiabilityOfTimeCon_ID", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string DH_LostTools_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(8);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IDLt";
                arr[2].Value = sIdTs;
                arr[3].ParameterName = "@Amount";
                arr[3].Value = sAmount;
                arr[4].ParameterName = "@PercentPay";
                arr[4].Value = sPercentPay;
                arr[5].ParameterName = "@PercentPayAdmon";
                arr[5].Value = sPercentPayAdmon;
                arr[6].ParameterName = "@Comments";
                arr[6].Value = sComments;
                arr[7].ParameterName = "@ID";
                arr[7].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_DrillLostTools_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Biability of Time. " + eX.Message); ;
            }
        }
        public DataTable getDH_LostTools_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_Drill_LostTools_ID", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string DH_BillableAdditives_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(5);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IDBA";
                arr[2].Value = sIdBA;
                arr[3].ParameterName = "@Amount";
                arr[3].Value = sAmount;
                arr[4].ParameterName = "@ID";
                arr[4].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_BillableAdditives_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Biability Additives. " + eX.Message); ;
            }
        }
        public DataTable getDH_BillableAdditives_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillBillableAdditives_ID", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string DH_DrillingTime_Add()
        {
            try
            {
                object oRes;



                SqlParameter[] arr = oData.GetParameters(8);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDDC";
                arr[1].Value = sIdDc;
                arr[2].ParameterName = "@IDDT";
                arr[2].Value = sIdDt;
                arr[3].ParameterName = "@ResTimeCont";
                arr[3].Value = sResTimeCont;
                arr[4].ParameterName = "@ResTimeComp";
                arr[4].Value = sResTimeComp;
                arr[5].ParameterName = "@TimeReportDrill";
                arr[5].Value = sTimeReportDrill;
                arr[6].ParameterName = "@TimeApprovedInter";
                arr[6].Value = sTimeApprovedInter;
                arr[7].ParameterName = "@ID";
                arr[7].Value = sID;
                oRes = oData.ExecuteScalar("usp_DH_DrillDrillingTime_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Drilling Time " + eX.Message); ;
            }
        }
        public DataTable getDH_DrillingTime_IdDc()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@IdDc";
                arr[0].Value = sIdDc;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillDrillingTime_IdDc", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_DrillingReport()
        { 
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);
                
                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_DrillingReport", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getDH_DrillingReportContractor()
        {
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_DrillingReportContractor", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getDH_DrillingReportProdGral()
        {
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_DrillingReportProdGral", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_EnvironmentReportProdGral()
        {
            try
            {
                //DateTime sDate_I = new DateTime();
                //sDate_I = DateTime.Parse(sDateini);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                arr[1].ParameterName = "@Year";
                arr[1].Value = sDateini;

                dtRegistro = oData.ExecuteDataset("usp_DH_Environment_Poll", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_EnvironmentReportProdGralGroup()
        {
            try
            {
                //DateTime sDate_I = new DateTime();
                //sDate_I = DateTime.Parse(sDateini);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                arr[1].ParameterName = "@Year";
                arr[1].Value = sDateini;

                dtRegistro = oData.ExecuteDataset("usp_DH_Environment_Poll_G", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_EnvironmentReportProdGralImpact()
        {
            try
            {
                //DateTime sDate_I = new DateTime();
                //sDate_I = DateTime.Parse(sDateini);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                arr[1].ParameterName = "@Year";
                arr[1].Value = sDateini;

                dtRegistro = oData.ExecuteDataset("usp_DH_Environment_Poll_Impact", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_DrillingReportRig()
        {
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillingReportRig", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getDH_DrillingReportCompany()
        {
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_Drilling_DrillReportContractor", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_DrillingReportProduccion()
        {
            try
            {


                string Fechai;
                string Fechaf;

                string sFechai = sDateini;
                string sFechaf = sDatefin;

                //DateTime sDate_I = new DateTime();
                //sDate_I = DateTime.Parse(sDateini);

                //DateTime sDate_F = new DateTime();
                //sDate_F = DateTime.Parse(sDatefin);


                Fechai = sFechai.Substring(0, 10);
                Fechaf = sFechaf.Substring(0, 10);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = string.Format("{0:MM/dd/yyyy}", Fechai);
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = string.Format("{0:MM/dd/yyyy}", Fechaf);
                
                //arr[2].ParameterName = "@Empresa";
                //arr[2].Value = sEmpresa;
                //arr[3].ParameterName = "@Maquina";
                //arr[3].Value = sMaquina;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillReport_Prod", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_DrillingReportEmp_Maq()
        {
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_DrillingReport_Emp_Maq", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataSet getDH_DrillingTime()
        {
            try
            {
                DateTime sDate_I = new DateTime();
                sDate_I = DateTime.Parse(sDateini);

                DateTime sDate_F = new DateTime();
                sDate_F = DateTime.Parse(sDatefin);

                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@DATEINI";
                arr[0].Value = sDate_I;
                arr[1].ParameterName = "@DATEFIN";
                arr[1].Value = sDate_F;
                dtRegistro = oData.ExecuteDataset("usp_DH_DrillingReportDaily", arr, CommandType.StoredProcedure);
                return dtRegistro;
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }

        public DataTable getDH_DrillingTime_Depth()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtRegistro = oData.ExecuteDataset("usp_DH_Platform_Depth", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_Environment_Poll_Platform()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@Platform";
                arr[0].Value = sPlatform;
                dtRegistro = oData.ExecuteDataset("usp_DH_Environment_Poll_Platform", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }

        }
        public DataTable getDH_Environment_Impact()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@ID";
                arr[0].Value = sID;
                dtRegistro = oData.ExecuteDataset("usp_DH_EnvironmentImpact_List", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }

        }
        public DataTable getDH_Environment_Poll_Select()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(3);
                arr[0].ParameterName = "@IDPH";
                arr[0].Value = sIDPH;
                arr[1].ParameterName = "@IDG";
                arr[1].Value = sIDG;
                arr[2].ParameterName = "@IDSG";
                arr[2].Value = sIDSG;
                dtRegistro = oData.ExecuteDataset("usp_DH_EnvironmentPoll_Select", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public DataTable getDH_Environment_Poll_Impact_Select()
        {
            try
            {
                DataSet dtRegistro = new DataSet();
                SqlParameter[] arr = oData.GetParameters(2);
                arr[0].ParameterName = "@IDI";
                arr[0].Value = sIDI;
                arr[1].ParameterName = "@IDH";
                arr[1].Value = sIDH;
                dtRegistro = oData.ExecuteDataset("usp_DH_EnvironmentPollImpact_Select", arr, CommandType.StoredProcedure);
                return dtRegistro.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
        public string getDH_Environment_Poll_Impact_Add()
        {
            try
            {
                object oRes;

                SqlParameter[] arr = oData.GetParameters(4);
                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@IDI";
                arr[1].Value = sIDI;
                arr[2].ParameterName = "@IDO";
                arr[2].Value = sIDO;
                arr[3].ParameterName = "@IDH";
                arr[3].Value = sIDPH;

                oRes = oData.ExecuteScalar("usp_DH_EnvironmentPollImpact_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();

             }
            catch (Exception eX)
            {
                throw new Exception("Error : " + eX.Message);
            }
        }
    }