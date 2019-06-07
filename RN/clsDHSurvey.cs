using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


    public class clsDHSurvey
    {
        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();

        public string sHoleID;
        public string sDepth;
        public string sAz;
        public string sDip;
        public string sMeaseredBy;
        public string sInstrument;
        public string sMethod;
        public string sTemp;
        public string sMagField;
        public string sGravFieald;
        public string sObservation;
        public string sDate;
        public string sInDate;
        public string sOpcion;

        public string DH_Survey_Add()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(13);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                arr[1].ParameterName = "@Depth";
                arr[1].Value = sDepth;
                arr[2].ParameterName = "@Az";
                arr[2].Value = sAz;
                arr[3].ParameterName = "@Dip";
                arr[3].Value = sDip;
                arr[4].ParameterName = "@MeaseredBy";
                arr[4].Value = sMeaseredBy;
                arr[5].ParameterName = "@Instrument";
                arr[5].Value = sInstrument;
                arr[6].ParameterName = "@Method";
                arr[6].Value = sMethod;
                arr[7].ParameterName = "@Temp";
                if (sTemp == "")
                        arr[7].Value = System.Data.SqlTypes.SqlInt32.Null;
                else    arr[7].Value = sTemp;
                arr[8].ParameterName = "@MagField";
                if (sMagField == "")
                        arr[8].Value = System.Data.SqlTypes.SqlInt32.Null;
                else    arr[8].Value = sMagField;
                arr[9].ParameterName = "@GravFieald";
                if (sGravFieald == "")
                        arr[9].Value = System.Data.SqlTypes.SqlInt32.Null;
                else    arr[9].Value = sGravFieald;
                arr[10].ParameterName = "@Observation";
                arr[10].Value = sObservation;


                string fechai = sDate;
                DateTime fechaSd = new DateTime();
                fechaSd = DateTime.Parse(fechai);

                arr[11].ParameterName = "@Date";
                //if (fechaSd == "")
                //    arr[11].Value = System.Data.SqlTypes.SqlInt32.Null;
                //else 
                arr[11].Value = fechaSd;
                //arr[12].ParameterName = "@InDate";
                //arr[12].Value = sInDate;
                arr[12].ParameterName = "@Opcion";
                arr[12].Value = sOpcion;
                oRes = oData.ExecuteScalar("usp_DH_Surveys_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Save error Survey " + eX.Message); ;
            }
        }
        public string DHSurveyDel()
        {
            try
            {
                object oRes;
                //DataSet dtDHHoleInProgress = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                oRes = oData.ExecuteScalar("usp_DH_Survey_Del", arr, CommandType.StoredProcedure);
                return oRes.ToString();

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Hole: " + eX.Message);
            }
        }
        public DataTable getDHSurvey_ID()
        {
            try
            {

                DataSet dtDHSurvey = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHSurvey = oData.ExecuteDataset("usp_DH_Survey_ID", arr, CommandType.StoredProcedure);
                return dtDHSurvey.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in Query: " + eX.Message);
            }
        }

    }

