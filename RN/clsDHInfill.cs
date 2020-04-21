using System;
using System.Data;
using System.Data.SqlClient;

namespace RN
{
    public class clsDHInfill
    {
        #region Variables
        public string sOpcion;
        public string sHoleID;
        public double dFrom;
        public double dTo;
        public int iDHInfillID;
        public int cuentaRegistros;

        public int ? Infill1Stage;
        public string Infill1Type;
        public double ? Infill1Number;
        public double ? Infill1Angle;
        public double ? Infill1StagePerc;

        public string Infill1MineralGange1;
        public string Infill1MineralGange1Texture;
        public double ? Infill1MineralGange1Perc;

        public string Infill1MineralGange2;
        public string Infill1MineralGange2Texture;
        public double ? Infill1MineralGange2Perc;

        public string Infill1MineralGange3;
        public string Infill1MineralGange3Texture;
        public double ? Infill1MineralGange3Perc;

        public string Infill1OreMineral1;
        public string Infill1OreMineral1Style;
        public double ? Infill1OreMineral1Perc;

        public string Infill1OreMineral2;
        public string Infill1OreMineral2Style;
        public double ? Infill1OreMineral2Perc;

        public string Infill1OreMineral3;
        public string Infill1OreMineral3Style;
        public double ? Infill1OreMineral3Perc;



        public int ? Infill2Stage;
        public string Infill2Type;
        public double ? Infill2Number;
        public double ? Infill2Angle;
        public double ? Infill2StagePerc;

        public string Infill2MineralGange1;
        public string Infill2MineralGange1Texture;
        public double ? Infill2MineralGange1Perc;

        public string Infill2MineralGange2;
        public string Infill2MineralGange2Texture;
        public double ? Infill2MineralGange2Perc;

        public string Infill2MineralGange3;
        public string Infill2MineralGange3Texture;
        public double ? Infill2MineralGange3Perc;

        public string Infill2OreMineral1;
        public string Infill2OreMineral1Style;
        public double ? Infill2OreMineral1Perc;

        public string Infill2OreMineral2;
        public string Infill2OreMineral2Style;
        public double ? Infill2OreMineral2Perc;

        public string Infill2OreMineral3;
        public string Infill2OreMineral3Style;
        public double ? Infill2OreMineral3Perc;



        public int ? Infill3Stage;
        public string Infill3Type;
        public double ? Infill3Number;
        public double ? Infill3Angle;
        public double ? Infill3StagePerc;

        public string Infill3MineralGange1;
        public string Infill3MineralGange1Texture;
        public double ? Infill3MineralGange1Perc;

        public string Infill3MineralGange2;
        public string Infill3MineralGange2Texture;
        public double ? Infill3MineralGange2Perc;

        public string Infill3MineralGange3;
        public string Infill3MineralGange3Texture;
        public double ? Infill3MineralGange3Perc;

        public string Infill3OreMineral1;
        public string Infill3OreMineral1Style;
        public double ? Infill3OreMineral1Perc;

        public string Infill3OreMineral2;
        public string Infill3OreMineral2Style;
        public double ? Infill3OreMineral2Perc;

        public string Infill3OreMineral3;
        public string Infill3OreMineral3Style;
        public double ? Infill3OreMineral3Perc;



        public int ? Infill4Stage;
        public string Infill4Type;
        public double ? Infill4Number;
        public double ? Infill4Angle;
        public double ? Infill4StagePerc;

        public string Infill4MineralGange1;
        public string Infill4MineralGange1Texture;
        public double ? Infill4MineralGange1Perc;

        public string Infill4MineralGange2;
        public string Infill4MineralGange2Texture;
        public double ? Infill4MineralGange2Perc;

        public string Infill4MineralGange3;
        public string Infill4MineralGange3Texture;
        public double ? Infill4MineralGange3Perc;

        public string Infill4OreMineral1;
        public string Infill4OreMineral1Style;
        public double ? Infill4OreMineral1Perc;

        public string Infill4OreMineral2;
        public string Infill4OreMineral2Style;
        public double ? Infill4OreMineral2Perc;

        public string Infill4OreMineral3;
        public string Infill4OreMineral3Style;
        public double ? Infill4OreMineral3Perc;

        public string wHoleInfill;

        private DataAccess.ManagerDA oData = new DataAccess.ManagerDA();
        #endregion

        public string DH_Infill_Add()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(97);

                arr[0].ParameterName = "@Opcion";
                arr[0].Value = sOpcion;
                arr[1].ParameterName = "@HoleID";
                arr[1].Value = sHoleID;
                arr[2].ParameterName = "@From";
                arr[2].Value = dFrom;
                arr[3].ParameterName = "@To";
                arr[3].Value = dTo;

                arr[4].ParameterName = "@Infill1Stage";
                if (Infill1Stage == null)
                    arr[4].Value = System.Data.SqlTypes.SqlInt32.Null;
                else arr[4].Value = Infill1Stage;

                arr[5].ParameterName = "@Infill1Type";
                if (Infill1Type == null)
                    arr[5].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[5].Value = Infill1Type;
                
                arr[6].ParameterName = "@Infill1Number";
                if (Infill1Number == null)
                    arr[6].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[6].Value = Infill1Number;

                arr[7].ParameterName = "@Infill1AngleToAxis";
                if (Infill1Angle == null)
                    arr[7].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[7].Value = Infill1Angle;

                arr[8].ParameterName = "@Infill1StagePerc";
                if (Infill1StagePerc == null)
                    arr[8].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[8].Value = Infill1StagePerc;

                arr[9].ParameterName = "@Infill1MineralGange1";
                if (Infill1MineralGange1 == null)
                    arr[9].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[9].Value = Infill1MineralGange1;

                arr[10].ParameterName = "@Infill1MineralGange1Texture";
                if (Infill1MineralGange1Texture == null)
                    arr[10].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[10].Value = Infill1MineralGange1Texture;

                arr[11].ParameterName = "@Infill1MineralGange1Perc";
                if (Infill1MineralGange1Perc == null)
                    arr[11].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[11].Value = Infill1MineralGange1Perc;

                arr[12].ParameterName = "@Infill1MineralGange2";
                if (Infill1MineralGange2 == null)
                    arr[12].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[12].Value = Infill1MineralGange2;

                arr[13].ParameterName = "@Infill1MineralGange2Texture";
                if (Infill1MineralGange2Texture == null)
                    arr[13].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[13].Value = Infill1MineralGange2Texture;

                arr[14].ParameterName = "@Infill1MineralGange2Perc";
                if (Infill1MineralGange2Perc == null)
                    arr[14].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[14].Value = Infill1MineralGange2Perc;

                arr[15].ParameterName = "@Infill1MineralGange3";
                if (Infill1MineralGange3 == null)
                    arr[15].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[15].Value = Infill1MineralGange3;

                arr[16].ParameterName = "@Infill1MineralGange3Texture";
                if (Infill1MineralGange3Texture == null)
                    arr[16].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[16].Value = Infill1MineralGange3Texture;

                arr[17].ParameterName = "@Infill1MineralGange3Perc";
                if (Infill1MineralGange3Perc == null)
                    arr[17].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[17].Value = Infill1MineralGange3Perc;

                arr[18].ParameterName = "@Infill1OreMineral1";
                if (Infill1OreMineral1 == null)
                    arr[18].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[18].Value = Infill1OreMineral1;

                arr[19].ParameterName = "@Infill1OreMineral1Style";
                if (Infill1OreMineral1Style == null)
                    arr[19].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[19].Value = Infill1OreMineral1Style;

                arr[20].ParameterName = "@Infill1OreMineral1Perc";
                if (Infill1OreMineral1Perc == null)
                    arr[20].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[20].Value = Infill1OreMineral1Perc;

                arr[21].ParameterName = "@Infill1OreMineral2";
                if (Infill1OreMineral2 == null)
                    arr[21].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[21].Value = Infill1OreMineral2;

                arr[22].ParameterName = "@Infill1OreMineral2Style";
                if (Infill1OreMineral2Style == null)
                    arr[22].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[22].Value = Infill1OreMineral2Style;

                arr[23].ParameterName = "@Infill1OreMineral2Perc";
                if (Infill1OreMineral2Perc == null)
                    arr[23].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[23].Value = Infill1OreMineral2Perc;

                arr[24].ParameterName = "@Infill1OreMineral3";
                if (Infill1OreMineral3 == null)
                    arr[24].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[24].Value = Infill1OreMineral3;

                arr[25].ParameterName = "@Infill1OreMineral3Style";
                if (Infill1OreMineral3Style == null)
                    arr[25].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[25].Value = Infill1OreMineral3Style;

                arr[26].ParameterName = "@Infill1OreMineral3Perc";
                if (Infill1OreMineral3Perc == null)
                    arr[26].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[26].Value = Infill1OreMineral3Perc;

                arr[27].ParameterName = "@Infill2Stage";
                if (Infill2Stage == null)
                    arr[27].Value = System.Data.SqlTypes.SqlInt32.Null;
                else arr[27].Value = Infill2Stage;

                arr[28].ParameterName = "@Infill2Type";
                if (Infill2Type == null)
                    arr[28].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[28].Value = Infill2Type;

                arr[29].ParameterName = "@Infill2Number";
                if (Infill2Number == null)
                    arr[29].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[29].Value = Infill2Number;

                arr[30].ParameterName = "@Infill2AngleToAxis";
                if (Infill2Angle == null)
                    arr[30].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[30].Value = Infill2Angle;

                arr[31].ParameterName = "@Infill2StagePerc";
                if (Infill2StagePerc == null)
                    arr[31].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[31].Value = Infill2StagePerc;

                arr[32].ParameterName = "@Infill2MineralGange1";
                if (Infill2MineralGange1 == null)
                    arr[32].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[32].Value = Infill2MineralGange1;

                arr[33].ParameterName = "@Infill2MineralGange1Texture";
                if (Infill2MineralGange1Texture == null)
                    arr[33].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[33].Value = Infill2MineralGange1Texture;

                arr[34].ParameterName = "@Infill2MineralGange1Perc";
                if (Infill2MineralGange1Perc == null)
                    arr[34].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[34].Value = Infill2MineralGange1Perc;

                arr[35].ParameterName = "@Infill2MineralGange2";
                if (Infill2MineralGange2 == null)
                    arr[35].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[35].Value = Infill2MineralGange2;

                arr[36].ParameterName = "@Infill2MineralGange2Texture";
                if (Infill2MineralGange2Texture == null)
                    arr[36].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[36].Value = Infill2MineralGange2Texture;

                arr[37].ParameterName = "@Infill2MineralGange2Perc";
                if (Infill2MineralGange2Perc == null)
                    arr[37].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[37].Value = Infill2MineralGange2Perc;

                arr[38].ParameterName = "@Infill2MineralGange3";
                if (Infill2MineralGange3 == null)
                    arr[38].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[38].Value = Infill2MineralGange3;

                arr[39].ParameterName = "@Infill2MineralGange3Texture";
                if (Infill2MineralGange3Texture == null)
                    arr[39].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[39].Value = Infill2MineralGange3Texture;

                arr[40].ParameterName = "@Infill2MineralGange3Perc";
                if (Infill2MineralGange3Perc == null)
                    arr[40].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[40].Value = Infill2MineralGange3Perc;

                arr[41].ParameterName = "@Infill2OreMineral1";
                if (Infill2OreMineral1 == null)
                    arr[41].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[41].Value = Infill2OreMineral1;

                arr[42].ParameterName = "@Infill2OreMineral1Style";
                if (Infill2OreMineral1Style == null)
                    arr[42].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[42].Value = Infill2OreMineral1Style;

                arr[43].ParameterName = "@Infill2OreMineral1Perc";
                if (Infill2OreMineral1Perc == null)
                    arr[43].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[43].Value = Infill2OreMineral1Perc;

                arr[44].ParameterName = "@Infill2OreMineral2";
                if (Infill2OreMineral2 == null)
                    arr[44].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[44].Value = Infill2OreMineral2;

                arr[45].ParameterName = "@Infill2OreMineral2Style";
                if (Infill2OreMineral2Style == null)
                    arr[45].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[45].Value = Infill2OreMineral2Style;

                arr[46].ParameterName = "@Infill2OreMineral2Perc";
                if (Infill2OreMineral2Perc == null)
                    arr[46].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[46].Value = Infill2OreMineral2Perc;

                arr[47].ParameterName = "@Infill2OreMineral3";
                if (Infill2OreMineral3 == null)
                    arr[47].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[47].Value = Infill2OreMineral3;

                arr[48].ParameterName = "@Infill2OreMineral3Style";
                if (Infill1Type == null)
                    arr[48].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[48].Value = Infill1Type;

                arr[49].ParameterName = "@Infill2OreMineral3Perc";
                if (Infill2OreMineral3Perc == null)
                    arr[49].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[49].Value = Infill2OreMineral3Perc;

                arr[50].ParameterName = "@Infill3Stage";
                if (Infill3Stage == null)
                    arr[50].Value = System.Data.SqlTypes.SqlInt32.Null;
                else arr[50].Value = Infill3Stage;

                arr[51].ParameterName = "@Infill3Type";
                if (Infill3Type == null)
                    arr[51].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[51].Value = Infill3Type;

                arr[52].ParameterName = "@Infill3Number";
                if (Infill3Number == null)
                    arr[52].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[52].Value = Infill3Number;

                arr[53].ParameterName = "@Infill3AngleToAxis";
                if (Infill3Angle == null)
                    arr[53].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[53].Value = Infill3Angle;

                arr[54].ParameterName = "@Infill3StagePerc";
                if (Infill3StagePerc == null)
                    arr[54].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[54].Value = Infill3StagePerc;

                arr[55].ParameterName = "@Infill3MineralGange1";
                if (Infill3MineralGange1 == null)
                    arr[55].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[55].Value = Infill3MineralGange1;

                arr[56].ParameterName = "@Infill3MineralGange1Texture";
                if (Infill3MineralGange1Texture == null)
                    arr[56].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[56].Value = Infill3MineralGange1Texture;

                arr[57].ParameterName = "@Infill3MineralGange1Perc";
                if (Infill3MineralGange1Perc == null)
                    arr[57].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[57].Value = Infill3MineralGange1Perc;

                arr[58].ParameterName = "@Infill3MineralGange2";
                if (Infill3MineralGange2 == null)
                    arr[58].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[58].Value = Infill3MineralGange2;

                arr[59].ParameterName = "@Infill3MineralGange2Texture";
                if (Infill3MineralGange2Texture == null)
                    arr[59].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[59].Value = Infill3MineralGange2Texture;

                arr[60].ParameterName = "@Infill3MineralGange2Perc";
                if (Infill3MineralGange2Perc == null)
                    arr[60].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[60].Value = Infill3MineralGange2Perc;

                arr[61].ParameterName = "@Infill3MineralGange3";
                if (Infill3MineralGange3 == null)
                    arr[61].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[61].Value = Infill3MineralGange3;

                arr[62].ParameterName = "@Infill3MineralGange3Texture";
                if (Infill3MineralGange3Texture == null)
                    arr[62].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[62].Value = Infill3MineralGange3Texture;

                arr[63].ParameterName = "@Infill3MineralGange3Perc";
                if (Infill3MineralGange3Perc == null)
                    arr[63].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[63].Value = Infill3MineralGange3Perc;

                arr[64].ParameterName = "@Infill3OreMineral1";
                if (Infill3OreMineral1 == null)
                    arr[64].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[64].Value = Infill3OreMineral1;

                arr[65].ParameterName = "@Infill3OreMineral1Style";
                if (Infill3OreMineral1Style == null)
                    arr[65].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[65].Value = Infill3OreMineral1Style;

                arr[66].ParameterName = "@Infill3OreMineral1Perc";
                if (Infill3OreMineral1Perc == null)
                    arr[66].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[66].Value = Infill3OreMineral1Perc;

                arr[67].ParameterName = "@Infill3OreMineral2";
                if (Infill3OreMineral2 == null)
                    arr[67].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[67].Value = Infill3OreMineral2;

                arr[68].ParameterName = "@Infill3OreMineral2Style";
                if (Infill3OreMineral2Style == null)
                    arr[68].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[68].Value = Infill3OreMineral2Style;

                arr[69].ParameterName = "@Infill3OreMineral2Perc";
                if (Infill3OreMineral2Perc == null)
                    arr[69].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[69].Value = Infill3OreMineral2Perc;

                arr[70].ParameterName = "@Infill3OreMineral3";
                if (Infill3OreMineral3 == null)
                    arr[70].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[70].Value = Infill3OreMineral3;

                arr[71].ParameterName = "@Infill3OreMineral3Style";
                if (Infill3OreMineral3Style == null)
                    arr[71].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[71].Value = Infill3OreMineral3Style;

                arr[72].ParameterName = "@Infill3OreMineral3Perc";
                if (Infill3OreMineral3Perc == null)
                    arr[72].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[72].Value = Infill3OreMineral3Perc;

                arr[73].ParameterName = "@Infill4Stage";
                if (Infill4Stage == null)
                    arr[73].Value = System.Data.SqlTypes.SqlInt32.Null;
                else arr[73].Value = Infill4Stage;

                arr[74].ParameterName = "@Infill4Type";
                if (Infill4Type == null)
                    arr[74].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[74].Value = Infill4Type;

                arr[75].ParameterName = "@Infill4Number";
                if (Infill4Number == null)
                    arr[75].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[75].Value = Infill4Number;

                arr[76].ParameterName = "@Infill4AngleToAxis";
                if (Infill4Angle == null)
                    arr[76].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[76].Value = Infill4Angle;

                arr[77].ParameterName = "@Infill4StagePerc";
                if (Infill4StagePerc == null)
                    arr[77].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[77].Value = Infill4StagePerc;

                arr[78].ParameterName = "@Infill4MineralGange1";
                if (Infill4MineralGange1 == null)
                    arr[78].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[78].Value = Infill4MineralGange1;

                arr[79].ParameterName = "@Infill4MineralGange1Texture";
                if (Infill4MineralGange1Texture == null)
                    arr[79].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[79].Value = Infill4MineralGange1Texture;

                arr[80].ParameterName = "@Infill4MineralGange1Perc";
                if (Infill4MineralGange1Perc == null)
                    arr[80].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[80].Value = Infill4MineralGange1Perc;

                arr[81].ParameterName = "@Infill4MineralGange2";
                if (Infill4MineralGange2 == null)
                    arr[81].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[81].Value = Infill4MineralGange2;

                arr[82].ParameterName = "@Infill4MineralGange2Texture";
                if (Infill4MineralGange2Texture == null)
                    arr[82].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[82].Value = Infill4MineralGange2Texture;

                arr[83].ParameterName = "@Infill4MineralGange2Perc";
                if (Infill4MineralGange2Perc == null)
                    arr[83].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[83].Value = Infill4MineralGange2Perc;

                arr[84].ParameterName = "@Infill4MineralGange3";
                if (Infill4MineralGange3 == null)
                    arr[84].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[84].Value = Infill4MineralGange3;

                arr[85].ParameterName = "@Infill4MineralGange3Texture";
                if (Infill4MineralGange3Texture == null)
                    arr[85].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[85].Value = Infill4MineralGange3Texture;

                arr[86].ParameterName = "@Infill4MineralGange3Perc";
                if (Infill4MineralGange3Perc == null)
                    arr[86].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[86].Value = Infill4MineralGange3Perc;

                arr[87].ParameterName = "@Infill4OreMineral1";
                if (Infill4OreMineral1 == null)
                    arr[87].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[87].Value = Infill4OreMineral1;

                arr[88].ParameterName = "@Infill4OreMineral1Style";
                if (Infill4OreMineral1Style == null)
                    arr[88].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[88].Value = Infill4OreMineral1Style;

                arr[89].ParameterName = "@Infill4OreMineral1Perc";
                if (Infill4OreMineral1Perc == null)
                    arr[89].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[89].Value = Infill4OreMineral1Perc;

                arr[90].ParameterName = "@Infill4OreMineral2";
                if (Infill4OreMineral2 == null)
                    arr[90].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[90].Value = Infill4OreMineral2;

                arr[91].ParameterName = "@Infill4OreMineral2Style";
                if (Infill4OreMineral2Style == null)
                    arr[91].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[91].Value = Infill4OreMineral2Style;

                arr[92].ParameterName = "@Infill4OreMineral2Perc";
                if (Infill4OreMineral2Perc == null)
                    arr[92].Value = System.Data.SqlTypes.SqlDouble.Null;
                else arr[92].Value = Infill4OreMineral2Perc;

                arr[93].ParameterName = "@Infill4OreMineral3";
                if (Infill4OreMineral3 == null)
                    arr[93].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[93].Value = Infill4OreMineral3;

                arr[94].ParameterName = "@Infill4OreMineral3Style";
                if (Infill4OreMineral3Style == null)
                    arr[94].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[94].Value = Infill4OreMineral3Style;

                arr[95].ParameterName = "@Infill4OreMineral3Perc";
                if (Infill4OreMineral3Perc == null)
                    arr[95].Value = System.Data.SqlTypes.SqlString.Null;
                else arr[95].Value = Infill4OreMineral3Perc;

                arr[96].ParameterName = "@SKDHIn";
                if (iDHInfillID == null)
                    arr[96].Value = System.Data.SqlTypes.SqlInt32.Null;
                else arr[96].Value = iDHInfillID;


                oRes = oData.ExecuteScalar("usp_DH_Infill_Insert", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Save error DH_Infill. " + ex.Message);
            }
        }

        public DataTable getDHInfillFromToValid()
        {
            try
            {
                DataSet dtDHMinFromToValid = new DataSet();
                SqlParameter[] arr = oData.GetParameters(3);
                arr[0].ParameterName = "@From";
                arr[0].Value = dFrom;
                arr[1].ParameterName = "@To";
                arr[1].Value = dTo;
                arr[2].ParameterName = "@HoleID";
                arr[2].Value = sHoleID;
                dtDHMinFromToValid = oData.ExecuteDataset("usp_DH_Mineraliz_ListFromToValid", arr, CommandType.StoredProcedure);
                return dtDHMinFromToValid.Tables[0];

            }
            catch (Exception eX)
            {
                throw new Exception("Error in dtDHInfillFromToValid: " + eX.Message);
            }
        }


        public DataTable getDHInfill()
        {
            try
            {
                DataSet dtDHMin = new DataSet();
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@HoleID";
                arr[0].Value = sHoleID;
                dtDHMin = oData.ExecuteDataset("usp_DH_Infill_List", arr, CommandType.StoredProcedure);
                return dtDHMin.Tables[0];
            }
            catch (Exception eX)
            {
                throw new Exception("Error in DH_Infill_List: " + eX.Message);
            }
        }

        public string DH_Samples_Delete()
        {
            try
            {
                object oRes;
                SqlParameter[] arr = oData.GetParameters(1);
                arr[0].ParameterName = "@SKDHInfill";
                arr[0].Value = iDHInfillID;
                oRes = oData.ExecuteScalar("usp_DH_Samples_DeleteInfill", arr, CommandType.StoredProcedure);
                return oRes.ToString();
            }
            catch (Exception eX)
            {
                throw new Exception("Delete error DH_Infill_Delete. " + eX.Message); ;
            }
        }


        public DataTable DH_Infill_Consulta(string holeId, double pfrom, double pto)
        {
            try
            {
                DataSet dataSet = new DataSet();
                SqlParameter[] parameters = oData.GetParameters(3);
                parameters[0].ParameterName = "@HoleID";
                parameters[0].Value = holeId;
                parameters[1].ParameterName = "@pFrom";
                parameters[1].Value = pfrom;
                parameters[2].ParameterName = "@pTo";
                parameters[2].Value = pto;
                dataSet = oData.ExecuteDataset("usp_DH_Infill_Consulta", parameters, CommandType.StoredProcedure);
                return dataSet.Tables[0];
            }
            catch (Exception ex)
            {
                throw new Exception("Delete error DH_Infill_Consulta. " + ex.Message);
            }
        }


        public string DH_Delete_Infill()
        {
            try
            {
                object oRes;
                SqlParameter[] parameters = oData.GetParameters(1);
                parameters[0].ParameterName = "@HoleInfill";
                parameters[0].Value = wHoleInfill;
                oRes = oData.ExecuteScalar("usp_DH_DeleteInfill_All", parameters, CommandType.StoredProcedure);

                return oRes.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Delete error DH_Infill_Delete. " + ex.Message);
            }
        }

    }
}