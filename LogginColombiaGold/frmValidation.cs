using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Configuration;

namespace LogginColombiaGold
{
    public partial class frmValidation : Form
    {

        clsRf oRf = new clsRf();
        clsDHCollars oCollars = new clsDHCollars();
        clsDHSamples oSamp = new clsDHSamples();
        clsDHGeotech oGeo = new clsDHGeotech();
        clsDHLithology oLit = new clsDHLithology();
        clsDH_Weathering oWeat = new clsDH_Weathering();
        clsDH_Structures oStr = new clsDH_Structures();
        clsDHMineraliz oMiner = new clsDHMineraliz();
        clsDHBox oBox = new clsDHBox();
        clsDHAlterations oAlt = new clsDHAlterations();
        clsDHOxides oOxid = new clsDHOxides();
        private DataTable dtCollars;

        Configuration conf = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        static string sValidLogging = "";

        public frmValidation()
        {
            InitializeComponent();
        }

        private void frmValidation_Load(object sender, EventArgs e)
        {
            try
            {
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                dtCollars = oCollars.getDHCollarsLogged();        
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnValidSamples_Click(object sender, EventArgs e)
        {
            
            try
            {
                sValidLogging = cmbTypeValidation.Text.ToString(); //Ejecuta los eventos bgw_DoWork, bgw_ProgressChanged y bgw_RunWorkerCompleted
                bgw.RunWorkerAsync();

            }
            catch (Exception ex)
            {    
                MessageBox.Show(ex.Message);
            }
            
            
        }

        private void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            Thread.Sleep(100);

            DateTime start = DateTime.Now;
            e.Result = "";
            for (int i = 0; i < 100; i++)
            {
                System.Threading.Thread.Sleep(50);

                bgw.ReportProgress(i, DateTime.Now);


                if (bgw.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
            }

            TimeSpan duration = DateTime.Now - start;

            e.Result = "Duration: " + duration.TotalMilliseconds.ToString() + " ms.";
        }

        private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbLogging.Visible = true;
            pbLogging.Value = e.ProgressPercentage; //actualizamos la barra de progreso
            DateTime time = Convert.ToDateTime(e.UserState); //obtenemos información adicional si procede

            if (pbLogging.Value > 98)
            {
                pbLogging.Visible = false;
            }
        }

        private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ValidLogging();
        }

        private void ValidLogging()
        {
            try
            {
                switch (sValidLogging)
                {
                       
                    case "Samples":
                        SamplesValid();
                        break;
                    case "Weathering":
                        WeathValid(); ;
                        break;
                    case "Lithology":
                        LiththValid();
                        break;
                    case "Alterations":
                        AlterationsValid();
                        break;
                    case "Structures":
                        StructuresValid();
                        break;
                    case "Mineralizations":
                        MineralizationsValid();
                        break;
                    case "Geotech":
                        GeotechValid();
                        break;
                    default:
                        Console.WriteLine("Default case");
                        break;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private DataTable FillStruct(string _sHoleid)
        {
            try
            {
                DataTable dtStruct = new DataTable();
                oStr.sOpcion = "2";
                oStr.sHoleID = _sHoleid.ToString();
                dtStruct = oStr.getDH_Structures();
                return dtStruct;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private DataTable FillAlterations(string _sHoleid)
        {
            try
            {
                DataTable dtAlterations = new DataTable();
                oAlt.sOpcion = "2";
                oAlt.sHoleID = _sHoleid.ToString();
                dtAlterations = oAlt.getDH_Alterations();
                return dtAlterations;

            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private DataTable FillWeathering(string _sHoleid)
        {
            try
            {
                DataTable dtWeat = new DataTable();
                oWeat.sHoleID = _sHoleid.ToString();
                oWeat.sOpcion = "2";
                oWeat.sHoleID = _sHoleid.ToString();
                dtWeat = oWeat.getDH_Weathering();
                return dtWeat;

            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private DataTable FillLithology(string _sHoleid)
        {
            try
            {
                DataTable dtLit = new DataTable();
                oLit.sOpcion = "2";
                oLit.sHoleID = _sHoleid.ToString();
                dtLit = oLit.getDH_Lithology();
                return dtLit;
            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private DataTable FillSample(string _sHoleid)
        {
            try
            {

                DataTable dtLoggin = new DataTable();
                oSamp.sHoleID = _sHoleid.ToString();
                dtLoggin = oSamp.getDHSamples();
                return dtLoggin;

            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private DataTable FillGeo(string _sHoleid)
        {
            try
            {
                DataTable dtGeo = new DataTable();
                oGeo.sOpcion = "2";
                oGeo.sHoleID = _sHoleid.ToString();
                dtGeo = oGeo.getDH_Geotech();
                return dtGeo;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private DataTable FillMineraliz(string _sHoleid)
        {
            try
            {
                DataTable dtMiner = new DataTable();
                oMiner.sOpcion = "2";
                oMiner.sHoleID = _sHoleid.ToString();
                dtMiner = oMiner.getDHMineraliz();
                return dtMiner;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void SamplesValid()
        {
            try
            {


                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oSamp.sOpcion = "2";
                oSamp.iFrom = 0; oSamp.iTo = 0; oSamp.sHoleID = "0"; oSamp.iDHSampID = 0;
                dtResult = oSamp.getDHSamplesValid();
                DataTable dtHoleID = new DataTable();
                DataTable dtFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillSample(dtCollars.Rows[a]["HoleId"].ToString());
                    //DataTable dtResult = dtResultTemp.Copy();

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oSamp.iFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oSamp.iTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oSamp.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oSamp.iDHSampID = long.Parse(dtHoleID.Rows[i]["SKDHSamples"].ToString());
                        dtValid = oSamp.getDHSamplesValid();

                        if (dtValid.Rows.Count > 0)
                        {
                            //DataRowView dv = (DataRowView)gdLoggin.Rows[i].DataBoundItem;
                            //DataRow dr = dv.Row;

                            //implementar ciclo de 1 hasta dtvalid.count
                            dtResult.ImportRow(dtValid.Rows[0]);

                        }
                    }


                    dgData.DataSource = dtResult;

                    //Exportar a excel los resultados de from to overlaps y from to next


                    oSamp.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtFromToNext = oSamp.getDHSamplesValidFromToNext();

                    //Fin Export Excel


                }


                if (dtResult.Rows.Count > 0 || dtFromToNext.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidSamples"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 4] = dtCollars.Rows[a]["HoleId"].ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["Sample"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["SampleType"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["DupDe"].ToString();
                        oSheet.Cells[iInicial, 6] = dtResult.Rows[i]["Lithology"].ToString();
                        oSheet.Cells[iInicial, 7] = "From To Overlaps";
                        oSheet.Cells[iInicial, 8] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }


                    //oSheet.Cells[iInicial, 1] = "From To Next Invalid";
                    //iInicial += 1;

                    for (int iF = 0; iF < dtFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtFromToNext.Rows[iF]["Sample"].ToString();
                        oSheet.Cells[iInicial, 2] = dtFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 3] = dtFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 4] = dtFromToNext.Rows[iF]["SampleType"].ToString();
                        oSheet.Cells[iInicial, 5] = dtFromToNext.Rows[iF]["DupDe"].ToString();
                        oSheet.Cells[iInicial, 6] = dtFromToNext.Rows[iF]["Lithology"].ToString();
                        oSheet.Cells[iInicial, 7] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 8] = dtResult.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }



                    oXL.Visible = true;
                    oXL.UserControl = true;


                }
                else
                {
                    MessageBox.Show("No Overlaps ", "Samples", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }



            }
            catch (Exception ex)
            {
                //return "Termino";

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Samples", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void LiththValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oSamp.sOpcion = "2";
                oLit.dFrom = 0; oLit.dTo = 0; oLit.sHoleID = "0"; oLit.iDHLithologyID = 0;
                dtResult = oLit.getDHLitValid();

                DataTable dtHoleID = new DataTable();
                DataTable dtLithFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillLithology(dtCollars.Rows[a]["HoleId"].ToString());

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oLit.dFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oLit.dTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oLit.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oLit.iDHLithologyID = long.Parse(dtHoleID.Rows[i]["SKDHLithology"].ToString());
                        dtValid = oLit.getDHLitValid();

                        if (dtValid.Rows.Count > 0)
                        {
                            dtResult.ImportRow(dtValid.Rows[0]);
                        }
                    }

                    dgData.DataSource = dtResult;

                    oLit.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtLithFromToNext = oLit.getDHLitFromToValidFromToNext();
                    

                    //Fin Export Excel

                }

                if (dtResult.Rows.Count > 0 || dtLithFromToNext.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidLitho"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 3] = cmbHoleIdLit.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["Litho"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["GSize"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Textures"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    for (int iF = 0; iF < dtLithFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtLithFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtLithFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtLithFromToNext.Rows[iF]["Litho"].ToString();
                        oSheet.Cells[iInicial, 4] = dtLithFromToNext.Rows[iF]["GSize"].ToString();
                        oSheet.Cells[iInicial, 5] = dtLithFromToNext.Rows[iF]["Textures"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 7] = dtLithFromToNext.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }



                    oXL.Visible = true;
                    oXL.UserControl = true;


                }
                else
                {
                    MessageBox.Show("No Overlaps ", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void WeathValid()
        {
            try
            {


                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oWeat.dFrom = 0; oWeat.dTo = 0; oWeat.sHoleID = "0"; oWeat.iDHWeatheringID = 0;
                dtResult = oWeat.getDHWeatValid();

                DataTable dtHoleID = new DataTable();
                DataTable dtWeatFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillWeathering(dtCollars.Rows[a]["HoleId"].ToString());

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oWeat.dFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oWeat.dTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oWeat.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oWeat.iDHWeatheringID = long.Parse(dtHoleID.Rows[i]["SKDHWeathering"].ToString());
                        dtValid = oWeat.getDHWeatValid();

                        if (dtValid.Rows.Count > 0)
                        {

                            dtResult.ImportRow(dtValid.Rows[0]);

                        }
                    }

                    dgData.DataSource = dtResult;

                    oWeat.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtWeatFromToNext = oWeat.getDHWeatValidFromToNext();

                }




                if (dtResult.Rows.Count > 0 || dtWeatFromToNext.Rows.Count > 0)
                {

                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidWeath"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 3] = dtCollars.Rows[a]["HoleId"].ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["Weathering"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["Oxidation"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Colour1"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    for (int iF = 0; iF < dtWeatFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtWeatFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtWeatFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtWeatFromToNext.Rows[iF]["Weathering"].ToString();
                        oSheet.Cells[iInicial, 4] = dtWeatFromToNext.Rows[iF]["Oxidation"].ToString();
                        oSheet.Cells[iInicial, 5] = dtWeatFromToNext.Rows[iF]["Colour1"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 7] = dtWeatFromToNext.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;


                }
                else 
                {
                    MessageBox.Show("No Overlaps ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            catch (Exception ex)
            {
                //return "Termino";

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void AlterationsValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oAlt.dFrom = 0; oAlt.dTo = 0; oAlt.sHoleID = "0"; oAlt.iSHDHAlterarions = 0;
                dtResult = oAlt.getDHAlterationsValid();

                DataTable dtHoleID = new DataTable();
                DataTable dtAlterFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillAlterations(dtCollars.Rows[a]["HoleId"].ToString());

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oAlt.dFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oAlt.dTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oAlt.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oAlt.iSHDHAlterarions = long.Parse(dtHoleID.Rows[i]["SKDHAlterarions"].ToString());
                        dtValid = oAlt.getDHAlterationsValid();

                        if (dtValid.Rows.Count > 0)
                        {
                            dtResult.ImportRow(dtValid.Rows[0]);
                        }
                    }

                    dgData.DataSource = dtResult;

                    oAlt.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtAlterFromToNext = oAlt.getDHAlterationsValidFromToNext();

                }


                
                if (dtResult.Rows.Count > 0 || dtAlterFromToNext.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidAlter"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 3] = cmbHoleIDAlt.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["A1Type"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["A1Int"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["A1Style"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    for (int iF = 0; iF < dtAlterFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtAlterFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtAlterFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtAlterFromToNext.Rows[iF]["A1Type"].ToString();
                        oSheet.Cells[iInicial, 4] = dtAlterFromToNext.Rows[iF]["A1Int"].ToString();
                        oSheet.Cells[iInicial, 5] = dtAlterFromToNext.Rows[iF]["A1Style"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }



                    oXL.Visible = true;
                    oXL.UserControl = true;
                }
                else
                {
                    MessageBox.Show("No Overlaps ", "Alterations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void GeotechValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oGeo.iFrom = 0; oGeo.iTo = 0; oGeo.sHoleID = "0"; oGeo.iDHGeotechID = 0;
                dtResult = oGeo.getDHGeotechValid();

                DataTable dtHoleID = new DataTable();
                DataTable dtGeoFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillGeo(dtCollars.Rows[a]["HoleId"].ToString());

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oGeo.iFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oGeo.iTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oGeo.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oGeo.iDHGeotechID = long.Parse(dtHoleID.Rows[i]["SKDHGeotech"].ToString());
                        dtValid = oGeo.getDHGeotechValid();

                        if (dtValid.Rows.Count > 0)
                        {
                            dtResult.ImportRow(dtValid.Rows[0]);
                        }
                    }

                    dgData.DataSource = dtResult;

                    oGeo.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtGeoFromToNext = oGeo.getDHGeotechValidFromToNext();

                }
               
                 
                if (dtResult.Rows.Count > 0 || dtGeoFromToNext.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidGeotech"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 3] = cmbHoleIdGeo.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["LithCod"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["Recm"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["RQDcm"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    for (int iF = 0; iF < dtGeoFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtGeoFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtGeoFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtGeoFromToNext.Rows[iF]["LithCod"].ToString();
                        oSheet.Cells[iInicial, 4] = dtGeoFromToNext.Rows[iF]["Recm"].ToString();
                        oSheet.Cells[iInicial, 5] = dtGeoFromToNext.Rows[iF]["RQDcm"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 7] = dtGeoFromToNext.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;



                }
                else
                {
                    MessageBox.Show("No Overlaps ", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void StructuresValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oStr.iFrom = 0; oStr.iTo = 0; oStr.sHoleID = "0"; oStr.iDHStructrueID = 0;
                dtResult = oStr.getDH_StructuresValid();

                DataTable dtHoleID = new DataTable();
                DataTable dtStrucFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillStruct(dtCollars.Rows[a]["HoleId"].ToString());

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oStr.iFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oStr.iTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oStr.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oStr.iDHStructrueID = long.Parse(dtHoleID.Rows[i]["SKDHStructrue"].ToString());
                        dtValid = oStr.getDH_StructuresValid();

                        if (dtValid.Rows.Count > 0)
                        {
                            dtResult.ImportRow(dtValid.Rows[0]);
                        }
                    }

                    dgData.DataSource = dtResult;

                    oStr.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtStrucFromToNext = oStr.getDH_StructuresValidFromToNext();
                }


               

                if (dtResult.Rows.Count > 0 || dtStrucFromToNext.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidStruct"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 3] = cmbHoleIDSt.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["Type"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["AngleToAxis"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Fill"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    for (int iF = 0; iF < dtStrucFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtStrucFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtStrucFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtStrucFromToNext.Rows[iF]["Type"].ToString();
                        oSheet.Cells[iInicial, 4] = dtStrucFromToNext.Rows[iF]["AngleToAxis"].ToString();
                        oSheet.Cells[iInicial, 5] = dtStrucFromToNext.Rows[iF]["Fill"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;

                }
                else
                {
                    MessageBox.Show("No Overlaps ", "Structures", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void MineralizationsValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oMiner.dFrom = 0; oMiner.dTo = 0; oMiner.sHoleID = "0"; oMiner.iDHMinID = 0;
                dtResult = oMiner.getDHMinValid();

                 DataTable dtHoleID = new DataTable();
                 DataTable dtMinerFromToNext = new DataTable();

                for (int a = 0; a < dtCollars.Rows.Count; a++)
                {

                    dtHoleID = FillMineraliz(dtCollars.Rows[a]["HoleId"].ToString());

                    for (int i = 0; i < dtHoleID.Rows.Count - 1; i++)
                    {
                        dtValid = new DataTable();
                        oMiner.dFrom = double.Parse(dtHoleID.Rows[i]["From"].ToString());
                        oMiner.dTo = double.Parse(dtHoleID.Rows[i]["To"].ToString());
                        oMiner.sHoleID = dtHoleID.Rows[i]["HoleID"].ToString();
                        oMiner.iDHMinID = long.Parse(dtHoleID.Rows[i]["SKDHMin"].ToString());
                        dtValid = oMiner.getDHMinValid();

                        if (dtValid.Rows.Count > 0)
                        {
                            dtResult.ImportRow(dtValid.Rows[0]);
                        }
                    }

                    dgData.DataSource = dtResult;

                    oMiner.sHoleID = dtCollars.Rows[a]["HoleId"].ToString();
                    dtMinerFromToNext = oMiner.getDHMinFromToValidFromToNext();

                }


                
                if (dtResult.Rows.Count > 0 || dtMinerFromToNext.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidMineral"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    //oSheet.Cells[4, 3] = cmbHoleIdMin.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["MZ1Mineral"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["MZ1Perc"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["MZ1Style"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[i]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    for (int iF = 0; iF < dtMinerFromToNext.Rows.Count; iF++)
                    {

                        oSheet.Cells[iInicial, 1] = dtMinerFromToNext.Rows[iF]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtMinerFromToNext.Rows[iF]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtMinerFromToNext.Rows[iF]["MZ1Mineral"].ToString();
                        oSheet.Cells[iInicial, 4] = dtMinerFromToNext.Rows[iF]["MZ1Perc"].ToString();
                        oSheet.Cells[iInicial, 5] = dtMinerFromToNext.Rows[iF]["MZ1Style"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                        oSheet.Cells[iInicial, 7] = dtResult.Rows[iF]["HoleId"].ToString();
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;

                }
                else
                {
                    MessageBox.Show("No Overlaps ", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

       

        private void cmbTypeValidation_SelectionChangeCommitted(object sender, EventArgs e)
        {
            
        }

        private void cmbTypeValidation_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
