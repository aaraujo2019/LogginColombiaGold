using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace LogginColombiaGold
{
    public partial class frmLoggin : Form
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
        clsDHDensity oDens = new clsDHDensity();

        bool bHoleValid = true;
        bool bInicio = false;
        
        static string sDHSamplesID = "0";
        static string sEdit = "0";
        static string sEditGeo = "0";
        static string sEditLit = "0";
        static string sEditWeat = "0";
        static string sEditStruct = "0";
        static string sEditMiner = "0";
        static string sEditBox = "0";
        static string sEditAlt = "0";
        static string sEditDens = "0";
        static string sEditDensM = "0";
        static string sValidLogging = ""; //variable para saber que pestaña validar

        static string SheetExcel = "";

        Configuration conf = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

        public frmLoggin()
        {
            InitializeComponent();
            bInicio = true;
            
            FillHoleIDForm();

            FillCmb();
            bInicio = false;
            DisableControls();
            oSamp.sOpcion = "1";
            FillLoggin();
            bHoleValid = true;

            FillCmbGeoTech();
            FillCmbLith();
            FillCmbWeath();
            FillCmbStruct();
            FillCmbMiner();
            FillCmbAlt();
            FillCmbBox();
            FillCmbAlterations();
            //FillCmbOxidation();
        }

        private void frmLoggin_Load(object sender, EventArgs e)
        {
            DataRow[] datoPest;
            DataRow[] dato = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'frmLoggin' ");

            CheckForIllegalCrossThreadCalls = false;

            if (dato.Length > 0)
            {

                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Samples'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbSamples);
                }
                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Box'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbBox);
                }
                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Alterations'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbAlteration);
                }
                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Lithology'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbLithology);
                }
                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Geotech'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbGeotech);
                }
                
                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Stuctures'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbStructure);
                }
                datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Mineralizations'");
                if (datoPest.Length == 0)
                {
                    TabPpal.TabPages.Remove(tbMineraliz);
                }

                /* No se usan solicitan quitarlos*/
                TabPpal.TabPages.Remove(tbWeathering);
                TabPpal.TabPages.Remove(tbDensity);

                //datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Weathering'");
                //if (datoPest.Length == 0)
                //{
                //    TabPpal.TabPages.Remove(tbWeathering);
                //}

                //datoPest = clsRf.dsPermisos.Tables[0].Select("nombre_Real_Form = 'Density'");
                //if (datoPest.Length == 0)
                //{
                //    TabPpal.TabPages.Remove(tbDensity);
                //}
            }

            //TabPpal.TabPages.Remove(tbDensity);
        }

        private void DisableControls()
        {
            try
            {
                txtFrom.Enabled = false;
                txtTo.Enabled = false;

                txtFrom.Text = "";
                txtTo.Text = "";
                cmbLithology.SelectedValue = "-1";
                cmbSampleType.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                throw new Exception("Error Disable Controls: " + ex.Message);
            }
        }

        private void EnableControls()
        {
            try
            {
                txtFrom.Enabled = true;
                txtTo.Enabled = true;
            }
            catch (Exception ex)
            {
                throw new Exception("Error Enable Controls: " + ex.Message);
            }
        }

        private void FillHoleIDForm()
        {
            try
            {
                //cmbHoleIDForm
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drC = dtCollars.NewRow();
                drC[0] = "Select an option..";
                dtCollars.Rows.Add(drC);
                cmbHoleIDForm.DisplayMember = "HoleID";
                cmbHoleIDForm.ValueMember = "HoleID";
                cmbHoleIDForm.DataSource = dtCollars;
                cmbHoleIDForm.SelectedValue = "Select an option..";


                cmbHoleIdDens.DisplayMember = "HoleID";
                cmbHoleIdDens.ValueMember = "HoleID";
                cmbHoleIdDens.DataSource = dtCollars.Copy();
                cmbHoleIdDens.SelectedValue = "Select an option..";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FillCmb()
        {
            try
            {
                DataTable dtSampleT = new DataTable();
                dtSampleT = oRf.getRfTypeSample();

                DataRow dr = dtSampleT.NewRow();
                dr[0] = "-1";
                dr[1] = "Select an option..";
                dtSampleT.Rows.Add(dr);

                cmbSampleType.DisplayMember = "Comb";
                cmbSampleType.ValueMember = "Code";
                cmbSampleType.DataSource = dtSampleT;
                cmbSampleType.SelectedValue = -1;


                DataTable dtLithology = new DataTable();
                dtLithology = oRf.getDsRfLithology().Tables[0];

                DataRow drL = dtLithology.NewRow();
                drL[0] = "-1";
                drL[1] = "Select an option..";
                dtLithology.Rows.Add(drL);

                cmbLithology.DisplayMember = "Comb";
                cmbLithology.ValueMember = "Code";
                cmbLithology.DataSource = dtLithology;
                cmbLithology.SelectedValue = -1;


                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drC = dtCollars.NewRow();
                drC[0] = "Select an option..";
                dtCollars.Rows.Add(drC);
                cmbHoleID.DisplayMember = "HoleID";
                cmbHoleID.ValueMember = "HoleID";
                cmbHoleID.DataSource = dtCollars;
                cmbHoleID.SelectedValue = "Select an option..";


                DataTable dtLocation = oRf.getLocation("");
                DataRow drLoc = dtLocation.NewRow();
                drLoc[1] = "Select an option...";
                dtLocation.Rows.Add(drLoc);
                cmbVeinNameDens.DisplayMember = "Description";
                cmbVeinNameDens.ValueMember = "Description";
                cmbVeinNameDens.DataSource = dtLocation;
                cmbVeinNameDens.SelectedValue = "Select an option...";

                DataTable dtLab = oRf.getRfCodeLab();
                cmbLabDensM.DisplayMember = "Code";
                cmbLabDensM.ValueMember = "Code";
                cmbLabDensM.DataSource = dtLab;
                cmbLabDensM.SelectedValue = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();

            }
            catch (Exception ex)
            {
                throw new Exception("Error Fill Sample TypD: " + ex.Message);
            }
        }



        private void txtFrom_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtTo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void cmbLithology_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void FillLoggin()
        {
            try
            {
                DataTable dtLoggin = new DataTable();
                oSamp.sHoleID = cmbHoleID.SelectedValue.ToString();
                dtLoggin = oSamp.getDHSamples();
                gdLoggin.DataSource = dtLoggin;

                gdLoggin.Columns["SKDHSamples"].Visible = false;

                
                foreach (DataGridViewColumn Col in gdLoggin.Columns)
                {
                    Col.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                //gdLoggin.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;

               
                
            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private DataTable SampleIdRepeat(string _sSampleId)
        {
            try
            {
                oSamp.sSample = _sSampleId;
                DataTable dtSamp = oSamp.getDHSamplesId();
                return dtSamp;
                //if (dtSamp.Rows.Count > 0)
                //{
                //    return false;
                //}
                //else
                //{ return true; }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        private void btnCancelSamp_Click(object sender, EventArgs e)
        {
            try
            {
                sEdit = "0";
                txtTo.Text = "";
                cmbSampleType.SelectedValue = "-1";
                txtDupDe.Text = "";
                cmbLithology.SelectedValue = "-1";
                txtCommentsSamp.Text = "";
                EnableControls();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {

                string sFrom = String.Format(txtFrom.Text.ToString(), "#########0.00");
                string sTo = String.Format(txtTo.Text.ToString(), "#########0.00");
                double dtotalFromTo = double.Parse(sTo.ToString()) - double.Parse(sFrom.ToString());


                //Valida que los datos en from to sean validos, mayor que cero o -99
                bool bFromtoValido = true;
                if (double.Parse(sTo.ToString()) >= 0 || double.Parse(sTo.ToString()) == -99)
                {
                    bFromtoValido = true;
                }
                else { bFromtoValido = false; }

                if (double.Parse(sFrom.ToString()) >= 0 || double.Parse(sFrom.ToString()) == -99)
                {
                    bFromtoValido = true;
                }
                else { bFromtoValido = false; }

                if (bFromtoValido == false)
                {
                    MessageBox.Show("Invalid To or From. (> 0 or -99)", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //Fin. Valida que los datos en from to sean validos, mayor que cero o -99

                //Validar si el holeid es valido para el usuario logueado          
                oCollars.sLogged = clsRf.sUser;
                DataTable dtLogg = oCollars.getDHCollarsLogged();
                DataRow[] datoL = dtLogg.Select("HoleID = '" + cmbHoleID.Text.ToString() + "'");
                if (datoL.Length > 0)
                { bHoleValid = true; }
                else { bHoleValid = false; }


                if (bHoleValid == false)
                {
                    MessageBox.Show("HoleId Invalid", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //Fin. Validar si el holeid es valido para el usuario logueado   



                if (sEdit == "0")
                {
                    //Valida que el sampleid no este repetido
                    DataTable dtValSamp = SampleIdRepeat(txtSampNo.Text.ToString());
                    bool bValSamp = false;
                    if (dtValSamp.Rows.Count > 0)
                    {
                        bValSamp = false;
                    }
                    else
                    {
                        bValSamp = true;
                    }

                    if (bValSamp == false)
                    {
                        MessageBox.Show("Sample duplicated to HoleID: " + dtValSamp.Rows[0]["HoleID"].ToString()
                            , "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                

                //validar lithology si se elige original
                string sLith = "";
                DataTable dtOri = dtOriginal();
                DataRow[] datoLith = dtOri.Select("Value = '" + cmbSampleType.SelectedValue.ToString() + "'");
                if (datoLith.Length > 0)
                {

                    if (cmbLithology.SelectedValue.ToString() == "-1" || cmbLithology.SelectedValue.ToString() == "")
                    {
                        MessageBox.Show("Selected an option Lithology", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    //Valida que to sea mayor que el from
                    if (double.Parse(sTo.ToString()) == double.Parse(sFrom.ToString()))
                    {
                        MessageBox.Show("From = To", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    //Valida que to sea mayor que el from
                    if (double.Parse(sTo.ToString()) < double.Parse(sFrom.ToString()))
                    {
                        MessageBox.Show("To <= From", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }


                    #region Valida Min - Max
                    //Se valida que el rango ingresado no supere el minimo o maximo establecido
                    if (dtotalFromTo <
                        double.Parse(ConfigurationSettings.AppSettings["MinMuestra"].ToString())
                        ||
                        dtotalFromTo >
                        double.Parse(ConfigurationSettings.AppSettings["MaxMuestra"].ToString()))
                    {
                        MessageBox.Show("Range 'From To' less than "
                                + ConfigurationSettings.AppSettings["MinMuestra"].ToString()
                                + " Or greater than "
                                + ConfigurationSettings.AppSettings["MaxMuestra"].ToString()
                            , "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        //Se valida que el rango ingresado no supere el minimo o maximo establecido informativamente
                        if (dtotalFromTo <
                        double.Parse(ConfigurationSettings.AppSettings["MinMenMuestra"].ToString())
                        ||
                        dtotalFromTo >
                        double.Parse(ConfigurationSettings.AppSettings["MaxMenMuestra"].ToString())
                        )
                        {
                            if (MessageBox.Show("Range 'From To' less than "
                                + ConfigurationSettings.AppSettings["MinMenMuestra"].ToString()
                                + " Or greater than "
                                + ConfigurationSettings.AppSettings["MaxMenMuestra"].ToString()
                            , "Logging", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                            {

                                goto ProcessAdd;
                                //MessageBox.Show(
                                //   (double.Parse(sTo.ToString()) - double.Parse(sFrom.ToString())).ToString()
                                //   );
                                //return;
                            }
                            else
                            {
                                return;
                            }
                        }
                        //Si entra por aca, no tiene problemas en el rango ingresado
                        else
                        {
                            goto ProcessAdd;
                        }

                    }
                    #endregion

                ProcessAdd:



                    //Valida que el rango sea valido para el pozo
                    DataTable dtValidRange = new DataTable();
                    oSamp.iFrom = double.Parse(txtFrom.Text.ToString());
                    oSamp.iTo = double.Parse(txtTo.Text.ToString());
                    oSamp.sHoleID = cmbHoleID.SelectedValue.ToString();
                    if (sEdit == "0")
                    {
                        oSamp.iDHSampID = 0;
                    }
                    else
                    { oSamp.iDHSampID = long.Parse(sDHSamplesID.ToString()); }

                    dtValidRange = oSamp.getDHSamplesFromToValid();
                    if (dtValidRange.Rows.Count > 0)
                    {
                        MessageBox.Show("Range 'From To' Overlaps", "Samples", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    //Fin. Valida que el rango sea valido para el pozo


                    //if (cmbLithology.SelectedValue.ToString() == "-1" || cmbLithology.SelectedValue.ToString() == "")
                    //    sLith = null;
                    //else 
                        
                    sLith = cmbLithology.SelectedValue.ToString();

                    clsDHSamples.sStaticFrom = txtTo.Text.ToString();


                }
                else
                {
                    txtFrom.Text = "-99";
                    txtTo.Text = "-99";

                    sLith = "";
                    clsDHSamples.sStaticFrom = "0";
                }




                if (cmbSampleType.SelectedValue.ToString() == "-1" ||
                    //cmbLithology.SelectedValue.ToString() == "-1" ||
                    cmbHoleID.SelectedValue.ToString() == "Select an option..")
                {
                    MessageBox.Show("Selected an option (Hole Id and SampleType)", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                oCollars.sHoleID = cmbHoleID.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length >= '" + txtTo.Text + "'");
                if (dato.Length > 0)
                {

                    if (sEdit == "0")
                    {
                        oSamp.sOpcion = "1";
                    }
                    else if (sEdit == "1")
                    {
                        oSamp.sOpcion = "2";
                    }

                    oSamp.sHoleID = cmbHoleID.SelectedValue.ToString();
                    oSamp.sSample = txtSampNo.Text.ToString().ToUpper();
                    oSamp.iFrom = double.Parse(txtFrom.Text.ToString());
                    oSamp.iTo = double.Parse(txtTo.Text.ToString());
                    oSamp.sSampleType = cmbSampleType.SelectedValue.ToString();
                    oSamp.sDupDe = txtDupDe.Text.ToString();
                    oSamp.sComments = txtCommentsSamp.Text.ToString();
                    oSamp.iDHSampID = Int64.Parse(sDHSamplesID.ToString());

                    oSamp.sLith = sLith; //cmbLithology.SelectedValue.ToString();

                    string sResp = oSamp.DHSamples_AddLoggin();

                    if (sResp.ToString() == "OK")
                    {
                        //MessageBox.Show("Insert Successful ", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    

                        ///Valida la informacion contra los datos de litologia
                        DataTable dtLit = new DataTable();
                        oLit.sOpcion = "2";
                        oLit.sHoleID = cmbHoleID.SelectedValue.ToString();
                        dtLit = oLit.getDH_Lithology();
                        DataRow[] myRowLth = dtLit.Select("[From] <= " + txtFrom.Text.ToString() + " and [To] >= " + txtTo.Text.ToString());
                        if (myRowLth.Length > 0)
                        {
                            if (myRowLth[0].Table.Rows[0]["Litho"].ToString() != cmbLithology.SelectedValue.ToString())
                            {
                                MessageBox.Show("Difference between litho-Samples and litho-Lithology ", "Samples", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        /* [Litho], [From], [To] */
                        ///Fin. Valida la informacion contra los datos de litologia



                        //Para limpiar la variable sDHSamplesID que se utiliza para modificar un dato
                        //cuando le doy doble clic en el registro
                        oSamp.sOpcion = "2";
                        FillLoggin();


                        //Implementar visualizar la ultima modificacion en pantalla
                        if (sEdit == "1")
                        {
                            if (gdLoggin.Rows.Count > 1)
                            {
                                DataTable dtSamp = (DataTable)gdLoggin.DataSource;
                                DataRow[] myRow = dtSamp.Select(@"SKDHSamples = '" + sDHSamplesID + "'");
                                int rowindex = dtSamp.Rows.IndexOf(myRow[0]);
                                gdLoggin.Rows[rowindex].Selected = true;
                                gdLoggin.CurrentCell = gdLoggin.Rows[rowindex].Cells[1];
                            }
                        }

                        sDHSamplesID = "0";


                        //Insertar el registro para el historial de transacciones por usuario
                        oRf.InsertTrans("DH_Collars", "Update", clsRf.sUser.ToString(),
                            "Hole ID: " + cmbHoleID.SelectedValue.ToString() + "." +
                            " Sample :" + txtSampNo.Text.ToString() + "." +
                            " From: " + txtFrom.Text.ToString() + "." +
                            " To: " + txtTo.Text.ToString() + "." +
                            "Sample TypD: " + cmbSampleType.SelectedValue.ToString());

                        if (sEdit == "0")
                        {
                            string sCons = clsDHSamples.sConsLoggin.ToString().Substring(
                                int.Parse(ConfigurationSettings.AppSettings["CantCaractLoggin"].ToString()));
                            sCons = (int.Parse(sCons.ToString()) + 1).ToString();

                            txtSampNo.Text = clsDHSamples.sConsLoggin.ToString().Substring(0, 1)
                                + sCons;
                            clsDHSamples.sConsLoggin = txtSampNo.Text.ToString();

                            txtTo.Text = "";
                            txtFrom.Text = clsDHSamples.sStaticFrom.ToString();
                            txtTo.Focus();
                            EnableControls();

                        }
                        else if (sEdit == "1")
                        {
                            DisableControls();
                            bHoleValid = false;
                        }

                        sEdit = "0";


                        /*  DataTable dtOri = dtOriginal();
                            DataRow[] drOrig = dtOri.Select("Value <> '" + cmbSampleType.SelectedValue.ToString() + "'");
                            if (drOrig.Length > 0)
                            {
                                txtFrom.Text = "-99";
                                txtTo.Text = "-99";
                            }
                         */


                        //DataTable dtOri = dtOriginal();
                        //DataRow[] drOrig = dtOri.Select("Value <> '" + cmbSampleType.SelectedValue.ToString() + "'");
                        //if (drOrig.Length > 0)
                        //{
                        //    clsDHSamples.sStaticFrom = txtTo.Text.ToString();
                        //}




                    }
                    else
                    {
                        MessageBox.Show("Error Insert: " + sResp.ToString(), "Samples", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }


                }
                else
                {
                    MessageBox.Show("'To' Invalid. 'To' greater than HoleId lenght");
                }
            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Samples", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        //private void ControlsClean()
        //{
        //    try
        //    {
        //        txtTo.Text = "";
        //        txtFrom.Text = "";
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message);
        //    }
        //}

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                sEdit = "0";


                //Valida que no este vacio
                if (txtSampNoIni.Text != "")
                {
                    EnableControls();
                }
                else if (txtSampNoIni.Text == "")
                {
                    MessageBox.Show("Empty Sample No. Init ", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DisableControls();
                    return;
                }


                //Valida que el sampleid no este repetido
                DataTable dtValSamp = SampleIdRepeat(txtSampNoIni.Text.ToString());
                bool bValSamp = false;
                if (dtValSamp.Rows.Count > 0)
                {
                    bValSamp = false;
                }
                else
                {
                    bValSamp = true;
                }

                if (bValSamp == false)
                {
                    MessageBox.Show("Sample duplicated to HoleID: " + dtValSamp.Rows[0]["HoleID"].ToString()
                    , "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //DisableControls();
                    return;
                }
                //bool bValSamp = SampleIdRepeat(txtSampNoIni.Text.ToString());
                //if (bValSamp == false)
                //{
                //    MessageBox.Show("Sample duplicated", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    DisableControls();
                //    return;
                //}

                //Valida que se haya seleccionado un registro en holeid
                if (cmbHoleID.SelectedValue.ToString() == "Select an option..".ToString())
                {
                    MessageBox.Show("Selected an option HoleId", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DisableControls();
                    return;
                }

                HoleIDValidate(cmbHoleID.SelectedValue.ToString());
                if (bHoleValid == false)
                {
                    MessageBox.Show("HoleId Invalid", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                DataTable dtInvSamp = oRf.getInvSamples();
                DataRow[] dato = dtInvSamp.Select("Sample = '" + txtSampNoIni.Text.ToString() + "'");
                if (dato.Length > 0)
                {
                    MessageBox.Show("Sample Valid", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    EnableControls();
                    txtFrom.Focus();

                    oSamp.sOpcion = "2";
                    FillLoggin();

                    clsDHSamples.sConsLoggin =  txtSampNoIni.Text.ToString().ToUpper();
                    txtSampNo.Text = clsDHSamples.sConsLoggin.ToString();

                }
                else
                {
                    MessageBox.Show("Sample Invalid", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DisableControls();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void HoleIDValidate(string _sHoleID)
        {
            try
            {
                oCollars.sHoleID = _sHoleID.ToString();
                oCollars.sLogged = clsRf.sUser;
                DataTable dtLogg = oCollars.getDHCollarsLogged();
                if (dtLogg.Rows.Count > 0)
                {
                    bHoleValid = true;
                    
                }
                else
                {
                    
                    bHoleValid = false;
                   
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void cmbHoleID_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                oSamp.sOpcion = "2";
                FillLoggin();


                if (sEdit == "1")
                {
                    return;
                }

                //Valido para que no me muestre mensajes de error al inicio del form
                if (bInicio == true)
                {
                    return;
                }

                if (cmbHoleID.SelectedValue.ToString() == "Select an option..".ToString())
                {
                    //MessageBox.Show("Selected an option HoleId", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DisableControls();
                    return;
                }

               

                //HoleIDValidate(cmbHoleID.SelectedValue.ToString());

                //oCollars.sHoleID = cmbHoleID.SelectedValue.ToString();
                //oCollars.sLogged = clsRf.sUser;
                //DataTable dtLogg = oCollars.getDHCollarsLogged();
                //if (dtLogg.Rows.Count > 0)
                //{
                //    bHoleValid = true;



                //    //Valida que no este vacio
                //    if (txtSampNoIni.Text != "")
                //    {
                //        EnableControls();
                //    }
                //    else if (txtSampNoIni.Text == "")
                //    {
                //        DisableControls(); 
                //        return;
                //    }
                    
                //    oSamp.sOpcion = "2";
                //    FillLoggin();

                //    //Valida que el campo Samp No Inicial no este repetido y que sea valido
                //    bool bValSamp = SampleIdRepeat(txtSampNoIni.Text.ToString());
                //    if (bValSamp == false)
                //    {
                //        MessageBox.Show("Sample No. duplicated", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //        return;
                //    }

                //}
                //else
                //{
                //    DisableControls();
                //    bHoleValid = false;
                //    MessageBox.Show("HoleId Invalid", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message); ;
            }
        }

        private DataTable dtOriginal()
        {
            DataTable dtOrig = new DataTable();
            dtOrig.Columns.Add("Key", typeof(String));
            dtOrig.Columns.Add("Value", typeof(String));


            for (int i = 0; i < conf.AppSettings.Settings.Count; i++)
            {
                if (conf.AppSettings.Settings.AllKeys[i].ToString().Contains("ORIGINAL"))
                {

                    DataRow drOrig = dtOrig.NewRow();
                    //drConect["Con"] = ;
                    drOrig["Key"] = conf.AppSettings.Settings.AllKeys[i].ToString();
                    drOrig["Value"] =
                        conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value.ToString();
                    dtOrig.Rows.Add(drOrig);

                }

            }

            return dtOrig;
        }

        private DataTable dtDupDe()
        {
            DataTable dtDup = new DataTable();
            dtDup.Columns.Add("Key", typeof(String));
            dtDup.Columns.Add("Value", typeof(String));


            for (int i = 0; i < conf.AppSettings.Settings.Count; i++)
            {
                if (conf.AppSettings.Settings.AllKeys[i].ToString().Contains("DupDe"))
                {

                    DataRow drDup = dtDup.NewRow();
                    //drConect["Con"] = ;
                    drDup["Key"] = conf.AppSettings.Settings.AllKeys[i].ToString();
                    drDup["Value"] =
                        conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value.ToString();
                    dtDup.Rows.Add(drDup);

                }

            }

            return dtDup;
        }

        private void cmbSampleType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (sEdit == "1")
                {
                    return;
                }


                //Valido para que no me muestre mensajes de error al inicio del form
                if (bInicio == true)
                {
                    return;
                }



                DataTable dtOri = dtOriginal();
                DataRow[] drOrig = dtOri.Select("Value <> '" + cmbSampleType.SelectedValue.ToString() + "'");
                if (drOrig.Length > 0)
                {
                    txtFrom.Text = "-99";
                    txtTo.Text = "-99";
                }
                

                DataTable dtDup = dtDupDe();
                DataRow[] dato = dtDup.Select("Value = '" + cmbSampleType.SelectedValue.ToString() + "'");
                if (dato.Length > 0)
                {

                    /*string sCons = clsDHSamples.sConsLoggin.ToString().Substring(
                            int.Parse(ConfigurationSettings.AppSettings["CantCaractLoggin"].ToString()));
                        sCons = (int.Parse(sCons.ToString()) + 1).ToString();

                        txtSampNo.Text = clsDHSamples.sConsLoggin.ToString().Substring(0, 1)
                            + sCons;
                        clsDHSamples.sConsLoggin = txtSampNo.Text.ToString();*/

                    string sCons = clsDHSamples.sConsLoggin.ToString().Substring(
                           int.Parse(ConfigurationSettings.AppSettings["CantCaractLoggin"].ToString()));
                    sCons = (int.Parse(sCons.ToString()) - 1).ToString();

                    txtDupDe.Text = clsDHSamples.sConsLoggin.ToString().Substring(0, 1)
                        + sCons;

                    txtFrom.Text = "-99";
                    txtTo.Text = "-99";
                    
                }
                else
                {
                    txtDupDe.Text = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                oXL = new Excel.Application();
                oXL.Visible = true;
                //Get a new workbook.
                //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                //oWB = oXL.Workbooks.Open(@"D:/Template_Shipment_Sgs.xls", 0, true, 5,


                oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_Logging"].ToString(),
                    0, false, 5,
                Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                Type.Missing, false, false, false);
                /*
                    0, true, 5,
                Type.Missing, Type.Missing, false, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, null, null);
                */


                //oXL.Workbooks.Add().SaveAs(sName.ToString(),
                //    Microsoft.Office.Interop.Excel.XlFileFormat.xlTextWindows,
                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, Type.Missing, Type.Missing);

                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 6] = cmbHoleID.SelectedValue.ToString();

                int iInicial = 6;
                for (int i = 0; i < gdLoggin.Rows.Count - 1; i++)
                {

                    oSheet.Cells[iInicial, 3] = gdLoggin.Rows[i].Cells["From"].Value.ToString();
                    oSheet.Cells[iInicial, 4] = gdLoggin.Rows[i].Cells["To"].Value.ToString();
                    oSheet.Cells[iInicial, 5] = gdLoggin.Rows[i].Cells["Sample"].Value.ToString();
                    oSheet.Cells[iInicial, 6] = gdLoggin.Rows[i].Cells["SampleType"].Value.ToString();

                    iInicial += 1;
                }

                oXL.Visible = true;
                oXL.UserControl = true;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        //[TestMethod()]
        //public void ExcelTest() 
        //{ 
        //    Microsoft.Office.Interop.Excel.Application excelApplication = new Application(); 
        //    string file = @"C:\testsheet.xls"; 
        //    Microsoft.Office.Interop.Excel.Workbook wkb = excelApplication.Workbooks.Open(file, 0, false, 5,
        //        Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
        //        Type.Missing, false, false, false); 
        //    Microsoft.Office.Interop.Excel.Worksheet wks = 
        //        wkb.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet; 
        //    wks.SaveAs(@"C:\savedFile.txt", Microsoft.Office.Interop.Excel.XlFileFormat.xlTextWindows, 
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
        //        Type.Missing, Type.Missing); } 
        
        private void gdLoggin_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //if (MessageBox.Show("Row Edit" + " '" + gdLoggin.Rows[e.RowIndex].Cells["Sample"].Value.ToString() + "' "
                //    , "Logging", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                //    MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                //{
                    sEdit = "1";
                    EnableControls();
                    //cmbHoleID.SelectedValue = gdLoggin.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                    HoleIDValidate(gdLoggin.Rows[e.RowIndex].Cells["HoleID"].Value.ToString());

                    if (bHoleValid == true)
                    {
                        sDHSamplesID = gdLoggin.Rows[e.RowIndex].Cells["SKDHSamples"].Value.ToString();
                        txtSampNo.Text = gdLoggin.Rows[e.RowIndex].Cells["Sample"].Value.ToString();
                        txtFrom.Text = gdLoggin.Rows[e.RowIndex].Cells["From"].Value.ToString();
                        txtTo.Text = gdLoggin.Rows[e.RowIndex].Cells["To"].Value.ToString();
                        txtDupDe.Text = gdLoggin.Rows[e.RowIndex].Cells["DupDe"].Value.ToString();
                        txtCommentsSamp.Text = gdLoggin.Rows[e.RowIndex].Cells["Comments"].Value.ToString();
                        cmbLithology.SelectedValue = gdLoggin.Rows[e.RowIndex].Cells["Lithology"].Value.ToString() == ""
                                ? "-1" : gdLoggin.Rows[e.RowIndex].Cells["Lithology"].Value.ToString();

                        cmbSampleType.SelectedValue = gdLoggin.Rows[e.RowIndex].Cells["SampleType"].Value.ToString();
                        cmbHoleID.SelectedValue = gdLoggin.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();

                        
                    }
                    else
                    {
                        MessageBox.Show("HoleId Invalid. Not Allow ", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                //}
            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }


        private void gdLoggin_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete" + " '" + gdLoggin.Rows[e.RowIndex].Cells["Sample"].Value.ToString() + "' "
                    , "Logging", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    sEdit = "0";
                    
                    HoleIDValidate(gdLoggin.Rows[e.RowIndex].Cells["HoleID"].Value.ToString());

                    if (bHoleValid == true)
                    {
                        
                        oSamp.iDHSampID = int.Parse(gdLoggin.Rows[e.RowIndex].Cells["SKDHSamples"].Value.ToString());
                        string sRespDel = oSamp.DHSamples_DeleteLoggin();
                        if (sRespDel == "OK")
                        {
                            FillLoggin();
                        }

                    }
                    else
                    {
                        MessageBox.Show("HoleId Invalid. Not Allow ", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void txtToGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }

            //TabEnter(e);

        }

        private void txtFromGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }


            //TabEnter(e);

        }


        #region Geotech

        private void cmbHoleIdGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
            //if (e.KeyChar == (char)(Keys.Enter))
            //{
            //    e.Handled = true;
            //    SendKeys.Send("{TAB}");
            //}
        }

        private void cmbLithGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtDifferGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtJnGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtJrGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtJaGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbDegreeBreak_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbHardness_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtComments_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private AutoCompleteStringCollection AutoCompleteCmb(DataTable _dtAutoComplete)
        {
            try
            {

                AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();
                foreach (DataRow row in _dtAutoComplete.Rows)
                {
                    stringCol.Add(Convert.ToString(row["Comb"]));
                }
                
                return stringCol;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FillCmbGeoTech()
        {
            try
            {
                //cmbHoleIdGeo
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCGeo = dtCollars.NewRow();
                drCGeo[0] = "Select an option..";
                dtCollars.Rows.Add(drCGeo);
                cmbHoleIdGeo.DisplayMember = "HoleID";
                cmbHoleIdGeo.ValueMember = "HoleID";
                cmbHoleIdGeo.DataSource = dtCollars;
                cmbHoleIdGeo.SelectedValue = "Select an option..";

                DataTable dtLithology = new DataTable();
                dtLithology = oRf.getDsRfLithology().Tables[1];

                DataRow drL = dtLithology.NewRow();
                drL[0] = "-1";
                drL[1] = "Select an option..";
                dtLithology.Rows.Add(drL);

                cmbLithGeo.DisplayMember = "Comb";
                cmbLithGeo.ValueMember = "Code";
                cmbLithGeo.DataSource = dtLithology;
                cmbLithGeo.SelectedValue = -1;

                DataTable dtDegreeBreak = new DataTable();
                dtDegreeBreak = oRf.getRfGeotechBreak();

                DataRow drD = dtDegreeBreak.NewRow();
                drD[0] = "-1";
                drD[1] = "Select an option..";
                dtDegreeBreak.Rows.Add(drD);

                cmbDegreeBreak.DisplayMember = "Comb";
                cmbDegreeBreak.ValueMember = "Category";
                cmbDegreeBreak.DataSource = dtDegreeBreak;
                cmbDegreeBreak.SelectedValue = -1;

                //Autocomplete cmbDegreeBreak
                //AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();
                //foreach (DataRow row in dtDegreeBreak.Rows)
                //{
                //    stringCol.Add(Convert.ToString(row["Comb"]));
                //}
                cmbDegreeBreak.AutoCompleteCustomSource = AutoCompleteCmb(dtDegreeBreak);
                //cmbDegreeBreak.AutoCompleteMode = AutoCompleteMode.Suggest;
                //cmbDegreeBreak.AutoCompleteSource = AutoCompleteSource.CustomSource;
                //Fin Autocomplete cmbDegreeBreak



                //getRfGeotechHardness cmbHardness
                DataTable dtHardness = new DataTable();
                dtHardness = oRf.getRfGeotechHardness();

                DataRow drH = dtHardness.NewRow();
                drH[0] = "-1";
                drH[1] = "Select an option..";
                dtHardness.Rows.Add(drH);

                cmbHardness.DisplayMember = "Comb";
                cmbHardness.ValueMember = "Id";
                cmbHardness.DataSource = dtHardness;
                cmbHardness.SelectedValue = -1;

            }
            catch (Exception ex)
            {
                throw new Exception("Error FillCmbGeoTech: " + ex.Message);
            }
        }

        private string ControlsValidate()
        {
            try
            {
                string sresp = "";

                if (cmbHoleIdGeo.SelectedValue.ToString() == "Select an option..")
                {
                    sresp = "Selected an option Hole ID";
                    return sresp;
                }
                if (txtFromGeo.Text == "" || txtToGeo.Text == "")
                {
                    sresp = "Empty From or To";
                    return sresp;
                }
                //if (txtFromGeo.Text != "-99")
                //{
                    //if (double.Parse(txtFromGeo.Text.ToString()) < 0 || double.Parse(txtToGeo.Text.ToString()) < 0)
                    //{
                        if (double.Parse(txtFromGeo.Text.ToString()) >= double.Parse(txtToGeo.Text.ToString()))
                        {
                            sresp = " 'From' greater than 'To'";
                            return sresp;
                        }
                    //}
                    //return sresp = "From or To must be greater than zero (0)";
                //}
                

                oCollars.sHoleID = cmbHoleIdGeo.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToGeo.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'To' greater than Hole Id lenght";
                    return sresp;
                }



                //if (txtJoinCondition.Text.ToString() == "")
                //{ txtJoinCondition.Text = "0"; }
                
                //if (double.Parse(txtJoinCondition.Text.ToString()) < 0
                //    || double.Parse(txtJoinCondition.Text.ToString()) > 25)
                //{
                //    sresp = "Join Condition less than 0 or greater than 25";
                //    return sresp;
                //}
                //if (txtRec_mGeo.Text == "" || txtRQD_cmGeo.Text == "")
                //{
                //    sresp = "Empty Rec m or RQD cm";
                //    return sresp;
                //}
                //if (txtRec_mGeo.Text.ToString() == "0" && double.Parse(txtNumOfFact.Text.ToString()) <= 30)
                //{
                //    sresp = "Illegal value when Rec is zero (0). It Must be greater than thirty (30)";
                //    return sresp;
                //}

                //if (double.Parse(txtRec_PorcGeo.Text.ToString()) > 110)
                //{
                //    sresp = "Illegal value Perc Rec cm";
                //    txtRQD_cmGeo.Text = "";
                //    return sresp;
                //}
                //if (double.Parse(txtRQD_PorcGeo.Text.ToString()) > 110)
                //{
                //    sresp = "Illegal value Perc RQD cm";
                //    txtRQD_cmGeo.Text = "";
                //    return sresp;
                //}



                //if (double.Parse(txtRQD_cmGeo.Text.ToString()) >= 
                //    (double.Parse(txtRec_mGeo.Text.ToString()) * 110))
                //{
                //    sresp = "Illegal value RQD cm";
                //    txtRQD_cmGeo.Text = "";
                //    return sresp;
                //}


                if (txtRQD_cmGeo.Text.ToString() != "")
                {
                    if (txtRQD_cmGeo.Text.ToString() == "-99")
                    {
                        goto ContinuarValid;
                    }
                    if (double.Parse(txtRQD_cmGeo.Text.ToString()) < 10
                    && double.Parse(txtRQD_cmGeo.Text.ToString()) > 0)
                    {
                        sresp = "Illegal value RQD cm. less than ten (10)";
                        txtRQD_cmGeo.Text = "";
                        return sresp;
                    }

                }
                ContinuarValid:
                

                return sresp;
            }
            catch (Exception ex)
            {
               throw new Exception(ex.Message);
            }
        }

        private void btnAddGeo_Click(object sender, EventArgs e)
        {
            try
            {

                if (double.Parse(txtFromGeo.Text.ToString()) == double.Parse(txtToGeo.Text.ToString()))
                {
                    MessageBox.Show(" 'From' equal to 'To'");
                    return;
                }

                if (double.Parse(txtFromGeo.Text.ToString()) > double.Parse(txtToGeo.Text.ToString()))
                {
                    MessageBox.Show("'From' greater than 'To'");
                    return;
                }

                if (txtRec_mGeo.Text.ToString() != "")
                {
                    double porc = (double.Parse(txtRec_mGeo.Text.ToString()) /
                        double.Parse(txtDifferGeo.Text.ToString()) * 100);
                    if (porc > 110)
                    {
                        MessageBox.Show("Rec (m) Invalid", "GeoTech", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    txtRec_PorcGeo.Text = (double.Parse(txtRec_mGeo.Text.ToString()) /
                        double.Parse(txtDifferGeo.Text.ToString()) * 100).ToString();
                }

                

                if (txtRec_mGeo.Text.ToString() == "-99")
                {
                    txtRec_PorcGeo.Text = "-99";
                }


                if (txtRQD_cmGeo.Text.ToString() != "")
                {
                    double porcRQ = (double.Parse(txtRQD_cmGeo.Text.ToString()) /
                        double.Parse(txtDifferGeo.Text.ToString()));
                    if (porcRQ > 110)
                    {
                        MessageBox.Show("RQD (m) Invalid", "GeoTech", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    txtRQD_PorcGeo.Text = (double.Parse(txtRQD_cmGeo.Text.ToString()) /
                        double.Parse(txtDifferGeo.Text.ToString())).ToString();
                }
                if (txtRQD_cmGeo.Text.ToString() == "-99")
                {
                    txtRQD_PorcGeo.Text = "-99";
                }

                string sResp = ControlsValidate().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "GeoTech", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (cmbLithGeo.SelectedValue.ToString() != "Select an option..")
                {
                    if (cmbLithGeo.SelectedValue.ToString() == "QB"
                    || cmbLithGeo.SelectedValue.ToString() == "QS"
                    || cmbLithGeo.SelectedValue.ToString() == "RBL"
                    || cmbLithGeo.SelectedValue.ToString() == "SRP")
                    {
                        txtRec_mGeo.Text = "-99";
                        txtRQD_cmGeo.Text = "-99";
                        txtNumOfFact.Text = "-99";
                        txtJoinCondition.Text = "-99";
                        txtJnGeo.Text = "-99";
                        txtJrGeo.Text = "-99";
                        txtJaGeo.Text = "-99";
                        cmbDegreeBreak.SelectedValue = "-1";
                        cmbHardness.SelectedValue = "-1";
                    }
                    /**/
                    else if (cmbLithGeo.SelectedValue.ToString() == "VOI"
                    || cmbLithGeo.SelectedValue.ToString() == "WRK")
                    {
                        txtRec_mGeo.Text = "0";
                        txtRQD_cmGeo.Text = "-99";
                        txtNumOfFact.Text = "-99";
                        txtJoinCondition.Text = "-99";
                        txtJnGeo.Text = "-99";
                        txtJrGeo.Text = "-99";
                        txtJaGeo.Text = "-99";
                        cmbDegreeBreak.SelectedValue = "-1";
                        cmbHardness.SelectedValue = "-1";
                    }
                }
                

                //Valida que el rango sea valido para el pozo
                DataTable dtValidRange = new DataTable();
                oGeo.iFrom = double.Parse(txtFromGeo.Text.ToString());
                oGeo.iTo = double.Parse(txtToGeo.Text.ToString());
                oGeo.sHoleID = cmbHoleIdGeo.SelectedValue.ToString();

                dtValidRange = oGeo.getDHGeotechFromToValid();
                if (dtValidRange.Rows.Count > 0)
                {
                    MessageBox.Show("Range 'From To' Overlaps", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                if (sEditGeo == "1")
                {
                    oGeo.sOpcion = "2";
                }
                else
                {
                    oGeo.iDHGeotechID = 0;
                    oGeo.sOpcion = "1";
                }
                
                oGeo.sHoleID = cmbHoleIdGeo.SelectedValue.ToString();
                if (dgGeotech.Rows.Count <= 1)
                {
                    oGeo.iFrom = 0;
                }
                else { oGeo.iFrom = double.Parse(txtFromGeo.Text.ToString()); }
                oGeo.iTo = double.Parse(txtToGeo.Text.ToString());

                if (cmbLithGeo.SelectedValue.ToString() == "-1" || cmbLithGeo.SelectedValue.ToString() == "")
                    oGeo.sLithCod = null;
                else oGeo.sLithCod = cmbLithGeo.SelectedValue.ToString();

                if (txtRec_mGeo.Text.ToString() == "")
                    oGeo.dRecm =  null;
                else oGeo.dRecm = double.Parse(txtRec_mGeo.Text.ToString());
             
                //oGeo.dRecm = txtRec_mGeo.Text.ToString() == "" ? null : double.Parse(txtRec_mGeo.Text.ToString());
                if (txtRQD_cmGeo.Text.ToString() == "" )
                    oGeo.dRQDcm = null;
                else oGeo.dRQDcm = double.Parse(txtRQD_cmGeo.Text.ToString());

                if (txtNumOfFact.Text.ToString() == "")
                    oGeo.dNoOfFract = null;
                else oGeo.dNoOfFract = double.Parse(txtNumOfFact.Text.ToString());

                if (txtJoinCondition.Text.ToString() == "")
                    oGeo.dJoinCond = null;
                else oGeo.dJoinCond = double.Parse(txtJoinCondition.Text.ToString());

                if (txtJrGeo.Text.ToString() == "")
                    oGeo.dJr =  null;
                else oGeo.dJr = double.Parse(txtJrGeo.Text.ToString());

                if (txtJnGeo.Text.ToString() == "" )
                    oGeo.dJn = null;
                else oGeo.dJn =  double.Parse(txtJnGeo.Text.ToString());

                if (txtJaGeo.Text.ToString() == "")
                    oGeo.dJa = null;
                else oGeo.dJa = double.Parse(txtJaGeo.Text.ToString());

                if (cmbDegreeBreak.SelectedValue.ToString() == "-1" || cmbDegreeBreak.SelectedValue.ToString() == "")
                    oGeo.sDegBreak = null;
                else oGeo.sDegBreak = cmbDegreeBreak.SelectedValue.ToString();

                if (cmbHardness.SelectedValue.ToString() == "-1" || cmbHardness.SelectedValue.ToString() == "")
                    oGeo.sHardness = null;
                else oGeo.sHardness = cmbHardness.SelectedValue.ToString();

                if (txtComments.Text.ToString() == "-1" || txtComments.Text.ToString() == "")
                    oGeo.sComments = null;
                else oGeo.sComments = txtComments.Text.ToString();


                oGeo.sAltWeath = null;

                clsDHGeotech.sStaticFrom = txtToGeo.Text.ToString();

                string sAdd = oGeo.DH_Geotech_Add();
                if (sAdd.ToString() == "OK")
                {
                    FilldtGeo("2");
                    //MessageBox.Show("Saved");
                    
                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Geotech", sEditGeo == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIdGeo.SelectedValue.ToString() + "." +
                        " From: " + txtFromGeo.Text.ToString() + "." +
                        " To: " + txtToGeo.Text.ToString() + "." +
                        " Lithology: " + cmbLithGeo.SelectedValue.ToString() == "Select an option.." || cmbLithGeo.SelectedValue.ToString() == ""
	                        ? "" : cmbLithGeo.SelectedValue.ToString() + "." +
                        " dRecm: " + txtRec_mGeo.Text.ToString() +"." +
                        " dRQDcm: " + txtRQD_cmGeo.Text.ToString() + "." +
                        " dNoFact: " + txtNumOfFact.Text.ToString() + "." +
                        " Join Condition: " + txtJoinCondition.Text.ToString() + "." +
                        " Degree BreakagD: " + cmbDegreeBreak.SelectedValue.ToString() == "Select an option.." || cmbDegreeBreak.SelectedValue.ToString() == ""
                            ? "" : cmbDegreeBreak.SelectedValue.ToString() + "." +
                        " Hardness: " + cmbHardness.SelectedValue.ToString() == "Select an option.." || cmbHardness.SelectedValue.ToString() == ""
                               ? "" : cmbHardness.SelectedValue.ToString());


                    sEditGeo = "0";

                    CleanControlsGeo();

                    txtFromGeo.Text = clsDHGeotech.sStaticFrom.ToString();

                    txtToGeo.Focus();

                }
                else
                {
                    MessageBox.Show("Error Insert: " + sAdd.ToString(), "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }
        }

        private void cmbHoleIDForm_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                cmbHoleID.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIdGeo.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIdLit.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIdWeat.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIDSt.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIdMin.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIDBox.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIDAlt.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                cmbHoleIdDens.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
                //cmbHoleIDOx.SelectedValue = cmbHoleIDForm.SelectedValue.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); ;
            }
        }

        
        private void FilldtGeo(string _sOpcion)
        {
            try
            {
                DataTable dtGeo = new DataTable();
                oGeo.sOpcion = _sOpcion;
                oGeo.sHoleID = cmbHoleIdGeo.SelectedValue.ToString();
                dtGeo = oGeo.getDH_Geotech();
                dgGeotech.DataSource = dtGeo;

                dgGeotech.Columns["SKDHGeotech"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void cmbHoleIdGeo_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldtGeo("2");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); ;
            }
        }

        private void CleanControlsGeo()
        {
            try
            {
                //cmbHoleIdGeo.SelectedValue = "";
                //txtFrom.Text = "";
                txtToGeo.Text = "";
                //cmbLithGeo.SelectedValue = "Select an option..";
                txtRec_mGeo.Text = "";
                txtRQD_cmGeo.Text ="";
                txtNumOfFact.Text = "";
                txtJoinCondition.Text = "";
                txtJrGeo.Text = "";
                txtJnGeo.Text = "";
                txtJaGeo.Text = "";
                cmbDegreeBreak.SelectedValue = "-1";
                cmbHardness.SelectedValue = "-1";
                txtComments.Text = "";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgGeotech_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                oGeo.iDHGeotechID = Int64.Parse(dgGeotech.Rows[e.RowIndex].Cells["SKDHGeotech"].Value.ToString());
                sEditGeo = "1";

                cmbHoleIdGeo.SelectedValue = dgGeotech.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtFromGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["To"].Value.ToString();
                txtRec_mGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["Recm"].Value.ToString();
                txtRQD_cmGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["RQDcm"].Value.ToString();
                txtNumOfFact.Text = dgGeotech.Rows[e.RowIndex].Cells["NoOfFract"].Value.ToString();
                txtJoinCondition.Text = dgGeotech.Rows[e.RowIndex].Cells["JointCond"].Value.ToString();
                txtJrGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["Jr"].Value.ToString();
                txtJnGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["Jn"].Value.ToString();
                txtJaGeo.Text = dgGeotech.Rows[e.RowIndex].Cells["Ja"].Value.ToString();
                txtComments.Text = dgGeotech.Rows[e.RowIndex].Cells["Comments"].Value.ToString();

                cmbLithGeo.SelectedValue = dgGeotech.Rows[e.RowIndex].Cells["LithCod"].Value.ToString() == ""
                    ? "-1" : dgGeotech.Rows[e.RowIndex].Cells["LithCod"].Value.ToString();
                cmbDegreeBreak.SelectedValue = dgGeotech.Rows[e.RowIndex].Cells["DegBreak"].Value.ToString() == ""
                    ? "-1" : dgGeotech.Rows[e.RowIndex].Cells["DegBreak"].Value.ToString();
                cmbHardness.SelectedValue = dgGeotech.Rows[e.RowIndex].Cells["Hardness"].Value.ToString() == ""
                    ? "-1" : dgGeotech.Rows[e.RowIndex].Cells["Hardness"].Value.ToString();

                GetDifferGeo();

            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControlsGeo();
                sEditGeo = "0";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgGeotech_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Hole Id" + dgGeotech.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                    + " From " + dgGeotech.Rows[e.RowIndex].Cells["From"].Value.ToString() 
                    + " To " + dgGeotech.Rows[e.RowIndex].Cells["To"].Value.ToString()
                    , "GeoTech", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        oGeo.sHoleID = dgGeotech.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                        oGeo.iFrom = double.Parse(dgGeotech.Rows[e.RowIndex].Cells["From"].Value.ToString());
                        string sDelete = oGeo.DH_Geotech_Delete();
                        if (sDelete == "OK")
                        {
                            MessageBox.Show("Row Deleted", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            FilldtGeo("2");
                            sEdit = "0";
                            CleanControlsGeo();
                        }
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtJoinCondition_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            //TabEnter(e);
        }

        private void txtRQD_PorcGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            //TabEnter(e);
        }

        private void txtRec_PorcGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            //TabEnter(e);
        }

        private void txtRec_mGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }

            //TabEnter(e);
        }

        private void txtRQD_cmGeo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            //TabEnter(e);
        }

        private void txtNumOfFact_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            //TabEnter(e);
        }

        #endregion

        #region Lithology

        private void FillCmbLith()
        {
            try
            {
                //cmbHoleIdGeo
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCGeo = dtCollars.NewRow();
                drCGeo[0] = "Select an option..";
                dtCollars.Rows.Add(drCGeo);
                cmbHoleIdLit.DisplayMember = "HoleID";
                cmbHoleIdLit.ValueMember = "HoleID";
                cmbHoleIdLit.DataSource = dtCollars;
                cmbHoleIdLit.SelectedValue = "Select an option..";

                DataTable dtLithology = new DataTable();
                dtLithology = oRf.getDsRfLithology().Tables[1];

                DataRow drL = dtLithology.NewRow();
                drL[0] = "-1";
                drL[1] = "Select an option..";
                dtLithology.Rows.Add(drL);

                cmbLithologyLit.DisplayMember = "Comb";
                cmbLithologyLit.ValueMember = "Code";
                cmbLithologyLit.DataSource = dtLithology;
                cmbLithologyLit.SelectedValue = -1;

                cmbLithoDens.DisplayMember = "Comb";
                cmbLithoDens.ValueMember = "Code";
                cmbLithoDens.DataSource = dtLithology.Copy();
                cmbLithoDens.SelectedValue = -1;


            }
            catch (Exception ex)
            {
                throw new Exception("Error FillCmbGeoTech: " + ex.Message);
            }
        }

        private string ControlsValidateLit()
        {
            try
            {
                string sresp = "";

                oCollars.sHoleID = cmbHoleIdLit.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToLit.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'To' greater than Hole Id lenght";
                    return sresp;
                }

                if (cmbHoleIdLit.SelectedValue.ToString() == "Select an option..")
                {
                    sresp = "Selected an option Hole ID";
                    return sresp;
                }
                if (txtFromLit.Text == "" || txtToLit.Text == "")
                {
                    sresp = "Empty From or To";
                    return sresp;
                }
                if (txtFromLit.Text != "-99")
                {
                    //if (double.Parse(txtFromLit.Text.ToString()) < 0 || double.Parse(txtToLit.Text.ToString()) < 0)
                    //{
                    if (double.Parse(txtFromLit.Text.ToString()) == double.Parse(txtToLit.Text.ToString()))
                    {
                        sresp = " 'From' equal to 'To'";
                        return sresp;
                    }

                    if (double.Parse(txtFromLit.Text.ToString()) > double.Parse(txtToLit.Text.ToString()))
                    {
                        sresp = " 'From' greater than 'To'";
                        return sresp;
                    }
                    //}
                    //return sresp = "From or To must be greater than zero (0)";
                }
                
                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FilldgLithology(string _sOpcion)
        {
            try
            {
                DataTable dtLit = new DataTable();
                oLit.sOpcion = _sOpcion;
                oLit.sHoleID = cmbHoleIdLit.SelectedValue.ToString();
                dtLit = oLit.getDH_Lithology();
                dgLithology.DataSource = dtLit;

                dgLithology.Columns["SKDHLithology"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private void btnAddLit_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateLit().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ///Implementar para que no permita que el from to se solape con algun otro registro
                if (sEditLit == "0")
                {
                    //Valida que el rango sea valido para el pozo
                    DataTable dtValidRange = new DataTable();
                    oLit.dFrom = double.Parse(txtFromLit.Text.ToString());
                    oLit.dTo = double.Parse(txtToLit.Text.ToString());
                    oLit.sHoleID = cmbHoleIdLit.SelectedValue.ToString();
                    dtValidRange = oLit.getDHLitFromToValid();
                    if (dtValidRange.Rows.Count > 0)
                    {
                        MessageBox.Show("Range 'From To' Overlaps", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                

                if (sEditLit == "1")
                {
                    oLit.sOpcion = "2";
                }
                else {
                    oLit.iDHLithologyID = 0;
                    oLit.sOpcion = "1"; 
                }
                
                oLit.sHoleID = cmbHoleIdLit.SelectedValue.ToString();
                
                
                if (txtObservLit.Text.ToString() == "")
                    oLit.sObservation = null;
                else oLit.sObservation = txtObservLit.Text.ToString();

                if (dgLithology.Rows.Count == 1)
                {
                    oLit.dFrom = 0;
                }
                else { oLit.dFrom = double.Parse(txtFromLit.Text.ToString()); }
                oLit.dTo = double.Parse(txtToLit.Text.ToString());
                oLit.sLithCode = cmbLithologyLit.SelectedValue.ToString();
                
                //oLit.sGSize = cmbGsizeLith.SelectedValue != null ? cmbGsizeLith.SelectedValue.ToString() : "-1";
                if (cmbGsizeLith.SelectedValue.ToString() == "-1" || cmbGsizeLith.SelectedValue.ToString() == "")
                    oLit.sGSize = null;
                else oLit.sGSize = cmbGsizeLith.SelectedValue.ToString();

                //oLit.sTextures = cmbTexturesLith.SelectedValue.ToString() != null ? cmbTexturesLith.SelectedValue.ToString() : "-1";
                if (cmbTexturesLith.SelectedValue.ToString() == "-1" || cmbTexturesLith.SelectedValue.ToString() == "")
                    oLit.sTextures = null;
                else oLit.sTextures = cmbTexturesLith.SelectedValue.ToString();

                if (cmbLithologyLit.SelectedValue.ToString() == "-1" || cmbLithologyLit.SelectedValue.ToString() == "")
                    oLit.sLithCode = "";
                else oLit.sLithCode = cmbLithologyLit.SelectedValue.ToString();

                clsDHLithology.sStaticFrom = txtToLit.Text.ToString();

                string sLit = oLit.DH_Lithology_Add();
                if (sLit == "OK")
                {

                    DataTable dtSamp = new DataTable();
                    oSamp.sOpcion = "2";
                    oSamp.sHoleID = cmbHoleIdLit.SelectedValue.ToString();
                    dtSamp = oSamp.getDHSamplesList();
                    DataRow[] myRowSamp = dtSamp.Select("[From] >= " + txtFromLit.Text.ToString() + " and [To] <= " + txtToLit.Text.ToString());
                    if (myRowSamp.Length > 0)
                    {
                        for (int i = 0; i < myRowSamp.Count(); i++)
                        {
                            if (myRowSamp[i].Table.Rows[i]["Lithology"].ToString() != cmbLithologyLit.SelectedValue.ToString())
                            {
                                MessageBox.Show("Difference between litho-Lithology and litho-Samples. SamplD: " +
                                    myRowSamp[0].ItemArray[2].ToString(), "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }

                    //MessageBox.Show("Saved");
                    FilldgLithology("2");

                    if (sEditLit == "1")
                    {
                        if (dgLithology.Rows.Count > 1)
                        {
                            DataTable dt = (DataTable)dgLithology.DataSource;
                            DataRow[] myRow = dt.Select(@"SKDHLithology = '" + oLit.iDHLithologyID + "'");
                            int rowindex = dt.Rows.IndexOf(myRow[0]);
                            dgLithology.Rows[rowindex].Selected = true;
                            dgLithology.CurrentCell = dgLithology.Rows[rowindex].Cells[1];
                        }
                    }



                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Lithology", sEditLit == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIdLit.SelectedValue.ToString() + "." +
                        //" sGSize :" + cmbGsizeLith.Text != "" ? cmbGsizeLith.SelectedValue.ToString() : "-1" + "." +
                        " From: " + txtFromLit.Text.ToString() + "." +
                        " To: " + txtToLit.Text.ToString() + "." +
                        " TexturD: " + cmbTexturesLith.SelectedValue.ToString());


                    CleanControlsLit();

                    txtFromLit.Text = clsDHLithology.sStaticFrom.ToString();
                    txtToLit.Focus();

                }
                else
                {
                    MessageBox.Show("Error Insert: " + sLit.ToString(), "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                sEditLit = "0";

            } 
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }       
        }

        private void CleanControlsLit()
        {
            try
            {
                txtToLit.Text = "0";
                txtObservLit.Text = "";
                cmbLithologyLit.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void txtFromLit_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //////TabEnter(e);
        }

        private void txtToLit_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //////TabEnter(e);
        }

        private void cmbHoleIdLit_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldgLithology("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgLithology_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditLit = "1";
                oLit.iDHLithologyID = Int64.Parse(dgLithology.Rows[e.RowIndex].Cells["SKDHLithology"].Value.ToString());

                cmbHoleIdLit.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtObservLit.Text = dgLithology.Rows[e.RowIndex].Cells["Observation"].Value.ToString();
                txtFromLit.Text = dgLithology.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToLit.Text = dgLithology.Rows[e.RowIndex].Cells["To"].Value.ToString();
                cmbLithologyLit.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Litho"].Value.ToString();

                cmbGsizeLith.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["GSize"].Value.ToString() == "" ?
                    "-1" : dgLithology.Rows[e.RowIndex].Cells["GSize"].Value.ToString();
                cmbTexturesLith.SelectedValue = dgLithology.Rows[e.RowIndex].Cells["Textures"].Value.ToString() == "" ?
                    "-1" : dgLithology.Rows[e.RowIndex].Cells["Textures"].Value.ToString();

            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelLit_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControlsLit();
                sEditLit = "0";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgLithology_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (MessageBox.Show("Row Delete. " + "Hole Id" + dgLithology.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                    + " From " + dgLithology.Rows[e.RowIndex].Cells["From"].Value.ToString()
                    + " To " + dgLithology.Rows[e.RowIndex].Cells["To"].Value.ToString()
                    , "Lithology", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                oLit.sHoleID = dgLithology.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                oLit.dFrom = double.Parse(dgLithology.Rows[e.RowIndex].Cells["From"].Value.ToString());
                string sDelete = oLit.DH_Lithology_Delete();
                if (sDelete == "OK")
                {
                    MessageBox.Show("Row Deleted", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FilldgLithology("2");
                    sEditLit = "0";
                    CleanControlsLit();
                }
            }
        }

        #endregion

        #region Weathering

        private void FillCmbWeath()
        {
            try
            {
                //cmbHoleIdGeo
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCGeo = dtCollars.NewRow();
                drCGeo[0] = "Select an option..";
                dtCollars.Rows.Add(drCGeo);
                cmbHoleIdWeat.DisplayMember = "HoleID";
                cmbHoleIdWeat.ValueMember = "HoleID";
                cmbHoleIdWeat.DataSource = dtCollars;
                cmbHoleIdWeat.SelectedValue = "Select an option..";

                DataTable dtWeathering = new DataTable();
                dtWeathering = oRf.getWeathering();

                DataRow drW = dtWeathering.NewRow();
                drW[0] = "-1";
                drW[1] = "Select an option..";
                dtWeathering.Rows.Add(drW);

                cmbWeatheringWeat.DisplayMember = "Comb";
                cmbWeatheringWeat.ValueMember = "Grade";
                cmbWeatheringWeat.DataSource = dtWeathering;
                cmbWeatheringWeat.SelectedValue = -1;

                DataTable dtOxidation = new DataTable();
                dtOxidation = oRf.getRfOxidation_List();

                DataRow drO = dtOxidation.NewRow();
                drO[0] = "-1";
                drO[1] = "Select an option..";
                dtOxidation.Rows.Add(drO);

                cmbOxidationWeat.DisplayMember = "Description";
                cmbOxidationWeat.ValueMember = "Code";
                cmbOxidationWeat.DataSource = dtOxidation;
                cmbOxidationWeat.SelectedValue = -1;

                //getRfColour_List cmbColourWeat
                DataTable dtColour = new DataTable();
                dtColour = oRf.getRfColour_List();

                DataRow drCol = dtColour.NewRow();
                drCol[0] = "-1";
                drCol[1] = "Select an option..";
                dtColour.Rows.Add(drCol);

                cmbColourWeat.DisplayMember = "Description";
                cmbColourWeat.ValueMember = "Code";
                cmbColourWeat.DataSource = dtColour;
                cmbColourWeat.SelectedValue = -1;

                //getRfPrefixW_List cmbSufixWeat
                DataTable dtPrefixW = new DataTable();
                dtPrefixW = oRf.getRfPrefixW_List();

                DataRow drPrW = dtPrefixW.NewRow();
                drPrW[0] = "-1";
                drPrW[1] = "Select an option..";
                dtPrefixW.Rows.Add(drPrW);

                cmbSufixWeat.DisplayMember = "Description";
                cmbSufixWeat.ValueMember = "Code";
                cmbSufixWeat.DataSource = dtPrefixW;
                cmbSufixWeat.SelectedValue = -1;

                DataTable dtMineralOxid = new DataTable();
                dtMineralOxid = oRf.getRfMinerMin_ListOxid();
                DataRow drM = dtMineralOxid.NewRow();
                drM[0] = "-1";
                drM[1] = "Select an option..";
                dtMineralOxid.Rows.Add(drM);

                CargarCombosWeath(dtMineralOxid, cmbMin1Oxid);
                CargarCombosWeath(dtMineralOxid, cmbMin2Oxid);
                CargarCombosWeath(dtMineralOxid, cmbMin3Oxid);
                CargarCombosWeath(dtMineralOxid, cmbMin4Oxid);

            }
            catch (Exception ex)
            {
                throw new Exception("Error FillCmbWeath: " + ex.Message);
            }
        }

        private void CargarCombosWeath(DataTable _dt, ComboBox _cbox)
        {
            try
            {
                if (_dt.Rows.Count > 0)
                {
                    _cbox.DataSource = _dt.Copy();
                    _cbox.ValueMember = _dt.Columns[0].ToString();
                    _cbox.DisplayMember = _dt.Columns[1].ToString();
                    _cbox.SelectedValue = "-1";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FilldgWeathering(string _sOpcion)
        {
            try
            {
                DataTable dtWeat = new DataTable();
                oWeat.sOpcion = _sOpcion;
                oWeat.sHoleID = cmbHoleIdWeat.SelectedValue.ToString();
                dtWeat = oWeat.getDH_Weathering();
                dgWeathering.DataSource = dtWeat;

                dgWeathering.Columns["SKDHWeathering"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private void cmbHoleIdWeat_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldgWeathering("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string ControlsValidateWeat()
        {
            try
            {
                string sresp = "";

                oCollars.sHoleID = cmbHoleIdWeat.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToWeat.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'To' greater than Hole Id lenght";
                    return sresp;
                }

                if (cmbHoleIdWeat.SelectedValue.ToString() == "Select an option..")
                {
                    sresp = "Selected an option Hole ID";
                    return sresp;
                }
                if (txtFromWeat.Text == "" || txtToWeat.Text == "")
                {
                    sresp = "Empty From or To";
                    return sresp;
                }
                if (txtFromWeat.Text != "-99")
                {
                    //if (double.Parse(txtFromWeat.Text.ToString()) < 0 || double.Parse(txtToWeat.Text.ToString()) < 0)
                    //{
                    if (double.Parse(txtFromWeat.Text.ToString()) == double.Parse(txtToWeat.Text.ToString()))
                    {
                        sresp = " 'From' equal to 'To'";
                        return sresp;
                    }

                    if (double.Parse(txtFromWeat.Text.ToString()) > double.Parse(txtToWeat.Text.ToString()))
                    {
                        sresp = " 'From' greater than 'To'";
                        return sresp;
                    }

                        //return sresp = "From or To must be greater than zero (0)";
                    //}                
                }
                
                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddWeat_Click(object sender, EventArgs e)
        {
            try
            {

                string sResp = ControlsValidateWeat().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (sEditWeat == "0")
                {
                    //Valida que el rango sea valido para el pozo
                    DataTable dtValidRange = new DataTable();
                    oWeat.dFrom = double.Parse(txtFromWeat.Text.ToString());
                    oWeat.dTo = double.Parse(txtToWeat.Text.ToString());
                    oWeat.sHoleID = cmbHoleIdWeat.SelectedValue.ToString();
                    dtValidRange = oWeat.getDHWeatFromToValid();
                    if (dtValidRange.Rows.Count > 0)
                    {
                        MessageBox.Show("Range 'From To' Overlaps", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }


                if (sEditWeat == "1")
                { oWeat.sOpcion = "2"; }
                else {
                    oWeat.iDHWeatheringID = 0;
                    oWeat.sOpcion = "1"; }

                oWeat.sHoleID = cmbHoleIdWeat.SelectedValue.ToString();

                if (dgWeathering.Rows.Count <= 1)
                {
                    oWeat.dFrom = 0;
                }
                else { oWeat.dFrom = double.Parse(txtFromWeat.Text.ToString()); }


                oWeat.dTo = double.Parse(txtToWeat.Text.ToString());
                oWeat.sWeathering = cmbWeatheringWeat.SelectedValue.ToString();


                if (cmbOxidationWeat.SelectedValue.ToString() == "" || cmbOxidationWeat.SelectedValue.ToString() == "-1")
                    oWeat.dOxidation = null;
                else oWeat.dOxidation = double.Parse(cmbOxidationWeat.SelectedValue.ToString());

                if (cmbColourWeat.SelectedValue.ToString() == "" || cmbColourWeat.SelectedValue.ToString() == "-1")
                    oWeat.sColour1 = null;
                else oWeat.sColour1 = cmbColourWeat.SelectedValue.ToString();


                if (cmbSufixWeat.SelectedValue.ToString() == "" || cmbSufixWeat.SelectedValue.ToString() == "-1")
                    oWeat.sSufix1 = null;
                else oWeat.sSufix1 = cmbSufixWeat.SelectedValue.ToString();
                
                oWeat.sColour2 = null;
                oWeat.sSufix2 = null;

                
                if (txtObservWeat.Text.ToString() == "")
                    oWeat.sObservation = null;
                else oWeat.sObservation = txtObservWeat.Text.ToString();


                if (cmbMin1Oxid.SelectedValue.ToString() == "" || cmbMin1Oxid.SelectedValue.ToString() == "-1")
                    oWeat.sMineral1 = null;
                else oWeat.sMineral1 = cmbMin1Oxid.SelectedValue.ToString();

                if (cmbMin2Oxid.SelectedValue.ToString() == "" || cmbMin2Oxid.SelectedValue.ToString() == "-1")
                    oWeat.sMineral2 = null;
                else oWeat.sMineral2 = cmbMin2Oxid.SelectedValue.ToString();

                if (cmbMin3Oxid.SelectedValue.ToString() == "" || cmbMin3Oxid.SelectedValue.ToString() == "-1")
                    oWeat.sMineral3 = null;
                else oWeat.sMineral3 = cmbMin3Oxid.SelectedValue.ToString();

                if (cmbMin4Oxid.SelectedValue.ToString() == "" || cmbMin4Oxid.SelectedValue.ToString() == "-1")
                    oWeat.sMineral4 = null;
                else oWeat.sMineral4 = cmbMin4Oxid.SelectedValue.ToString();

                clsDH_Weathering.sStaticFrom = txtToWeat.Text.ToString();

                string sWeatAdd = oWeat.DH_Weathering_Add();
                if (sWeatAdd == "OK")
                {
                    //MessageBox.Show("Saved");

                    FilldgWeathering("2");

                    if (sEditWeat == "1")
                    {
                        if (dgWeathering.Rows.Count > 1)
                        {
                            DataTable dt = (DataTable)dgWeathering.DataSource;
                            DataRow[] myRow = dt.Select(@"SKDHWeathering = '" + oWeat.iDHWeatheringID + "'");
                            int rowindex = dt.Rows.IndexOf(myRow[0]);
                            dgWeathering.Rows[rowindex].Selected = true;
                            dgWeathering.CurrentCell = dgWeathering.Rows[rowindex].Cells[1];
                        }
                    }

                    cleanControlsWeat();


                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Weathering", sEditWeat == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIdWeat.SelectedValue.ToString() + "." +
                        " From: " + txtFromWeat.Text.ToString() + "." +
                        " To: " + txtToWeat.Text.ToString() + "." +
                        " Oxidation: " + cmbOxidationWeat.SelectedValue.ToString() + "." +
                        " Weathering: " + cmbWeatheringWeat.SelectedValue.ToString() + "." +
                        " Colour: " + cmbColourWeat.SelectedValue.ToString() + "." +
                        " Sufix: " + cmbSufixWeat.SelectedValue.ToString());
                    
                    
                    

                    txtFromWeat.Text = clsDH_Weathering.sStaticFrom.ToString();
                    txtToWeat.Focus();

                }
                else
                {
                    MessageBox.Show("Error Insert: " + sWeatAdd.ToString(), "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                sEditWeat = "0";
            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }
        }

        private void dgWeathering_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditWeat = "1";
                oWeat.iDHWeatheringID = Int64.Parse(dgWeathering.Rows[e.RowIndex].Cells["SKDHWeathering"].Value.ToString());

                cmbHoleIdWeat.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtFromWeat.Text = dgWeathering.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToWeat.Text = dgWeathering.Rows[e.RowIndex].Cells["To"].Value.ToString();
                cmbWeatheringWeat.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Weathering"].Value.ToString();

                cmbOxidationWeat.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Oxidation"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Oxidation"].Value.ToString();

                cmbColourWeat.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Colour1"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Colour1"].Value.ToString();

                cmbSufixWeat.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Sufix1"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Sufix1"].Value.ToString();

                txtObservWeat.Text = dgWeathering.Rows[e.RowIndex].Cells["Observation"].Value.ToString();

                cmbMin1Oxid.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Mineral1"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Mineral1"].Value.ToString();

                cmbMin2Oxid.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Mineral2"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Mineral2"].Value.ToString();

                cmbMin3Oxid.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Mineral3"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Mineral3"].Value.ToString();

                cmbMin4Oxid.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["Mineral4"].Value.ToString() == ""
                    ? "-1" : dgWeathering.Rows[e.RowIndex].Cells["Mineral4"].Value.ToString();

            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void cleanControlsWeat()
        {
            try
            {
                //cmbHoleIdWeat.SelectedValue = dgWeathering.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                //txtFromWeat.Text = dgWeathering.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToWeat.Text = "";
                cmbWeatheringWeat.SelectedValue = "-1";
                cmbOxidationWeat.SelectedValue = "-1";
                cmbColourWeat.SelectedValue = "-1";
                cmbSufixWeat.SelectedValue = "-1";
                txtObservWeat.Text = "";
                cmbMin1Oxid.SelectedValue = "-1";
                cmbMin2Oxid.SelectedValue = "-1";
                cmbMin3Oxid.SelectedValue = "-1";
                cmbMin4Oxid.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnCancelWeat_Click(object sender, EventArgs e)
        {
            try
            {
                sEditWeat = "0";
                cleanControlsWeat();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgWeathering_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Hole Id" + dgWeathering.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                    + " From " + dgWeathering.Rows[e.RowIndex].Cells["From"].Value.ToString()
                    + " To " + dgWeathering.Rows[e.RowIndex].Cells["To"].Value.ToString()
                    , "Weathering", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {

                    oWeat.sHoleID = dgWeathering.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                    oWeat.dFrom = double.Parse(dgWeathering.Rows[e.RowIndex].Cells["From"].Value.ToString());

                    string sWeatDel = oWeat.DH_Weathering_Delete();

                    if (sWeatDel == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FilldgWeathering("2");
                    }
                    sEditWeat = "0";

                }

            }
            catch (Exception)
            {

                throw;
            }
        }

        #endregion

        #region Structure

        private string ControlsValidateStr()
        {
            try
            {
                string sresp = "";

                oCollars.sHoleID = cmbHoleIDSt.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToSt.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'Depth' greater than Hole Id lenght";
                    return sresp;
                }

                if (double.Parse(txtFromSt.Text.ToString()) == double.Parse(txtToSt.Text.ToString()))
                {
                    sresp = " 'From' equal to 'To'";
                    return sresp;
                }

                if (double.Parse(txtFromSt.Text.ToString()) > double.Parse(txtToSt.Text.ToString()))
                {
                    sresp = " 'From' greater than 'To'";
                    return sresp;
                }

                //if (txtAngleToCorest.Text == "")
                //{
                //    sresp = "Angle To Axis must greater than zero (0)";
                //    return sresp;
                //}

                if (txtAngleToCorest.Text != "")
                {
                    if (double.Parse(txtAngleToCorest.Text.ToString()) < 0
                    || double.Parse(txtAngleToCorest.Text.ToString()) > 90)
                    {
                        sresp = "Angle To Axis less than 0 or greater than 90";
                        return sresp;
                    }  
                }

                if (txtUpAngleSt.Text != "")
                {
                    if (double.Parse(txtUpAngleSt.Text.ToString()) < 0)
                    {
                        sresp = "Up Angle less than 0";
                        return sresp;
                    }
                }

                if (txtBtnAngleSt.Text != "")
                {
                    if (double.Parse(txtBtnAngleSt.Text.ToString()) < 0)
                    {
                        sresp = "Btn Angle less than 0";
                        return sresp;
                    }
                }

                if (txtAppThickSt.Text != "")
                {
                    if (double.Parse(txtAppThickSt.Text.ToString()) < 0)
                    {
                        sresp = "App Thick less than 0";
                        return sresp;
                    }
                }

                if (txtNumberSt.Text != "")
                {
                    if (double.Parse(txtNumberSt.Text.ToString()) < 0)
                    {
                        sresp = "Number less than 0";
                        return sresp;
                    }
                }
                

                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FillCmbStruct()
        {
            try
            {
                //cmbHoleIdGeo
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCGeo = dtCollars.NewRow();
                drCGeo[0] = "Select an option..";
                dtCollars.Rows.Add(drCGeo);
                cmbHoleIDSt.DisplayMember = "HoleID";
                cmbHoleIDSt.ValueMember = "HoleID";
                cmbHoleIDSt.DataSource = dtCollars;
                cmbHoleIDSt.SelectedValue = "Select an option..";

                DataTable dtStructType = new DataTable();
                dtStructType = oRf.getRfTypeStructure_List();

                DataRow drS = dtStructType.NewRow();
                drS[0] = "-1";
                drS[1] = "Select an option..";
                dtStructType.Rows.Add(drS);

                cmbStructureTypeSt.DisplayMember = "Comb";
                cmbStructureTypeSt.ValueMember = "Code";
                cmbStructureTypeSt.DataSource = dtStructType;
                cmbStructureTypeSt.SelectedValue = "-1";

                cmbStructDens.DisplayMember = "Comb";
                cmbStructDens.ValueMember = "Code";
                cmbStructDens.DataSource = dtStructType.Copy();
                cmbStructDens.SelectedValue = "-1";

                //getRfFillStructure_List
                DataTable dtFillStr = new DataTable();
                dtFillStr = oRf.getRfFillStructure_List();
                DataRow drFill = dtFillStr.NewRow();
                drFill[0] = "-1";
                drFill[1] = "Select an option..";
                dtFillStr.Rows.Add(drFill);
                cmbFillSt.DisplayMember = "Comb";
                cmbFillSt.ValueMember = "Code";
                cmbFillSt.DataSource = dtFillStr;
                cmbFillSt.SelectedValue = "-1";

                cmbFillSt2.DisplayMember = "Comb";
                cmbFillSt2.ValueMember = "Code";
                cmbFillSt2.DataSource = dtFillStr.Copy();
                cmbFillSt2.SelectedValue = "-1";

                cmbFillSt3.DisplayMember = "Comb";
                cmbFillSt3.ValueMember = "Code";
                cmbFillSt3.DataSource = dtFillStr.Copy();
                cmbFillSt3.SelectedValue = "-1";

                cmbFillSt4.DisplayMember = "Comb";
                cmbFillSt4.ValueMember = "Code";
                cmbFillSt4.DataSource = dtFillStr.Copy();
                cmbFillSt4.SelectedValue = "-1";

            }
            catch (Exception ex)
            {
                throw new Exception("Error FillCmbStr: " + ex.Message);
            }
        }

        private void FilldtStruct(string _sOpcion)
        {
            try
            {
                DataTable dtStruct = new DataTable();
                oStr.sOpcion = _sOpcion;
                oStr.sHoleID = cmbHoleIDSt.SelectedValue.ToString();
                dtStruct = oStr.getDH_Structures();
                dgStructure.DataSource = dtStruct;

                dgStructure.Columns["SKDHStructrue"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void cmbHoleIDSt_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldtStruct("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtDepthSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //if (Char.IsNumber(e.KeyChar))
            //{
            //    e.Handled = false;
            //}
            //if (Char.IsLetter(e.KeyChar))
            //{
            //    e.Handled = true;
            //}

            ////TabEnter(e);
        }

        private void txtAngleToCorest_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //if (Char.IsNumber(e.KeyChar))
            //{
            //    e.Handled = false;
            //}
            //if (Char.IsLetter(e.KeyChar))
            //{
            //    e.Handled = true;
            //}
            ////TabEnter(e);
        }

        private void btnAddSt_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateStr().ToString();
               
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataTable dtValidRange = new DataTable();
                oStr.iFrom = double.Parse(txtFromSt.Text.ToString());
                oStr.iTo = double.Parse(txtToSt.Text.ToString());
                oStr.sHoleID = cmbHoleIDSt.SelectedValue.ToString();
                dtValidRange = oStr.getDH_StructuresValid();
                if (dtValidRange.Rows.Count > 0)
                {
                    MessageBox.Show("Range 'From To' Overlaps", "Structures", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                if (sEditStruct == "1")
                { oStr.sOpcion = "2"; }
                else { 
                        oStr.sOpcion = "1";
                        oStr.iDHStructrueID = 0;
                    }
                
                oStr.sType = cmbStructureTypeSt.SelectedValue.ToString();


                if (txtAngleToCorest.Text.ToString() == "")
                    oStr.dAngleToCore = null;
                else oStr.dAngleToCore = double.Parse(txtAngleToCorest.Text.ToString());

                if (txtCommentsSt.Text.ToString() == "")
                    oStr.sComments = null;
                else oStr.sComments = txtCommentsSt.Text.ToString();
                
                oStr.dLenght = 0;

                if (txtUpAngleSt.Text.ToString() == "")
                    oStr.dUpAngle = null;
                else oStr.dUpAngle = double.Parse(txtUpAngleSt.Text.ToString());

                if (txtBtnAngleSt.Text.ToString() == "")
                    oStr.dBtonAngle = null;
                else oStr.dBtonAngle = double.Parse(txtBtnAngleSt.Text.ToString());

                if (txtAppThickSt.Text.ToString() == "")
                    oStr.dAppThick = null;
                else oStr.dAppThick = double.Parse(txtAppThickSt.Text.ToString());

                if (cmbFillSt.SelectedValue.ToString() ==  "-1" || cmbFillSt.SelectedValue.ToString() == "")
                    oStr.sFill = null;
                else oStr.sFill = cmbFillSt.SelectedValue.ToString();

                if (txtNumberSt.Text.ToString() == "")
                    oStr.dNumber = null;
                else oStr.dNumber = double.Parse(txtNumberSt.Text.ToString());

                if (cmbFillSt2.SelectedValue.ToString() == "-1" || cmbFillSt2.SelectedValue.ToString() == "")
                    oStr.sFill2 = null;
                else oStr.sFill2 = cmbFillSt2.SelectedValue.ToString();

                if (cmbFillSt3.SelectedValue.ToString() == "-1" || cmbFillSt3.SelectedValue.ToString() == "")
                    oStr.sFill3 = null;
                else oStr.sFill3 = cmbFillSt3.SelectedValue.ToString();

                if (cmbFillSt4.SelectedValue.ToString() == "-1" || cmbFillSt4.SelectedValue.ToString() == "")
                    oStr.sFill4 = null;
                else oStr.sFill4 = cmbFillSt4.SelectedValue.ToString();

                clsDH_Structures.sStaticFrom = txtToSt.Text.ToString();

                string sRespStr = oStr.DH_Structures_Add();
                if (sRespStr == "OK")
                {
                    //MessageBox.Show("Saved");
                    FilldtStruct("2");

                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Structures", sEditStruct == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIDSt.SelectedValue.ToString() + "." +
                        " From: " + txtFromSt.Text.ToString() + "." +
                        " To: " + txtToSt.Text.ToString() + "." +
                        " Type St: " + cmbStructureTypeSt.SelectedValue.ToString() + "." +
                        " Angle To Axis: " + txtAngleToCorest.Text.ToString() + "." +
                        " Up AnglD: " + txtUpAngleSt.Text.ToString() + "." +
                        " Btn AnglD: " + txtBtnAngleSt.Text.ToString() + "." +
                        " App Thick: " + txtAppThickSt.Text.ToString() + "." +
                        " Number: " + txtNumberSt.Text.ToString());


                    if (sEditStruct == "1")
                    {
                        if (dgStructure.Rows.Count > 1)
                        {
                            DataTable dt = (DataTable)dgStructure.DataSource;
                            DataRow[] myRow = dt.Select(@"SKDHStructrue = '" + oStr.iDHStructrueID + "'");
                            int rowindex = dt.Rows.IndexOf(myRow[0]);
                            dgStructure.Rows[rowindex].Selected = true;
                            dgStructure.CurrentCell = dgStructure.Rows[rowindex].Cells[1];
                        }
                    }


                    CleanControlsSt();

                    txtFromSt.Text = clsDH_Structures.sStaticFrom.ToString();
                    txtToSt.Focus();
                }
                else
                {
                    MessageBox.Show("Error Insert: " + sRespStr.ToString(), "Structures", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CleanControlsSt();
                }

                sEditStruct = "0";

            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }
        }

        private void CleanControlsSt()
        {
            try
            {
                sEditStruct = "0";
                txtAngleToCorest.Text = "";
                txtBtnAngleSt.Text = "";
                txtUpAngleSt.Text = "";
                txtAppThickSt.Text = "";
                txtNumberSt.Text = "";
                txtCommentsSt.Text = "";
                txtToSt.Text = "";
                cmbStructureTypeSt.SelectedValue = "-1";
                cmbFillSt.SelectedValue = "-1";
                cmbFillSt2.SelectedValue = "-1";
                cmbFillSt3.SelectedValue = "-1";
                cmbFillSt4.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelSt_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControlsSt();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgStructure_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                sEditStruct = "1";
                oStr.iDHStructrueID = Int64.Parse(dgStructure.Rows[e.RowIndex].Cells["SKDHStructrue"].Value.ToString());

                cmbHoleIDSt.SelectedValue = dgStructure.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtFromSt.Text = dgStructure.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToSt.Text = dgStructure.Rows[e.RowIndex].Cells["To"].Value.ToString();


                cmbStructureTypeSt.SelectedValue = dgStructure.Rows[e.RowIndex].Cells["Type"].Value.ToString() == "" ?
                    "-1" : dgStructure.Rows[e.RowIndex].Cells["Type"].Value.ToString();

                txtAngleToCorest.Text = dgStructure.Rows[e.RowIndex].Cells["AngleToAxis"].Value.ToString() == "" ?
                    "" : dgStructure.Rows[e.RowIndex].Cells["AngleToAxis"].Value.ToString();

                txtCommentsSt.Text = dgStructure.Rows[e.RowIndex].Cells["Comments"].Value.ToString() == "" ?
                    "" : dgStructure.Rows[e.RowIndex].Cells["Comments"].Value.ToString();

                txtUpAngleSt.Text = dgStructure.Rows[e.RowIndex].Cells["UpAngle"].Value.ToString() == "" ?
                    "" : dgStructure.Rows[e.RowIndex].Cells["UpAngle"].Value.ToString();

                txtBtnAngleSt.Text = dgStructure.Rows[e.RowIndex].Cells["BtonAngle"].Value.ToString() == "" ?
                    "" : dgStructure.Rows[e.RowIndex].Cells["BtonAngle"].Value.ToString();

                txtAppThickSt.Text = dgStructure.Rows[e.RowIndex].Cells["AppThick"].Value.ToString() == "" ?
                    "" : dgStructure.Rows[e.RowIndex].Cells["AppThick"].Value.ToString();

                cmbFillSt.SelectedValue = dgStructure.Rows[e.RowIndex].Cells["Fill"].Value.ToString() == "" ?
                    "-1" : dgStructure.Rows[e.RowIndex].Cells["Fill"].Value.ToString();

                txtNumberSt.Text = dgStructure.Rows[e.RowIndex].Cells["Number"].Value.ToString() == "" ?
                    "" : dgStructure.Rows[e.RowIndex].Cells["Number"].Value.ToString();

                cmbFillSt2.SelectedValue = dgStructure.Rows[e.RowIndex].Cells["Fill2"].Value.ToString() == "" ?
                    "-1" : dgStructure.Rows[e.RowIndex].Cells["Fill2"].Value.ToString();

                cmbFillSt3.SelectedValue = dgStructure.Rows[e.RowIndex].Cells["Fill3"].Value.ToString() == "" ?
                    "-1" : dgStructure.Rows[e.RowIndex].Cells["Fill3"].Value.ToString();

                cmbFillSt4.SelectedValue = dgStructure.Rows[e.RowIndex].Cells["Fill4"].Value.ToString() == "" ?
                    "-1" : dgStructure.Rows[e.RowIndex].Cells["Fill4"].Value.ToString();

            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Structures", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void dgStructure_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Hole Id" + dgStructure.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                   + " From " + dgStructure.Rows[e.RowIndex].Cells["From"].Value.ToString()
                   + " Type " + dgStructure.Rows[e.RowIndex].Cells["Type"].Value.ToString()
                   , "Structure", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oStr.iDHStructrueID = Int64.Parse(dgStructure.Rows[e.RowIndex].Cells["SKDHStructrue"].Value.ToString());

                    string sRespDel = oStr.DH_Structures_Delete();
                    if (sRespDel == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FilldtStruct("2");
                    }
                    sEditStruct = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        #endregion
        

        #region Mineralization

        private void CargarCombosMin(DataTable _dt, ComboBox _cbox)
        {
            try
            {
                if (_dt.Rows.Count > 0)
                {
                    _cbox.DataSource = _dt.Copy();
                    _cbox.ValueMember = _dt.Columns[0].ToString();
                    _cbox.DisplayMember = _dt.Columns[1].ToString();
                    _cbox.SelectedValue = "-1";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FillCmbMiner()
        {
            try
            {
                //getRfMinerMin_List
                oCollars.sHoleID = "";
                oCollars.sLogged = clsRf.sUser;
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCGeo = dtCollars.NewRow();
                drCGeo[0] = "Select an option..";
                dtCollars.Rows.Add(drCGeo);
                cmbHoleIdMin.DisplayMember = "HoleID";
                cmbHoleIdMin.ValueMember = "HoleID";
                cmbHoleIdMin.DataSource = dtCollars;
                cmbHoleIdMin.SelectedValue = "Select an option..";


                DataTable dtMineral = new DataTable();
                dtMineral = oRf.getRfMinerMin_List();
                DataRow drM = dtMineral.NewRow();
                drM[0] = "-1";
                drM[1] = "Select an option..";
                dtMineral.Rows.Add(drM);

                CargarCombosMin(dtMineral, cmbM1Z1);
                CargarCombosMin(dtMineral, cmbM1Z2);
                CargarCombosMin(dtMineral, cmbM1Z3);

                CargarCombosMin(dtMineral,cmbMineral1Dens);
                CargarCombosMin(dtMineral, cmbMineral2Dens);

                CargarCombosMin(dtMineral, cmbM2Z1);
                CargarCombosMin(dtMineral, cmbM2Z2);
                CargarCombosMin(dtMineral, cmbM2Z3);

                CargarCombosMin(dtMineral, cmbM3Z1);
                CargarCombosMin(dtMineral, cmbM3Z2);
                CargarCombosMin(dtMineral, cmbM3Z3);

                DataTable dtMinStyle = new DataTable();
                dtMinStyle = oRf.getRfMinerMinSt_List();
                DataRow drMin = dtMinStyle.NewRow();
                drMin[0] = "-1";
                drMin[1] = "Select an option..";
                dtMinStyle.Rows.Add(drMin);

                CargarCombosMin(dtMinStyle, cmbStyleM1);
                CargarCombosMin(dtMinStyle, cmbStyleM2);
                CargarCombosMin(dtMinStyle, cmbStyleM3);

                DataTable dtMinPerc = new DataTable();
                dtMinPerc = oRf.getRfMinerPercent_List(ConfigurationSettings.AppSettings["IDProjectGC"].ToString()); //Id Proyecto Gran Colombia. Ej GSG, GZG ...
                DataRow drMinPerc = dtMinPerc.NewRow();
                drMinPerc[0] = "-1";
                drMinPerc[1] = "Select an option..";
                dtMinPerc.Rows.Add(drMinPerc);

                //CargarCombosMin(dtMinPerc, cmbPorcM1);
                //CargarCombosMin(dtMinPerc, cmbPorcM2);
                //CargarCombosMin(dtMinPerc, cmbPorcM3);


                DataTable dtGSizeMin = new DataTable();
                dtGSizeMin = oRf.getRfGSize_ListMin("1");
                DataRow drG = dtGSizeMin.NewRow();
                drG[0] = "-1";
                drG[1] = "Select an option..";
                dtGSizeMin.Rows.Add(drG);

                CargarCombosMin(dtGSizeMin, cmbGSizeMin1);
                CargarCombosMin(dtGSizeMin, cmbGSizeMin2);
                CargarCombosMin(dtGSizeMin, cmbGSizeMin3);

            }
            catch (Exception ex)
            {
                throw new Exception("Error FillCmbWeath: " + ex.Message);
            }
        }

        private string ControlsValidateMin()
        {
            try
            {
                string sresp = "";

                oCollars.sHoleID = cmbHoleIdMin.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToMin.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'To' greater than Hole Id lenght";
                    return sresp;
                }

                if (cmbHoleIdMin.SelectedValue.ToString() == "Select an option..")
                {
                    sresp = "Selected an option Hole ID";
                    return sresp;
                }
                if (txtFromMin.Text == "" || txtToMin.Text == "")
                {
                    sresp = "Empty From or To";
                    return sresp;
                }
                if (txtFromMin.Text != "-99")
                {
                    //if (double.Parse(txtFromMin.Text.ToString()) < 0 || double.Parse(txtToMin.Text.ToString()) < 0)
                    //{
                    if (double.Parse(txtFromMin.Text.ToString()) == double.Parse(txtToMin.Text.ToString()))
                    {
                        sresp = " 'From' equal to 'To'";
                        return sresp;
                    }

                    if (double.Parse(txtFromMin.Text.ToString()) > double.Parse(txtToMin.Text.ToString()))
                    {
                        sresp = " 'From' greater than 'To'";
                        return sresp;
                    }
                    //}
                    //return sresp = "From or To must be greater than zero (0)";
                }

               
                if (txtMinPerc1.Text != "")
                {
                    if (double.Parse(txtMinPerc1.Text) > 100)
                    {
                        sresp ="Percentage 1 isn´t more than 100";
                        return sresp;
                    }
                }

                if (txtMinPerc2.Text != "")
                {
                    if (double.Parse(txtMinPerc2.Text) > 100)
                    {
                        sresp = "Percentage 2 isn´t more than 100";
                        return sresp;
                    }
                }

                if (txtMinPerc3.Text != "")
                {
                    if (double.Parse(txtMinPerc3.Text) > 100)
                    {
                        sresp = "Percentage 3 isn´t more than 100";
                        return sresp;
                    }
                }
                
               
                if (cmbM1Z1.SelectedValue.ToString() == "-1")
                {
                    sresp = "Selected an option Mineral 1";
                    return sresp;
                }

                //if (cmbPorcM1.SelectedValue.ToString() == "-1")
                //{
                //    sresp = "Empty Percent Mineralization 1 (#m%)";
                //    return sresp;
                //}

                //if (cmbStyleM1.SelectedValue.ToString() == "-1")
                //{
                //    sresp = "Selected an option Style 1";
                //    return sresp;
                //}
                



                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private DataTable dtCryst_QX()
        {
            DataTable dtCryst_QX = new DataTable();
            dtCryst_QX.Columns.Add("Key", typeof(String));
            dtCryst_QX.Columns.Add("Value", typeof(String));


            for (int i = 0; i < conf.AppSettings.Settings.Count; i++)
            {
                if (conf.AppSettings.Settings.AllKeys[i].ToString().Contains("Cryst_QX"))
                {

                    DataRow drDup = dtCryst_QX.NewRow();
                    //drConect["Con"] = ;
                    drDup["Key"] = conf.AppSettings.Settings.AllKeys[i].ToString();
                    drDup["Value"] =
                        conf.AppSettings.Settings[conf.AppSettings.Settings.AllKeys[i].ToString()].Value.ToString();
                    dtCryst_QX.Rows.Add(drDup);

                }

            }

            return dtCryst_QX;
        }

        private void txtFromMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void txtToMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void btnAddMin_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateMin().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (sEditMiner == "1"){
                    oMiner.sOpcion = "2"; }
                else {

                    DataTable dtValidRange = new DataTable();
                    oMiner.dFrom = double.Parse(txtFromMin.Text.ToString());
                    oMiner.dTo = double.Parse(txtToMin.Text.ToString());
                    oMiner.sHoleID = cmbHoleIdMin.SelectedValue.ToString();
                    dtValidRange = oMiner.getDHMinFromToValid();
                    if (dtValidRange.Rows.Count > 0)
                    {
                        MessageBox.Show("Range 'From To' Overlaps", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    oMiner.iDHMinID = 0;
                    oMiner.sOpcion = "1"; 
                }
                
                if (dgMineraliz.Rows.Count <= 1)
                { oMiner.dFrom = 0;  }
                else { oMiner.dFrom = double.Parse(txtFromMin.Text.ToString()); }

                //Se valida si el mineral elegido es Crystalline quartz
                DataTable dtCryst = dtCryst_QX();
                DataRow[] dato = dtCryst.Select("Value = '" + cmbM1Z1.SelectedValue.ToString() + "'");
                if (dato.Length > 0)
                {
                    MessageBox.Show(cmbM1Z1.SelectedValue.ToString());
                    return;
                }

                oMiner.dTo = double.Parse(txtToMin.Text.ToString());
                oMiner.sHoleID = cmbHoleIdMin.SelectedValue.ToString();
                oMiner.sMZ1Mineral = cmbM1Z1.SelectedValue.ToString();

                /*if (txtUpAngleSt.Text.ToString() == "")
                    oStr.dUpAngle = null;
                else oStr.dUpAngle = double.Parse(txtUpAngleSt.Text.ToString());*/

                if (cmbM1Z2.SelectedValue.ToString() == "-1" || cmbM1Z2.SelectedValue.ToString() == "")
                    oMiner.sMZ1Mineral2 = null;
                else oMiner.sMZ1Mineral2 = cmbM1Z2.SelectedValue.ToString();
                
                if (cmbM1Z3.SelectedValue.ToString() == "-1" || cmbM1Z3.SelectedValue.ToString() == "")
                    oMiner.sMZ1Mineral3 = null;
                else oMiner.sMZ1Mineral3 = cmbM1Z3.SelectedValue.ToString();

                if (txtMinPerc1.Text.ToString() == "")
                    oMiner.dMZ1Perc = null;
                else oMiner.dMZ1Perc = double.Parse(txtMinPerc1.Text.ToString());

                if (cmbStyleM1.SelectedValue.ToString() == "-1" || cmbStyleM1.SelectedValue.ToString() == "")
                    oMiner.sMZ1Style = null;
                else oMiner.sMZ1Style = cmbStyleM1.SelectedValue.ToString();

                if (cmbM2Z1.SelectedValue.ToString() == "-1" || cmbM2Z1.SelectedValue.ToString() == "")
                    oMiner.sMZ2Mineral = null;
                else oMiner.sMZ2Mineral = cmbM2Z1.SelectedValue.ToString();

                if (cmbM2Z2.SelectedValue.ToString() == "-1" || cmbM2Z2.SelectedValue.ToString() == "")
                    oMiner.sMZ2Mineral2 = null;
                else oMiner.sMZ2Mineral2 = cmbM2Z2.SelectedValue.ToString();

                if (cmbM2Z3.SelectedValue.ToString() == "-1" || cmbM2Z3.SelectedValue.ToString() == "")
                    oMiner.sMZ2Mineral3 = null;
                else oMiner.sMZ2Mineral3 = cmbM2Z3.SelectedValue.ToString();

                if (txtMinPerc2.Text.ToString() == "")
                    oMiner.dMZ2Perc = null;
                else oMiner.dMZ2Perc = double.Parse(txtMinPerc2.Text.ToString());

                if (cmbStyleM2.SelectedValue.ToString() == "-1" || cmbStyleM2.SelectedValue.ToString() == "")
                    oMiner.sMZ2Style = null;
                else oMiner.sMZ2Style = cmbStyleM2.SelectedValue.ToString();

                if (cmbM3Z1.SelectedValue.ToString() == "-1" || cmbM3Z1.SelectedValue.ToString() == "")
                    oMiner.sMZ3Mineral = null;
                else oMiner.sMZ3Mineral = cmbM3Z1.SelectedValue.ToString();

                if (cmbM3Z2.SelectedValue.ToString() == "-1" || cmbM3Z2.SelectedValue.ToString() == "")
                    oMiner.sMZ3Mineral2 = null;
                else oMiner.sMZ3Mineral2 = cmbM3Z2.SelectedValue.ToString();

                if (cmbM3Z3.SelectedValue.ToString() == "-1" || cmbM3Z3.SelectedValue.ToString() == "")
                    oMiner.sMZ3Mineral3 = null;
                else oMiner.sMZ3Mineral3 = cmbM3Z3.SelectedValue.ToString();

                if (txtMinPerc3.Text.ToString() == "")
                    oMiner.dMZ3Perc = null;
                else oMiner.dMZ3Perc = double.Parse(txtMinPerc3.Text.ToString());

                if (cmbStyleM3.SelectedValue.ToString() == "-1" || cmbStyleM3.SelectedValue.ToString() == "")
                    oMiner.sMZ3Style = null;
                else oMiner.sMZ3Style = cmbStyleM3.SelectedValue.ToString();

                if (txtCommentsMin.Text.ToString() == "")
                    oMiner.sComments = null;
                else oMiner.sComments = txtCommentsMin.Text.ToString();



                if (cmbGSizeMin1.SelectedValue.ToString() == "-1" || cmbGSizeMin1.SelectedValue.ToString() == "")
                    oMiner.sGSize1 = null;
                else oMiner.sGSize1 = cmbGSizeMin1.SelectedValue.ToString();

                if (cmbGSizeMin2.SelectedValue.ToString() == "-1" || cmbGSizeMin2.SelectedValue.ToString() == "")
                    oMiner.sGSize2 = null;
                else oMiner.sGSize2 = cmbGSizeMin2.SelectedValue.ToString();

                if (cmbGSizeMin3.SelectedValue.ToString() == "-1" || cmbGSizeMin3.SelectedValue.ToString() == "")
                    oMiner.sGSize3 = null;
                else oMiner.sGSize3 = cmbGSizeMin3.SelectedValue.ToString();



                clsDHMineraliz.sStaticFrom = txtToMin.Text.ToString();

                string sRespMin = oMiner.DH_Mineraliz_Add();
                if (sRespMin == "OK")
                {
                    //MessageBox.Show("Saved");
                    FilldgMineraliz("2");


                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Mineralizations", sEditMiner == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIdMin.SelectedValue.ToString() + "." +
                        " From: " + txtFromMin.Text.ToString() + "." +
                        " To: " + txtToMin.Text.ToString() + "." +
                        " Mineral 1: " + cmbM1Z1.SelectedValue.ToString() + "." +
                        " Mineral 2: " + cmbM1Z2.SelectedValue.ToString() + "." +
                        " Mineral 3: " + cmbM1Z3.SelectedValue.ToString() + "." +
                        " Style Min: " + cmbStyleM1.SelectedValue.ToString() + "." +
                        " Porcentaje /m: " + txtMinPerc1.Text.ToString());


                    if (sEditMiner == "1")
                    {
                        if (dgAlterations.Rows.Count > 1)
                        {
                            DataTable dt = (DataTable)dgMineraliz.DataSource;
                            DataRow[] myRow = dt.Select(@"SKDHMin = '" + oMiner.iDHMinID + "'");
                            int rowindex = dt.Rows.IndexOf(myRow[0]);
                            dgMineraliz.Rows[rowindex].Selected = true;
                            dgMineraliz.CurrentCell = dgMineraliz.Rows[rowindex].Cells[1];
                        }
                    }


                    CleanControlsMin();

                    txtFromMin.Text = clsDHMineraliz.sStaticFrom.ToString();
                    txtToMin.Focus();

                }
                else
                {
                    MessageBox.Show("Error Insert: " + sRespMin.ToString(), "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                
                sEditMiner = "0";
            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }
        }

        private void cmbHoleIdMin_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldgMineraliz("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FilldgMineraliz(string _sOpcion)
        {
            try
            {
                DataTable dtMiner = new DataTable();
                oMiner.sOpcion = _sOpcion;
                oMiner.sHoleID = cmbHoleIdMin.SelectedValue.ToString();
                dtMiner = oMiner.getDHMineraliz();
                dgMineraliz.DataSource = dtMiner;

                dgMineraliz.Columns["SKDHMin"].Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgMineraliz_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditMiner = "1";
                oMiner.iDHMinID = Int64.Parse(dgMineraliz.Rows[e.RowIndex].Cells["SKDHMin"].Value.ToString());
                cmbHoleIdMin.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtFromMin.Text = dgMineraliz.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToMin.Text = dgMineraliz.Rows[e.RowIndex].Cells["To"].Value.ToString();

                cmbM1Z1.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ1Mineral"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ1Mineral"].Value.ToString();

                cmbM1Z2.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ1Mineral2"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ1Mineral2"].Value.ToString();

                cmbM1Z3.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ1Mineral3"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ1Mineral3"].Value.ToString();

                cmbM2Z1.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ2Mineral"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ2Mineral"].Value.ToString();

                cmbM2Z2.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ2Mineral2"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ2Mineral2"].Value.ToString();

                cmbM2Z3.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ2Mineral3"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ2Mineral3"].Value.ToString();

                cmbM3Z1.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ3Mineral"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ3Mineral"].Value.ToString();

                cmbM3Z2.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ3Mineral2"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ3Mineral2"].Value.ToString();

                cmbM3Z3.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ3Mineral3"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ3Mineral3"].Value.ToString();

                cmbStyleM1.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ1Style"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ1Style"].Value.ToString();

                cmbStyleM2.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ2Style"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ2Style"].Value.ToString();

                cmbStyleM3.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["MZ3Style"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["MZ3Style"].Value.ToString();

               txtMinPerc1.Text = dgMineraliz.Rows[e.RowIndex].Cells["MZ1Perc"].Value.ToString() == "" ?
                    "" : dgMineraliz.Rows[e.RowIndex].Cells["MZ1Perc"].Value.ToString();

               txtMinPerc2.Text = dgMineraliz.Rows[e.RowIndex].Cells["MZ2Perc"].Value.ToString() == "" ?
                    "" : dgMineraliz.Rows[e.RowIndex].Cells["MZ2Perc"].Value.ToString();

               txtMinPerc3.Text = dgMineraliz.Rows[e.RowIndex].Cells["MZ3Perc"].Value.ToString() == "" ?
                    "": dgMineraliz.Rows[e.RowIndex].Cells["MZ3Perc"].Value.ToString();

                txtCommentsMin.Text = dgMineraliz.Rows[e.RowIndex].Cells["Comments"].Value.ToString() == "" ?
                    "" : dgMineraliz.Rows[e.RowIndex].Cells["Comments"].Value.ToString();



                cmbGSizeMin1.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["Gsize"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["Gsize"].Value.ToString();

                cmbGSizeMin2.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["Gsize2"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["Gsize2"].Value.ToString();

                cmbGSizeMin3.SelectedValue = dgMineraliz.Rows[e.RowIndex].Cells["Gsize3"].Value.ToString() == "" ?
                    "-1" : dgMineraliz.Rows[e.RowIndex].Cells["Gsize3"].Value.ToString();
            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanControlsMin()
        {
            try
            {
                sEditMiner = "0";
                txtToMin.Text = "";
                txtCommentsMin.Text = "";

                txtMinPerc1.Text = "";
                txtMinPerc2.Text = "";
                txtMinPerc3.Text = "";


                cmbM1Z1.SelectedValue = "-1";
                cmbM1Z2.SelectedValue = "-1"; 
                cmbM1Z3.SelectedValue = "-1";
                cmbM2Z1.SelectedValue = "-1"; 
                cmbM2Z2.SelectedValue = "-1";
                cmbM2Z3.SelectedValue = "-1";
                cmbM3Z1.SelectedValue = "-1";
                cmbM3Z2.SelectedValue = "-1";
                cmbM3Z3.SelectedValue = "-1";

                cmbStyleM1.SelectedValue = "-1";
                cmbStyleM2.SelectedValue = "-1";
                cmbStyleM3.SelectedValue = "-1";

                cmbGSizeMin1.SelectedValue = "-1";
                cmbGSizeMin2.SelectedValue = "-1";
                cmbGSizeMin3.SelectedValue = "-1";
                //cmbHoleIdMin.SelectedValue = "Select an option..";
                
                txtToMin.Text = "";
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelMin_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControlsMin();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgMineraliz_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Hole Id" + dgMineraliz.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                    + " From " + dgMineraliz.Rows[e.RowIndex].Cells["From"].Value.ToString()
                    + " To " + dgMineraliz.Rows[e.RowIndex].Cells["To"].Value.ToString()
                    , "Mineralizations", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oMiner.iDHMinID = Int64.Parse(dgMineraliz.Rows[e.RowIndex].Cells["SKDHMin"].Value.ToString());
                    string sRespDel = oMiner.DH_Mineraliz_Delete();
                    if (sRespDel == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FilldgMineraliz("2");
                    }
                    sEditMiner = "0";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        #endregion

        //[BrowsableAttribute(false)]
        //public event EventHandler LostFocus;
        //private void txtFromGeo_LostFocus(object sender, System.EventArgs e)
        //{
        //    try
        //    {
        //        if (txtFromGeo.Text.ToString() == "-99")
        //        {
        //            txtToGeo.Text = "-99";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        private void txtFromGeo_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtFromGeo.Text.ToString() == "-99")
                {
                    txtToGeo.Text = "-99";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GetDifferGeo()
        {
            try
            {
                if (txtFromGeo.Text.ToString() == "-99")
                {
                    txtDifferGeo.Text = "-99";
                }
                else
                {
                    txtDifferGeo.Text = (double.Parse(txtToGeo.Text.ToString()) -
                        double.Parse(txtFromGeo.Text.ToString())).ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void txtToGeo_Leave(object sender, EventArgs e)
        {
            try
            {
                GetDifferGeo();

                //if (txtFromGeo.Text.ToString() == "-99")
                //{
                //    txtDifferGeo.Text = "-99";
                //}
                //else
                //{
                //    txtDifferGeo.Text = (double.Parse(txtToGeo.Text.ToString()) -
                //        double.Parse(txtFromGeo.Text.ToString())).ToString();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string ValidGreaterThanZero(string _sValor)
        {
            try
            {
                string sValid = "";

                if (_sValor.ToString() == "")
                {
                    MessageBox.Show("Value must greater than zero (0)", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return sValid;
                }
                if (double.Parse(_sValor.ToString()) <= 0)
                {
                    MessageBox.Show("Value must greater than zero (0)", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return sValid;
                }
                return sValid;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void txtDepthSt_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtFromSt.Text.ToString() == "")
                {
                    MessageBox.Show("Empty Depth", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtFromSt.Focus();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtRec_mGeo_Leave(object sender, EventArgs e)
            {
            try
            {
                if (txtRec_mGeo.Text.ToString() != "")
                {
                    txtRec_PorcGeo.Text = (double.Parse(txtRec_mGeo.Text.ToString()) / 
                        double.Parse(txtDifferGeo.Text.ToString()) * 100).ToString();
                }
                if (txtRec_mGeo.Text.ToString() == "-99")
                {
                    txtRec_PorcGeo.Text = "-99";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtRQD_cmGeo_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtRQD_cmGeo.Text.ToString() != "")
                {
                    txtRQD_PorcGeo.Text = (double.Parse(txtRQD_cmGeo.Text.ToString()) /
                        double.Parse(txtDifferGeo.Text.ToString())).ToString();
                }
                if (txtRQD_cmGeo.Text.ToString() == "-99")
                {
                    txtRQD_PorcGeo.Text = "-99";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtFromLit_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtFromLit.Text == "-99")
                {
                    txtToLit.Text = "-99";
                    return;
                }


                if (txtFromLit.Text.ToString() == "")
                {
                    MessageBox.Show("From must greater than zero (0)", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtFromLit.Focus();
                    return;
                }
                if (double.Parse(txtFromLit.Text.ToString()) < 0)
                {
                    MessageBox.Show("From must greater than zero (0)", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtFromLit.Focus();
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtToLit_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtToLit.Text == "-99")
                {
                    txtFromLit.Text = "-99";
                }

                //if (txtToLit.Text.ToString() == "")
                //{
                //    MessageBox.Show("To must greater than zero (0)", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtToLit.Focus();
                //    return;
                //}
                //if (double.Parse(txtToLit.Text.ToString()) < 0)
                //{
                //    MessageBox.Show("To must greater than zero (0)", "Lithology", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtToLit.Focus();
                //    return;
                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtFromWeat_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtFromWeat.Text == "-99")
                {
                    txtToWeat.Text = "-99";
                }

                if (txtFromWeat.Text.ToString() == "")
                {
                    MessageBox.Show("From must greater than zero (0)", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtFromWeat.Focus();
                    return;
                }
                if (double.Parse(txtFromWeat.Text.ToString()) < 0)
                {
                    MessageBox.Show("From must greater than zero (0)", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtFromWeat.Focus();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtToWeat_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtToWeat.Text == "-99")
                {
                    txtFromWeat.Text = "-99";
                }

                //if (txtToWeat.Text.ToString() == "")
                //{
                //    MessageBox.Show("To must greater than zero (0)", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtToWeat.Focus();
                //    return;
                //}
                //if (double.Parse(txtToWeat.Text.ToString()) < 0)
                //{
                //    MessageBox.Show("To must greater than zero (0)", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtToWeat.Focus();
                //    return;
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtFromMin_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtFromMin.Text == "-99")
                {
                    txtToMin.Text = "-99";
                }

                //if (txtFromMin.Text.ToString() == "")
                //{
                //    MessageBox.Show("From must greater than zero (0)", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtFromMin.Focus();
                //    return;
                //}

                //if (double.Parse(txtFromMin.Text.ToString()) < 0)
                //{
                //    MessageBox.Show("From must greater than zero (0)", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtFromMin.Focus();
                //    return;
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtToMin_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtToMin.Text == "-99")
                {
                    txtFromMin.Text = "-99";
                }

                //if (txtToMin.Text.ToString() == "")
                //{
                //    MessageBox.Show("To must greater than zero (0)", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtToMin.Focus();
                //    return;
                //}
                //if (double.Parse(txtToMin.Text.ToString()) < 0)
                //{
                //    MessageBox.Show("To must greater than zero (0)", "Mineralizations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    txtToMin.Focus();
                //    return;
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool Keypress(KeyPressEventArgs e)
        {
            
            if (Char.IsNumber(e.KeyChar))
            {
                return false;
            }
            if (Char.IsLetter(e.KeyChar))
            {
                return true;
            }

            return false;
        }

        #region Box

        private void FilldgBox(string _sOpcion)
        {
            try
            {
                DataTable dtBox = new DataTable();
                oBox.sOpcion = _sOpcion;
                oBox.sHoleID = cmbHoleIDBox.SelectedValue.ToString();
                dtBox = oBox.getDH_Box();
                dgBox.DataSource = dtBox;

                dgBox.Columns["SKDHBox"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private void txtFromBox_KeyPress(object sender, KeyPressEventArgs e)
        {

            e.Handled = Keypress(e);
            ////TabEnter(e);

            //if (Char.IsNumber(e.KeyChar))
            //{
            //    e.Handled = false;
            //}
            //if (Char.IsLetter(e.KeyChar))
            //{
            //    e.Handled = true;
            //}
        }

        private void txtToBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            ////TabEnter(e);
        }

        private void txtNoBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            ////TabEnter(e);
        }

        private void txtStand_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            ////TabEnter(e);
        }

        private string ControlsValidateBox()
        {
            try
            {
                string sresp = "";

                if (cmbHoleIDBox.SelectedValue.ToString() == "Select an option..")
                {
                    sresp = "Selected an option Hole ID";
                    return sresp;
                }

                if (txtFromBox.Text == "" || txtToBox.Text == "")
                {
                    sresp = "Empty From or To";
                    return sresp;
                }

                if (double.Parse(txtFromBox.Text.ToString()) == double.Parse(txtToBox.Text.ToString()))
                {
                    sresp = " 'From' equal to 'To'";
                    return sresp;
                }

                if (double.Parse(txtFromBox.Text.ToString()) > double.Parse(txtToBox.Text.ToString()))
                {
                    sresp = " 'From' greater than 'To'";
                    return sresp;
                }
               


                oCollars.sHoleID = cmbHoleIDBox.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToBox.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'To' greater than Hole Id lenght";
                    return sresp;
                }


                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddBox_Click(object sender, EventArgs e)
        {
            try
            {

                string sResp = ControlsValidateBox().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Box", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataTable dtValidRange = new DataTable();
                oBox.dFrom = double.Parse(txtFromBox.Text.ToString());
                oBox.dTo = double.Parse(txtToBox.Text.ToString());
                oBox.sHoleID = cmbHoleIDBox.SelectedValue.ToString();
                dtValidRange = oBox.getDHBoxFromToValid();
                if (dtValidRange.Rows.Count > 0)
                {
                    MessageBox.Show("Range 'From To' Overlaps", "Box", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }



                if (sEditBox == "1")
                { oBox.sOpcion = "2"; }
                else
                {
                    oBox.sOpcion = "1";
                    oBox.iSKDHBox = 0;
                }

                if (dgBox.Rows.Count <= 1)
                {
                    oBox.dFrom = 0;
                }
                else { oBox.dFrom = double.Parse(txtFromBox.Text.ToString()); }

                oBox.dTo = double.Parse(txtToBox.Text.ToString());
                oBox.sHoleID = cmbHoleIDBox.SelectedValue.ToString();
                oBox.iBox = int.Parse(txtNoBox.Text.ToString());

                if (txtStand.Text.ToString() == "")
                    oBox.iStand =   null;
                else oBox.iStand = int.Parse(txtStand.Text.ToString());

                if (txtColumnBox.Text.ToString() == "")
                    oBox.sColumn = null;
                else oBox.sColumn = txtColumnBox.Text.ToString();

                if (txtRowBox.Text.ToString() == "")
                    oBox.sRow = null;
                else oBox.sRow = txtRowBox.Text.ToString();

                if (txtPhotoBox.Text.ToString() == "")
                    oBox.iPhoto = null;
                else oBox.iPhoto = int.Parse(txtPhotoBox.Text.ToString());

                if (txtEditPhotoBox.Text.ToString() == "")
                    oBox.iEditPhoto = null;
                else oBox.iEditPhoto = int.Parse(txtEditPhotoBox.Text.ToString());


                string sAddBox = oBox.DH_Box_Add();
                if (sAddBox == "OK")
                {
                    FilldgBox("2");


                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Box", sEditBox == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIDBox.SelectedValue.ToString() + "." +
                        " From: " + txtFromBox.Text.ToString() + "." +
                        " To: " + txtToBox.Text.ToString() + "." +
                        " Box: " + txtNoBox.Text.ToString() + "." +
                        " Stand: " + txtStand.Text.ToString() + "." +
                        " Column: " + txtColumnBox.Text.ToString() + "." +
                        " Row: " + txtRowBox.Text.ToString());


                    //sEditBox = "0";
                    clsDHBox.sStaticFrom = txtToBox.Text.ToString();
                    
                    CleanControlsBox();

                    txtFromBox.Text = clsDHBox.sStaticFrom.ToString();
                    txtToBox.Focus();
                }
                else
                {
                    MessageBox.Show("Error Insert: " + sAddBox.ToString(), "Box", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show("Add Box Error: "+ ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }
        }

        private void CleanControlsBox()
        {
            try
            {
                oBox.iSKDHBox = 0;
                sEditBox = "0";

                txtToBox.Text ="";
                txtNoBox.Text = "";
                txtStand.Text = "";
                txtColumnBox.Text = "";
                txtRowBox.Text = "";
                txtPhotoBox.Text = "";
                txtEditPhotoBox.Text = "";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void cmbHoleIDBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldgBox("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FillCmbBox()
        {
            try
            {
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCBox = dtCollars.NewRow();
                drCBox[0] = "Select an option..";
                dtCollars.Rows.Add(drCBox);
                cmbHoleIDBox.DisplayMember = "HoleID";
                cmbHoleIDBox.ValueMember = "HoleID";
                cmbHoleIDBox.DataSource = dtCollars;
                cmbHoleIDBox.SelectedValue = "Select an option..";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void dgBox_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                oBox.iSKDHBox = Int64.Parse(dgBox.Rows[e.RowIndex].Cells["SKDHBox"].Value.ToString());
                sEditBox = "1";

                txtFromBox.Text = dgBox.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToBox.Text = dgBox.Rows[e.RowIndex].Cells["To"].Value.ToString();
                cmbHoleIDBox.SelectedValue = dgBox.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtNoBox.Text = dgBox.Rows[e.RowIndex].Cells["Box"].Value.ToString();
                txtStand.Text = dgBox.Rows[e.RowIndex].Cells["Stand"].Value.ToString();
                txtColumnBox.Text = dgBox.Rows[e.RowIndex].Cells["Column"].Value.ToString();
                txtRowBox.Text = dgBox.Rows[e.RowIndex].Cells["Row"].Value.ToString();
                txtEditPhotoBox.Text = dgBox.Rows[e.RowIndex].Cells["EditPhoto"].Value.ToString();
                txtPhotoBox.Text = dgBox.Rows[e.RowIndex].Cells["Photo"].Value.ToString();

            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Box", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void dgBox_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Hole Id" + dgBox.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                   + " From " + dgBox.Rows[e.RowIndex].Cells["From"].Value.ToString()
                   + " To " + dgBox.Rows[e.RowIndex].Cells["To"].Value.ToString()
                   , "Box", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oBox.iSKDHBox = Int64.Parse(dgBox.Rows[e.RowIndex].Cells["SKDHBox"].Value.ToString());
                    string sDelete = oBox.DH_Box_Delete();
                    if (sDelete == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Box", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FilldgBox("2");
                        sEditBox = "0";
                        //CleanControlsGeo();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        #endregion


        #region Alterations

        private void CargarCombosAlt(DataTable _dt, ComboBox _cbox)
        {
            try
            {
                if (_dt.Rows.Count > 0)
                {
                    _cbox.DataSource = _dt.Copy();
                    _cbox.ValueMember = _dt.Columns[0].ToString();
                    _cbox.DisplayMember = _dt.Columns[1].ToString();
                    _cbox.SelectedValue = "-1";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FilldgAlterations(string _sOpcion)
        {
            try
            {
                DataTable dtAlterations = new DataTable();
                oAlt.sOpcion = _sOpcion;
                oAlt.sHoleID = cmbHoleIDAlt.SelectedValue.ToString();
                dtAlterations = oAlt.getDH_Alterations();
                dgAlterations.DataSource = dtAlterations;

                dgAlterations.Columns["SKDHAlterarions"].Visible = false;

            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }


        private void FillCmbAlt()
        {
            try
            {
                DataTable dtAlt = new DataTable();
                dtAlt = oRf.getRfTypeAlt_List();
                DataRow drAlt = dtAlt.NewRow();
                drAlt[0] = "-1";
                drAlt[1] = "Select an option..";
                dtAlt.Rows.Add(drAlt);

                CargarCombosAlt(dtAlt, cmbTypeAlt);
                CargarCombosAlt(dtAlt, cmbTypeAlt2);

                CargarCombosAlt(dtAlt, cmbAltTypeDens);

                DataTable dtIntensity = new DataTable();
                dtIntensity = oRf.getRfIntensityAlt_List(ConfigurationSettings.AppSettings["IDProjectGC"].ToString());
                DataRow drInt = dtIntensity.NewRow();
                drInt[0] = "-1";
                drInt[1] = "Select an option..";
                dtIntensity.Rows.Add(drInt);

                CargarCombosAlt(dtIntensity, cmbIntAlt);
                CargarCombosAlt(dtIntensity, cmbIntAlt2);

                CargarCombosAlt(dtIntensity, cmbAltIntensityDens);

                DataTable dtMinAlt = new DataTable();
                dtMinAlt = oRf.getRfMinerAlt_List();
                DataRow drMinA = dtMinAlt.NewRow();
                drMinA[0] = "-1";
                drMinA[1] = "Select an option..";
                dtMinAlt.Rows.Add(drMinA);

                CargarCombosAlt(dtMinAlt, cmbMin1Alt);
                CargarCombosAlt(dtMinAlt, cmbMin1Alt2);
                CargarCombosAlt(dtMinAlt, cmbMin2Alt1);
                CargarCombosAlt(dtMinAlt, cmbMin2Alt2);
                CargarCombosAlt(dtMinAlt, cmbMin3Alt1);
                CargarCombosAlt(dtMinAlt, cmbMin3Alt2);


                DataTable dtStyleAlt = new DataTable();
                dtStyleAlt = oRf.getRfStyleAlt_List();
                DataRow drStyleA = dtStyleAlt.NewRow();
                drStyleA[0] = "-1";
                drStyleA[1] = "Select an option..";
                dtStyleAlt.Rows.Add(drStyleA);

                CargarCombosAlt(dtStyleAlt, cmbStyleAlt1);
                CargarCombosAlt(dtStyleAlt, cmbStyleAlt2);
                CargarCombosAlt(dtStyleAlt, cmbStyleAlt12);
                CargarCombosAlt(dtStyleAlt, cmbStyleAlt22);


            }
            catch (Exception ex)
            {
                throw new Exception("Error FillCmbAlt: " + ex.Message);
            }
        }

        private void FillCmbAlterations()
        {
            try
            {
                DataTable dtCollars = oCollars.getDHCollarsLogged();
                DataRow drCBox = dtCollars.NewRow();
                drCBox[0] = "Select an option..";
                dtCollars.Rows.Add(drCBox);
                cmbHoleIDAlt.DisplayMember = "HoleID";
                cmbHoleIDAlt.ValueMember = "HoleID";
                cmbHoleIDAlt.DataSource = dtCollars;
                cmbHoleIDAlt.SelectedValue = "Select an option..";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private string ControlsValidateAlt()
        {
            try
            {
                string sresp = "";

                if (cmbHoleIDAlt.SelectedValue.ToString() == "Select an option..")
                {
                    sresp = "Selected an option Hole ID";
                    return sresp;
                }

                if (txtFromAlt.Text == "" || txtToAlt.Text == "")
                {
                    sresp = "Empty From or To";
                    return sresp;
                }

                if (double.Parse(txtFromAlt.Text.ToString()) == double.Parse(txtToAlt.Text.ToString()))
                {
                    sresp = " 'From' equal to 'To'";
                    return sresp;
                }

                if (double.Parse(txtFromAlt.Text.ToString()) > double.Parse(txtToAlt.Text.ToString()))
                {
                    sresp = " 'From' greater than 'To'";
                    return sresp;
                }

                if (cmbTypeAlt.SelectedValue.ToString() == "-1" ) //||
                    //cmbIntAlt.SelectedValue.ToString() == "-1" ||
                    //cmbMin1Alt.SelectedValue.ToString() == "-1" ||
                    //cmbStyleAlt1.SelectedValue.ToString() == "-1")
                {
                    sresp = "You must fill Alteration 1";
                    return sresp;
                }



                oCollars.sHoleID = cmbHoleIDBox.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtToAlt.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'To' greater than Hole Id lenght";
                    return sresp;
                }


                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void CleanControlsAlt()
        {
            try
            {
                oAlt.iSHDHAlterarions = 0;
                sEditAlt = "0";

                txtToAlt.Text = "";
                txtCommentsAlt.Text = "";

                cmbTypeAlt.SelectedValue = "-1";
                cmbTypeAlt2.SelectedValue = "-1";
                cmbIntAlt.SelectedValue = "-1";
                cmbIntAlt2.SelectedValue = "-1";
                cmbStyleAlt1.SelectedValue = "-1";
                cmbStyleAlt2.SelectedValue = "-1";
                cmbMin1Alt.SelectedValue = "-1";
                cmbMin1Alt2.SelectedValue = "-1";
                cmbMin2Alt1.SelectedValue = "-1";
                cmbStyleAlt12.SelectedValue = "-1";
                cmbStyleAlt22.SelectedValue = "-1";
                cmbMin2Alt2.SelectedValue = "-1";
                cmbMin3Alt1.SelectedValue = "-1";
                cmbMin3Alt2.SelectedValue = "-1";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddAlt_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateAlt().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Alterations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                if (sEditAlt == "1")
                {
                    oAlt.sOpcion = "2";
                }
                else
                {
                    oAlt.iSHDHAlterarions = 0;
                    oAlt.sOpcion = "1";
                }

                oAlt.sHoleID = cmbHoleIDAlt.SelectedValue.ToString();
                if (dgAlterations.Rows.Count <= 1)
                {
                    oAlt.dFrom = 0;
                }
                else { oAlt.dFrom = double.Parse(txtFromAlt.Text.ToString()); }
                oAlt.dTo = double.Parse(txtToAlt.Text.ToString());

                DataTable dtValidRange = new DataTable();
                dtValidRange = oAlt.getDHAlterationsFromToValid();
                if (dtValidRange.Rows.Count > 0)
                {
                    MessageBox.Show("Range 'From To' Overlaps", "Alterations", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                oAlt.sA1Type = cmbTypeAlt.SelectedValue.ToString();
                
                
                
                if (cmbTypeAlt2.SelectedValue.ToString() == "-1" || cmbTypeAlt2.SelectedValue.ToString() == "")
                    oAlt.sA2Type = null;
                else oAlt.sA2Type = cmbTypeAlt2.SelectedValue.ToString();

                if (cmbIntAlt.SelectedValue.ToString() == "-1" || cmbIntAlt.SelectedValue.ToString() == "")
                    oAlt.sA1Int = null;
                else oAlt.sA1Int = cmbIntAlt.SelectedValue.ToString();

                if (cmbIntAlt2.SelectedValue.ToString() == "-1" || cmbIntAlt2.SelectedValue.ToString() == "")
                    oAlt.sA2Int = null;
                else oAlt.sA2Int = cmbIntAlt2.SelectedValue.ToString();

                if (cmbStyleAlt1.SelectedValue.ToString() == "-1" || cmbStyleAlt1.SelectedValue.ToString() == "")
                    oAlt.sA1Style = null;
                else oAlt.sA1Style = cmbStyleAlt1.SelectedValue.ToString();

                if (cmbStyleAlt2.SelectedValue.ToString() == "-1" || cmbStyleAlt2.SelectedValue.ToString() == "")
                    oAlt.sA2Style = null;
                else oAlt.sA2Style = cmbStyleAlt2.SelectedValue.ToString();

                if (cmbMin1Alt.SelectedValue.ToString() == "-1" || cmbMin1Alt.SelectedValue.ToString() == "")
                    oAlt.sA1Min = null;
                else oAlt.sA1Min = cmbMin1Alt.SelectedValue.ToString();

                if (cmbMin1Alt2.SelectedValue.ToString() == "-1" || cmbMin1Alt2.SelectedValue.ToString() == "")
                    oAlt.sA2Min = null;
                else oAlt.sA2Min = cmbMin1Alt2.SelectedValue.ToString();

                if (txtCommentsAlt.Text.ToString() == "")
                    oAlt.sComments = null;
                else oAlt.sComments = oAlt.sComments = txtCommentsAlt.Text.ToString();

                if (cmbMin2Alt1.SelectedValue.ToString() == "-1" || cmbMin2Alt1.SelectedValue.ToString() == "")
                    oAlt.sA1Min2 = null;
                else oAlt.sA1Min2 = cmbMin2Alt1.SelectedValue.ToString();

                if (cmbMin2Alt2.SelectedValue.ToString() == "-1" || cmbMin2Alt2.SelectedValue.ToString() == "")
                    oAlt.sA2Min2 = null;
                else oAlt.sA2Min2 = cmbMin2Alt2.SelectedValue.ToString();

                if (cmbStyleAlt12.SelectedValue.ToString() == "-1" || cmbStyleAlt12.SelectedValue.ToString() == "")
                    oAlt.sA1Style2 = null;
                else oAlt.sA1Style2 = cmbStyleAlt12.SelectedValue.ToString();

                if (cmbStyleAlt22.SelectedValue.ToString() == "-1" || cmbStyleAlt22.SelectedValue.ToString() == "")
                    oAlt.sA2Style2 = null;
                else oAlt.sA2Style2 = cmbStyleAlt22.SelectedValue.ToString();

                if (cmbMin3Alt1.SelectedValue.ToString() == "-1" || cmbMin3Alt1.SelectedValue.ToString() == "")
                    oAlt.sA1Min3 = null;
                else oAlt.sA1Min3 = cmbMin3Alt1.SelectedValue.ToString();

                if (cmbMin3Alt2.SelectedValue.ToString() == "-1" || cmbMin3Alt2.SelectedValue.ToString() == "")
                    oAlt.sA2Min3 = null;
                else oAlt.sA2Min3 = cmbMin3Alt2.SelectedValue.ToString();

                clsDHAlterations.sStaticFrom = txtToAlt.Text.ToString();

                string sRespAltAdd = oAlt.DH_Alterations_Add();
                if (sRespAltAdd.ToString() == "OK")
                {
                    FilldgAlterations("2");
                    
                    //sEditAlt = "0";

                    //Insertar el registro para el historial de transacciones por usuario
                    oRf.InsertTrans("DH_Alterations", sEditAlt == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                        "Hole ID: " + cmbHoleIDAlt.SelectedValue.ToString() + "." +
                        " From: " + txtFromAlt.Text.ToString() + "." +
                        " To: " + txtToAlt.Text.ToString() + "." +
                        " Type Alt: " + cmbTypeAlt.SelectedValue.ToString() + "." +
                        " Intensity Alt: " + cmbIntAlt.SelectedValue.ToString() + "." +
                        " Mineral Alt: " + cmbMin1Alt.SelectedValue.ToString() + "." +
                        " Style Alt: " + cmbStyleAlt1.SelectedValue.ToString());

                    if (sEditAlt == "1")
                    {
                        if (dgAlterations.Rows.Count > 1)
                        {
                            DataTable dt = (DataTable)dgAlterations.DataSource;
                            DataRow[] myRow = dt.Select(@"SKDHAlterarions = '" + oAlt.iSHDHAlterarions + "'");
                            int rowindex = dt.Rows.IndexOf(myRow[0]);
                            dgAlterations.Rows[rowindex].Selected = true;
                            dgAlterations.CurrentCell = dgAlterations.Rows[rowindex].Cells[1];
                        }
                    }


                    CleanControlsAlt();

                    txtFromAlt.Text = clsDHAlterations.sStaticFrom.ToString();
                    txtToAlt.Focus();

                }
                else
                {
                    MessageBox.Show("Error Insert: " + sRespAltAdd.ToString(), "Alterations", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }


            }
            catch (Exception ex)
            {
                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("You must enter all required records", "Structure", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                
            }
        }

        private void cmbHoleIDAlt_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldgAlterations("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgAlterations_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                oAlt.iSHDHAlterarions = Int64.Parse(dgAlterations.Rows[e.RowIndex].Cells["SKDHAlterarions"].Value.ToString());
                sEditAlt = "1";


                cmbHoleIDAlt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtFromAlt.Text = dgAlterations.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToAlt.Text = dgAlterations.Rows[e.RowIndex].Cells["To"].Value.ToString();
                cmbTypeAlt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Type"].Value.ToString();

                cmbTypeAlt2.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Type"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Type"].Value.ToString();

                cmbIntAlt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Int"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A1Int"].Value.ToString();

                cmbIntAlt2.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Int"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Int"].Value.ToString();

                cmbStyleAlt1.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Style"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A1Style"].Value.ToString();

                cmbStyleAlt2.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Style"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Style"].Value.ToString();

                cmbMin1Alt.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Min"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A1Min"].Value.ToString();

                cmbMin1Alt2.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Min"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Min"].Value.ToString();

                txtCommentsAlt.Text = dgAlterations.Rows[e.RowIndex].Cells["Comments"].Value.ToString() == "" ?
                    "" : dgAlterations.Rows[e.RowIndex].Cells["Comments"].Value.ToString();

                cmbMin2Alt1.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Min2"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A1Min2"].Value.ToString();

                cmbMin2Alt2.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Min2"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Min2"].Value.ToString();

                cmbStyleAlt12.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Style2"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A1Style2"].Value.ToString();

                cmbStyleAlt22.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Style2"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Style2"].Value.ToString();


                cmbMin3Alt1.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A1Min3"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A1Min3"].Value.ToString();

                cmbMin3Alt2.SelectedValue = dgAlterations.Rows[e.RowIndex].Cells["A2Min3"].Value.ToString() == "" ?
                    "-1" : dgAlterations.Rows[e.RowIndex].Cells["A2Min3"].Value.ToString();
            }
            catch (Exception ex)
            {
                if (ex.GetType().Name == "FormatException")
                {
                    MessageBox.Show("Invalid Data", "Geotech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelAlt_Click(object sender, EventArgs e)
        {
            CleanControlsAlt();
        }

        #endregion

        private void txtFromAlt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);

        }

        private void txtToAlt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void dgAlterations_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "Hole Id" + dgAlterations.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                   + " From " + dgAlterations.Rows[e.RowIndex].Cells["From"].Value.ToString()
                   + " To " + dgAlterations.Rows[e.RowIndex].Cells["To"].Value.ToString()
                   , "Box", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oAlt.iSHDHAlterarions = Int64.Parse(dgAlterations.Rows[e.RowIndex].Cells["SKDHAlterarions"].Value.ToString());
                    string sDelete = oAlt.DH_Alterations_Delete();
                    if (sDelete == "OK")
                    {
                        MessageBox.Show("Row Deleted", "Alterations", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FilldgAlterations("2");
                        sEditAlt = "0";
                        //CleanControlsGeo();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtUpAngleSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void txtBtnAngleSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void txtAppThickSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void txtNumberSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void cmbLithologyLit_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                oRf.sCodeLith = cmbLithologyLit.SelectedValue.ToString();

                /*DataRow dr = dtSampleT.NewRow();
                dr[0] = "-1";
                dr[1] = "Select an option..";
                dtSampleT.Rows.Add(dr);
                 * 
                cmbHoleIDBox.SelectedValue = "Select an option..";*/

                DataTable dtGSize = new DataTable();
                dtGSize = oRf.getRFGsize_List();

                DataRow drG = dtGSize.NewRow();
                drG[0] = "-1";
                drG[1] = "Select an option..";
                dtGSize.Rows.Add(drG);
                cmbGsizeLith.DisplayMember = "Comb";
                cmbGsizeLith.ValueMember = "Code";
                cmbGsizeLith.DataSource = dtGSize;
                cmbGsizeLith.SelectedValue = "-1";

                DataTable dtTextures = new DataTable();
                dtTextures = oRf.getRfTextures_List();
                DataRow drTx = dtTextures.NewRow();
                drTx[0] = "-1";
                drTx[1] = "Select an option..";
                dtTextures.Rows.Add(drTx);
                cmbTexturesLith.DisplayMember = "Comb";
                cmbTexturesLith.ValueMember = "Code";
                cmbTexturesLith.DataSource = dtTextures;
                cmbTexturesLith.SelectedValue = "-1";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        #region Export Excel

        private void ExcelGenerateSamples(Excel._Worksheet _oSheet)
        {
            try
            {
                //oSheet.Cells[1, 27] = cmbHoleIDForm.SelectedValue.ToString();
                _oSheet.get_Range("AA1", "AD1").MergeCells = true;
                _oSheet.get_Range("AA1", "AD1").Value2 = cmbHoleIDForm.SelectedValue.ToString();

                DataTable dtLogging = (DataTable)gdLoggin.DataSource;

                int iInicial = 6;
                for (int i = 0; i < dtLogging.Rows.Count - 1; i++)
                {

                    _oSheet.Cells[iInicial, 1] = dtLogging.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtLogging.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtLogging.Rows[i]["Sample"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtLogging.Rows[i]["SampleType"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtLogging.Rows[i]["DupDe"].ToString();
                    _oSheet.Cells[iInicial, 6] = dtLogging.Rows[i]["Lithology"].ToString();
                    _oSheet.Cells[iInicial, 7] = dtLogging.Rows[i]["Comments"].ToString();
                    iInicial += 1;
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateWeathering(Excel._Worksheet _oSheet)
        {
            try
            {

                DataTable dtWeathering = (DataTable)dgWeathering.DataSource;

                int iInicial = 2;
                for (int i = 0; i < dtWeathering.Rows.Count - 1; i++)
                {
                    _oSheet.Cells[iInicial, 1] = dtWeathering.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtWeathering.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtWeathering.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtWeathering.Rows[i]["Weathering"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtWeathering.Rows[i]["Oxidation"].ToString();

                    _oSheet.Cells[iInicial, 6] = dtWeathering.Rows[i]["Mineral1"].ToString();
                    _oSheet.Cells[iInicial, 7] = dtWeathering.Rows[i]["Mineral2"].ToString();
                    _oSheet.Cells[iInicial, 8] = dtWeathering.Rows[i]["Mineral3"].ToString();
                    _oSheet.Cells[iInicial, 9] = dtWeathering.Rows[i]["Mineral4"].ToString();


                    _oSheet.Cells[iInicial, 10] = dtWeathering.Rows[i]["Colour1"].ToString();
                    _oSheet.Cells[iInicial, 11] = dtWeathering.Rows[i]["Sufix1"].ToString();
                    _oSheet.Cells[iInicial, 12] = dtWeathering.Rows[i]["Observation"].ToString();
                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateAlterations(Excel._Worksheet _oSheet)
        {
            try
            {

                DataTable dtAlterations = (DataTable)dgAlterations.DataSource;

                int iInicial = 3;
                for (int i = 0; i < dtAlterations.Rows.Count - 1; i++)
                {

                    _oSheet.Cells[iInicial, 1] = dtAlterations.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtAlterations.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtAlterations.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtAlterations.Rows[i]["A1Type"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtAlterations.Rows[i]["A1Int"].ToString();
                    _oSheet.Cells[iInicial, 6] = dtAlterations.Rows[i]["A1Style"].ToString();
                    _oSheet.Cells[iInicial, 7] = dtAlterations.Rows[i]["A1Style2"].ToString()
                           == "-1" ? "" : dtAlterations.Rows[i]["A1Style2"].ToString();
                    _oSheet.Cells[iInicial, 8] = dtAlterations.Rows[i]["A1Min"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A1Min"].ToString();
                    _oSheet.Cells[iInicial, 9] = dtAlterations.Rows[i]["A1Min2"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A1Min2"].ToString(); 
                    _oSheet.Cells[iInicial, 10] = dtAlterations.Rows[i]["A1Min3"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A1Min3"].ToString();
                    _oSheet.Cells[iInicial, 11] = dtAlterations.Rows[i]["A2Type"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A2Type"].ToString();
                    _oSheet.Cells[iInicial, 12] = dtAlterations.Rows[i]["A2Int"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A2Int"].ToString();
                    _oSheet.Cells[iInicial, 13] = dtAlterations.Rows[i]["A2Style"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Style"].ToString();
                    _oSheet.Cells[iInicial, 14] = dtAlterations.Rows[i]["A2Style2"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Style2"].ToString(); 
                    _oSheet.Cells[iInicial, 15] = dtAlterations.Rows[i]["A2Min"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Min"].ToString();
                    _oSheet.Cells[iInicial, 16] = dtAlterations.Rows[i]["A2Min"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Min2"].ToString();
                    _oSheet.Cells[iInicial, 17] = dtAlterations.Rows[i]["A2Min3"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Min3"].ToString();

                    _oSheet.Cells[iInicial, 18] = dtAlterations.Rows[i]["Comments"].ToString();

                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateStructures(Excel._Worksheet _oSheet)
        {
            try
            {

                DataTable dtStr = (DataTable)dgStructure.DataSource;

                int iInicial = 2;
                for (int i = 0; i < dtStr.Rows.Count - 1; i++)
                {

                    _oSheet.Cells[iInicial, 1] = dtStr.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtStr.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtStr.Rows[i]["To"].ToString(); ;
                    _oSheet.Cells[iInicial, 4] = dtStr.Rows[i]["Type"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtStr.Rows[i]["AngleToAxis"].ToString();
                    _oSheet.Cells[iInicial, 6] = dtStr.Rows[i]["UpAngle"].ToString();

                    _oSheet.Cells[iInicial, 7] = dtStr.Rows[i]["BtonAngle"].ToString();
                    _oSheet.Cells[iInicial, 8] = dtStr.Rows[i]["AppThick"].ToString();
                    _oSheet.Cells[iInicial, 9] = dtStr.Rows[i]["Fill"].ToString();
                    _oSheet.Cells[iInicial, 10] = dtStr.Rows[i]["Fill2"].ToString();
                    _oSheet.Cells[iInicial, 11] = dtStr.Rows[i]["Fill3"].ToString();
                    _oSheet.Cells[iInicial, 12] = dtStr.Rows[i]["Fill4"].ToString();
                    _oSheet.Cells[iInicial, 13] = dtStr.Rows[i]["Number"].ToString();
                    _oSheet.Cells[iInicial, 14] = dtStr.Rows[i]["Comments"].ToString();
                    _oSheet.Cells[iInicial, 15] = dtStr.Rows[i]["Lenght"].ToString();

                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateMineralizations(Excel._Worksheet _oSheet)
        {
            try
            {
                DataTable dtMiner = (DataTable)dgMineraliz.DataSource;

                int iInicial = 3;
                for (int i = 0; i < dtMiner.Rows.Count - 1; i++)
                {

                    /*Gsize, GSize2, GSize3*/
                    _oSheet.Cells[iInicial, 1] = dtMiner.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtMiner.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtMiner.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtMiner.Rows[i]["MZ1Mineral"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtMiner.Rows[i]["MZ1Mineral2"].ToString();
                    _oSheet.Cells[iInicial, 6] = dtMiner.Rows[i]["MZ1Mineral3"].ToString();
                    _oSheet.Cells[iInicial, 7] = dtMiner.Rows[i]["MZ1Style"].ToString();
                    _oSheet.Cells[iInicial, 8] = dtMiner.Rows[i]["MZ1Perc"].ToString();

                    _oSheet.Cells[iInicial, 9] = dtMiner.Rows[i]["Gsize"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["Gsize"].ToString();

                    //sEditGeo == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                    _oSheet.Cells[iInicial, 10] = dtMiner.Rows[i]["MZ2Mineral"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ2Mineral"].ToString();
                    _oSheet.Cells[iInicial, 11] = dtMiner.Rows[i]["MZ2Mineral2"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ2Mineral2"].ToString();
                    _oSheet.Cells[iInicial, 12] = dtMiner.Rows[i]["MZ2Mineral3"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ2Mineral3"].ToString();
                    _oSheet.Cells[iInicial, 13] = dtMiner.Rows[i]["MZ2Style"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ2Style"].ToString();
                    _oSheet.Cells[iInicial, 14] = dtMiner.Rows[i]["MZ2Perc"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ2Perc"].ToString();

                    _oSheet.Cells[iInicial, 15] = dtMiner.Rows[i]["GSize2"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["GSize2"].ToString();

                    _oSheet.Cells[iInicial, 16] = dtMiner.Rows[i]["MZ3Mineral"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ3Mineral"].ToString();
                    _oSheet.Cells[iInicial, 17] = dtMiner.Rows[i]["MZ3Mineral2"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ3Mineral2"].ToString();
                    _oSheet.Cells[iInicial, 18] = dtMiner.Rows[i]["MZ3Mineral3"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ3Mineral3"].ToString();
                    _oSheet.Cells[iInicial, 19] = dtMiner.Rows[i]["MZ3Style"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ3Style"].ToString();
                    _oSheet.Cells[iInicial, 20] = dtMiner.Rows[i]["MZ3Perc"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ3Perc"].ToString();

                    _oSheet.Cells[iInicial, 21] = dtMiner.Rows[i]["GSize3"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["GSize3"].ToString();

                    _oSheet.Cells[iInicial, 22] = dtMiner.Rows[i]["Comments"].ToString();

                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateLithology(Excel._Worksheet _oSheet)
        {
            try
            {
                DataTable dtLitho = (DataTable)dgLithology.DataSource;

                int iInicial = 2;
                for (int i = 0; i < dtLitho.Rows.Count - 1; i++)
                {
                    _oSheet.Cells[iInicial, 1] = dtLitho.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtLitho.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtLitho.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtLitho.Rows[i]["Litho"].ToString();
                    
                    _oSheet.Cells[iInicial, 5] = dtLitho.Rows[i]["Textures"].ToString();
                    _oSheet.Cells[iInicial, 6] = dtLitho.Rows[i]["GSize"].ToString();
                    
                    _oSheet.Cells[iInicial, 7] = dtLitho.Rows[i]["Observation"].ToString();
                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateGeotech(Excel._Worksheet _oSheet)
        {
            try
            {

                DataTable dtGeo = (DataTable)dgGeotech.DataSource;

                int iInicial = 2;
                for (int i = 0; i < dtGeo.Rows.Count - 1; i++)
                {
                    _oSheet.Cells[iInicial, 1] = dtGeo.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtGeo.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtGeo.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtGeo.Rows[i]["LithCod"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtGeo.Rows[i]["Recm"].ToString();
                    
                    _oSheet.Cells[iInicial, 6] = dtGeo.Rows[i]["RQDcm"].ToString();
                    _oSheet.Cells[iInicial, 7] = dtGeo.Rows[i]["NoOfFract"].ToString();
                    _oSheet.Cells[iInicial, 8] = dtGeo.Rows[i]["JointCond"].ToString();
                    _oSheet.Cells[iInicial, 9] = dtGeo.Rows[i]["Jn"].ToString();
                    _oSheet.Cells[iInicial, 10] = dtGeo.Rows[i]["Jr"].ToString();
                    _oSheet.Cells[iInicial, 11] = dtGeo.Rows[i]["Ja"].ToString();
                    _oSheet.Cells[iInicial, 12] = dtGeo.Rows[i]["DegBreak"].ToString();
                    _oSheet.Cells[iInicial, 13] = dtGeo.Rows[i]["Hardness"].ToString();
                    _oSheet.Cells[iInicial, 14] = dtGeo.Rows[i]["Comments"].ToString();
                    _oSheet.Cells[iInicial, 15] = dtGeo.Rows[i]["AltWeath"].ToString();
            

                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelGenerateBox(Excel._Worksheet _oSheet)
        {
            try
            {

                DataTable dtBox = (DataTable)dgBox.DataSource;

                int iInicial = 5;
                for (int i = 0; i < dtBox.Rows.Count - 1; i++)
                {
                    _oSheet.Cells[iInicial, 1] = dtBox.Rows[i]["HoleID"].ToString();
                    _oSheet.Cells[iInicial, 2] = dtBox.Rows[i]["From"].ToString();
                    _oSheet.Cells[iInicial, 3] = dtBox.Rows[i]["To"].ToString();
                    _oSheet.Cells[iInicial, 4] = dtBox.Rows[i]["Box"].ToString();
                    _oSheet.Cells[iInicial, 5] = dtBox.Rows[i]["Stand"].ToString();
                    _oSheet.Cells[iInicial, 6] = dtBox.Rows[i]["column"].ToString();
                    _oSheet.Cells[iInicial, 7] = dtBox.Rows[i]["row"].ToString();
                    
                    iInicial += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }



        private void ExcelHeader(string _sSheetExcel)
        {
            try
            {
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                oXL = new Excel.Application();
                //oXL.Visible = true;

                oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_LoggingAll"].ToString(),
                    0, false, 5,
                Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                Type.Missing, false, false, false);


                oSheet = (Excel._Worksheet)oWB.Sheets[_sSheetExcel];//(Excel._Worksheet)oWB.ActiveSheet;

                switch (_sSheetExcel)
                {
                    case "Alterations":
                        ExcelGenerateAlterations(oSheet); ;
                        break;
                    case "Geotech":
                        ExcelGenerateGeotech(oSheet);
                        break;
                    case "Mineraliz":
                        ExcelGenerateMineralizations(oSheet);
                        break;
                    case "Lithology":
                        ExcelGenerateLithology(oSheet);
                        break;
                    case "Structures":
                        ExcelGenerateStructures(oSheet);
                        break;
                    case "Weatering":
                        ExcelGenerateWeathering(oSheet);
                        break;
                }



                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void ExcelHeader()
        {
            try
            {
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                string _sSheetExcel = SheetExcel;

                oXL = new Excel.Application();
                //oXL.Visible = true;

                oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_LoggingAll"].ToString(),
                    0, false, 5,
                Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                Type.Missing, false, false, false);


                oSheet = (Excel._Worksheet)oWB.Sheets[_sSheetExcel];//(Excel._Worksheet)oWB.ActiveSheet;

                switch (_sSheetExcel)
                {
                    case "Alterations":
                        ExcelGenerateAlterations(oSheet); ;
                        break;
                    case "Geotech":
                        ExcelGenerateGeotech(oSheet);
                        break;
                    case "Mineraliz":
                        ExcelGenerateMineralizations(oSheet);
                        break;
                    case "Lithology":
                        ExcelGenerateLithology(oSheet);
                        break;
                    case "Structures":
                        ExcelGenerateStructures(oSheet);
                        break;
                    case "Weatering":
                        ExcelGenerateWeathering(oSheet);
                        break;
                }



                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnExporExcelWeath_Click(object sender, EventArgs e)
        {
            try
            {

                ExcelHeader("Weatering");


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Export Excel: " + ex.Message);
            }
        }

        private void btnExporExcelLith_Click(object sender, EventArgs e)
        {
            try
            {


                pCargando.Visible = true;

                SheetExcel = "Lithology";

                Thread oThread = new Thread(new ThreadStart(ExcelHeader));
                oThread.Start();

                // Wait for foreground thread to end.
                oThread.Join();

                pCargando.Visible = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Export Excel: " + ex.Message);
            }
        }

        private void btnExporExcelAlt_Click(object sender, EventArgs e)
        {
            try
            {

                pCargando.Visible = true;

                SheetExcel = "Alterations";

                Thread oThread = new Thread(new ThreadStart(ExcelHeader));
                oThread.Start();

                // Wait for foreground thread to end.
                oThread.Join();

                pCargando.Visible = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Export Excel: " + ex.Message);
            }
        }

        private void btnExporExcestr_Click(object sender, EventArgs e)
        {
            try
            {

                pCargando.Visible = true;

                SheetExcel = "Structures";

                Thread oThread = new Thread(new ThreadStart(ExcelHeader));
                oThread.Start();

                // Wait for foreground thread to end.
                oThread.Join();

                pCargando.Visible = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Export Excel: " + ex.Message);
            }
        }

        private void btnExporExcelMin_Click(object sender, EventArgs e)
        {
            try
            {

                pCargando.Visible = true;

                SheetExcel = "Mineraliz";

                Thread oThread = new Thread(new ThreadStart(ExcelHeader));
                oThread.Start();

                // Wait for foreground thread to end.
                oThread.Join();

                pCargando.Visible = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Export Excel: " + ex.Message);
            }
        }

        private void btnExporExcelGeo_Click(object sender, EventArgs e)
        {
            try
            {

                pCargando.Visible = true;

                SheetExcel = "Geotech";

                Thread oThread = new Thread(new ThreadStart(ExcelHeader));
                oThread.Start();

                // Wait for foreground thread to end.
                oThread.Join();

                pCargando.Visible = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Export Excel: " + ex.Message);
            }
        }

        #endregion

        //private void TabEnter(KeyPressEventArgs _e)
        //{
        //    if (_e.KeyChar == (char)(Keys.Enter))
        //    {
        //        _e.Handled = true;
        //        SendKeys.Send("{TAB}");
        //    }
        //}

        private void txtSampNoIni_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbHoleID_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }




        # region Oxidation

        //private void FillCmbOxidation()
        //{
        //    try
        //    {
        //        DataTable dtCollars = oCollars.getDHCollarsLogged();
        //        DataRow drCBox = dtCollars.NewRow();
        //        drCBox[0] = "Select an option..";
        //        dtCollars.Rows.Add(drCBox);
        //        cmbHoleIDOx.DisplayMember = "HoleID";
        //        cmbHoleIDOx.ValueMember = "HoleID";
        //        cmbHoleIDOx.DataSource = dtCollars;
        //        cmbHoleIDOx.SelectedValue = "Select an option..";
                
        //        DataTable dtOxidationPerc = oRf.getRfOxides_List();
        //        DataRow drOx = dtOxidationPerc.NewRow();
        //        drOx[0] = "-1";
        //        drOx[1] = "Select an option..";
        //        dtOxidationPerc.Rows.Add(drOx);
        //        cmbOxidesPerc.DisplayMember = "Description";
        //        cmbOxidesPerc.ValueMember = "Code";
        //        cmbOxidesPerc.DataSource = dtOxidationPerc;
        //        cmbOxidesPerc.SelectedValue = "-1";

        //        DataTable dtOxidationInt = oRf.getRfOxidation_List();
        //        DataRow drOxI = dtOxidationInt.NewRow();
        //        drOxI[0] = "-1";
        //        drOxI[1] = "Select an option..";
        //        dtOxidationInt.Rows.Add(drOxI);
        //        cmbOxidesIntOx.DisplayMember = "Description";
        //        cmbOxidesIntOx.ValueMember = "Code";
        //        cmbOxidesIntOx.DataSource = dtOxidationInt;
        //        cmbOxidesIntOx.SelectedValue = "-1";

        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message);
        //    }
        //}

        //private void txtFromOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (Char.IsNumber(e.KeyChar))
        //    {
        //        e.Handled = false;
        //    }
        //    if (Char.IsLetter(e.KeyChar))
        //    {
        //        e.Handled = true;
        //    }


        //    TabEnter(e);
        //}

        //private void txtToOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (Char.IsNumber(e.KeyChar))
        //    {
        //        e.Handled = false;
        //    }
        //    if (Char.IsLetter(e.KeyChar))
        //    {
        //        e.Handled = true;
        //    }


        //    TabEnter(e);
        //}

        //private void cmbHoleIDOx_KeyPress(object sender, KeyPressEventArgs e)
        //{

        //    TabEnter(e);
        //}

        //private void txtHemOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtGtOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtJarOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtLimOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtCuOOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void cmbOxidesPerc_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtOtherOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtOtherGrOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void txtDistOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        //private void cmbOxidesIntOx_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    TabEnter(e);
        //}

        #endregion

        private void cmbHoleIDSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtToSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void cmbStructureTypeSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbFillSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbFillSt2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbFillSt3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbFillSt4_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtCommentsSt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbHoleIdWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            ////TabEnter(e);
        }

        private void txtFromWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            ////TabEnter(e);
        }

        private void txtToWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
            //TabEnter(e);
        }

        private void cmbWeatheringWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbOxidationWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin1Oxid_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin2Oxid_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin3Oxid_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin4Oxid_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbColourWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbSufixWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtObservWeat_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbHoleIDBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ////TabEnter(e);
        }

        private void txtColumnBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ////TabEnter(e);
        }

        private void txtRowBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ////TabEnter(e);
        }

        private void btnCancelBox_Click(object sender, EventArgs e)
        {
            CleanControlsBox();
        }

        private void cmbHoleIdLit_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbLithologyLit_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtObservLit_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbTexturesLith_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbGsizeLith_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbHoleIDAlt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtCommentsAlt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbTypeAlt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbIntAlt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin1Alt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbStyleAlt1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin2Alt1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbTypeAlt2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbIntAlt2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin1Alt2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbStyleAlt2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbMin2Alt2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbHoleIdMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void txtCommentsMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM1Z1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM1Z2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM1Z3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbStyleM1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbPorcM1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM2Z1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM2Z2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM2Z3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbStyleM2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbPorcM2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM3Z1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM3Z2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbM3Z3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbStyleM3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }

        private void cmbPorcM3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //TabEnter(e);
        }


        private void SamplesValid()
        {
            try
            {


                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oSamp.iFrom = 0; oSamp.iTo = 0; oSamp.sHoleID = "0"; oSamp.iDHSampID = 0;
                dtResult = oSamp.getDHSamplesValid();

                for (int i = 0; i < gdLoggin.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oSamp.iFrom = double.Parse(gdLoggin.Rows[i].Cells["From"].Value.ToString());
                    oSamp.iTo = double.Parse(gdLoggin.Rows[i].Cells["To"].Value.ToString());
                    oSamp.sHoleID = gdLoggin.Rows[i].Cells["HoleID"].Value.ToString();
                    oSamp.iDHSampID = long.Parse(gdLoggin.Rows[i].Cells["SKDHSamples"].Value.ToString());
                    dtValid = oSamp.getDHSamplesValid();

                    if (dtValid.Rows.Count > 0)
                    {
                        //DataRowView dv = (DataRowView)gdLoggin.Rows[i].DataBoundItem;
                        //DataRow dr = dv.Row;

                        //implementar ciclo de 1 hasta dtvalid.count
                        dtResult.ImportRow(dtValid.Rows[0]);

                    }
                }


                gdLoggin.DataSource = dtResult;

                

                //Exportar a excel los resultados de from to overlaps y from to next


                oSamp.sHoleID = cmbHoleID.SelectedValue.ToString();
                DataTable dtFromToNext = oSamp.getDHSamplesValidFromToNext();

                DataTable dtFromToLithoValid = oSamp.getDHSamples_Litho_ListValid();


                if (dtResult.Rows.Count > 0 || dtFromToNext.Rows.Count > 0 || dtFromToLithoValid.Rows.Count > 0)
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

                    oSheet.Cells[4, 4] = cmbHoleID.SelectedValue.ToString();

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
                        iInicial += 1;
                    }


                    IEnumerable<DataRow> query =
                    from fromValid in dtFromToLithoValid.AsEnumerable()
                    where fromValid.Field<String>("HoleId") == cmbHoleID.SelectedValue.ToString()
                    select fromValid;

                    // Create a table from the query.
                    DataTable filterTableSampLitho = query.CopyToDataTable<DataRow>();
                    for (int iL = 0; iL < filterTableSampLitho.Rows.Count; iL++)
                    {
                        oSheet.Cells[iInicial, 1] = filterTableSampLitho.Rows[iL]["Sample"].ToString();
                        oSheet.Cells[iInicial, 2] = filterTableSampLitho.Rows[iL]["From"].ToString();
                        oSheet.Cells[iInicial, 3] = filterTableSampLitho.Rows[iL]["To"].ToString();
                        oSheet.Cells[iInicial, 4] = filterTableSampLitho.Rows[iL]["SampleType"].ToString();
                        oSheet.Cells[iInicial, 5] = ""; //dtFromToNext.Rows[iL]["DupDe"].ToString();
                        oSheet.Cells[iInicial, 6] = filterTableSampLitho.Rows[iL]["Lithology"].ToString();
                        oSheet.Cells[iInicial, 7] = filterTableSampLitho.Rows[iL]["error"].ToString();
                        iInicial += 1;
                    }


                    oXL.Visible = true;
                    oXL.UserControl = true;

               
                }


                //Fin Export Excel



            }
            catch (Exception ex)
            {
                //return "Termino";

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Samples", MessageBoxButtons.OK, MessageBoxIcon.Error); }

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

                for (int i = 0; i < dgWeathering.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oWeat.dFrom = double.Parse(dgWeathering.Rows[i].Cells["From"].Value.ToString());
                    oWeat.dTo = double.Parse(dgWeathering.Rows[i].Cells["To"].Value.ToString());
                    oWeat.sHoleID = dgWeathering.Rows[i].Cells["HoleID"].Value.ToString();
                    oWeat.iDHWeatheringID = long.Parse(dgWeathering.Rows[i].Cells["SKDHWeathering"].Value.ToString());
                    dtValid = oWeat.getDHWeatValid();

                    if (dtValid.Rows.Count > 0)
                    {

                        dtResult.ImportRow(dtValid.Rows[0]);

                    }
                }

                
                dgWeathering.DataSource = dtResult;

                oWeat.sHoleID = cmbHoleIdWeat.SelectedValue.ToString();
                DataTable dtWeatFromToNext = oWeat.getDHWeatValidFromToNext();
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

                    oSheet.Cells[4, 3] = cmbHoleIdWeat.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["Weathering"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["Oxidation"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Colour1"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
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
                        iInicial += 1;
                    }



                    oXL.Visible = true;
                    oXL.UserControl = true;

               


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
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void LiththValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oLit.dFrom = 0; oLit.dTo = 0; oLit.sHoleID = "0"; oLit.iDHLithologyID = 0;
                dtResult = oLit.getDHLitValid();

                for (int i = 0; i < dgLithology.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oLit.dFrom = double.Parse(dgLithology.Rows[i].Cells["From"].Value.ToString());
                    oLit.dTo = double.Parse(dgLithology.Rows[i].Cells["To"].Value.ToString());
                    oLit.sHoleID = dgLithology.Rows[i].Cells["HoleID"].Value.ToString();
                    oLit.iDHLithologyID = long.Parse(dgLithology.Rows[i].Cells["SKDHLithology"].Value.ToString());
                    dtValid = oLit.getDHLitValid();

                    if (dtValid.Rows.Count > 0)
                    {
                        dtResult.ImportRow(dtValid.Rows[0]);
                    }
                }

                dgLithology.DataSource = dtResult;

                oLit.sHoleID = cmbHoleIdLit.SelectedValue.ToString();
                DataTable dtLithFromToNext = oLit.getDHLitFromToValidFromToNext();
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

                    oSheet.Cells[4, 3] = cmbHoleIdLit.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["Litho"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["GSize"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Textures"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
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
                        iInicial += 1;
                    }



                    oXL.Visible = true;
                    oXL.UserControl = true;

                    
                }


            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        //getDHAlterationsValid
        private void AlterationsValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oAlt.dFrom = 0; oAlt.dTo = 0; oAlt.sHoleID = "0"; oAlt.iSHDHAlterarions = 0;
                dtResult = oAlt.getDHAlterationsValid();

                for (int i = 0; i < dgAlterations.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oAlt.dFrom = double.Parse(dgAlterations.Rows[i].Cells["From"].Value.ToString());
                    oAlt.dTo = double.Parse(dgAlterations.Rows[i].Cells["To"].Value.ToString());
                    oAlt.sHoleID = dgAlterations.Rows[i].Cells["HoleID"].Value.ToString();
                    oAlt.iSHDHAlterarions = long.Parse(dgAlterations.Rows[i].Cells["SKDHAlterarions"].Value.ToString());
                    dtValid = oAlt.getDHAlterationsValid();

                    if (dtValid.Rows.Count > 0)
                    {
                        dtResult.ImportRow(dtValid.Rows[0]);
                    }
                }

                dgAlterations.DataSource = dtResult;

                oAlt.sHoleID = cmbHoleIDAlt.SelectedValue.ToString();
                DataTable dtAlterFromToNext = oAlt.getDHAlterationsValidFromToNext();
                if (dtResult.Rows.Count > 0 || dtAlterFromToNext.Rows.Count>0)
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
                    
                    oSheet.Cells[4, 3] = cmbHoleIDAlt.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["A1Type"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["A1Int"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["A1Style"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
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
                        iInicial += 1;
                    }



                    oXL.Visible = true;
                    oXL.UserControl = true;
                }

            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        //getDH_StructuresValid
        private void StructuresValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oStr.iFrom = 0; oStr.iTo = 0; oStr.sHoleID = "0"; oStr.iDHStructrueID = 0;
                dtResult = oStr.getDH_StructuresValid();

                for (int i = 0; i < dgStructure.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oStr.iFrom = double.Parse(dgStructure.Rows[i].Cells["From"].Value.ToString());
                    oStr.iTo = double.Parse(dgStructure.Rows[i].Cells["To"].Value.ToString());
                    oStr.sHoleID = dgStructure.Rows[i].Cells["HoleID"].Value.ToString();
                    oStr.iDHStructrueID = long.Parse(dgStructure.Rows[i].Cells["SKDHStructrue"].Value.ToString());
                    dtValid = oStr.getDH_StructuresValid();

                    if (dtValid.Rows.Count > 0)
                    {
                        dtResult.ImportRow(dtValid.Rows[0]);
                    }
                }

                dgStructure.DataSource = dtResult;

                oStr.sHoleID = cmbHoleIDSt.SelectedValue.ToString();
                DataTable dtStrucFromToNext = oStr.getDH_StructuresValidFromToNext();

                if (dtResult.Rows.Count > 0 || dtStrucFromToNext.Rows.Count> 0)
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

                    oSheet.Cells[4, 3] = cmbHoleIDSt.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["Type"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["AngleToAxis"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Fill"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
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
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;
                    
                }
            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

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

                for (int i = 0; i < dgMineraliz.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oMiner.dFrom = double.Parse(dgMineraliz.Rows[i].Cells["From"].Value.ToString());
                    oMiner.dTo = double.Parse(dgMineraliz.Rows[i].Cells["To"].Value.ToString());
                    oMiner.sHoleID = dgMineraliz.Rows[i].Cells["HoleID"].Value.ToString();
                    oMiner.iDHMinID = long.Parse(dgMineraliz.Rows[i].Cells["SKDHMin"].Value.ToString());
                    dtValid = oMiner.getDHMinValid();

                    if (dtValid.Rows.Count > 0)
                    {
                        dtResult.ImportRow(dtValid.Rows[0]);
                    }
                }

                dgMineraliz.DataSource = dtResult;

                oMiner.sHoleID = cmbHoleIdMin.SelectedValue.ToString();
                DataTable dtMinerFromToNext = oMiner.getDHMinFromToValidFromToNext();
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

                    oSheet.Cells[4, 3] = cmbHoleIdMin.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["MZ1Mineral"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["MZ1Perc"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["MZ1Style"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
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
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;
                    
                }


            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void BoxValid()
        {
            try
            {
                
                DataTable dtResult = new DataTable();
                oBox.dFrom = 0; oBox.dTo = 0; oBox.sHoleID = "0"; oBox.iSKDHBox = 0;
                dtResult = oBox.getDHBoxValidExport();

                dgGeotech.DataSource = dtResult;

                //oGeo.sHoleID = cmbHoleIdGeo.SelectedValue.ToString();
                //DataTable dtGeoFromToNext = oGeo.getDHGeotechValidFromToNext();

                //if (dtResult.Rows.Count > 0 || dtGeoFromToNext.Rows.Count > 0)
                if (dtResult.Rows.Count > 0)
                {
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    Excel._Worksheet oSheet;
                    Excel.Range oRng;

                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidBox"].ToString(),
                        0, false, 5,
                    Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                    Type.Missing, false, false, false);


                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                    oSheet.Cells[4, 3] = cmbHoleIdGeo.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["HoleId"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["Box"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["Photo"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
                        iInicial += 1;
                    }

                    //for (int iF = 0; iF < dtGeoFromToNext.Rows.Count; iF++)
                    //{
                    //oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["HoleId"].ToString();
                    //oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["From"].ToString();
                    //oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["To"].ToString();
                    //oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["Box"].ToString();
                    //oSheet.Cells[iInicial, 6] = "From To Next Invalid";
                    //    iInicial += 1;
                    //}

                    oXL.Visible = true;
                    oXL.UserControl = true;



                }
                else 
                {
                    MessageBox.Show("No overlaps");
                }

            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        //getDHGeotechValid
        private void GeotechValid()
        {
            try
            {
                DataTable dtValid;
                DataTable dtResult = new DataTable();
                oGeo.iFrom = 0; oGeo.iTo = 0; oGeo.sHoleID = "0"; oGeo.iDHGeotechID = 0;
                dtResult = oGeo.getDHGeotechValid();

                for (int i = 0; i < dgGeotech.Rows.Count - 1; i++)
                {
                    dtValid = new DataTable();
                    oGeo.iFrom = double.Parse(dgGeotech.Rows[i].Cells["From"].Value.ToString());
                    oGeo.iTo = double.Parse(dgGeotech.Rows[i].Cells["To"].Value.ToString());
                    oGeo.sHoleID = dgGeotech.Rows[i].Cells["HoleID"].Value.ToString();
                    oGeo.iDHGeotechID = long.Parse(dgGeotech.Rows[i].Cells["SKDHGeotech"].Value.ToString());
                    dtValid = oGeo.getDHGeotechValid();

                    if (dtValid.Rows.Count > 0)
                    {
                        dtResult.ImportRow(dtValid.Rows[0]);
                    }
                }

                dgGeotech.DataSource = dtResult;

                oGeo.sHoleID = cmbHoleIdGeo.SelectedValue.ToString();
                DataTable dtGeoFromToNext = oGeo.getDHGeotechValidFromToNext();
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

                    oSheet.Cells[4, 3] = cmbHoleIdGeo.SelectedValue.ToString();

                    int iInicial = 6;
                    for (int i = 0; i < dtResult.Rows.Count; i++)
                    {

                        oSheet.Cells[iInicial, 1] = dtResult.Rows[i]["From"].ToString();
                        oSheet.Cells[iInicial, 2] = dtResult.Rows[i]["To"].ToString();
                        oSheet.Cells[iInicial, 3] = dtResult.Rows[i]["LithCod"].ToString();
                        oSheet.Cells[iInicial, 4] = dtResult.Rows[i]["Recm"].ToString();
                        oSheet.Cells[iInicial, 5] = dtResult.Rows[i]["RQDcm"].ToString();
                        oSheet.Cells[iInicial, 6] = "From To Overlaps";
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
                        iInicial += 1;
                    }

                    oXL.Visible = true;
                    oXL.UserControl = true;
                    
                    
                    
                }

            }
            catch (Exception ex)
            {

                if (ex.GetType().ToString() != "System.NullReferenceException")
                {
                    MessageBox.Show(ex.Message);
                }
                else
                { MessageBox.Show("Error ", "Weathering", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            }
        }

        private void ExportAllLogging()
        {
            try
            {


                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                oXL = new Excel.Application();
                //oXL.Visible = true;

                oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_LoggingAll"].ToString(),
                    0, false, 5,
                Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                Type.Missing, false, false, false);
                //Samples Minerals

                int iInicial = 0;

                pCargando.Visible = true;

                #region Samples

                oSheet = (Excel._Worksheet)oWB.Sheets["Samples Minerals"];
                //ExcelGenerateSamples(oSheet);

                //oSheet.Cells[1, 27] = cmbHoleIDForm.SelectedValue.ToString();
                oSheet.get_Range("AA1", "AD1").MergeCells = true;
                oSheet.get_Range("AA1", "AD1").Value2 = cmbHoleIDForm.SelectedValue.ToString();

                oXL.Visible = true;

                DataTable dtLogging = (DataTable)gdLoggin.DataSource;

                iInicial = 6;
                for (int i = 0; i < dtLogging.Rows.Count; i++)
                {

                    oSheet.Cells[iInicial, 1] = dtLogging.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 2] = dtLogging.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 3] = dtLogging.Rows[i]["Sample"].ToString();
                    oSheet.Cells[iInicial, 4] = dtLogging.Rows[i]["SampleType"].ToString();
                    oSheet.Cells[iInicial, 5] = dtLogging.Rows[i]["DupDe"].ToString();
                    oSheet.Cells[iInicial, 6] = dtLogging.Rows[i]["Lithology"].ToString();
                    oSheet.Cells[iInicial, 7] = dtLogging.Rows[i]["Comments"].ToString();
                    iInicial += 1;
                }

                #endregion

                #region weathering

                oSheet = (Excel._Worksheet)oWB.Sheets["Weathering"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateWeathering(oSheet);

                DataTable dtWeathering = (DataTable)dgWeathering.DataSource;

                iInicial = 2;
                for (int i = 0; i < dtWeathering.Rows.Count ; i++)
                {
                    oSheet.Cells[iInicial, 1] = dtWeathering.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtWeathering.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtWeathering.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 4] = dtWeathering.Rows[i]["Weathering"].ToString();
                    oSheet.Cells[iInicial, 5] = dtWeathering.Rows[i]["Oxidation"].ToString();

                    oSheet.Cells[iInicial, 6] = dtWeathering.Rows[i]["Mineral1"].ToString();
                    oSheet.Cells[iInicial, 7] = dtWeathering.Rows[i]["Mineral2"].ToString();
                    oSheet.Cells[iInicial, 8] = dtWeathering.Rows[i]["Mineral3"].ToString();
                    oSheet.Cells[iInicial, 9] = dtWeathering.Rows[i]["Mineral4"].ToString();


                    oSheet.Cells[iInicial, 10] = dtWeathering.Rows[i]["Colour1"].ToString();
                    oSheet.Cells[iInicial, 11] = dtWeathering.Rows[i]["Sufix1"].ToString();
                    oSheet.Cells[iInicial, 12] = dtWeathering.Rows[i]["Observation"].ToString();
                    iInicial += 1;
                }

                #endregion

                #region Alterations

                oSheet = (Excel._Worksheet)oWB.Sheets["Alterations"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateAlterations(oSheet);

                DataTable dtAlterations = (DataTable)dgAlterations.DataSource;

                iInicial = 3;
                for (int i = 0; i < dtAlterations.Rows.Count; i++)
                {

                    oSheet.Cells[iInicial, 1] = dtAlterations.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtAlterations.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtAlterations.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 4] = dtAlterations.Rows[i]["A1Type"].ToString();
                    oSheet.Cells[iInicial, 5] = dtAlterations.Rows[i]["A1Int"].ToString();
                    oSheet.Cells[iInicial, 6] = dtAlterations.Rows[i]["A1Style"].ToString();
                    oSheet.Cells[iInicial, 7] = dtAlterations.Rows[i]["A1Style2"].ToString()
                           == "-1" ? "" : dtAlterations.Rows[i]["A1Style2"].ToString();
                    oSheet.Cells[iInicial, 8] = dtAlterations.Rows[i]["A1Min"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A1Min"].ToString();
                    oSheet.Cells[iInicial, 9] = dtAlterations.Rows[i]["A1Min2"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A1Min2"].ToString();
                    oSheet.Cells[iInicial, 10] = dtAlterations.Rows[i]["A1Min3"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A1Min3"].ToString();
                    oSheet.Cells[iInicial, 11] = dtAlterations.Rows[i]["A2Type"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A2Type"].ToString();
                    oSheet.Cells[iInicial, 12] = dtAlterations.Rows[i]["A2Int"].ToString()
                        == "-1" ? "" : dtAlterations.Rows[i]["A2Int"].ToString();
                    oSheet.Cells[iInicial, 13] = dtAlterations.Rows[i]["A2Style"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Style"].ToString();
                    oSheet.Cells[iInicial, 14] = dtAlterations.Rows[i]["A2Style2"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Style2"].ToString();
                    oSheet.Cells[iInicial, 15] = dtAlterations.Rows[i]["A2Min"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Min"].ToString();
                    oSheet.Cells[iInicial, 16] = dtAlterations.Rows[i]["A2Min"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Min2"].ToString();
                    oSheet.Cells[iInicial, 17] = dtAlterations.Rows[i]["A2Min3"].ToString()
                         == "-1" ? "" : dtAlterations.Rows[i]["A2Min3"].ToString();

                    oSheet.Cells[iInicial, 18] = dtAlterations.Rows[i]["Comments"].ToString();

                    iInicial += 1;
                }

                #endregion

                #region Structures

                oSheet = (Excel._Worksheet)oWB.Sheets["Structures"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateStructures(oSheet);

                DataTable dtStr = (DataTable)dgStructure.DataSource;

                iInicial = 2;
                for (int i = 0; i < dtStr.Rows.Count; i++)
                {

                    oSheet.Cells[iInicial, 1] = dtStr.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtStr.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtStr.Rows[i]["To"].ToString(); ;
                    oSheet.Cells[iInicial, 4] = dtStr.Rows[i]["Type"].ToString();
                    oSheet.Cells[iInicial, 5] = dtStr.Rows[i]["AngleToAxis"].ToString();
                    oSheet.Cells[iInicial, 6] = dtStr.Rows[i]["UpAngle"].ToString();
                    oSheet.Cells[iInicial, 7] = dtStr.Rows[i]["BtonAngle"].ToString();
                    oSheet.Cells[iInicial, 8] = dtStr.Rows[i]["AppThick"].ToString();
                    oSheet.Cells[iInicial, 9] = dtStr.Rows[i]["Fill"].ToString();
                    oSheet.Cells[iInicial, 10] = dtStr.Rows[i]["Fill2"].ToString();
                    oSheet.Cells[iInicial, 11] = dtStr.Rows[i]["Fill3"].ToString();
                    oSheet.Cells[iInicial, 12] = dtStr.Rows[i]["Fill4"].ToString();
                    oSheet.Cells[iInicial, 13] = dtStr.Rows[i]["Number"].ToString();
                    oSheet.Cells[iInicial, 14] = dtStr.Rows[i]["Comments"].ToString();
                    oSheet.Cells[iInicial, 15] = dtStr.Rows[i]["Lenght"].ToString();

                    iInicial += 1;
                }

                #endregion

                #region Mineralizations

                oSheet = (Excel._Worksheet)oWB.Sheets["Mineraliz"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateMineralizations(oSheet);

                DataTable dtMiner = (DataTable)dgMineraliz.DataSource;

                iInicial = 3;
                for (int i = 0; i < dtMiner.Rows.Count; i++)
                {

                    /*Gsize, GSize2, GSize3*/
                    oSheet.Cells[iInicial, 1] = dtMiner.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtMiner.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtMiner.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 4] = dtMiner.Rows[i]["MZ1Mineral"].ToString();
                    oSheet.Cells[iInicial, 5] = dtMiner.Rows[i]["MZ1Mineral2"].ToString();
                    oSheet.Cells[iInicial, 6] = dtMiner.Rows[i]["MZ1Mineral3"].ToString();
                    oSheet.Cells[iInicial, 7] = dtMiner.Rows[i]["MZ1Style"].ToString();
                    oSheet.Cells[iInicial, 8] = dtMiner.Rows[i]["MZ1Perc"].ToString();

                    oSheet.Cells[iInicial, 9] = dtMiner.Rows[i]["Gsize"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["Gsize"].ToString();

                    //sEditGeo == "1" ? "Update" : "Insert", clsRf.sUser.ToString(),
                    oSheet.Cells[iInicial, 10] = dtMiner.Rows[i]["MZ2Mineral"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ2Mineral"].ToString();
                    oSheet.Cells[iInicial, 11] = dtMiner.Rows[i]["MZ2Mineral2"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ2Mineral2"].ToString();
                    oSheet.Cells[iInicial, 12] = dtMiner.Rows[i]["MZ2Mineral3"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ2Mineral3"].ToString();
                    oSheet.Cells[iInicial, 13] = dtMiner.Rows[i]["MZ2Style"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ2Style"].ToString();
                    oSheet.Cells[iInicial, 14] = dtMiner.Rows[i]["MZ2Perc"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ2Perc"].ToString();

                    oSheet.Cells[iInicial, 15] = dtMiner.Rows[i]["GSize2"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["GSize2"].ToString();

                    oSheet.Cells[iInicial, 16] = dtMiner.Rows[i]["MZ3Mineral"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ3Mineral"].ToString();
                    oSheet.Cells[iInicial, 17] = dtMiner.Rows[i]["MZ3Mineral2"].ToString()
                        == "-1" ? "" : dtMiner.Rows[i]["MZ3Mineral2"].ToString();
                    oSheet.Cells[iInicial, 18] = dtMiner.Rows[i]["MZ3Mineral3"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ3Mineral3"].ToString();
                    oSheet.Cells[iInicial, 19] = dtMiner.Rows[i]["MZ3Style"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ3Style"].ToString();
                    oSheet.Cells[iInicial, 20] = dtMiner.Rows[i]["MZ3Perc"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["MZ3Perc"].ToString();

                    oSheet.Cells[iInicial, 21] = dtMiner.Rows[i]["GSize3"].ToString()
                         == "-1" ? "" : dtMiner.Rows[i]["GSize3"].ToString();

                    oSheet.Cells[iInicial, 22] = dtMiner.Rows[i]["Comments"].ToString();

                    iInicial += 1;
                }

                #endregion

                #region Lithology

                oSheet = (Excel._Worksheet)oWB.Sheets["Lithology"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateLithology(oSheet);

                DataTable dtLitho = (DataTable)dgLithology.DataSource;

                iInicial = 2;
                for (int i = 0; i < dtLitho.Rows.Count; i++)
                {
                    oSheet.Cells[iInicial, 1] = dtLitho.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtLitho.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtLitho.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 4] = dtLitho.Rows[i]["Litho"].ToString();

                    oSheet.Cells[iInicial, 5] = dtLitho.Rows[i]["Textures"].ToString();
                    oSheet.Cells[iInicial, 6] = dtLitho.Rows[i]["GSize"].ToString();

                    oSheet.Cells[iInicial, 7] = dtLitho.Rows[i]["Observation"].ToString();
                    iInicial += 1;
                }

                #endregion

                #region Geotech

                oSheet = (Excel._Worksheet)oWB.Sheets["Geotech"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateGeotech(oSheet);

                DataTable dtGeo = (DataTable)dgGeotech.DataSource;

                iInicial = 2;
                for (int i = 0; i < dtGeo.Rows.Count; i++)
                {
                    oSheet.Cells[iInicial, 1] = dtGeo.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtGeo.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtGeo.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 4] = dtGeo.Rows[i]["LithCod"].ToString();
                    oSheet.Cells[iInicial, 5] = dtGeo.Rows[i]["Recm"].ToString();

                    oSheet.Cells[iInicial, 6] = dtGeo.Rows[i]["RQDcm"].ToString();
                    oSheet.Cells[iInicial, 7] = dtGeo.Rows[i]["NoOfFract"].ToString();
                    oSheet.Cells[iInicial, 8] = dtGeo.Rows[i]["JointCond"].ToString();
                    oSheet.Cells[iInicial, 9] = dtGeo.Rows[i]["Jn"].ToString();
                    oSheet.Cells[iInicial, 10] = dtGeo.Rows[i]["Jr"].ToString();
                    oSheet.Cells[iInicial, 11] = dtGeo.Rows[i]["Ja"].ToString();
                    oSheet.Cells[iInicial, 12] = dtGeo.Rows[i]["DegBreak"].ToString();
                    oSheet.Cells[iInicial, 13] = dtGeo.Rows[i]["Hardness"].ToString();
                    oSheet.Cells[iInicial, 14] = dtGeo.Rows[i]["Comments"].ToString();
                    oSheet.Cells[iInicial, 15] = dtGeo.Rows[i]["AltWeath"].ToString();


                    iInicial += 1;
                }

                #endregion

                #region Box

                oSheet = (Excel._Worksheet)oWB.Sheets["Box"];
                //ExcelGenerateBox(oSheet);

                DataTable dtBox = (DataTable)dgBox.DataSource;

                iInicial = 5;
                for (int i = 0; i < dtBox.Rows.Count; i++)
                {
                    oSheet.Cells[iInicial, 1] = dtBox.Rows[i]["HoleID"].ToString();
                    oSheet.Cells[iInicial, 2] = dtBox.Rows[i]["From"].ToString();
                    oSheet.Cells[iInicial, 3] = dtBox.Rows[i]["To"].ToString();
                    oSheet.Cells[iInicial, 4] = dtBox.Rows[i]["Box"].ToString();
                    oSheet.Cells[iInicial, 5] = dtBox.Rows[i]["Stand"].ToString();
                    oSheet.Cells[iInicial, 6] = dtBox.Rows[i]["column"].ToString();
                    oSheet.Cells[iInicial, 7] = dtBox.Rows[i]["row"].ToString();

                    iInicial += 1;
                }

                #endregion

                #region cover logging 

                oSheet = (Excel._Worksheet)oWB.Sheets["Cover Logging"];//(Excel._Worksheet)oWB.ActiveSheet;
                //ExcelGenerateGeotech(oSheet);
                DataTable dtData = oRf.getCollarsPlatf(cmbHoleIDForm.SelectedValue.ToString());

                /*C.PlatformId, C.HoleID ,P.EastPlanned, P.NorthPlanned, P.ElevationPlanned, P.Location, P.StartDate, 
                 * P.FinalDate,
		            P.EastST, P.NorthST, P.ElevationST, P.Zone, 'Datum', C.Azimuth, C.Dip, 'Lenght',
		            P.Contractor, P.RigUsed, LoggedBy, LoggedBy1, LoggedBy2, LoggedBy3, ReLoggedBy ,  'geotech by', 
                 * 'Porpuse'
	            */

                if (dtData != null)
                {
                    if (dtData.Rows.Count > 0)
                    {
                        oSheet.Cells[10, 6] = dtData.Rows[0]["EastPlanned"].ToString();
                        oSheet.Cells[11, 6] = dtData.Rows[0]["NorthPlanned"].ToString();
                        oSheet.Cells[12, 6] = dtData.Rows[0]["ElevationPlanned"].ToString();
                        oSheet.Cells[13, 6] = dtData.Rows[0]["EastST"].ToString();
                        oSheet.Cells[14, 6] = dtData.Rows[0]["NorthST"].ToString();
                        oSheet.Cells[15, 6] = dtData.Rows[0]["ElevationST"].ToString();
                        oSheet.Cells[16, 6] = dtData.Rows[0]["Zone"].ToString();
                        oSheet.Cells[18, 6] = dtData.Rows[0]["Azimuth"].ToString();
                        oSheet.Cells[19, 6] = dtData.Rows[0]["Dip"].ToString();
                        oSheet.Cells[8, 10] = dtData.Rows[0]["Location"].ToString();
                        oSheet.Cells[10, 10] = dtData.Rows[0]["StartDate"].ToString();
                        oSheet.Cells[11, 10] = dtData.Rows[0]["FinalDate"].ToString();
                        oSheet.Cells[15, 10] = dtData.Rows[0]["Contractor"].ToString();
                        oSheet.Cells[16, 10] = dtData.Rows[0]["RigUsed"].ToString();
                    }
                }
                
                #endregion

                //oXL.Visible = true;
                oXL.UserControl = true;

                //oXL.Quit();

                pCargando.Visible = false;

                MessageBox.Show("Successful process");

            }
            catch (Exception exExportAll)
            {
                MessageBox.Show("Error Export Excel: " + exExportAll.Message);
            }
        }

        

        private void btnExporExcelAll_Click(object sender, EventArgs e)
        {
            try
            {
                //Implementar hilos en background

                //Thread tExport = new Thread(new ThreadStart(ExportAllLogging));
                //tExport.IsBackground = true;
                //tExport.Start();
                //tExport.Join();

                sValidLogging = "ExportAll";
                bgw.RunWorkerAsync();

                //switch (sValidLogging)
                //{
                //    case "ExportAll":
                //        ExportAllLogging();
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
                sValidLogging = "samples"; //Ejecuta los eventos bgw_DoWork, bgw_ProgressChanged y bgw_RunWorkerCompleted
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
            //SamplesValid();pbLogging

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
            try
            {
                ValidLogging();
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }  

        }

        private void ValidLogging()
        {
            try
            {
                switch (sValidLogging)
                {
                    case "ExportAll":
                        ExportAllLogging();
                        break;
                    case "samples":
                        SamplesValid();
                        break;
                    case "weathering":
                        WeathValid();
                        break;
                    case "lithology":
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
                    case "box":
                        BoxValid();
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

        private void btnValidWeath_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "weathering";
                bgw.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message);
            }
        }

        private void btnValidLith_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "lithology";
                bgw.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnValidAlt_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "Alterations";
                bgw.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnValidStr_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "Structures";
                bgw.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnValidMin_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "Mineralizations";
                bgw.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btnValidGeo_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "Geotech";
                bgw.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TabPpal_KeyDown(object sender, KeyEventArgs e)
        {

            /*if (_e.KeyChar == (char)(Keys.Enter))
            {
                _e.Handled = true;
                SendKeys.Send("{TAB}");
            }*/
            if (e.KeyValue == (char)Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void txtBoxDens_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtFromDens_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtToDens_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtLenghtDens_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtDiameterDens_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        #region Density

        private string ControlsValidateDensity()
        {
            try
            {
                string sresp = "";

                oCollars.sHoleID = cmbHoleIDSt.SelectedValue.ToString();
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow[] dato = dtCollars.Select("Length < '" + txtFromDens.Text + "'");
                if (dato.Length > 0)
                {
                    sresp = " 'Depth' greater than Hole Id lenght";
                    return sresp;
                }

                if (double.Parse(txtFromDens.Text.ToString()) == double.Parse(txtToDens.Text.ToString()))
                {
                    sresp = " 'From' equal to 'To'";
                    return sresp;
                }

                if (double.Parse(txtFromDens.Text.ToString()) > double.Parse(txtToDens.Text.ToString()))
                {
                    sresp = " 'From' greater than 'To'";
                    return sresp;
                }

                if (txtSampleNoDens.ToString() == "" )
                {
                    sresp = " Empty Sample";
                    return sresp;
                }


                return sresp;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnAddDens_Click(object sender, EventArgs e)
        {
            try
            {
                string sResp = ControlsValidateDensity().ToString();
                if (sResp.ToString() != "")
                {
                    MessageBox.Show(sResp.ToString(), "Density", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (sEditDens == "0")
                {
                    oDens.iSKDHDensity = 0;
                    oDens.sOpcion = "1";
                }
                else if (sEditDens == "1")
                {
                    oDens.sOpcion = "2";
                }
                
                oDens.sHoleID = cmbHoleIdDens.SelectedValue.ToString();
                oDens.sBox = txtBoxDens.Text.ToString();
                oDens.dFrom = double.Parse(txtFromDens.Text.ToString());
                oDens.dTo = double.Parse(txtToDens.Text.ToString());
                oDens.dLenght = double.Parse(txtLenghtDens.Text.ToString());
                oDens.dDiameter = double.Parse(txtDiameterDens.Text.ToString());
                oDens.sSample = txtSampleNoDens.Text.ToString();
                oDens.sLith = cmbLithoDens.SelectedValue.ToString();
                oDens.sComments = txtCommentsDens.Text.ToString();
                oDens.sVeinName = cmbVeinNameDens.SelectedValue.ToString();
                oDens.sTexture = cmbTextureDens.SelectedValue.ToString();
                oDens.sStructure = cmbStructDens.SelectedValue.ToString();
                oDens.sMineral1 = cmbMineral1Dens.SelectedValue.ToString();
                oDens.sMineral2 = cmbMineral2Dens.SelectedValue.ToString();
                oDens.sSulfphides = txtSulphideDens.Text.ToString();
                oDens.sAltType = cmbAltTypeDens.SelectedValue.ToString();
                oDens.sAltInt = cmbAltIntensityDens.SelectedValue.ToString();

                clsDHDensity.sStaticFrom = txtToDens.Text.ToString();

                string sRespAdd = oDens.DH_Dens_Add();

                if (int.Parse(sRespAdd.ToString()) > 0)
                {

                    txtFromDens.Text = clsDHDensity.sStaticFrom;
                    FilldgDensity("2");

                    //Implementar visualizar la ultima modificacion en pantalla
                    if (sEditDens == "1")
                    {
                        if (dgDensity.Rows.Count > 1)
                        {
                            DataTable dtDens = (DataTable)dgDensity.DataSource;
                            DataRow[] myRow = dtDens.Select(@"SKDHDensity = '" + oDens.iSKDHDensity + "'");
                            int rowindex = dtDens.Rows.IndexOf(myRow[0]);
                            dgDensity.Rows[rowindex].Selected = true;
                            dgDensity.CurrentCell = dgDensity.Rows[rowindex].Cells[1] ;
                        }
                    }

                    CleanControlsDens();
                    sEditDens = "0";


                }
                else
                {
                    MessageBox.Show("Insert error:" + sRespAdd.ToString());
                }           


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Density", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbLithoDens_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                oRf.sCodeLith = cmbLithoDens.SelectedValue.ToString();

                DataTable dtTextures = new DataTable();
                dtTextures = oRf.getRfTextures_List();
                DataRow drTx = dtTextures.NewRow();
                drTx[0] = "-1";
                drTx[1] = "Select an option..";
                dtTextures.Rows.Add(drTx);
                cmbTextureDens.DisplayMember = "Comb";
                cmbTextureDens.ValueMember = "Code";
                cmbTextureDens.DataSource = dtTextures;
                cmbTextureDens.SelectedValue = "-1";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FilldgDensity(string _sOpc)
        {
            try
            {
                oDens.sOpcion = _sOpc.ToString();
                oDens.sHoleID = cmbHoleIdDens.SelectedValue.ToString();
                DataTable dtDens = oDens.getDHDensity();
                dgDensity.DataSource = dtDens;

                dgDensity.Columns["SKDHDensity"].Visible = false;


            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        private void cmbHoleIdDens_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                FilldgDensity("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgDensity_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Row Delete. " + "HoleId " + dgDensity.Rows[e.RowIndex].Cells["HoleID"].Value.ToString()
                    + " From " + dgDensity.Rows[e.RowIndex].Cells["From"].Value.ToString()
                    + " To " + dgDensity.Rows[e.RowIndex].Cells["To"].Value.ToString()
                    + " Lenght " + dgDensity.Rows[e.RowIndex].Cells["Lenght"].Value.ToString()
                    , "Density", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {

                    oDens.iSKDHDensity = int.Parse(dgDensity.Rows[e.RowIndex].Cells["SKDHDensity"].Value.ToString());
                    string sRespDel = oDens.DH_Dens_Delete();
                    if (sRespDel.ToString() == "OK")
                    {
                        FilldgDensity("2");
                        sEditDens = "0";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgDensity_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                oDens.iSKDHDensity = int.Parse(dgDensity.Rows[e.RowIndex].Cells["SKDHDensity"].Value.ToString());
                cmbHoleIdDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["HoleID"].Value.ToString();
                txtBoxDens.Text = dgDensity.Rows[e.RowIndex].Cells["Box"].Value.ToString();
                txtFromDens.Text = dgDensity.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToDens.Text = dgDensity.Rows[e.RowIndex].Cells["To"].Value.ToString();
                txtLenghtDens.Text = dgDensity.Rows[e.RowIndex].Cells["Lenght"].Value.ToString();
                txtDiameterDens.Text = dgDensity.Rows[e.RowIndex].Cells["Diameter"].Value.ToString();
                txtSampleNoDens.Text = dgDensity.Rows[e.RowIndex].Cells["Sample"].Value.ToString();
                cmbLithoDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["Litho"].Value.ToString();
                txtCommentsDens.Text = dgDensity.Rows[e.RowIndex].Cells["Comments"].Value.ToString();
                cmbVeinNameDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["VeinName"].Value.ToString();
                cmbTextureDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["Texture"].Value.ToString();
                cmbStructDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["Structure"].Value.ToString();
                cmbMineral1Dens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["Mineralization_1"].Value.ToString();
                cmbMineral2Dens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["Mineralization_2"].Value.ToString();
                txtSulphideDens.Text = dgDensity.Rows[e.RowIndex].Cells["Sulfphides_per"].Value.ToString();
                cmbAltTypeDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["AltType"].Value.ToString();
                cmbAltIntensityDens.SelectedValue = dgDensity.Rows[e.RowIndex].Cells["AltInt"].Value.ToString();
                sEditDens = "1";

                FilldgDensityMethod("2");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FilldgDensityMethod(string _sOpcion)
        {
            try
            {
                oDens.sOpcionM = _sOpcion;
                DataTable dtDensMet = oDens.getDHDensityMethod();
                dgDensityM.DataSource = dtDensMet;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CleanControlsDens()
        {
            try
            {
                sEditDens = "0";
                oDens.iSKDHDensity = 0;
                //cmbHoleIdDens.SelectedValue = "Select an option..";
                txtBoxDens.Text = "";
                //txtFromDens.Text = dgDensity.Rows[e.RowIndex].Cells["From"].Value.ToString();
                txtToDens.Text = "";
                txtLenghtDens.Text = "";
                txtSampleNoDens.Text = "";
                cmbLithoDens.SelectedValue = "-1";
                txtCommentsDens.Text = "";
                cmbVeinNameDens.SelectedValue = "Select an option...";
                cmbTextureDens.SelectedValue = "-1";
                cmbStructDens.SelectedValue = "-1";
                cmbMineral1Dens.SelectedValue = "-1";
                cmbMineral2Dens.SelectedValue = "-1";
                txtSulphideDens.Text = "";
                cmbAltTypeDens.SelectedValue = "-1";
                cmbAltIntensityDens.SelectedValue = "-1";


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnCancelDens_Click(object sender, EventArgs e)
        {
            try
            {
                CleanControlsDens();
                FilldgDensityMethod("2");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        

        private void btnAddDensM_Click(object sender, EventArgs e)
        {
            try
            {
                if (oDens.iSKDHDensity == null)
                {
                    MessageBox.Show("Select a row Density");
                    return;
                }
                if (oDens.iSKDHDensity <= 0)
                {
                    MessageBox.Show("Select a row Density");
                    return;
                }

                if (sEditDensM == "0")
                {
                    oDens.iSKDHDensityMethod = 0;
                    oDens.sOpcionM = "1";
                }
                else
                {
                    oDens.sOpcionM = "2";
                }
                
                oDens.sLab = cmbLabDensM.SelectedValue.ToString();
                oDens.dDrySamp = double.Parse(txtDrySampDensM.Text.ToString());
                oDens.dImmersedSamp = double.Parse(txtInmersedDensM.Text.ToString());
                oDens.dDensity = double.Parse(txtDensityDensM.Text.ToString());
                oDens.sMethod = txtMethodDensM.Text.ToString();
                oDens.iPriority = 1;
                string sResp = oDens.DH_DensMethod_Add();
                if (sResp == "OK")
                {
                    //MessageBox.Show("Si");
                    FilldgDensityMethod("2");

                    //Implementar visualizar la ultima modificacion en pantalla
                    if (sEditDensM == "1")
                    {
                        if (dgDensityM.Rows.Count > 1)
                        {
                            DataTable dtDens = (DataTable)dgDensityM.DataSource;
                            DataRow[] myRow = dtDens.Select(@"SKDHDensityMethod = '" + oDens.iSKDHDensityMethod + "'");
                            int rowindex = dtDens.Rows.IndexOf(myRow[0]);
                            dgDensityM.Rows[rowindex].Selected = true;
                            dgDensityM.CurrentCell = dgDensityM.Rows[rowindex].Cells[1];
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Error Insert: " + sResp.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cleanControlsDensMet()
        {
            try
            {

                cmbLabDensM.SelectedValue = ConfigurationSettings.AppSettings["IDProjectGC"].ToString();
                txtDrySampDensM.Text = "";
                txtInmersedDensM.Text = "";
                txtDensityDensM.Text = "";
                txtMethodDensM.Text = "";
                sEditDensM = "0";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void btnCancelDensM_Click(object sender, EventArgs e)
        {
            try
            {
                oDens.iSKDHDensity = 0;
                FilldgDensityMethod("2");
                cleanControlsDensMet();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgDensityM_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                sEditDensM = "1";
                oDens.iSKDHDensityMethod = int.Parse(dgDensityM.Rows[e.RowIndex].Cells["SKDHDensityMethod"].Value.ToString());
                cmbLabDensM.SelectedValue = dgDensityM.Rows[e.RowIndex].Cells["Lab"].Value.ToString();
                txtDrySampDensM.Text = dgDensityM.Rows[e.RowIndex].Cells["DrySamp_g"].Value.ToString();
                txtInmersedDensM.Text = dgDensityM.Rows[e.RowIndex].Cells["ImmersedSamp_g"].Value.ToString();
                txtDensityDensM.Text = dgDensityM.Rows[e.RowIndex].Cells["Density"].Value.ToString();
                txtMethodDensM.Text = dgDensityM.Rows[e.RowIndex].Cells["Method"].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgDensityM_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (MessageBox.Show("Row Delete. " + "Lab" + dgDensityM.Rows[e.RowIndex].Cells["Lab"].Value.ToString()
                    + " DrySamp_g " + dgDensityM.Rows[e.RowIndex].Cells["DrySamp_g"].Value.ToString()
                    + " ImmersedSamp_g " + dgDensityM.Rows[e.RowIndex].Cells["ImmersedSamp_g"].Value.ToString()
                    + " Density " + dgDensityM.Rows[e.RowIndex].Cells["Density"].Value.ToString()
                    , "Density", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    oDens.iSKDHDensityMethod = int.Parse(dgDensityM.Rows[e.RowIndex].Cells["SKDHDensityMethod"].Value.ToString());
                    string sRespDel = oDens.DH_DensMethod_Delete();
                    if (sRespDel.ToString() == "OK")
                    {
                        FilldgDensityMethod("2");
                        sEditDensM = "0";
                    }
                }


                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion         

        private void btnValidBox_Click(object sender, EventArgs e)
        {
            try
            {
                sValidLogging = "box"; //Ejecuta los eventos bgw_DoWork, bgw_ProgressChanged y bgw_RunWorkerCompleted
                bgw.RunWorkerAsync();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExcelBox_Click(object sender, EventArgs e)
        {
            try
            {
                
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                oXL = new Excel.Application();
                oXL.Visible = true;
                //Get a new workbook.
                //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                //oWB = oXL.Workbooks.Open(@"D:/Template_Shipment_Sgs.xls", 0, true, 5,


                oWB = oXL.Workbooks.Open(ConfigurationSettings.AppSettings["Ruta_ValidBox"].ToString(),
                    0, false, 5,
                Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, true, false,
                Type.Missing, false, false, false);

                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //oSheet.Cells[1, 6] = cmbHoleIDBox.SelectedValue.ToString();

                int iInicial = 6;
                for (int i = 0; i < dgBox.Rows.Count - 1; i++)
                {

                    oSheet.Cells[iInicial, 1] = dgBox.Rows[i].Cells["HoleId"].Value.ToString();
                    oSheet.Cells[iInicial, 2] = dgBox.Rows[i].Cells["From"].Value.ToString();
                    oSheet.Cells[iInicial, 3] = dgBox.Rows[i].Cells["To"].Value.ToString();
                    oSheet.Cells[iInicial, 4] = dgBox.Rows[i].Cells["Box"].Value.ToString();
                    oSheet.Cells[iInicial, 5] = dgBox.Rows[i].Cells["Photo"].Value.ToString();

                    iInicial += 1;
                }

                oXL.Visible = true;
                oXL.UserControl = true;


            }
            catch (Exception ex)
            {

                 MessageBox.Show("Error: " + ex.Message, "Box", MessageBoxButtons.OK, MessageBoxIcon.Error); 

            }
        
        }

        private void txtPhotoBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtEditPhotoBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void cmbTypeAlt_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //DataTable dtMinAlt = new DataTable();
                //dtMinAlt = oRf.getRfMinerAlt_ListMin(cmbTypeAlt.SelectedValue.ToString());
                //DataRow drMinA = dtMinAlt.NewRow();
                //drMinA[0] = "-1";
                //drMinA[1] = "Select an option..";
                //dtMinAlt.Rows.Add(drMinA);

                //CargarCombosAlt(dtMinAlt, cmbMin1Alt);
                //CargarCombosAlt(dtMinAlt, cmbMin2Alt1);
                //CargarCombosAlt(dtMinAlt, cmbMin3Alt1);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        

        private void cmbTypeAlt2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //DataTable dtMinAlt = new DataTable();
                //dtMinAlt = oRf.getRfMinerAlt_ListMin(cmbTypeAlt2.SelectedValue.ToString());
                //DataRow drMinA = dtMinAlt.NewRow();
                //drMinA[0] = "-1";
                //drMinA[1] = "Select an option..";
                //dtMinAlt.Rows.Add(drMinA);

                //CargarCombosAlt(dtMinAlt, cmbMin1Alt2);
                //CargarCombosAlt(dtMinAlt, cmbMin2Alt2);
                //CargarCombosAlt(dtMinAlt, cmbMin3Alt2);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtMinPerc1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtMinPerc2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtMinPerc3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Keypress(e);
        }

        private void txtMinPerc1_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtMinPerc1.Text != "")
                {
                    if (double.Parse(txtMinPerc1.Text) > 100)
                    {
                        MessageBox.Show("Percentage isn´t more than 100");
                        txtMinPerc1.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtMinPerc2_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtMinPerc2.Text != "")
                {
                    if (double.Parse(txtMinPerc2.Text) > 100)
                    {
                        MessageBox.Show("Percentage isn´t more than 100");
                        txtMinPerc2.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtMinPerc3_Leave(object sender, EventArgs e)
        {
            try
            {
                if (txtMinPerc3.Text != "")
                {
                    if (double.Parse(txtMinPerc3.Text) > 100)
                    {
                        MessageBox.Show("Percentage isn´t more than 100");
                        txtMinPerc3.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void ValidFTT_Structures()
        //{
        //    try
        //    {
        //        switch (cmbStructureTypeSt.SelectedValue.ToString())
        //        {
        //            case "VEN":
        //                break;
        //            case "VNA":
        //                break;
        //            case "VNS":
        //                break;
        //            default:
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message);
        //    }
        //}
    

        
     

    }
}
