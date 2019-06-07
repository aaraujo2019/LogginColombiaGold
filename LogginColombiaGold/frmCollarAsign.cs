using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LogginColombiaGold
{
    public partial class frmCollarAsign : Form
    {
        
        
        bool bInicio = false;

        clsDHCollars oCollars = new clsDHCollars();
        clsRf oRf = new clsRf();

        public frmCollarAsign()
        {
            InitializeComponent();
            bInicio = true;
            FillCmb();
            bInicio = false;
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private DataTable LoadAssign(string _sAsign)
        {
            try
            {
                DataTable dtAssign = new DataTable();
                dtAssign = oRf.getUsuarios(_sAsign);
                return dtAssign;
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


                DataTable dtAs1 = LoadAssign("");//oRf.getUsuarios("");
                DataRow drA1 = dtAs1.NewRow();
                drA1[1] = "-1";
                drA1[2] = "Select an option..";
                dtAs1.Rows.Add(drA1);
                cmbAss1.DisplayMember = "cmb";
                cmbAss1.ValueMember = "login";
                cmbAss1.DataSource = dtAs1;
                cmbAss1.SelectedValue = "-1";

                DataTable dtAs2 = LoadAssign("");//oRf.getUsuarios("");
                DataRow drA2 = dtAs2.NewRow();
                drA2[1] = "-1";
                drA2[2] = "Select an option..";
                dtAs2.Rows.Add(drA2);
                cmbAss2.DisplayMember = "cmb";
                cmbAss2.ValueMember = "login";
                cmbAss2.DataSource = dtAs2;
                cmbAss2.SelectedValue = "-1";

                DataTable dtAs3 = LoadAssign("");//oRf.getUsuarios("");
                DataRow drA3 = dtAs3.NewRow();
                drA3[1] = "-1";
                drA3[2] = "Select an option..";
                dtAs3.Rows.Add(drA3);
                cmbAss3.DisplayMember = "cmb";
                cmbAss3.ValueMember = "login";
                cmbAss3.DataSource = dtAs3;
                cmbAss3.SelectedValue = "-1";

                //cmbAss1.DataSource = oRf.getUsuarios("");
                //cmbAss2.DataSource = oRf.getUsuarios("");
                //cmbAss3.DataSource = oRf.getUsuarios("");



                oCollars.sHoleID = "";
                DataTable dtCollars = oCollars.getDHCollars();
                DataRow drC = dtCollars.NewRow();
                drC[0] = "Select an option..";
                dtCollars.Rows.Add(drC);
                cmbHoleID.DisplayMember = "HoleID";
                cmbHoleID.ValueMember = "HoleID";
                cmbHoleID.DataSource = dtCollars;
                cmbHoleID.SelectedValue = "Select an option..";



            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void LoadAssign()
        {
            try
            {
                oCollars.sHoleID = cmbHoleID.SelectedValue.ToString();
                DataTable dtAssign = oCollars.getDHCollarsListAssign();
                gdAssign.DataSource = dtAssign;
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
                //Valido para que no me muestre mensajes de error al inicio del form
                if (bInicio == true)
                {
                    return;
                }
                LoadAssign();

                //oCollars.sHoleID = cmbHoleID.SelectedValue.ToString();
                //DataTable dtAssign = oCollars.getDHCollarsListAssign();
                //gdAssign.DataSource = dtAssign;
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
                //Validar que haya seleccionado el holeid y como minimo un usuario asignado
                if (cmbHoleID.SelectedValue.ToString() == "-1")
                {
                    MessageBox.Show("Select an option HoleId");
                    return;
                }

                if (cmbAss1.SelectedValue.ToString() == "-1" &&
                    cmbAss2.SelectedValue.ToString() == "-1" &&
                    cmbAss3.SelectedValue.ToString() == "-1")
                {
                    MessageBox.Show("Select an option Asign");
                    return;
                }

                oCollars.sHoleID = cmbHoleID.SelectedValue.ToString();
                string sLog1 = cmbAss1.SelectedValue.ToString();
                oCollars.sLoggedBy1 = (cmbAss1.SelectedValue.ToString() != "-1") ?
                    cmbAss1.SelectedValue.ToString() : ""; //cmbAss1.SelectedValue.ToString();
                oCollars.sLoggedBy2 = (cmbAss2.SelectedValue.ToString() != "-1") ?
                    cmbAss2.SelectedValue.ToString() : ""; //cmbAss2.SelectedValue.ToString();
                oCollars.sLoggedBy3 = (cmbAss3.SelectedValue.ToString() != "-1") ?
                    cmbAss3.SelectedValue.ToString() : ""; //cmbAss3.SelectedValue.ToString();
                string sResp = oCollars.DHSamples_UpdateAssign();
                if (sResp.ToString() == "OK")
                {
                    MessageBox.Show("Assign Successful ", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadAssign();
                }
                else
                {
                    MessageBox.Show("Assign Error", "Logging", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void frmCollarAsign_Load(object sender, EventArgs e)
        {

        }
    }
}
