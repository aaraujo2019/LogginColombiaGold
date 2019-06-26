using System;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace LogginColombiaGold
{
    public partial class frmPpal : Form
    {
        clsRf oRf = new clsRf();
        DataTable dtFormsAllowed = new DataTable();

        public frmPpal()
        {
            InitializeComponent();
        }

        private void logginToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //clsRf.sIdGrupo = "3";
            //dtFormsAllowed = oRf.getFormsByGrupo(clsRf.sIdGrupo, ConfigurationSettings.AppSettings["IDProject"].ToString());
            //clsRf.dsPermisos = oRf.getFormsByGrupoAll(clsRf.sIdGrupo);
            try
            {
                //***** Este fragmento es utilizado para realización de modificaciones y pruebas. AAA ****

                //frmLoggin oLog = new frmLoggin();
                //oLog.MdiParent = this;
                //oLog.Show();

                //****************************************************************
                // Para pruebas se debe comentar el siguiente fragmento de codigo.

                DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'frmLoggin'");
                if (dato.Length > 0)
                {
                    frmLoggin oLog = new frmLoggin();
                    oLog.MdiParent = this;
                    oLog.Show();
                }
                else
                {
                    MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void frmPpal_Load(object sender, EventArgs e)
        {
            try
            {
                mnuValidate.Visible = true;
                mnuValidate.Enabled = true;

                dtFormsAllowed = oRf.getFormsByGrupo(clsRf.sIdGrupo, ConfigurationSettings.AppSettings["IDProject"].ToString());
                clsRf.dsPermisos = oRf.getFormsByGrupoAll(clsRf.sIdGrupo);

                MdiClient ctlMDI = default(MdiClient);
                foreach (Control ctl in this.Controls)
                {
                    try
                    {
                        ctlMDI = (MdiClient)ctl;
                        ctlMDI.BackColor = Color.White;
                    }
                    catch (InvalidCastException ex)
                    {
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void logOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void frmPpal_FormClosed(object sender, FormClosedEventArgs e)
        {
            clsRf.dsPermisos = new DataSet();
            Application.Exit();
        }

        private void collarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'frmCollarAsign'");
                if (dato.Length > 0)
                {
                    frmCollarAsign oAsign = new frmCollarAsign();
                    oAsign.MdiParent = this;
                    oAsign.Show();
                }
                else
                {
                    MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void passwordChangeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                frmChangeLoggin oCh = new frmChangeLoggin();
                oCh.MdiParent = this;
                oCh.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void reportTransactionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DataRow[] dato = dtFormsAllowed.Select("nombre_Real_Form = 'frmReportTrans'");
                if (dato.Length > 0)
                {
                    frmReportTrans oRpt = new frmReportTrans();
                    oRpt.MdiParent = this;
                    oRpt.Show();
                }
                else
                {
                    MessageBox.Show("Form is not allowed", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void validationLoggingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmValidation oVal = new frmValidation();
            oVal.MdiParent = this;
            oVal.Show();
        }
    }
}
