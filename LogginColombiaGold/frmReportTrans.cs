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
    public partial class frmReportTrans : Form
    {
        clsRf oRf = new clsRf();

        public frmReportTrans()
        {
            InitializeComponent();
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

        private void frmReportTrans_Load(object sender, EventArgs e)
        {
            try
            {
                DataTable dtAs1 = LoadAssign("");//oRf.getUsuarios("");
                DataRow drA1 = dtAs1.NewRow();
                drA1[1] = "-1";
                drA1[7] = "Select an option..";
                dtAs1.Rows.Add(drA1);
                cmbsearch.DisplayMember = "cmb";
                cmbsearch.ValueMember = "login";
                cmbsearch.DataSource = dtAs1;
                cmbsearch.SelectedValue = "-1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtTrans = new DataTable();
                string sSearch;
                if (cmbsearch.SelectedValue.ToString() != "-1")
                {
                    sSearch = cmbsearch.SelectedValue.ToString();
                }
                else { sSearch = ""; }
                dtTrans = oRf.getTransList(sSearch.ToString());
                dgUsers.DataSource = dtTrans;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
