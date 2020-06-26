using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TB1300
{
    public partial class Form1 : Form
    {
        SAPbobsCOM.Company oCompany;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                oCompany = new SAPbobsCOM.Company();
				oCompany.Server = "ServerName";
				oCompany.SLDServer = "ServerName:40000";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
				oCompany.CompanyDB = "SBODemoSK";
				oCompany.UserName = "manager";
				oCompany.Password = "manager";
				oCompany.DbUserName = "sa";
				oCompany.DbPassword = "1234";


                int ret = oCompany.Connect();
                if (ret == 0)
                    MessageBox.Show("Connection ok");
                else
                    MessageBox.Show("Connection failed: " + oCompany.GetLastErrorCode().ToString() + " - " + oCompany.GetLastErrorDescription());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection error: " + ex.Message);
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (oCompany.Connected == true)
                {
                    oCompany.Disconnect();
                    MessageBox.Show("You are now disconnected");
                }
                else MessageBox.Show("You are not connected to the company.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
                oCompany = null;
                Application.Exit();
            }
        }
    }
}
