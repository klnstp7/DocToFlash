using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Daemon
{
    public partial class SettingFrm : Form
    {
        private MainFrm mainForm = null;

        //Setting
        public string watchFolderPath;
        public string swfToolsPath;
        public int intervalTime;
        public int previousDay;

        public SettingFrm()
        {
            InitializeComponent();
        }

        private void SettingFrm_Load(object sender, EventArgs e)
        {
            mainForm = (MainFrm)this.Owner;

            this.tbxFilePath.Text = mainForm.watchFolderPath;
            this.tbxSWFToolPath.Text = mainForm.swfToolsPath;
            this.tbxIntervalTime.Text = mainForm.intervalTime.ToString();
           

        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (this.tbxFilePath.Text.Trim().Length == 0 ||
                this.tbxSWFToolPath.Text.Trim().Length == 0 ||
                this.tbxIntervalTime.Text.Trim().Length == 0 )
            {
                MessageBox.Show("Every Field should be filled");
                return;
            }

            mainForm.watchFolderPath=this.tbxFilePath.Text;
            mainForm.swfToolsPath=this.tbxSWFToolPath.Text;
            mainForm.intervalTime=Convert.ToInt32(this.tbxIntervalTime.Text);

         
            this.DialogResult = DialogResult.OK;
            this.Close();
        }



        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

       
    }
}
