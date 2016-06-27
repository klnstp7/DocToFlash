using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using System.Diagnostics;
using MySql.Data.MySqlClient;

namespace Daemon
{
    public partial class MainFrm : Form
    {
        private string applicationPath;
        private const string pdfFilesPath=@"\pdffiles\";
      

        //Setting
        public string watchFolderPath;
        public string swfToolsPath;
        public int intervalTime;
    
        private delegate void AppendDelegate(TextBox textBox, string text);

        private DateTime startDate;

        private SettingFrm settingForm = null;
        private QueryFrm queryForm = null;
        private FormWindowState preState;

        public MainFrm()
        {
            InitializeComponent();
        }

        #region Initial
        private void FrmMain_Load(object sender, EventArgs e)
        {
            applicationPath = Application.StartupPath;
            watchFolderPath =this.GetParmValue("watchfolderpath");
            swfToolsPath = this.GetParmValue("swftoolspath");
            string strInterValTime=this.GetParmValue("intervalTime");
            if(strInterValTime.Length>0){
                  intervalTime =Convert.ToInt32(strInterValTime);
            }
              
            if (Directory.Exists(watchFolderPath+"/resource") == false)
            {
                Directory.CreateDirectory(watchFolderPath + "/resource");
            }
            if (Directory.Exists(watchFolderPath + "/flash") == false)
            {
                Directory.CreateDirectory(watchFolderPath + "/flash");
            }
            
            if (Directory.Exists(applicationPath + pdfFilesPath) == false)
            {
                Directory.CreateDirectory(applicationPath + pdfFilesPath);
            }

            startDate = DateTime.Now;
            //this.StartJob();
           //this.WindowState = FormWindowState.Minimized;

        }

        private string GetParmValue(string parmKey)
        {
            string temp="";
            string sql="select datavalue from uh_config where var='"+parmKey+"'";
            MySqlDataReader reader=MysqlHelper.ExecuteReader(sql);
            if(reader.HasRows){
                reader.Read();
                temp=Convert.ToString(reader["datavalue"]);
            }
            reader.Close();
            return  temp;
        }


        #endregion

        #region Button Click
        private void btnSetting_Click(object sender, EventArgs e)
        {
            this.timer1.Enabled = false;
            this.btnPause.Enabled = false;


            if (settingForm == null || settingForm.IsDisposed)
            {
                settingForm = new SettingFrm();
                settingForm.StartPosition = FormStartPosition.CenterScreen;
                settingForm.Owner = this;
            }

            DialogResult result = settingForm.ShowDialog();
            if (result==DialogResult.OK)
            {
                this.SetParmValue("watchfolderpath", watchFolderPath);
                this.SetParmValue("swftoolspath", swfToolsPath);
                this.SetParmValue("intervalTime", intervalTime.ToString());
                this.btnStart.Enabled = true;
                this.btnPause.Enabled = true;
                this.btnQuery.Enabled = true;
                MessageBox.Show("Save successfully");

            }
            else if (result==DialogResult.Cancel)
            {
                settingForm = null;
            }
 
        }

        private void SetParmValue(string parmKey, string parmValue)
        {
            MySqlParameter[] parms = { new MySqlParameter("?svar", parmKey), new MySqlParameter("?sdatavalue", parmValue) };
            if (this.GetParmValue(parmKey).Length == 0)
            {
                MysqlHelper.ExecuteNonQuery("Insert into uh_config(var,datavalue) values(?svar,?sdatavalue) ", parms);
            }
            else
            {
                MysqlHelper.ExecuteNonQuery("update uh_config set datavalue=?sdatavalue where var=?svar", parms);
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            this.StartJob();
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            this.timer1.Enabled = false;
            this.btnSetting.Enabled = true;
            this.btnStart.Enabled = true;
            this.btnPause.Enabled = false;
            this.btnQuery.Enabled = true;
           
            WriteLog("[" + DateTime.Now + "] ********************* Job Pause ********************* ");
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (watchFolderPath.Length == 0 || swfToolsPath.Length == 0 || intervalTime <= 0 )
            {
                MessageBox.Show("请先配置系统参数");
                return;
            }

            this.timer1.Enabled = false;
            this.btnSetting.Enabled = true;
            this.btnStart.Enabled = true;
            this.btnPause.Enabled = true;
            this.btnQuery.Enabled = true; ;

            if (queryForm == null)
            {
                queryForm = new QueryFrm();
                queryForm.StartPosition = FormStartPosition.CenterParent;
                queryForm.Owner = this;
            }
            DialogResult result= queryForm.ShowDialog() ;
            
            if (result==DialogResult.Cancel)
            {
                queryForm = null;
            };

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定退出?", "提示",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        #endregion

        #region "Convert"
        private void timer1_Tick(object sender, EventArgs e)
        {
            this.timer1.Enabled = false;

            if (startDate.AddDays(7) < DateTime.Now)
            {
                this.tbxResult.Text = "";
                WriteLog("[" + DateTime.Now + "] ********************* Job Start ********************* ");
            }

            this.ConvertToSwf();

            this.ConvertToSwfTask();

            this.timer1.Enabled = true;
        }

        private void ConvertToSwf()
        {
           
            string Sql = "SELECT * FROM uh_co_resource " +
                        " WHERE tranresult IS NULL  " +
                        " order by dateline desc " +
                        " LIMIT 0,1";
            MySqlDataReader reader = MysqlHelper.ExecuteReader(Sql);

            string[] allowExt = { ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".pdf" };

            int id=0;

            string filepath = string.Empty;
            string fileName = string.Empty;
            string fileExtension = string.Empty;

            string sourcefile = string.Empty;
            string pdffile = string.Empty;
            string targetfile = string.Empty;

            bool tranresult = false;
            string errmsg = string.Empty;
            try
            {
                if (reader.Read())
                { 
                    id = Convert.ToInt32(reader["rid"]);
                    filepath = watchFolderPath + reader["filepath"];
                    fileName = Path.GetFileNameWithoutExtension(filepath);
                    fileExtension = Path.GetExtension(filepath).ToLower();

                    sourcefile = watchFolderPath + reader["filepath"];
                    pdffile = applicationPath + pdfFilesPath + fileName + fileExtension + ".pdf";
                    targetfile = watchFolderPath + reader["flashpath"];
                   
               
                    
                    if (allowExt.Contains(fileExtension) == true)
                    {
                        string cmd = string.Empty;

                        if (fileExtension == ".doc" || fileExtension == ".docx"
                            || fileExtension == ".ppt" || fileExtension == ".pptx"
                            || fileExtension == ".xls" || fileExtension == ".xlsx")
                        { //Doc,PPT,xls
                            bool convertToPdf = false;
                            ConvertToPDF convertor = new ConvertToPDF();
                            if (fileExtension == ".doc" || fileExtension == ".docx")
                            {
                                convertToPdf = convertor.DOCConvertToPDF(sourcefile, pdffile);
                            }
                            if (fileExtension == ".ppt" || fileExtension == ".pptx")
                            {
                                convertToPdf = convertor.PPTConvertToPDF(sourcefile, pdffile);
                            }
                            if (fileExtension == ".xls" || fileExtension == ".xlsx")
                            {
                                convertToPdf = convertor.XLSConvertToPDF(sourcefile, pdffile);
                            }

                            if (convertToPdf == true && File.Exists(pdffile) == true)
                            {
                                WriteLog("[" + DateTime.Now + "] " + sourcefile + " To PDF " + pdffile + " Success.");
                                cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t \"" + pdffile + "\" -o " + "\"" + targetfile + "\" -s poly2bitmap -s flashversion=9";
                                ExecuteCmd(cmd);
                                if (File.Exists(pdffile) == true)
                                {
                                    File.Delete(pdffile);
                                }
                            }

                        }
                        else //PDF
                        {
                            if (fileExtension == ".pdf")
                            {
                                cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t  \"" + filepath + "\" -o " + "\"" + targetfile + "\" -s poly2bitmap -s flashversion=9 ";

                                ExecuteCmd(cmd);

                            }
                        }
                    }
                    tranresult = true; 
                }
             }
            catch (Exception ex)
            {
                errmsg = ex.Message;
            }
            finally
            {
                reader.Close();
            }

            if (id > 0)
            {
                if (tranresult == true)
                {
                    Sql = "UPDATE uh_co_resource set trandate=UNIX_TIMESTAMP(),tranresult='S' where rid='" + id.ToString() + "' ";
                    WriteLog("[" + DateTime.Now + "] " + sourcefile + " To Flash  Success.");

                }
                else
                {
                    Sql = "UPDATE uh_co_resource set trandate=UNIX_TIMESTAMP(),tranresult='F',errormsg='" + errmsg + "' where rid='" + id.ToString() + "' ";
                    WriteLog("[" + DateTime.Now + "] " + sourcefile + " To Flash  Fail.");
                }
                MysqlHelper.ExecuteNonQuery(Sql);
            }
        }


        private void ConvertToSwfTask()
        {

            string Sql = "SELECT * FROM uh_co_task " +
                        " WHERE tranresult IS NULL  " +
                        " order by publishdate desc " +
                        " LIMIT 0,1";
            MySqlDataReader reader = MysqlHelper.ExecuteReader(Sql);

            string[] allowExt = { ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".pdf" };

            int id = 0;

            string filepath = string.Empty;
            string fileName = string.Empty;
            string fileExtension = string.Empty;

            string sourcefile = string.Empty;
            string pdffile = string.Empty;
            string targetfile = string.Empty;

            bool tranresult = false;
            string errmsg = string.Empty;
            try
            {
                if (reader.Read())
                {
                    id = Convert.ToInt32(reader["taskid"]);
                    filepath = watchFolderPath + reader["attachment"];
                    fileName = Path.GetFileNameWithoutExtension(filepath);
                    fileExtension = Path.GetExtension(filepath).ToLower();

                    sourcefile = watchFolderPath + reader["attachment"];
                    pdffile = applicationPath + pdfFilesPath + fileName + fileExtension + ".pdf";
                    targetfile = watchFolderPath + reader["attachflashpath"];

                    

                    if (allowExt.Contains(fileExtension) == true)
                    {
                        string cmd = string.Empty;

                        if (fileExtension == ".doc" || fileExtension == ".docx"
                            || fileExtension == ".ppt" || fileExtension == ".pptx"
                            || fileExtension == ".xls" || fileExtension == ".xlsx")
                        { //Doc,PPT,xls
                            bool convertToPdf = false;
                            ConvertToPDF convertor = new ConvertToPDF();
                            if (fileExtension == ".doc" || fileExtension == ".docx")
                            {
                                convertToPdf = convertor.DOCConvertToPDF(sourcefile, pdffile);
                            }
                            if (fileExtension == ".ppt" || fileExtension == ".pptx")
                            {
                                convertToPdf = convertor.PPTConvertToPDF(sourcefile, pdffile);
                            }
                            if (fileExtension == ".xls" || fileExtension == ".xlsx")
                            {
                                convertToPdf = convertor.XLSConvertToPDF(sourcefile, pdffile);
                            }

                            if (convertToPdf == true && File.Exists(pdffile) == true)
                            {
                                WriteLog("[" + DateTime.Now + "] " + sourcefile + " To PDF " + pdffile + " Success.");
                                cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t \"" + pdffile + "\" -o " + "\"" + targetfile + "\" -s poly2bitmap -s flashversion=9";
                                ExecuteCmd(cmd);
                                if (File.Exists(pdffile) == true)
                                {
                                    File.Delete(pdffile);
                                }
                            }

                        }
                        else //PDF
                        {
                            if (fileExtension == ".pdf")
                            {
                                cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t  \"" + filepath + "\" -o " + "\"" + targetfile + "\" -s poly2bitmap -s flashversion=9 ";

                                ExecuteCmd(cmd);

                            }
                        }
                    }
                    tranresult = true;
                }
            }
            catch (Exception ex)
            {
                errmsg = ex.Message;
            }
            finally
            {
                reader.Close();
            }

            if (id > 0)
            {
                if (tranresult == true)
                {
                    Sql = "UPDATE uh_co_task set trandate=UNIX_TIMESTAMP(),tranresult='S' where taskid='" + id.ToString() + "' ";
                    WriteLog("[" + DateTime.Now + "] " + sourcefile + " To Flash  Success.");

                }
                else
                {
                    Sql = "UPDATE uh_co_task set trandate=UNIX_TIMESTAMP(),tranresult='F',errormsg='" + errmsg + "' where taskid='" + id.ToString() + "' ";
                    WriteLog("[" + DateTime.Now + "] " + sourcefile + " To Flash  Fail.");
                }
                MysqlHelper.ExecuteNonQuery(Sql);
            }
        }
        #endregion

        #region Log File
        private void WriteLog(string str)
        {
            UpdateResult(this.tbxResult, str);
        }

        private void UpdateResult(TextBox textBox, string text)
        {
            if (!textBox.InvokeRequired)
            {
                textBox.AppendText(text + "\n");
            }
            else
            {
                AppendDelegate del = new AppendDelegate(UpdateResult);
                this.Invoke(del, textBox, text);
            }
        }
        #endregion

        #region Execute Cmd
        private void ExecuteCmd(string cmd)
        {
            try
            {
                using (Process p = new Process())
                {
                    p.StartInfo.FileName = "cmd";
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.CreateNoWindow = true;
                    p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    p.Start();
                    string strOutput = null;
                    p.StandardInput.WriteLine(cmd);
                    p.StandardInput.WriteLine("exit");
                    strOutput = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Job
        private void StartJob()
        {
            if (watchFolderPath.Length == 0 || swfToolsPath.Length == 0 || intervalTime <= 0)
            {
                MessageBox.Show("请先配置系统参数");
                return;
            }

            this.timer1.Enabled = true;
            this.btnSetting.Enabled = false;
            this.btnStart.Enabled = false;
            this.btnPause.Enabled = true;
            this.btnQuery.Enabled = true;

           
            WriteLog("[" + DateTime.Now + "] ********************* Job Start ********************* ");
      
        }
        #endregion

        #region NotifyIcon
        private void MainFrm_Resize(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Minimized)
           {
                preState = this.WindowState;            
           }
        }

        private void MainFrm_Deactivate(object sender, EventArgs e)
        {
             if (this.WindowState == FormWindowState.Minimized)
               this.Visible = false;
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
             if (this.Visible)
                this.Visible = false;
           else
               this.Visible = true;
            if (this.WindowState == FormWindowState.Minimized)
               this.WindowState = preState;
           else
                this.WindowState = FormWindowState.Minimized;
        }

        private void toolStripMenuItemShow_Click(object sender, EventArgs e)
        {
            this.Show();
        }

        private void toolStripMenuItemHide_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void toolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

    }
}
