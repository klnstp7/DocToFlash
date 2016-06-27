using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using System.Configuration;
using System.Diagnostics;

namespace Daemon
{
    public partial class QueryFrm : Form
    {
        //Setting
        public string watchFolderPath;
        public string swfToolsPath;
        //application 
        private string applicationPath;
        private string pdfFilesPath;
  
       
        private MainFrm mainForm = null;

        public QueryFrm()
        {
            InitializeComponent();
        }

        private void QueryFrm_Load(object sender, EventArgs e)
        {
            mainForm = (MainFrm)this.Owner;
           
            watchFolderPath = mainForm.watchFolderPath;
            swfToolsPath = mainForm.swfToolsPath;

            applicationPath = Application.StartupPath;
            pdfFilesPath = ConfigurationSettings.AppSettings["pdffilespath"];

            DataTable dtTranResult = new DataTable();
            dtTranResult.Columns.Add("Value", typeof(string));
            dtTranResult.Columns.Add("Text", typeof(string));
            for (int i = 0; i < 3; i++)
            {
                DataRow row = dtTranResult.NewRow();
                switch (i)
                {
                    case 0:
                        row["Value"] = "";
                        row["Text"] = "";
                        break;
                    case 1:
                        row["Value"] = "S";
                        row["Text"] = "成功";
                        break;
                    case 2:
                        row["Value"] = "F";
                        row["Text"] = "失败";
                        break;
                    default:
                        break;
                }

                dtTranResult.Rows.Add(row);
            }

            this.comboBox1.DataSource = dtTranResult;
            this.comboBox1.ValueMember = "Value";
            this.comboBox1.DisplayMember = "Text";

            this.btnTry.Enabled = false;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            this.ShowResult();
            this.Enabled = true;
           
        }

        private void gvResult_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            int currentRow = e.RowIndex;
            if (currentRow > -1)
            {
                int id = Convert.ToInt32(this.gvResult.Rows[currentRow].Cells[0].Value);
                string tranresult = this.gvResult.Rows[currentRow].Cells[4].Value.ToString();

                if (tranresult != "成功")
                {
                    this.btnTry.Enabled = true;
                }
                else
                {
                    this.btnTry.Enabled = false;
                }
            }
        }

        private void btnTry_Click(object sender, EventArgs e)
        {
            this.btnTry.Enabled = false;
            int currentRow = Convert.ToInt32(this.gvResult.CurrentRow.Index);
            if (currentRow == -1)
            {
                MessageBox.Show("请选择资源");
                return;
            }
            if (MessageBox.Show("确定将该文档重新转化成flash?", "提示", MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes)
            {

                int id = Convert.ToInt32(this.gvResult.Rows[currentRow].Cells[0].Value);
                this.TryConvertToSwf(id);
               
                MessageBox.Show("重新转换成功");
            }
            this.ShowResult();
            this.btnTry.Enabled = true;
        }

        private void ShowResult()
        {
            
            string wheresql = "1";
            wheresql += " AND DATE_FORMAT(FROM_UNIXTIME( a.dateline),'%Y-%m-%d')>='" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
            wheresql += " AND DATE_FORMAT(FROM_UNIXTIME( a.dateline),'%Y-%m-%d')<='" + this.dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'";
            if (this.comboBox1.SelectedValue.ToString() != "")
            {
                wheresql += " AND a.Tranresult='" + this.comboBox1.SelectedValue + "'";
            }
            if (this.textBox1.Text.Trim().Length > 0)
            {
                wheresql += " AND a.title like '%" + this.comboBox1.SelectedValue + "%'";
            }
            string Sql = "SELECT a.rid,a.title AS rname,b.title AS cname, " +
                        "DATE_FORMAT(FROM_UNIXTIME( a.dateline),'%Y-%m-%d') as upldate, " +
                        "CASE IFNULL(a.tranresult,'') WHEN  '' THEN '等待' WHEN 'S' THEN '成功' WHEN 'F' THEN '失败' ELSE '未知' END  AS tranresult, " +
                        "a.errormsg " +
                         "FROM uh_co_resource a INNER JOIN uh_co_course b " +
                        "ON a.cid=b.cid WHere " + wheresql + " ORder by a.dateline desc";

            DataTable dtResult = MysqlHelper.ExecuteDataTable(Sql);
            this.gvResult.DataSource = dtResult;
            this.gvResult.Refresh();
            
        }
        private void TryConvertToSwf(int id)
        {
            //string Sql = "SELECT * FROM uh_co_resource WHERE  DATE_FORMAT(FROM_UNIXTIME( dateline),'%y-%m-%d')=  DATE_FORMAT( SYSDATE(),'%y-%m-%d')";
            string Sql = "SELECT * FROM uh_co_resource " +
                        "WHERE rid='"+id.ToString()+"'";
            MySqlDataReader reader = MysqlHelper.ExecuteReader(Sql);

            string[] allowExt = { ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".pdf" };

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
                    filepath = watchFolderPath + reader["filepath"];
                    fileName = Path.GetFileNameWithoutExtension(filepath);
                    fileExtension = Path.GetExtension(filepath).ToLower();

                    sourcefile = watchFolderPath + reader["filepath"];
                    pdffile = applicationPath + pdfFilesPath + fileName + fileExtension + ".pdf";
                    targetfile = watchFolderPath + reader["flashpath"];
                    if (File.Exists(targetfile) == true)
                    {
                        File.Delete(targetfile);
                    }
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
                    Sql = "UPDATE uh_co_resource set trandate=UNIX_TIMESTAMP(),tranresult='S',errormsg=null,trytimes=ifnull(trytimes,0)+1 where rid='" + id.ToString() + "' ";

                }
                else
                {
                    Sql = "UPDATE uh_co_resource set trandate=UNIX_TIMESTAMP(),tranresult='F',errormsg='" + errmsg + "',trytimes=ifnull(trytimes,0)+1 where rid='" + id.ToString() + "' ";
                   
                }
                MysqlHelper.ExecuteNonQuery(Sql);
            }
        }

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
     
    }
}
