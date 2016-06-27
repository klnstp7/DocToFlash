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

namespace TransformToFlash_ScheduleJob
{
    public partial class Form1 : Form
    {
        private string applicationPath;
        private string logFileName;
        private string logFilePath;

        private string watchFolderPath;
        private string pdfFilesPath;
        private string swfToolsPath;

        private FileStream stream;
        private StreamWriter writer;
        private delegate void AppendDelegate(TextBox textBox, string text);
        private static object _sysobj = new object();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            applicationPath = Application.StartupPath;
            watchFolderPath = ConfigurationSettings.AppSettings["watchfolderpath"];
            pdfFilesPath = ConfigurationSettings.AppSettings["pdffilespath"];
            swfToolsPath = ConfigurationSettings.AppSettings["swftoolspath"];
            logFilePath = ConfigurationSettings.AppSettings["logfilepath"];

  
            if (Directory.Exists(applicationPath + logFilePath) == false)
            {
                Directory.CreateDirectory(applicationPath + logFilePath);
            }
            if (Directory.Exists(applicationPath + pdfFilesPath) == false)
            {
                Directory.CreateDirectory(applicationPath + pdfFilesPath);
            }

            string Sql = "SELECT * FROM uh_co_resource WHERE  DATE_FORMAT(FROM_UNIXTIME( dateline),'%y-%m-%d')=  DATE_FORMAT( SYSDATE(),'%y-%m-%d')";
            System.Data.DataTable dataTable = MysqlHelper.ExecuteDataTable(Sql);

            
           //遍历文件
            try
            {
                OpenLogFile();
                WriteLog("[" + DateTime.Now + "] ********************* Schedule Job Start ********************* ");
                WriteLog(" ");
                foreach (DataRow row in dataTable.Rows)
                {
                    lock (_sysobj)
                    {
                       
                        string filepath = watchFolderPath+row["filepath"];
                        string fileName = Path.GetFileNameWithoutExtension(filepath);
                        string fileExtension = Path.GetExtension(filepath).ToLower();
                        string[] allowExt = { ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".pdf" };

                        if (allowExt.Contains(fileExtension) == true)
                        {
                            string sourcefile = watchFolderPath + row["filepath"];
                            string pdffile = applicationPath + pdfFilesPath + fileName + fileExtension + ".pdf";
                            string targetfile = watchFolderPath + row["flashpath"];

                            string cmd = string.Empty;

                            if ((fileExtension == ".doc" || fileExtension == ".docx"
                                || fileExtension == ".ppt" || fileExtension == ".pptx"
                                || fileExtension == ".xls" || fileExtension == ".xlsx"
                                ) && File.Exists(targetfile) == false)
                            {
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
                                    if (File.Exists(targetfile) == true)
                                    {
                                        WriteLog("[" + DateTime.Now + "] " + pdffile + " To Flash " + targetfile + " Success.");
                                    }
                                    else
                                    {
                                        WriteLog("[" + DateTime.Now + "] " + pdffile + " To Flash " + targetfile + " Fail.");
                                    }
                                    if (File.Exists(pdffile) == true)
                                    {
                                        File.Delete(pdffile);
                                    }
                                }
                                else
                                {
                                    WriteLog("[" + DateTime.Now + "] " + sourcefile + " To PDF " + pdffile + " Fail.");
                                }
                            }
                            else
                            {
                                if (fileExtension == ".pdf" && File.Exists(targetfile) == false)
                                {
                                    cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t  \"" + filepath + "\" -o " + "\"" + targetfile + "\" -s poly2bitmap -s flashversion=9 ";

                                    ExecuteCmd(cmd);
                                    if (File.Exists(targetfile) == true)
                                    {
                                        WriteLog("[" + DateTime.Now + "] " + filepath + " To Flash " + targetfile + " Success.");
                                    }
                                    else
                                    {
                                        WriteLog("[" + DateTime.Now + "] " + filepath + " To Flash " + targetfile + " Fail.");
                                    }
                                }
                            }
                        }
                      
                    }
                   
                }
                WriteLog(" ");
                WriteLog("[" + DateTime.Now + "] ********************* Schedule Job End ********************* ");
            }
            catch (Exception ex)
            {
                WriteLog("[" + DateTime.Now + "] Error Message:" + ex.Message);

            }
            finally
            {
                CloseLogFile();
            }
            Application.Exit();
        }


        #region Log File
        private void OpenLogFile()
        {
            string todayStr = DateTime.Today.ToString("yyyyMMdd");
            logFileName = todayStr + ".txt";
            stream = new FileStream(applicationPath + @"\" + logFilePath + @"\" + logFileName, FileMode.Append, FileAccess.Write);
            writer = new StreamWriter(stream);
        }

        private void WriteLog(string str)
        {
            writer.WriteLine(str);
            writer.WriteLine("");

            UpdateResult(this.tbxResult, str);
        }

        private void CloseLogFile()
        {
            writer.Close();
            stream.Close();
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
        #endregion
    }
}
