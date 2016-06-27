using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using System.IO;

using Microsoft.Office;

using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;


namespace TransformToFlashService
{
    public partial class Service1 : ServiceBase
    {
        private string applicationPath;
        private string logFileName;
        private string logFilePath;

        private string watchFolderPath;
        private string targetFolderPath;
        private string pdfFilesPath;
        private string swfToolsPath;

        private FileStream stream;
        private StreamWriter writer;

        private static object _sysobj = new object();

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            applicationPath = AppDomain.CurrentDomain.BaseDirectory;
            watchFolderPath = ConfigurationSettings.AppSettings["watchfolderpath"];
            targetFolderPath = ConfigurationSettings.AppSettings["targetfolderpath"];
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
            this.fileSystemWatcher1.Path = watchFolderPath;
            this.fileSystemWatcher1.Filter = "*.*";
            this.fileSystemWatcher1.EnableRaisingEvents = true;
        }

       
        private void fileSystemWatcher1_Created(object sender, FileSystemEventArgs e)
        {
            lock (_sysobj)
            {
                Convert(e.FullPath);
            }
        }

        #region Main Method
        private void Convert(string filepath)
        {
            try
            {
                OpenLogFile();

                string filePath = Path.GetFileName(filepath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string fileExtension = Path.GetExtension(filePath).ToLower();
                string[] allowExt = { ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".pdf" };

                if (allowExt.Contains(fileExtension) == true)
                {
                    string sourcefile = watchFolderPath + fileName + fileExtension;
                    string pdffile = applicationPath + pdfFilesPath + fileName + fileExtension + ".pdf";
                    string targetfile = targetFolderPath + fileName + fileExtension + ".swf";

                    string cmd = string.Empty;

                    if (fileExtension == ".doc" || fileExtension == ".docx" || fileExtension == ".ppt" || fileExtension == ".pptx" || fileExtension == ".xls" || fileExtension == ".xlsx")
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
                            if (File.Exists(targetfile) == false)
                            {
                                cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t \"" + pdffile + "\" -o " + "\"" + targetfile + "\"  -s poly2bitmap -s flashversion=9";
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
                            cmd = "\"" + swfToolsPath + "pdf2swf.exe" + "\" -t  \"" + filepath + "\" -o " + "\"" + targetfile + "\" -s flashversion=9 ";
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
            catch (Exception ex)
            {
                WriteLog("[" + DateTime.Now + "] Error Message:" + ex.Message);
                return;
            }
            finally
            {
                CloseLogFile();
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

        #region Log
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
        }

        private void CloseLogFile()
        {
            writer.Close();
            stream.Close();
        }
        #endregion

        protected override void OnStop()
        {
            this.fileSystemWatcher1.EnableRaisingEvents = false;
        }

    }
}
