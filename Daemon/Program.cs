using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Reflection;


namespace Daemon
{
    static class Program
    {
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(System.IntPtr hWnd, int cmdShow);
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(System.IntPtr hWnd);
        private const int WS_SHOWNORMAL = 1;

        [STAThread]
        static void Main()
        {
            Process instance = RunningInstance();
            if (instance == null)
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                System.Windows.Forms.Application.DoEvents();
                Application.Run(new MainFrm());
            }
            else
            {
                HandleRunningInstance(instance);
            }

            
        }

        public static Process RunningInstance()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);
            foreach (Process process in processes)
            {
                if (process.Id != current.Id)
                {
                    if (Assembly.GetExecutingAssembly().Location.Replace("/", "\\") == current.MainModule.FileName)
                    {
                        return process;
                    }
                }
            }
            return null;
        }
        private static void HandleRunningInstance(Process instance)
        {
            MessageBox.Show("The system is running.", "Warning Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ShowWindowAsync(instance.MainWindowHandle, WS_SHOWNORMAL);
            SetForegroundWindow(instance.MainWindowHandle);
        } 
    }
}
