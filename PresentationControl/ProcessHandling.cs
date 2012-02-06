using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace PresentationControl
{
    static class ProcessHandling
    {
        #region Private

        #region External Functions

        [DllImport("user32.dll")]
        static extern int GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32")]
        static extern UInt32 GetWindowThreadProcessId(Int32 hWnd, out Int32 lpdwProcessId);

        #endregion

        #endregion

        #region Public

        /* Metoda zwracająca ID procesu */
        public static Int32 GetWindowProcessID(Int32 hwnd)
        {
            Int32 pid = 1;
            GetWindowThreadProcessId(hwnd, out pid);
            return pid;
        }

        /* Metoda zwracająca nazwę procesu aktywnej aplikacji */
        public static string GetActiveWindowProcessName()
        {
            Int32 hwnd = GetForegroundWindow();
            string appProcessName = Process.GetProcessById(GetWindowProcessID(hwnd)).ProcessName;
            return appProcessName;
        }

        /* Metoda ustawiająca aktywne okno */
        public static void SetActiveWindow(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);

            foreach (Process p in processes)
            {
                IntPtr hWnd = p.MainWindowHandle;
                SetForegroundWindow(hWnd);
            }
        }

        #endregion
    }
}
