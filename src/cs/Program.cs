using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDiffAggregator
{
    static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main()
        {
            // Without this, Windows stretches the window as a bitmap on scaled displays and text turns blurry
            SetProcessDPIAware();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
