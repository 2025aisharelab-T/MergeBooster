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
            DebugLog.Write(string.Format("===== MergeBooster 시작 (v{0}, OS {1}, .NET {2}) 로그 파일: {3}",
                typeof(Program).Assembly.GetName().Version, Environment.OSVersion, Environment.Version, DebugLog.CurrentPath));
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) => {
                DebugLog.Error("UI 스레드 처리되지 않은 예외", e.Exception);
                MessageBox.Show("예기치 않은 오류가 발생했습니다:\n" + e.Exception.Message + "\n\n상세 로그: " + DebugLog.CurrentPath, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) => DebugLog.Error("처리되지 않은 예외 (종료됨)", e.ExceptionObject as Exception);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            DebugLog.Write("===== MergeBooster 종료");
        }
    }
}
