# 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator GUI App) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Windows 단일 독립 실행 파일 (`ExcelDiffAggregator.exe`)을 C# .NET WinForms 기반으로 빌드하여, 원본 대비 수정본 엑셀 문서의 변경사항만 추출하고 노란색 하이라이트를 적용하는 GUI 프로그램 생성

**Architecture:** C# WinForms 모던 Flat UI + BackgroundWorker 비동기 프로세서 + Dynamic Excel COM Interop 엔진. 메모리 2D 배열 상에서 필터/숨김 행 무관 1:1 절대좌표 매칭.

**Tech Stack:** C# (.NET Framework 4.0 / WinForms), Native Windows `csc.exe` compiler.

## Global Constraints
- Target Binary: `ExcelDiffAggregator.exe`
- Compiler: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`
- Highlight Color: Light Yellow `#FFF2CC` (RGB: 255, 242, 204)
- Zero External Dependencies: Runs natively on Windows without requiring Python or third-party DLLs.

---

### Task 1: C# UI 및 엔진 소스 작성 (`Program.cs`, `MainForm.cs`, `ExcelDiffEngine.cs`)

**Files:**
- Create: `src/cs/Program.cs`
- Create: `src/cs/MainForm.cs`
- Create: `src/cs/ExcelDiffEngine.cs`

- [ ] **Step 1: Write `src/cs/ExcelDiffEngine.cs`**

```csharp
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace ExcelDiffAggregator
{
    public class ExcelDiffEngine
    {
        public delegate void LogHandler(string message);

        public static int ProcessDiff(
            string basePath,
            string modPath,
            string sheetName,
            bool applyHighlight,
            bool addComment,
            LogHandler logger
        )
        {
            logger?.Invoke("엑셀 프로세스 시작 중...");
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new Exception("Microsoft Excel이 이 컴퓨터에 설치되어 있지 않습니다.");
            }

            dynamic excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;

            int diffCount = 0;
            dynamic wbBase = null;
            dynamic wbMod = null;

            try
            {
                logger?.Invoke("원본 파일 열기: " + System.IO.Path.GetFileName(basePath));
                wbBase = excel.Workbooks.Open(basePath, ReadOnly: false);

                logger?.Invoke("수정본 파일 열기: " + System.IO.Path.GetFileName(modPath));
                wbMod = excel.Workbooks.Open(modPath, ReadOnly: true);

                dynamic wsBase = null;
                dynamic wsMod = null;

                if (string.IsNullOrEmpty(sheetName) || sheetName == "(첫번째 시트)")
                {
                    wsBase = wbBase.Worksheets[1];
                    wsMod = wbMod.Worksheets[1];
                }
                else
                {
                    wsBase = wbBase.Worksheets[sheetName];
                    wsMod = wbMod.Worksheets[sheetName];
                }

                logger?.Invoke("데이터 범위 읽는 중...");
                dynamic usedRange = wsBase.UsedRange;
                int rMax = usedRange.Rows.Count;
                int cMax = usedRange.Columns.Count;

                object[,] baseArr = (object[,])wsBase.Range[wsBase.Cells[1, 1], wsBase.Cells[rMax, cMax]].Value2;
                object[,] modArr = (object[,])wsMod.Range[wsMod.Cells[1, 1], wsMod.Cells[rMax, cMax]].Value2;

                logger?.Invoke("배열 데이터 1:1 비교 중 (" + rMax + "행 x " + cMax + "열)...");

                int colorYellow = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204));

                for (int r = 1; r <= rMax; r++)
                {
                    for (int c = 1; c <= cMax; c++)
                    {
                        object oBase = baseArr[r, c];
                        object oMod = modArr[r, c];

                        string vBase = oBase != null ? oBase.ToString().Trim() : "";
                        string vMod = oMod != null ? oMod.ToString().Trim() : "";

                        if (!string.IsNullOrEmpty(vMod) && vMod != vBase)
                        {
                            wsBase.Cells[r, c].Value2 = oMod;

                            if (applyHighlight)
                            {
                                wsBase.Cells[r, c].Interior.Color = colorYellow;
                            }

                            if (addComment)
                            {
                                try
                                {
                                    wsBase.Cells[r, c].ClearComments();
                                    wsBase.Cells[r, c].AddComment("기존 값: " + vBase + "\n수정일시: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                                }
                                catch { }
                            }
                            diffCount++;
                        }
                    }
                }

                logger?.Invoke("변경사항 저장 중...");
                wbMod.Close(false);
                wbBase.Save();
                wbBase.Close(true);

                return diffCount;
            }
            finally
            {
                excel.Quit();
                Marshal.ReleaseComObject(excel);
            }
        }
    }
}
```

- [ ] **Step 2: Write `src/cs/MainForm.cs`**

```csharp
using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

namespace ExcelDiffAggregator
{
    public class MainForm : Form
    {
        private TextBox txtBasePath;
        private TextBox txtModPath;
        private TextBox txtSheetName;
        private CheckBox chkHighlight;
        private CheckBox chkComment;
        private Button btnRun;
        private TextBox txtLog;
        private BackgroundWorker worker;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "📊 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator)";
            this.Size = new Size(620, 500);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("맑은 고딕", 9F, FontStyle.Regular);

            Label lblTitle = new Label
            {
                Text = "⚡ 엑셀 변경사항 비교 및 자동 덮어쓰기 취합기",
                Font = new Font("맑은 고딕", 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 78, 120),
                Location = new Point(20, 15),
                AutoSize = true
            };
            this.Controls.Add(lblTitle);

            // Group 1: Files
            GroupBox grpFiles = new GroupBox
            {
                Text = "1. 파일 지정",
                Location = new Point(20, 50),
                Size = new Size(565, 105)
            };

            Label lblBase = new Label { Text = "원본(기준) 파일:", Location = new Point(15, 28), AutoSize = true };
            txtBasePath = new TextBox { Location = new Point(120, 25), Size = new Size(330, 23), ReadOnly = true };
            Button btnSelectBase = new Button { Text = "파일 찾기...", Location = new Point(458, 24), Size = new Size(90, 25) };
            btnSelectBase.Click += (s, e) => {
                OpenFileDialog dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm;*.xls" };
                if (dlg.ShowDialog() == DialogResult.OK) txtBasePath.Text = dlg.FileName;
            };

            Label lblMod = new Label { Text = "수정본 파일:", Location = new Point(15, 63), AutoSize = true };
            txtModPath = new TextBox { Location = new Point(120, 60), Size = new Size(330, 23), ReadOnly = true };
            Button btnSelectMod = new Button { Text = "파일 찾기...", Location = new Point(458, 59), Size = new Size(90, 25) };
            btnSelectMod.Click += (s, e) => {
                OpenFileDialog dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm;*.xls" };
                if (dlg.ShowDialog() == DialogResult.OK) txtModPath.Text = dlg.FileName;
            };

            grpFiles.Controls.AddRange(new Control[] { lblBase, txtBasePath, btnSelectBase, lblMod, txtModPath, btnSelectMod });
            this.Controls.Add(grpFiles);

            // Group 2: Options
            GroupBox grpOpts = new GroupBox
            {
                Text = "2. 취합 옵션",
                Location = new Point(20, 165),
                Size = new Size(565, 65)
            };

            Label lblSheet = new Label { Text = "비교 시트명:", Location = new Point(15, 27), AutoSize = true };
            txtSheetName = new TextBox { Text = "(첫번째 시트)", Location = new Point(95, 24), Size = new Size(130, 23) };
            chkHighlight = new CheckBox { Text = "노란색 하이라이트 표시", Checked = true, Location = new Point(245, 26), AutoSize = true };
            chkComment = new CheckBox { Text = "이전 값 주석 남기기", Checked = true, Location = new Point(410, 26), AutoSize = true };

            grpOpts.Controls.AddRange(new Control[] { lblSheet, txtSheetName, chkHighlight, chkComment });
            this.Controls.Add(grpOpts);

            // Run Button
            btnRun = new Button
            {
                Text = "🚀 변경사항 취합 및 덮어쓰기 실행",
                Location = new Point(20, 240),
                Size = new Size(565, 40),
                BackColor = Color.FromArgb(31, 78, 120),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat
            };
            btnRun.Click += BtnRun_Click;
            this.Controls.Add(btnRun);

            // Log Textbox
            txtLog = new TextBox
            {
                Location = new Point(20, 295),
                Size = new Size(565, 150),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.Black,
                ForeColor = Color.LightGreen,
                Font = new Font("Consolas", 9F)
            };
            this.Controls.Add(txtLog);

            worker = new BackgroundWorker();
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        private void AppendLog(string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendLog), msg);
            }
            else
            {
                txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + "\r\n");
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBasePath.Text))
            {
                MessageBox.Show("원본 파일 경로를 선택해 주세요.", "경로 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(txtModPath.Text))
            {
                MessageBox.Show("수정본 파일 경로를 선택해 주세요.", "경로 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnRun.Enabled = false;
            txtLog.Clear();
            AppendLog("취합 작업 시작...");

            string basePath = txtBasePath.Text;
            string modPath = txtModPath.Text;
            string sheetName = txtSheetName.Text.Trim();
            bool highlight = chkHighlight.Checked;
            bool comment = chkComment.Checked;

            worker.RunWorkerAsync(new object[] { basePath, modPath, sheetName, highlight, comment });
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] args = (object[])e.Argument;
            int count = ExcelDiffEngine.ProcessDiff(
                (string)args[0],
                (string)args[1],
                (string)args[2],
                (bool)args[3],
                (bool)args[4],
                AppendLog
            );
            e.Result = count;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnRun.Enabled = true;
            if (e.Error != null)
            {
                AppendLog("❌ 오류 발생: " + e.Error.Message);
                MessageBox.Show("취합 중 오류가 발생했습니다:\n" + e.Error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                int count = (int)e.Result;
                AppendLog("🎉 취합 완료! 총 " + count + "개 셀 반영됨.");
                MessageBox.Show("성공적으로 취합이 완료되었습니다!\n총 " + count + "개 셀이 반영 및 하이라이트 되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
```

- [ ] **Step 3: Write `src/cs/Program.cs`**

```csharp
using System;
using System.Windows.Forms;

namespace ExcelDiffAggregator
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
```

- [ ] **Step 4: Commit Task 1**

```bash
git add src/cs/ExcelDiffEngine.cs src/cs/MainForm.cs src/cs/Program.cs
git commit -m "feat(cs): add C# WinForms source files for ExcelDiffAggregator"
```

---

### Task 2: C# 빌드 스크립트 작성 및 `ExcelDiffAggregator.exe` 컴파일

**Files:**
- Create: `scripts/build_exe.bat`
- Build Output: `ExcelDiffAggregator.exe`

- [ ] **Step 1: Write `scripts/build_exe.bat`**

```bat
@echo off
chcp 65001 > nul
echo Building ExcelDiffAggregator.exe using csc.exe...

set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe

if not exist "%CSC%" (
    echo Error: csc.exe not found at %CSC%
    exit /b 1
)

"%CSC%" /target:winexe /out:ExcelDiffAggregator.exe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll src\cs\Program.cs src\cs\MainForm.cs src\cs\ExcelDiffEngine.cs

if %ERRORLEVEL% equ 0 (
    echo Build successful! Created ExcelDiffAggregator.exe
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
```

- [ ] **Step 2: Run `scripts/build_exe.bat` to compile `.exe`**

Run: `scripts\build_exe.bat`
Expected: `ExcelDiffAggregator.exe` generated cleanly.

- [ ] **Step 3: Commit Task 2**

```bash
git add scripts/build_exe.bat ExcelDiffAggregator.exe
git commit -m "build: compile standalone C# ExcelDiffAggregator.exe binary"
```
