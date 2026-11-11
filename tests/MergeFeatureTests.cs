using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDiffAggregator.Tests
{
    // Stage 1: formula protection, key-column row matching. Stage 2: preview, change report, backup prune/restore.
    static class MergeFeatureTests
    {
        public static int Failures;
        public static int Passes;

        private static void Check(bool ok, string what)
        {
            if (ok) { Passes++; Console.WriteLine("  PASS  " + what); }
            else { Failures++; Console.WriteLine("  FAIL  " + what); }
        }

        // ---------- workbook helpers ----------

        // cells: strings starting with "=" become formulas; everything else is written as a value
        private static void WriteBook(Type excelType, string path, string sheetName, object[,] cells)
        {
            dynamic excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                dynamic wb = excel.Workbooks.Add();
                dynamic ws = wb.Worksheets[1];
                ws.Name = sheetName;
                for (int r = 0; r < cells.GetLength(0); r++)
                {
                    for (int c = 0; c < cells.GetLength(1); c++)
                    {
                        object v = cells[r, c];
                        if (v == null) continue;
                        string s = v as string;
                        if (s != null && s.StartsWith("=")) ws.Cells[r + 1, c + 1].Formula = s;
                        else if (s != null && s.StartsWith("'"))
                        {
                            // "'1,000" -> stored as text, like a user typing into a text-formatted cell
                            ws.Cells[r + 1, c + 1].NumberFormat = "@";
                            ws.Cells[r + 1, c + 1].Value2 = s.Substring(1);
                        }
                        else ws.Cells[r + 1, c + 1].Value2 = v;
                    }
                }
                if (File.Exists(path)) File.Delete(path);
                wb.SaveAs(path, 51);
                wb.Close(false);
                Marshal.ReleaseComObject((object)wb);
            }
            finally
            {
                try { excel.Quit(); } catch { }
                try { Marshal.ReleaseComObject((object)excel); } catch { }
            }
        }

        private class CellInfo
        {
            public string Text;
            public bool HasFormula;
            public string Formula;
        }

        private static Dictionary<string, CellInfo> ReadCells(Type excelType, string path, string sheetName, params string[] addresses)
        {
            Dictionary<string, CellInfo> map = new Dictionary<string, CellInfo>();
            dynamic excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                dynamic wb = excel.Workbooks.Open(path, 0, true);
                dynamic ws = wb.Worksheets[sheetName];
                foreach (string a in addresses)
                {
                    dynamic cell = ws.Range[a];
                    CellInfo info = new CellInfo();
                    object v = cell.Value2;
                    info.Text = v == null ? "" : v.ToString();
                    info.HasFormula = (bool)cell.HasFormula;
                    info.Formula = Convert.ToString(cell.Formula);
                    map[a] = info;
                }
                wb.Close(false);
                Marshal.ReleaseComObject((object)wb);
            }
            finally
            {
                try { excel.Quit(); } catch { }
                try { Marshal.ReleaseComObject((object)excel); } catch { }
            }
            return map;
        }

        private static List<string> SheetNames(Type excelType, string path)
        {
            List<string> names = new List<string>();
            dynamic excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                dynamic wb = excel.Workbooks.Open(path, 0, true);
                foreach (dynamic ws in wb.Worksheets) names.Add((string)ws.Name);
                wb.Close(false);
                Marshal.ReleaseComObject((object)wb);
            }
            finally
            {
                try { excel.Quit(); } catch { }
                try { Marshal.ReleaseComObject((object)excel); } catch { }
            }
            return names;
        }

        private static int UsedRows(Type excelType, string path, string sheetName)
        {
            dynamic excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                dynamic wb = excel.Workbooks.Open(path, 0, true);
                dynamic ws = wb.Worksheets[sheetName];
                int rows = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count - 1;
                wb.Close(false);
                Marshal.ReleaseComObject((object)wb);
                return rows;
            }
            finally
            {
                try { excel.Quit(); } catch { }
                try { Marshal.ReleaseComObject((object)excel); } catch { }
            }
        }

        private static MergeOptions Options(string basePath, params string[] mods)
        {
            MergeOptions o = new MergeOptions();
            o.BasePath = basePath;
            o.ModPaths = mods;
            o.ApplyHighlight = true;
            o.AddComment = true;
            o.AutoBackup = false;
            o.CreateReport = false;
            o.Preview = null;
            return o;
        }

        private static MergeResult Run(MergeOptions o)
        {
            return ExcelDiffEngine.ProcessMerge(o, msg => Console.WriteLine("        [LOG] " + msg));
        }

        private static int CountStatus(MergeResult r, ChangeStatus s)
        {
            return r.Count(s);
        }

        // ---------- data sets ----------

        private static object[,] FormulaBase()
        {
            return new object[,] {
                { "항목", "값" },
                { "가", 10 },
                { "나", 20 },
                { "합계", "=SUM(B2:B3)" }
            };
        }

        private static object[,] FormulaMod()
        {
            return new object[,] {
                { "항목", "값" },
                { "가", 15 },
                { "나", 20 },
                { "합계", "=SUM(B2:B3)" }   // 저장 시 계산값 35 → 예전 엔진은 이 값을 원본 수식 칸에 덮어씀
            };
        }

        private static object[,] StaffBase()
        {
            return new object[,] {
                { "사번", "성명", "부서" },
                { 1001, "홍길동", "영업1팀" },
                { 1002, "김철수", "개발1팀" },
                { 1003, "이영희", "기획팀" }
            };
        }

        // Columns reordered, rows reordered, 1002 changed, 1004 added, 1001 missing
        private static object[,] StaffModShuffled()
        {
            return new object[,] {
                { "부서", "사번", "성명" },
                { "기획팀", 1003, "이영희" },
                { "인사팀", 1004, "신규" },
                { "AI개발팀", 1002, "김철수" }
            };
        }

        private static object[,] TitledBase()
        {
            return new object[,] {
                { "직원 명단 (2026)", null, null },
                { "사번", "성명", "부서" },
                { 1001, "홍길동", "영업1팀" },
                { 1002, "김철수", "개발1팀" },
                { 1003, "이영희", "기획팀" }
            };
        }

        private static object[,] TitledMod(string dept1002, string dept1004)
        {
            return new object[,] {
                { "직원 명단 (2026)", null, null },
                { "사번", "성명", "부서" },
                { 1001, "홍길동", "영업1팀" },
                { 1002, "김철수", dept1002 },
                { 1003, "이영희", "기획팀" },
                { 1004, "신규", dept1004 }
            };
        }

        // =====================================================================
        public static void RunStage1(Type excelType, string root)
        {
            string dir = Path.Combine(root, "stage1");
            Directory.CreateDirectory(dir);
            Console.WriteLine("\n===== 1단계 검증: 수식 칸 보호 =====");

            string fBase = Path.Combine(dir, "formula_base.xlsx");
            string fBase2 = Path.Combine(dir, "formula_base_noprotect.xlsx");
            string fMod = Path.Combine(dir, "formula_mod.xlsx");
            WriteBook(excelType, fBase, "Data", FormulaBase());
            WriteBook(excelType, fBase2, "Data", FormulaBase());
            WriteBook(excelType, fMod, "Data", FormulaMod());

            MergeOptions o = Options(fBase, fMod);
            o.ProtectFormulas = true;
            MergeResult r = Run(o);
            Dictionary<string, CellInfo> c = ReadCells(excelType, fBase, "Data", "B2", "B4");
            Check(c["B2"].Text == "15", "수식 보호 ON: 일반 칸 B2는 15로 반영 (실제 " + c["B2"].Text + ")");
            Check(c["B4"].HasFormula && c["B4"].Formula == "=SUM(B2:B3)", "수식 보호 ON: B4 수식 =SUM(B2:B3) 유지 (실제 " + c["B4"].Formula + ")");
            Check(r.AppliedCells == 1, "수식 보호 ON: 반영 칸 수 1 (실제 " + r.AppliedCells + ")");
            Check(CountStatus(r, ChangeStatus.FormulaProtected) == 1, "수식 보호 ON: '수식 보호로 건너뜀' 1건 기록");

            o = Options(fBase2, fMod);
            o.ProtectFormulas = false;
            r = Run(o);
            c = ReadCells(excelType, fBase2, "Data", "B4");
            Check(!c["B4"].HasFormula && c["B4"].Text == "35", "수식 보호 OFF: 예전 동작 재현 - B4 수식이 값 35로 바뀜 (보호 기능이 실제로 차이를 만든다는 대조군)");

            Console.WriteLine("\n===== 1단계 검증: 키 열 기준 행 매칭 =====");
            string kBasePos = Path.Combine(dir, "staff_base_position.xlsx");
            string kBaseKey = Path.Combine(dir, "staff_base_key.xlsx");
            string kMod = Path.Combine(dir, "staff_mod_shuffled.xlsx");
            WriteBook(excelType, kBasePos, "직원", StaffBase());
            WriteBook(excelType, kBaseKey, "직원", StaffBase());
            WriteBook(excelType, kMod, "직원", StaffModShuffled());

            o = Options(kBasePos, kMod);
            r = Run(o);
            Check(r.AppliedCells >= 8, "대조군(위치 비교): 행·열 순서만 바뀐 파일인데 " + r.AppliedCells + "칸이 '변경'으로 덮어써짐 (기존 문제 재현)");

            o = Options(kBaseKey, kMod);
            o.KeyColumns = "사번";
            o.HeaderRow = 1;
            r = Run(o);
            c = ReadCells(excelType, kBaseKey, "직원", "A2", "B2", "C2", "A3", "C3", "A4", "C4", "A5", "B5", "C5");
            Check(r.AppliedCells == 4, "키 매칭: 반영 칸 = 1002 부서 1칸 + 새 행 3칸 = 4 (실제 " + r.AppliedCells + ")");
            Check(c["C3"].Text == "AI개발팀", "키 매칭: 순서가 바뀌어도 1002의 부서를 찾아 C3에 반영 (실제 " + c["C3"].Text + ")");
            Check(c["A2"].Text == "1001" && c["C2"].Text == "영업1팀", "키 매칭: 수정본에 없는 1001 행은 그대로 (삭제 안 함)");
            Check(c["A4"].Text == "1003" && c["C4"].Text == "기획팀", "키 매칭: 바뀌지 않은 1003 행은 그대로");
            Check(c["A5"].Text == "1004" && c["B5"].Text == "신규" && c["C5"].Text == "인사팀", "키 매칭: 새 키 1004는 5행에 머리글 위치에 맞춰 추가 (A=1004, B=신규, C=인사팀)");
            Check(r.NewRows == 1, "키 매칭: 새 행 수 1 (실제 " + r.NewRows + ")");
            Check(CountStatus(r, ChangeStatus.MissingInMod) == 1, "키 매칭: '수정본에 없는 행' 1건(1001) 기록");

            Console.WriteLine("\n===== 1단계 검증: 제목 행 + 여러 파일 충돌 + 같은 새 키 =====");
            string tBaseFirst = Path.Combine(dir, "titled_base_first.xlsx");
            string tBaseLast = Path.Combine(dir, "titled_base_last.xlsx");
            string tMod1 = Path.Combine(dir, "titled_mod1.xlsx");
            string tMod2 = Path.Combine(dir, "titled_mod2.xlsx");
            WriteBook(excelType, tBaseFirst, "명단", TitledBase());
            WriteBook(excelType, tBaseLast, "명단", TitledBase());
            WriteBook(excelType, tMod1, "명단", TitledMod("A팀", "인사팀"));
            WriteBook(excelType, tMod2, "명단", TitledMod("B팀", "총무팀"));

            o = Options(tBaseFirst, tMod1, tMod2);
            o.KeyColumns = "사번";
            o.HeaderRow = 2;
            o.ConflictPolicy = ConflictPolicy.FirstWins;
            r = Run(o);
            c = ReadCells(excelType, tBaseFirst, "명단", "A1", "C4", "A6", "C6", "A7");
            Check(c["A1"].Text == "직원 명단 (2026)", "머리글 2행: 1행 제목은 비교·변경하지 않음");
            Check(c["C4"].Text == "A팀", "첫 파일 우선: 1002 부서는 첫 파일 값 A팀 (실제 " + c["C4"].Text + ")");
            Check(c["A6"].Text == "1004" && c["C6"].Text == "인사팀", "첫 파일 우선: 새 키 1004는 6행 하나만, 부서는 첫 파일 값 인사팀");
            Check(c["A7"].Text == "", "두 파일이 같은 새 키를 보내도 행이 중복 추가되지 않음 (7행 비어 있음)");
            Check(CountStatus(r, ChangeStatus.ConflictKept) == 2, "첫 파일 우선: '충돌 - 기존 값 유지' 2건 (실제 " + CountStatus(r, ChangeStatus.ConflictKept) + ")");

            o = Options(tBaseLast, tMod1, tMod2);
            o.KeyColumns = "사번";
            o.HeaderRow = 2;
            o.ConflictPolicy = ConflictPolicy.LastWins;
            r = Run(o);
            c = ReadCells(excelType, tBaseLast, "명단", "C4", "C6", "A7");
            Check(c["C4"].Text == "B팀" && c["C6"].Text == "총무팀", "마지막 파일 우선: 1002=B팀, 1004=총무팀");
            Check(c["A7"].Text == "", "마지막 파일 우선: 새 행 중복 없음");
            Check(r.NewRows == 1, "마지막 파일 우선: 새 행 수 1");

            string missingKeyBase = Path.Combine(dir, "badkey_base.xlsx");
            WriteBook(excelType, missingKeyBase, "직원", StaffBase());
            byte[] before = File.ReadAllBytes(missingKeyBase);
            o = Options(missingKeyBase, kMod);
            o.KeyColumns = "없는열이름";
            r = Run(o);
            Check(r.AppliedCells == 0 && Equal(before, File.ReadAllBytes(missingKeyBase)), "없는 키 열 이름: 아무것도 바꾸지 않고 원본 파일 그대로 (저장도 안 함)");
        }

        // =====================================================================
        public static void RunStage2(Type excelType, string root)
        {
            string dir = Path.Combine(root, "stage2");
            Directory.CreateDirectory(dir);

            Console.WriteLine("\n===== 2단계 검증: 미리보기 묶음 제외 =====");
            string pBase = Path.Combine(dir, "preview_base.xlsx");
            string pModA = Path.Combine(dir, "A부서.xlsx");
            string pModB = Path.Combine(dir, "B부서.xlsx");
            WriteBook(excelType, pBase, "직원", StaffBase());
            WriteBook(excelType, pModA, "직원", new object[,] { { "사번", "성명", "부서" }, { 1001, "홍길동", "영업2팀" }, { 1002, "김철수", "개발1팀" }, { 1003, "이영희", "기획팀" } });
            WriteBook(excelType, pModB, "직원", new object[,] { { "사번", "성명", "부서" }, { 1001, "홍길동", "영업2팀" }, { 1002, "김철수", "QA팀" }, { 1003, "이영희", "기획팀" } });

            int previewCalls = 0;
            int previewItems = 0;
            List<string> groupLabels = null;
            MergeOptions o = Options(pBase, pModA, pModB);
            o.KeyColumns = "사번";
            o.Preview = items =>
            {
                previewCalls++;
                previewItems = items.Count;
                using (PreviewDialog dlg = new PreviewDialog(items))
                {
                    groupLabels = dlg.GroupLabels;
                    dlg.SetGroup("파일: B부서.xlsx", false);   // B부서 변경은 반영하지 않음
                }
                return true;
            };
            MergeResult r = Run(o);
            Dictionary<string, CellInfo> c = ReadCells(excelType, pBase, "직원", "C2", "C3");
            Check(previewCalls == 1 && previewItems == 2, "미리보기는 원본에 쓰기 전에 1번 호출, 변경 예정 2건 (실제 " + previewCalls + "번, " + previewItems + "건)");
            Check(groupLabels != null && groupLabels.Contains("파일: A부서.xlsx") && groupLabels.Contains("열: C (부서)"), "미리보기 묶음 목록에 파일·열 묶음이 있음");
            Check(c["C2"].Text == "영업2팀", "승인한 A부서 변경(1001 → 영업2팀)은 반영");
            Check(c["C3"].Text == "개발1팀", "묶음 제외한 B부서 변경(1002 → QA팀)은 반영 안 됨, 원래 값 유지");
            Check(CountStatus(r, ChangeStatus.Excluded) == 1 && r.AppliedCells == 1, "결과: 반영 1건, 제외 1건");

            Console.WriteLine("\n===== 2단계 검증: 같은 칸, 나중 파일 제외 시 앞 파일 값 반영 =====");
            string sBase = Path.Combine(dir, "supersede_base.xlsx");
            WriteBook(excelType, sBase, "직원", StaffBase());
            string sMod1 = Path.Combine(dir, "첫째.xlsx");
            string sMod2 = Path.Combine(dir, "둘째.xlsx");
            WriteBook(excelType, sMod1, "직원", new object[,] { { "사번", "성명", "부서" }, { 1001, "홍길동", "X팀" } });
            WriteBook(excelType, sMod2, "직원", new object[,] { { "사번", "성명", "부서" }, { 1001, "홍길동", "Y팀" } });
            o = Options(sBase, sMod1, sMod2);
            o.Preview = items =>
            {
                foreach (ChangeItem it in items) if (it.FileName == "둘째.xlsx") it.Approved = false;
                return true;
            };
            r = Run(o);
            c = ReadCells(excelType, sBase, "직원", "C2");
            Check(c["C2"].Text == "X팀", "나중 파일(Y팀)을 제외하면 앞 파일 값 X팀이 반영 (실제 " + c["C2"].Text + ")");

            Console.WriteLine("\n===== 2단계 검증: 미리보기 엑셀 표 보기 =====");
            string gBase = Path.Combine(dir, "grid_base.xlsx");
            WriteBook(excelType, gBase, "직원", StaffBase());
            o = Options(gBase, pModA, pModB);
            o.KeyColumns = "사번";
            string gC2 = null, gC3Before = null, gC3After = null, gB2 = null, gA1 = null;
            bool gSheetAttached = false, gToggled = false, gToggleNoChange = true;
            o.Preview = items =>
            {
                gSheetAttached = items.Count > 0 && items.TrueForAll(it => it.Sheet != null && it.Sheet.Rows == 4 && it.Sheet.Cols == 3);
                using (PreviewDialog dlg = new PreviewDialog(items))
                {
                    gA1 = dlg.DisplayText("직원", 1, 1);
                    gB2 = dlg.DisplayText("직원", 2, 2);
                    gC2 = dlg.DisplayText("직원", 2, 3);
                    gC3Before = dlg.DisplayText("직원", 3, 3);
                    gToggled = dlg.ToggleCell("직원", 3, 3);          // C3(B부서 QA팀) 더블클릭 → 제외
                    gToggleNoChange = dlg.ToggleCell("직원", 2, 2);   // 바뀌지 않는 칸은 무시
                    gC3After = dlg.DisplayText("직원", 3, 3);
                    Capture(dlg, Path.Combine(dir, "ui_preview_grid.png"));
                }
                return true;
            };
            r = Run(o);
            c = ReadCells(excelType, gBase, "직원", "C2", "C3");
            Check(gSheetAttached, "미리보기 항목마다 원본 시트 값(4행×3열)이 함께 전달됨");
            Check(gA1 == "사번" && gB2 == "홍길동", "표에 바뀌지 않는 원본 칸도 그대로 보임 (A1=" + gA1 + ", B2=" + gB2 + ")");
            Check(gC2 == "영업2팀" && gC3Before == "QA팀", "표의 변경 칸은 반영될 새 값으로 보임 (C2=" + gC2 + ", C3=" + gC3Before + ")");
            Check(gToggled && !gToggleNoChange && gC3After == "개발1팀", "변경 칸 더블클릭 시 제외되어 원래 값으로 표시, 안 바뀌는 칸은 무시 (C3=" + gC3After + ")");
            Check(c["C2"].Text == "영업2팀" && c["C3"].Text == "개발1팀", "표에서 제외한 C3은 원본 유지, C2만 반영 (C2=" + c["C2"].Text + ", C3=" + c["C3"].Text + ")");

            Console.WriteLine("\n===== 2단계 검증: 미리보기 취소 =====");
            string cBase = Path.Combine(dir, "cancel_base.xlsx");
            WriteBook(excelType, cBase, "직원", StaffBase());
            byte[] before = File.ReadAllBytes(cBase);
            int backupsBefore = BackupManager.List(cBase).Count;
            o = Options(cBase, pModA);
            o.AutoBackup = true;
            o.Preview = items => false;
            r = Run(o);
            Check(r.Cancelled && r.AppliedCells == 0, "취소하면 결과가 '취소됨', 반영 0칸");
            Check(Equal(before, File.ReadAllBytes(cBase)), "취소하면 원본 파일 바이트가 그대로");
            Check(BackupManager.List(cBase).Count == backupsBefore, "취소하면 이번에 만든 백업도 지워져 백업이 쌓이지 않음");

            Console.WriteLine("\n===== 2단계 검증: 변경 내역 보고서 =====");
            string rBase = Path.Combine(dir, "report_base.xlsx");
            WriteBook(excelType, rBase, "직원", StaffBase());
            string rMod = Path.Combine(dir, "report_mod.xlsx");
            WriteBook(excelType, rMod, "직원", StaffModShuffled());
            o = Options(rBase, rMod);
            o.KeyColumns = "사번";
            o.CreateReport = true;
            r = Run(o);
            bool reportExists = r.ReportPath != null && File.Exists(r.ReportPath);
            Check(reportExists, "보고서 파일 생성: " + (r.ReportPath == null ? "(없음)" : Path.GetFileName(r.ReportPath)));
            if (reportExists)
            {
                List<string> names = SheetNames(excelType, r.ReportPath);
                Check(names.Count == 2 && names[0] == "요약" && names[1] == "변경내역", "보고서 시트 구성: 요약, 변경내역");
                int rows = UsedRows(excelType, r.ReportPath, "변경내역");
                Check(rows == r.Items.Count + 1, "변경내역 행 수 = 항목 " + r.Items.Count + "건 + 머리글 1 (실제 " + rows + ")");
                Dictionary<string, CellInfo> head = ReadCells(excelType, r.ReportPath, "변경내역", "A1", "J1", "J2");
                Check(head["A1"].Text == "순번" && head["J1"].Text == "처리 결과", "변경내역 머리글: 순번 … 처리 결과");
                bool hasApplied = false, hasMissing = false;
                foreach (ChangeItem it in r.Items)
                {
                    if (it.Status == ChangeStatus.Applied) hasApplied = true;
                    if (it.Status == ChangeStatus.MissingInMod) hasMissing = true;
                }
                Check(hasApplied && hasMissing, "보고서 항목에 '반영됨'과 '수정본에 없는 행'이 함께 기록");
                Check(!names.Contains("취합내역") && SheetNames(excelType, rBase).Count == 1, "보고서는 별도 파일 - 원본에는 시트가 추가되지 않음");
            }

            Console.WriteLine("\n===== 2단계 검증: 백업 정리 · 복원 =====");
            string bDir = Path.Combine(dir, "backup");
            if (Directory.Exists(bDir)) Directory.Delete(bDir, true);
            Directory.CreateDirectory(bDir);
            string bBase = Path.Combine(bDir, "월간보고.xlsx");
            WriteBook(excelType, bBase, "직원", StaffBase());
            for (int i = 1; i <= 12; i++)
            {
                File.Copy(bBase, Path.Combine(bDir, string.Format("월간보고_backup_20260101_0000{0:00}.xlsx", i)));
            }
            File.WriteAllText(Path.Combine(bDir, "월간보고_backup_메모.txt"), "백업 형식이 아닌 파일");
            File.Copy(bBase, Path.Combine(bDir, "다른파일_backup_20260101_000001.xlsx"));

            int deleted = BackupManager.Prune(bBase, 10, msg => Console.WriteLine("        [LOG] " + msg));
            List<BackupManager.BackupEntry> left = BackupManager.List(bBase);
            Check(deleted == 2 && left.Count == 10, "백업 12개 → 최근 10개만 남기고 2개 삭제 (삭제 " + deleted + ", 남음 " + left.Count + ")");
            Check(!File.Exists(Path.Combine(bDir, "월간보고_backup_20260101_000001.xlsx")) && !File.Exists(Path.Combine(bDir, "월간보고_backup_20260101_000002.xlsx")),
                "가장 오래된 2개(000001, 000002)가 삭제됨");
            Check(left[0].FileName == "월간보고_backup_20260101_000012.xlsx", "목록 맨 위가 가장 최근 백업");
            Check(File.Exists(Path.Combine(bDir, "다른파일_backup_20260101_000001.xlsx")) && File.Exists(Path.Combine(bDir, "월간보고_backup_메모.txt")),
                "다른 원본의 백업이나 형식이 다른 파일은 건드리지 않음");

            string b1 = ExcelDiffEngine.CreateAutoBackup(bBase, null);
            string b2 = ExcelDiffEngine.CreateAutoBackup(bBase, null);
            Check(b1 != null && b2 != null && b1 != b2 && File.Exists(b1) && File.Exists(b2), "같은 초에 백업 2번 만들어도 서로 덮어쓰지 않음");

            // Restore: base currently = staff; put a different workbook in a backup and restore it
            string otherBook = Path.Combine(dir, "other_content.xlsx");
            WriteBook(excelType, otherBook, "직원", new object[,] { { "사번", "성명", "부서" }, { 9999, "복원확인", "테스트팀" } });
            string chosen = Path.Combine(bDir, "월간보고_backup_20250101_090000.xlsx");
            File.Copy(otherBook, chosen, true);
            byte[] currentBase = File.ReadAllBytes(bBase);
            string safety = BackupManager.Restore(bBase, chosen, msg => Console.WriteLine("        [LOG] " + msg));
            Check(Equal(File.ReadAllBytes(bBase), File.ReadAllBytes(otherBook)), "복원 후 원본 = 선택한 백업 내용");
            Check(safety != null && Equal(File.ReadAllBytes(safety), currentBase), "복원 직전 원본이 새 백업으로 보관됨 (되돌리기 가능)");
            Check(File.Exists(chosen), "복원에 사용한 백업 파일은 그대로 남음");
            c = ReadCells(excelType, bBase, "직원", "B2");
            Check(c["B2"].Text == "복원확인", "복원된 원본을 엑셀로 열어 값 확인 (B2=복원확인)");
        }

        // =====================================================================
        // =====================================================================
        // 6번: 비교 규칙
        public static void RunRules(Type excelType, string root)
        {
            string dir = Path.Combine(root, "rules");
            Directory.CreateDirectory(dir);

            Console.WriteLine("\n===== 6번 검증: 비교 규칙 단위 테스트 =====");
            MergeOptions none = new MergeOptions();
            MergeOptions all = new MergeOptions { IgnoreWhitespace = true, IgnoreCase = true, CompareNumbers = true };
            MergeOptions tol = new MergeOptions { CompareNumbers = true, NumberTolerance = 0.5 };
            Check(!CompareRules.Same("홍 길동", "홍길동", none) && CompareRules.Same("홍 길동", "홍길동", all), "공백 무시: '홍 길동' = '홍길동' (규칙 끄면 다름)");
            Check(CompareRules.Same("a b\n", "ab", all) && CompareRules.Same("서울　시", "서울시", all), "특수 공백(NBSP·전각 공백·줄바꿈)도 무시");
            Check(!CompareRules.Same("ABC", "abc", none) && CompareRules.Same("ABC", "abc", all), "대소문자 무시: 'ABC' = 'abc'");
            Check(!CompareRules.Same("1,000", "1000", none) && CompareRules.Same("1,000", "1000", all), "숫자 값 비교: '1,000' = 1000 (규칙 끄면 다름)");
            Check(CompareRules.Same("10%", "0.1", all) && CompareRules.Same("1.50", "1.5", all) && CompareRules.Same("-1,234.5", "-1234.5", all), "숫자 형식 차이: 10% = 0.1, 1.50 = 1.5, -1,234.5 = -1234.5");
            Check(CompareRules.Same("0.30000000000000004", "0.3", all), "오차 0이어도 부동소수 계산 오차는 같은 값");
            Check(!CompareRules.Same("1001", "1000", all), "오차 0이면 1001 ≠ 1000 (실제 변경은 그대로 잡음)");
            Check(CompareRules.Same("100.4", "100", tol) && CompareRules.Same("99.5", "100", tol) && !CompareRules.Same("100.6", "100", tol), "허용 오차 0.5: 100.4·99.5 = 100, 100.6 ≠ 100");
            Check(!CompareRules.Same("1,5", "15", all), "'1,5'는 천 단위 구분 위치가 틀려 숫자로 보지 않음 → 15와 다름");
            Check(!CompareRules.Same("abc", "abd", all) && !CompareRules.Same("12a", "12", all), "규칙을 다 켜도 내용이 다르면 변경으로 잡음");
            double n;
            Check(!CompareRules.TryNumber("", out n) && !CompareRules.TryNumber("abc", out n) && CompareRules.TryNumber("1E3", out n) && n == 1000, "숫자 판별: 빈값·문자 거부, 1E3 = 1000");
            List<string> unknown = new List<string>();
            Dictionary<string, int> hdr = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            hdr["비고"] = 4; hdr["A"] = 9;
            HashSet<int> cols = CompareRules.ResolveColumns("비고, H, 없는열, a", hdr, unknown);
            Check(cols.Count == 3 && cols.Contains(4) && cols.Contains(8) && cols.Contains(9) && unknown.Count == 1 && unknown[0] == "없는열",
                "제외 열 해석: 머리글 이름 우선(머리글 'A' → 9열), 열 문자 H → 8열, 못 찾은 이름은 따로 알림");

            Console.WriteLine("\n===== 6번 검증: 실제 엑셀 - 규칙 끔 vs 켬 =====");
            object[,] baseCells = {
                { "이름", "금액", "비율", "비고", "코드", "점수" },
                { "홍길동", 1000, 0.1, "확인", "AB-1", 80 },
                { "김철수", 2500, 0.25, null, "cd-2", 90 }
            };
            object[,] modCells = {
                { "이름", "금액", "비율", "비고", "코드", "점수" },
                { "'홍 길동", "'1,000", "'10%", "수정함", "ab-1", 80.3 },
                { "김철수", 2600, 0.25, "메모", "CD-2", 91 }
            };
            string mod = Path.Combine(dir, "rules_mod.xlsx");
            WriteBook(excelType, mod, "점검", modCells);

            string offBase = Path.Combine(dir, "rules_off.xlsx");
            WriteBook(excelType, offBase, "점검", baseCells);
            MergeOptions o = Options(offBase, mod);
            MergeResult r = Run(o);
            Check(r.AppliedCells == 10 && r.RuleSameCells == 0, "규칙을 끄면 표시 차이까지 10칸이 모두 변경으로 반영 (실제 " + r.AppliedCells + "칸)");

            string onBase = Path.Combine(dir, "rules_on.xlsx");
            WriteBook(excelType, onBase, "점검", baseCells);
            o = Options(onBase, mod);
            o.IgnoreWhitespace = true; o.IgnoreCase = true; o.CompareNumbers = true; o.NumberTolerance = 0.5;
            o.ExcludeColumns = "비고";
            o.CreateReport = true;
            r = Run(o);
            Dictionary<string, CellInfo> c = ReadCells(excelType, onBase, "점검", "A2", "B2", "C2", "D2", "E2", "F2", "B3", "D3", "E3", "F3");
            Check(r.AppliedCells == 2, "규칙을 켜면 실제 변경 2칸만 반영 (실제 " + r.AppliedCells + "칸)");
            Check(c["B3"].Text == "2600" && c["F3"].Text == "91", "진짜 변경은 반영: 금액 2500→2600, 점수 90→91 (오차 0.5 초과)");
            Check(c["A2"].Text == "홍길동" && c["E2"].Text == "AB-1" && c["E3"].Text == "cd-2", "공백·대소문자만 다른 칸은 원본 그대로 (A2=" + c["A2"].Text + ", E2=" + c["E2"].Text + ", E3=" + c["E3"].Text + ")");
            Check(c["B2"].Text == "1000" && c["C2"].Text == "0.1" && c["F2"].Text == "80", "숫자 형식·오차 범위 칸은 원본 숫자 그대로 (B2=" + c["B2"].Text + ", C2=" + c["C2"].Text + ", F2=" + c["F2"].Text + ")");
            Check(c["D2"].Text == "확인" && c["D3"].Text == "", "제외 열 '비고'는 수정본 값이 달라도 원본 유지 (D2=" + c["D2"].Text + ", D3='" + c["D3"].Text + "')");
            Check(r.RuleSameCells == 6 && r.ExcludedColumnCells == 2, "결과 집계: 규칙으로 같은 값 6칸, 제외 열 2칸 (실제 " + r.RuleSameCells + ", " + r.ExcludedColumnCells + ")");
            List<string> summaryText = new List<string>();
            if (r.ReportPath != null && File.Exists(r.ReportPath))
            {
                List<string> addrs = new List<string>();
                for (int i = 1; i <= 30; i++) { addrs.Add("A" + i); addrs.Add("B" + i); }
                Dictionary<string, CellInfo> rep = ReadCells(excelType, r.ReportPath, "요약", addrs.ToArray());
                foreach (CellInfo ci in rep.Values) summaryText.Add(ci.Text);
            }
            string joined = string.Join("|", summaryText.ToArray());
            Check(joined.Contains("비교 규칙") && joined.Contains("공백 차이 무시") && joined.Contains("제외 열 [비고]"), "보고서 요약에 비교 규칙이 기록됨");

            Console.WriteLine("\n===== 6번 검증: 제외 열을 열 문자로 지정 + 없는 이름 =====");
            string letterBase = Path.Combine(dir, "rules_letter.xlsx");
            WriteBook(excelType, letterBase, "점검", baseCells);
            o = Options(letterBase, mod);
            o.ExcludeColumns = "D, 없는열";
            List<string> logs = new List<string>();
            r = ExcelDiffEngine.ProcessMerge(o, msg => logs.Add(msg));
            c = ReadCells(excelType, letterBase, "점검", "D2", "B3");
            bool warned = logs.Exists(l => l.Contains("없는열") && l.Contains("찾지 못해"));
            Check(c["D2"].Text == "확인" && c["B3"].Text == "2600" && r.ExcludedColumnCells == 2, "열 문자 D 제외: 비고는 유지, 다른 변경은 반영");
            Check(warned && r.Completed, "없는 제외 열 이름은 경고만 하고 취합은 계속 진행");

            Console.WriteLine("\n===== 6번 검증: 키 열에도 공백 규칙 적용 =====");
            object[,] keyMod = { { "사번", "성명", "부서" }, { 1001, "홍 길동", "영업2팀" } };
            string keyModPath = Path.Combine(dir, "key_space_mod.xlsx");
            WriteBook(excelType, keyModPath, "직원", keyMod);
            string kOff = Path.Combine(dir, "key_space_off.xlsx");
            WriteBook(excelType, kOff, "직원", StaffBase());
            o = Options(kOff, keyModPath);
            o.KeyColumns = "성명";
            r = Run(o);
            int rowsOff = UsedRows(excelType, kOff, "직원");
            string kOn = Path.Combine(dir, "key_space_on.xlsx");
            WriteBook(excelType, kOn, "직원", StaffBase());
            o = Options(kOn, keyModPath);
            o.KeyColumns = "성명";
            o.IgnoreWhitespace = true;
            r = Run(o);
            c = ReadCells(excelType, kOn, "직원", "C2");
            int rowsOn = UsedRows(excelType, kOn, "직원");
            Check(rowsOff == 5, "규칙 끔: '홍 길동'을 다른 사람으로 보고 새 행 추가 (행 수 " + rowsOff + ")");
            Check(rowsOn == 4 && c["C2"].Text == "영업2팀", "공백 무시: '홍 길동' = 홍길동 행으로 찾아 부서만 수정, 새 행 없음 (행 수 " + rowsOn + ", C2=" + c["C2"].Text + ")");
        }

        // Renders the main window and the preview dialog to PNG so the new layout can be inspected
        public static void RenderUi(string outDir)
        {
            Directory.CreateDirectory(outDir);
            using (MainForm form = new MainForm())
            {
                Capture(form, Path.Combine(outDir, "ui_main.png"));
            }

            List<ChangeItem> sample = new List<ChangeItem>();
            string[,] rows = {
                { "A부서.xlsx", "직원", "3", "3", "1002", "부서", "개발1팀", "AI개발팀", "", "u" },
                { "A부서.xlsx", "직원", "5", "1", "1004", "사번", "", "1004", "", "n" },
                { "A부서.xlsx", "직원", "5", "2", "1004", "성명", "", "신규", "", "n" },
                { "B부서.xlsx", "직원", "3", "3", "1002", "부서", "AI개발팀", "QA팀", "충돌 - 마지막 파일 우선으로 A부서.xlsx 값을 덮어씀", "u" }
            };
            for (int i = 0; i < rows.GetLength(0); i++)
            {
                ChangeItem it = new ChangeItem();
                it.Seq = i + 1;
                it.FileName = rows[i, 0];
                it.SheetName = rows[i, 1];
                it.Row = int.Parse(rows[i, 2]);
                it.Col = int.Parse(rows[i, 3]);
                it.RowKey = rows[i, 4];
                it.ColumnHeader = rows[i, 5];
                it.OldText = rows[i, 6];
                it.NewText = rows[i, 7];
                it.Note = rows[i, 8];
                it.Kind = rows[i, 9] == "n" ? ChangeKind.NewRow : ChangeKind.Update;
                sample.Add(it);
            }
            // Base sheet snapshot (1-based like an Excel range array) so the grid tab shows the whole table
            object[,] values = (object[,])Array.CreateInstance(typeof(object), new int[] { 4, 3 }, new int[] { 1, 1 });
            object[,] src = StaffBase();
            for (int r0 = 0; r0 < 4; r0++) for (int c0 = 0; c0 < 3; c0++) values[r0 + 1, c0 + 1] = src[r0, c0];
            PreviewSheet sheet = new PreviewSheet { Name = "직원", Values = values, Rows = 4, Cols = 3, HeaderRow = 1 };
            foreach (ChangeItem it in sample) it.Sheet = sheet;
            using (PreviewDialog dlg = new PreviewDialog(sample))
            {
                Capture(dlg, Path.Combine(outDir, "ui_preview.png"));
            }

            // Wide population-style sheet: 40 rows × 9 columns, changes only from row 12 down
            object[,] big = (object[,])Array.CreateInstance(typeof(object), new int[] { 40, 9 }, new int[] { 1, 1 });
            for (int c = 1; c <= 9; c++) big[1, c] = "항목" + c;
            for (int r = 2; r <= 40; r++) { big[r, 1] = "지역" + r; for (int c = 2; c <= 9; c++) big[r, c] = (double)(r * 100003 + c * 7919); }
            PreviewSheet bigSheet = new PreviewSheet { Name = "연령별인구현황", Values = big, Rows = 40, Cols = 9 };
            List<ChangeItem> bigItems = new List<ChangeItem>();
            for (int r = 12; r <= 30; r++)
            {
                for (int c = 2; c <= 7; c++)
                {
                    ChangeItem it = new ChangeItem { FileName = "연령별인구현황_A.xlsx", SheetName = "연령별인구현황", Row = r, Col = c, OldText = "", NewText = (r * 1000 + c).ToString(), Sheet = bigSheet };
                    bigItems.Add(it);
                }
            }
            using (PreviewDialog dlg = new PreviewDialog(bigItems))
            {
                Capture(dlg, Path.Combine(outDir, "ui_preview_big.png"));
            }
            using (PreviewDialog dlg = new PreviewDialog(bigItems))
            {
                dlg.Shown += (s, e) => dlg.Size = dlg.MinimumSize;
                Capture(dlg, Path.Combine(outDir, "ui_preview_small.png"));
            }
            Console.WriteLine("UI 렌더링 저장: " + outDir);
        }

        // Child controls only paint once the window is really shown, so show it briefly, pump messages, then capture
        private static void Capture(Form form, string file)
        {
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(20, 20);
            form.Show();
            DateTime until = DateTime.Now.AddMilliseconds(800);
            while (DateTime.Now < until)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(20);
            }
            using (Bitmap bmp = new Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(bmp, new Rectangle(0, 0, form.Width, form.Height));
                bmp.Save(file);
            }
            form.Close();
        }

        private static bool Equal(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }
    }
}
