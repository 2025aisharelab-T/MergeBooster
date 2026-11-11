using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Microsoft.CSharp.RuntimeBinder;

namespace ExcelDiffAggregator
{
    public enum ConflictPolicy
    {
        LastWins = 0,    // 마지막 파일 우선 (기본값 - 덮어쓰기)
        FirstWins = 1,   // 첫 번째 파일 우선 (최초 수정값 보존)
        PromptUser = 2   // 사용자 대화상자 선택 (충돌 시 팝업 선택) — 콤보박스 인덱스와 1:1 대응
    }

    public class ExcelDiffEngine
    {
        public delegate void LogHandler(string message);

        // Minimum gap between row-level progress updates sent to the UI
        private const int ProgressIntervalMs = 100;

        // Share of the bar reserved for startup (0..5) and for saving/cleanup (95..100)
        private const double WorkStart = 5.0;
        private const double WorkSpan = 90.0;

        // Merge layout inside 5..95: analysing files 5..80, writing approved cells 80..95
        private const double AnalyzeSpan = 75.0;
        private const double ApplyStart = WorkStart + AnalyzeSpan;
        private const double ApplySpan = WorkSpan - AnalyzeSpan;

        private static void ReportRow(ProgressTracker tracker, double start, double span, int row, int rowCount, string label)
        {
            if (!tracker.Due) return;
            tracker.Report(
                start + span * (row - 1) / rowCount,
                string.Format("{0} · {1:N0}/{2:N0}행", label, row, rowCount),
                false);
        }

        public static string CreateAutoBackup(string basePath, LogHandler logger)
        {
            try
            {
                basePath = System.IO.Path.GetFullPath(basePath);
                if (!System.IO.File.Exists(basePath)) return null;

                string dir = System.IO.Path.GetDirectoryName(basePath);
                string name = System.IO.Path.GetFileNameWithoutExtension(basePath);
                string ext = System.IO.Path.GetExtension(basePath);
                string backupFileName = string.Format("{0}_backup_{1}{2}", name, DateTime.Now.ToString("yyyyMMdd_HHmmss"), ext);
                string backupPath = System.IO.Path.Combine(dir, backupFileName);
                // Two backups in the same second must not overwrite each other
                for (int n = 1; System.IO.File.Exists(backupPath); n++)
                {
                    backupFileName = string.Format("{0}_backup_{1}_{2}{3}", name, DateTime.Now.ToString("yyyyMMdd_HHmmss"), n, ext);
                    backupPath = System.IO.Path.Combine(dir, backupFileName);
                }

                System.IO.File.Copy(basePath, backupPath, false);
                if (logger != null) logger("🛡️ 원본 백업 파일 자동 생성 완료: " + backupFileName);
                return backupPath;
            }
            catch (Exception ex)
            {
                if (logger != null) logger("⚠️ 원본 백업 파일 생성 실패: " + ex.Message);
                return null;
            }
        }

        internal static bool IsFileLocked(string filePath)
        {
            try
            {
                filePath = System.IO.Path.GetFullPath(filePath);
                if (!System.IO.File.Exists(filePath)) return false;
                using (System.IO.FileStream stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch
            {
                return true;
            }
        }

        private static dynamic OpenWorkbook(dynamic excel, string filePath, bool readOnly)
        {
            filePath = System.IO.Path.GetFullPath(filePath);
            if (!System.IO.File.Exists(filePath))
            {
                throw new System.IO.FileNotFoundException("엑셀 파일이 존재하지 않습니다: " + filePath);
            }

            dynamic workbooks = excel.Workbooks;

            try { excel.DisplayAlerts = false; } catch {}
            try { excel.AskToUpdateLinks = false; } catch {}
            try { excel.EnableEvents = false; } catch {}

            try
            {
                return workbooks.Open(filePath, 0, readOnly);
            }
            catch (Exception ex1)
            {
                try
                {
                    return workbooks.Open(filePath);
                }
                catch
                {
                    try
                    {
                        return workbooks.Open(filePath, 0, true);
                    }
                    catch
                    {
                        string msg = string.Format(
                            "엑셀 파일({0})을 열 수 없습니다.\r\n\r\n" +
                            "■ 원인 및 해결 방법:\r\n" +
                            "1. 엑셀 프로그램에서 해당 파일이 열려 있다면 닫고 다시 시도해 주세요.\r\n" +
                            "2. 파일 우클릭 -> [속성] -> 하단 '차단 해제(Unblock)'가 있다면 체크 후 [확인].\r\n" +
                            "3. 다른 프로그램이 파일을 사용 중이거나 암호가 설정되어 있는지 확인해 주세요.\r\n\r\n" +
                            "■ 상세 예외: {1}",
                            System.IO.Path.GetFileName(filePath),
                            ex1.Message
                        );
                        throw new Exception(msg, ex1);
                    }
                }
            }
        }

        private static void GetSheetBounds(dynamic ws, out int maxRow, out int maxCol)
        {
            maxRow = 1;
            maxCol = 1;
            if (ws == null) return;

            try
            {
                dynamic usedRange = ws.UsedRange;
                if ((object)usedRange != null)
                {
                    int rStart = Convert.ToInt32(usedRange.Row);
                    int cStart = Convert.ToInt32(usedRange.Column);
                    int rCount = Convert.ToInt32(usedRange.Rows.Count);
                    int cCount = Convert.ToInt32(usedRange.Columns.Count);

                    maxRow = Math.Max(1, rStart + rCount - 1);
                    maxCol = Math.Max(1, cStart + cCount - 1);

                    try { Marshal.ReleaseComObject((object)usedRange); } catch {}
                }
            }
            catch
            {
                maxRow = 1;
                maxCol = 1;
            }
        }

        private static object[,] GetRangeValueArray(dynamic ws, int rMax, int cMax)
        {
            if (rMax < 1 || cMax < 1) return new object[2, 2];

            dynamic rng = null;
            object val = null;
            try
            {
                rng = ws.Range[ws.Cells[1, 1], ws.Cells[rMax, cMax]];
                val = rng.Value2;
            }
            catch
            {
                return new object[rMax + 1, cMax + 1];
            }
            finally
            {
                if ((object)rng != null) { try { Marshal.ReleaseComObject((object)rng); } catch {} }
            }

            if (val is object[,])
            {
                return (object[,])val;
            }

            object[,] arr = new object[rMax + 1, cMax + 1];
            arr[1, 1] = val;
            return arr;
        }

        // Kept for existing callers and tests: position matching, formula protection on, no preview, no report
        public static int ProcessDiffMultiple(
            string basePath,
            string[] modPaths,
            string sheetName,
            bool applyHighlight,
            bool addComment,
            int highlightColorOle,
            bool autoBackup,
            ConflictPolicy conflictPolicy,
            LogHandler logger,
            ProgressHandler progress = null
        )
        {
            MergeOptions opt = new MergeOptions();
            opt.BasePath = basePath;
            opt.ModPaths = modPaths;
            opt.SheetName = sheetName;
            opt.ApplyHighlight = applyHighlight;
            opt.AddComment = addComment;
            opt.HighlightColorOle = highlightColorOle;
            opt.AutoBackup = autoBackup;
            opt.ConflictPolicy = conflictPolicy;
            return ProcessMerge(opt, logger, progress).AppliedCells;
        }

        public static MergeResult ProcessMerge(MergeOptions opt, LogHandler logger, ProgressHandler progress = null)
        {
            MergeResult result = new MergeResult();
            Exception threadEx = null;
            ProgressTracker tracker = new ProgressTracker(progress, ProgressIntervalMs, null);

            System.Threading.Thread staThread = new System.Threading.Thread(() =>
            {
                try
                {
                    ProcessMergeInternal(opt, result, logger, tracker);
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            });

            staThread.SetApartmentState(System.Threading.ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            if (threadEx != null)
            {
                DebugLog.Error("취합 스레드 예외", threadEx);
                string errStr = threadEx.ToString();
                bool disconnected = errStr.Contains("0x80010114") || errStr.Contains("RPC_E_DISCONNECTED") || errStr.Contains("요청한 개체가 없습니다");
                // Only a disconnect after the save finished is harmless; earlier it means the merge did not happen
                if (disconnected && result.Completed)
                {
                    if (logger != null) logger("  ℹ️ COM 최종 정리 중 0x80010114 예외가 발생했으나, 파일 취합 및 디스크 저장은 완료되었습니다.");
                    tracker.Report(100, "취합 완료", true);
                    return result;
                }
                throw threadEx;
            }
            return result;
        }

        private static bool IsAutoAllSheets(string sheetName)
        {
            return string.IsNullOrEmpty(sheetName) ||
                   sheetName == "(전체 시트 자동 인식)" ||
                   sheetName == "(전체 시트)" ||
                   sheetName == "(모든 시트)";
        }

        private static List<string> ResolveTargetSheets(dynamic wbBase, string sheetName)
        {
            List<string> names = new List<string>();
            if (IsAutoAllSheets(sheetName))
            {
                dynamic sheets = wbBase.Worksheets;
                int count = sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic ws = sheets[i];
                    names.Add((string)ws.Name);
                    try { Marshal.ReleaseComObject((object)ws); } catch {}
                }
                try { Marshal.ReleaseComObject((object)sheets); } catch {}
            }
            else if (sheetName == "(첫번째 시트)")
            {
                dynamic ws1 = wbBase.Worksheets[1];
                names.Add((string)ws1.Name);
                try { Marshal.ReleaseComObject((object)ws1); } catch {}
            }
            else
            {
                names.Add(sheetName);
            }
            return names;
        }

        private static string PolicyText(ConflictPolicy policy)
        {
            switch (policy)
            {
                case ConflictPolicy.LastWins: return "마지막 파일 우선";
                case ConflictPolicy.FirstWins: return "첫 번째 파일 우선";
                case ConflictPolicy.PromptUser: return "사용자 선택";
            }
            return policy.ToString();
        }

        private static void ProcessMergeInternal(MergeOptions opt, MergeResult result, LogHandler logger, ProgressTracker tracker)
        {
            string basePath = System.IO.Path.GetFullPath(opt.BasePath);
            string[] modPaths = opt.ModPaths ?? new string[0];
            if (!System.IO.File.Exists(basePath))
            {
                throw new System.IO.FileNotFoundException("원본 파일이 존재하지 않습니다: " + basePath);
            }

            if (IsFileLocked(basePath))
            {
                throw new Exception(string.Format("원본 엑셀 파일({0})이 현재 Excel이나 다른 프로그램에서 열려 있습니다.\n\n해결 방법: 엑셀 프로그램에서 원본 파일을 닫으신 후 다시 실행해 주세요.", System.IO.Path.GetFileName(basePath)));
            }

            tracker.Report(0, "준비 중…", true);

            // 1. Auto Backup Snapshot (+ keep only the newest N)
            if (opt.AutoBackup)
            {
                tracker.Report(0, "원본 백업 중…", true);
                result.BackupPath = CreateAutoBackup(basePath, logger);
                if (result.BackupPath != null && opt.BackupKeepCount > 0)
                {
                    BackupManager.Prune(basePath, opt.BackupKeepCount, logger);
                }
            }

            if (logger != null) logger("엑셀 애플리케이션 시작 중...");
            tracker.Report(1, "엑셀을 시작하는 중…", true);
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new Exception("Microsoft Excel이 이 컴퓨터에 설치되어 있지 않습니다.");
            }

            dynamic excel = Activator.CreateInstance(excelType);
            try { excel.Visible = false; } catch {}
            try { excel.DisplayAlerts = false; } catch {}
            try { excel.ScreenUpdating = false; } catch {}
            try { excel.EnableEvents = false; } catch {}
            try { excel.AskToUpdateLinks = false; } catch {}
            try { excel.AutomationSecurity = 3; } catch {}
            try { excel.FeatureInstall = 0; } catch {}
            try { excel.Calculation = -4135; } catch {}

            dynamic wbBase = null;
            bool isFileSuccessfullySaved = false;
            Dictionary<string, SheetState> states = new Dictionary<string, SheetState>(StringComparer.OrdinalIgnoreCase);
            ConflictResolutionChoice? userRememberedChoice = null;
            int seq = 0;

            try
            {
                if (logger != null) logger("원본 파일 열기: " + System.IO.Path.GetFileName(basePath));
                tracker.Report(3, "원본 파일을 여는 중…", true);
                try
                {
                    wbBase = OpenWorkbook(excel, basePath, false);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("원본 엑셀 파일({0})을 열 수 없습니다. 파일 경로, 다른 프로그램 사용 여부, 보안/암호 보호 여부를 확인하세요.\n오류 상세: {1}", System.IO.Path.GetFileName(basePath), ex.Message), ex);
                }

                bool isBaseReadOnly = false;
                try { isBaseReadOnly = (bool)wbBase.ReadOnly; } catch {}
                if (isBaseReadOnly)
                {
                    throw new Exception(string.Format("원본 엑셀 파일({0})이 읽기 전용 모드로 열렸습니다.\n엑셀에서 해당 파일이 열려 있는지 확인하고 닫은 후 다시 시도하세요.", System.IO.Path.GetFileName(basePath)));
                }

                if (logger != null)
                {
                    logger(opt.UseKeyMatching
                        ? string.Format("🔑 키 열 기준 행 매칭: [{0}] (머리글 {1}행) · 열은 머리글 이름으로 맞춤", opt.KeyColumns.Trim(), opt.HeaderRow)
                        : "📍 같은 위치(행·열)끼리 비교");
                    if (opt.ProtectFormulas) logger("🧮 수식 칸 보호 사용: 원본에 수식이 있는 칸은 바꾸지 않습니다.");
                    if (opt.HasCompareRules) logger("📏 비교 규칙: " + CompareRules.Describe(opt));
                }
                DebugLog.Write("DEBUG", "비교 규칙: " + CompareRules.Describe(opt));
                DebugLog.Write("DEBUG", string.Format("병합 옵션: 키='{0}' 머리글행={1} 수식보호={2} 미리보기={3} 보고서={4} 정책={5}",
                    opt.KeyColumns, opt.HeaderRow, opt.ProtectFormulas, opt.Preview != null, opt.CreateReport, opt.ConflictPolicy));

                List<string> targetSheetNames = ResolveTargetSheets(wbBase, opt.SheetName);
                bool isAutoAllSheets = IsAutoAllSheets(opt.SheetName);

                // ===== STEP 1: analyse every file in memory (nothing is written to the base yet) =====
                for (int idx = 0; idx < modPaths.Length; idx++)
                {
                    string modPath = System.IO.Path.GetFullPath(modPaths[idx]);
                    string modFileName = System.IO.Path.GetFileName(modPath);

                    if (!System.IO.File.Exists(modPath))
                    {
                        if (logger != null) logger(string.Format("  ❌ 수정본 파일이 존재하지 않음: {0}", modFileName));
                        continue;
                    }

                    if (logger != null) logger(string.Format("[{0}/{1}] 수정본 비교 시작: {2}", idx + 1, modPaths.Length, modFileName));

                    // Each modified file owns an equal slice of the analysis part of the bar
                    double fileStart = WorkStart + AnalyzeSpan * idx / modPaths.Length;
                    double fileSpan = AnalyzeSpan / modPaths.Length;
                    string fileLabel = string.Format("파일 {0}/{1} · {2}", idx + 1, modPaths.Length, modFileName);
                    tracker.Report(fileStart, fileLabel + " · 여는 중…", true);

                    dynamic wbMod = null;
                    try
                    {
                        try
                        {
                            wbMod = OpenWorkbook(excel, modPath, true);
                        }
                        catch (Exception ex)
                        {
                            if (logger != null) logger(string.Format("  ❌ 수정본 파일({0}) 열기 실패: {1}", modFileName, ex.Message));
                            continue;
                        }

                        int filePlanned = 0;
                        int sheetNo = 0;

                        foreach (string sName in targetSheetNames)
                        {
                            // The first 10% of a file's slice is opening it, the rest is split evenly between its sheets
                            double sheetStart = fileStart + fileSpan * (0.1 + 0.9 * sheetNo / targetSheetNames.Count);
                            double sheetEnd = fileStart + fileSpan * (0.1 + 0.9 * (sheetNo + 1) / targetSheetNames.Count);
                            double sheetSpan = sheetEnd - sheetStart;
                            string sheetLabel = string.Format("{0} · 시트 [{1}]", fileLabel, sName);
                            sheetNo++;

                            if (logger != null) logger(string.Format("  ├ 시트 [{0}] 비교 시작...", sName));
                            tracker.Report(sheetStart, sheetLabel + " · 읽는 중…", true);

                            SheetState st = GetSheetState(states, wbBase, sName, opt, logger);
                            if (st == null) continue;

                            dynamic wsMod = null;
                            try
                            {
                                wsMod = wbMod.Worksheets[sName];
                            }
                            catch
                            {
                                if (logger != null && isAutoAllSheets)
                                {
                                    logger(string.Format("    ⚠️ 수정본({0})에 시트 [{1}]가 존재하지 않아 스킵합니다.", modFileName, sName));
                                }
                                continue;
                            }

                            object[,] modArr;
                            try
                            {
                                int rMaxMod, cMaxMod;
                                GetSheetBounds(wsMod, out rMaxMod, out cMaxMod);
                                modArr = GetRangeValueArray(wsMod, rMaxMod, cMaxMod);
                            }
                            finally
                            {
                                try { Marshal.ReleaseComObject((object)wsMod); } catch {}
                            }

                            int sheetPlanned = AnalyzeSheet(st, modArr, modFileName, opt, result.Items, ref seq, ref userRememberedChoice,
                                logger, tracker, sheetStart, sheetSpan, sheetLabel);
                            filePlanned += sheetPlanned;

                            tracker.Report(sheetEnd, sheetLabel + " · 분석 완료", true);
                            if (logger != null && isAutoAllSheets)
                            {
                                logger(string.Format("    ├ 시트 [{0}]: {1}개 칸 변경 예정", sName, sheetPlanned));
                            }
                        }

                        if (logger != null) logger(string.Format("  └ {0}: 총 {1}개 칸 변경 예정", modFileName, filePlanned));
                    }
                    finally
                    {
                        if ((object)wbMod != null)
                        {
                            try { wbMod.Close(false); } catch (Exception exCloseMod) { if (logger != null) logger("    ⚠️ wbMod.Close 예외: " + exCloseMod.Message); }
                            try { Marshal.ReleaseComObject((object)wbMod); } catch {}
                            wbMod = null;
                        }
                    }
                }

                List<ChangeItem> planned = new List<ChangeItem>();
                foreach (ChangeItem it in result.Items)
                {
                    if (it.Status == ChangeStatus.Planned) planned.Add(it);
                }
                foreach (SheetState sst in states.Values)
                {
                    if (sst == null) continue;
                    result.RuleSameCells += sst.RuleSame;
                    result.ExcludedColumnCells += sst.ExcludedSkipped;
                }
                LogPlanSummary(result, logger);
                if (logger != null && (result.RuleSameCells > 0 || result.ExcludedColumnCells > 0))
                {
                    logger(string.Format("📏 비교 규칙으로 같은 값 처리 {0}칸 · 제외 열이라 건너뜀 {1}칸", result.RuleSameCells, result.ExcludedColumnCells));
                }

                // ===== STEP 2: preview (user picks what to apply) =====
                if (planned.Count > 0 && opt.Preview != null)
                {
                    if (logger != null) logger("🔍 미리보기 창에서 반영할 항목을 확인해 주세요...");
                    tracker.Report(WorkStart + AnalyzeSpan, "미리보기에서 반영할 항목을 고르는 중…", true);
                    Dictionary<string, PreviewSheet> snapshots = new Dictionary<string, PreviewSheet>(StringComparer.OrdinalIgnoreCase);
                    foreach (ChangeItem it in planned)
                    {
                        SheetState sst;
                        if (!states.TryGetValue(it.SheetName, out sst) || sst == null) continue;
                        PreviewSheet ps;
                        if (!snapshots.TryGetValue(it.SheetName, out ps))
                        {
                            ps = new PreviewSheet
                            {
                                Name = sst.Name,
                                Values = sst.BaseArr,
                                Rows = sst.BaseRows,
                                Cols = sst.BaseCols,
                                HeaderRow = opt.UseKeyMatching ? opt.HeaderRow : 0
                            };
                            snapshots[it.SheetName] = ps;
                        }
                        it.Sheet = ps;
                    }
                    bool approved;
                    try
                    {
                        approved = opt.Preview(planned);
                    }
                    catch (Exception exPreview)
                    {
                        DebugLog.Error("미리보기 창 오류", exPreview);
                        throw new Exception("미리보기 창에서 오류가 발생해 취합을 중단했습니다. 원본은 바뀌지 않았습니다.\n" + exPreview.Message, exPreview);
                    }

                    if (!approved)
                    {
                        foreach (ChangeItem it in planned)
                        {
                            it.Status = ChangeStatus.Excluded;
                            it.Note = "미리보기에서 취소";
                        }
                        result.Cancelled = true;
                        DebugLog.Write("DEBUG", "미리보기에서 취소 → 원본 저장 안 함");
                        if (logger != null) logger("⏹️ 미리보기에서 취소했습니다. 원본은 바뀌지 않았습니다.");

                        try { wbBase.Close(false); } catch {}
                        try { Marshal.ReleaseComObject((object)wbBase); } catch {}
                        wbBase = null;

                        // Nothing changed, so this run's backup is just a duplicate of the original
                        if (result.BackupPath != null)
                        {
                            try
                            {
                                System.IO.File.Delete(result.BackupPath);
                                if (logger != null) logger("  └ 변경이 없어 이번에 만든 백업(" + System.IO.Path.GetFileName(result.BackupPath) + ")을 삭제했습니다.");
                                result.BackupPath = null;
                            }
                            catch (Exception exDel)
                            {
                                DebugLog.Error("취소 후 백업 삭제 실패", exDel);
                            }
                        }
                        result.Completed = true;
                        tracker.Report(100, "취소됨 · 원본 변경 없음", true);
                        return;
                    }
                }

                // ===== STEP 3: apply approved items =====
                foreach (ChangeItem it in planned)
                {
                    if (!it.Approved)
                    {
                        it.Status = ChangeStatus.Excluded;
                        it.Note = "미리보기에서 제외";
                    }
                }

                // A cell may have approved changes from several files; the latest approved one wins
                Dictionary<string, ChangeItem> finalByCell = new Dictionary<string, ChangeItem>();
                List<string> cellOrder = new List<string>();
                foreach (ChangeItem it in result.Items)
                {
                    if (it.Status != ChangeStatus.Planned) continue;
                    ChangeItem prev;
                    if (finalByCell.TryGetValue(it.CellKey, out prev))
                    {
                        prev.Status = ChangeStatus.Superseded;
                        prev.Note = string.Format("같은 칸에 나중 파일({0}) 값이 반영됨", it.FileName);
                    }
                    else
                    {
                        cellOrder.Add(it.CellKey);
                    }
                    finalByCell[it.CellKey] = it;
                }

                int total = cellOrder.Count;
                int done = 0;
                HashSet<string> newRowIds = new HashSet<string>();
                if (logger != null) logger(string.Format("[STEP 2] 승인된 {0}개 칸을 원본에 반영합니다...", total));
                tracker.Report(ApplyStart, string.Format("원본에 반영 중 · 총 {0:N0}칸", total), true);

                foreach (string key in cellOrder)
                {
                    ChangeItem it = finalByCell[key];
                    SheetState st = states[it.SheetName];
                    dynamic cell = null;
                    try
                    {
                        if (st.Ws == null)
                        {
                            st.Ws = wbBase.Worksheets[st.Name];
                            try { st.Ws.Unprotect(); } catch {}
                        }
                        cell = st.Ws.Cells[it.Row, it.Col];
                        cell.Value2 = it.NewValue;

                        if (opt.ApplyHighlight)
                        {
                            try { cell.Interior.Color = opt.HighlightColorOle; } catch {}
                        }

                        if (opt.AddComment)
                        {
                            try
                            {
                                try { cell.ClearComments(); } catch {}
                                string original = st.OriginalText(it.Row, it.Col);
                                string commentText = string.Format("[{0}] 변경(수정값): {1} (기존: {2})\n수정일시: {3}",
                                    it.FileName,
                                    it.NewText,
                                    string.IsNullOrEmpty(original) ? "(빈값)" : original,
                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                                cell.AddComment(commentText);
                            }
                            catch { }
                        }

                        it.Status = ChangeStatus.Applied;
                        result.AppliedCells++;
                        if (it.Kind == ChangeKind.NewRow) newRowIds.Add(it.SheetName + "!" + it.Row);
                    }
                    catch (Exception exCell)
                    {
                        it.Status = ChangeStatus.Failed;
                        it.Note = exCell.Message;
                        if (logger != null) logger(string.Format("    ⚠️ 셀 [{0}!{1}] 변경 예외: {2}", it.SheetName, it.Address, exCell.Message));
                        DebugLog.Error(string.Format("셀 변경 예외 [{0}] 파일={1}", it.CellKey, it.FileName), exCell);
                    }
                    finally
                    {
                        if ((object)cell != null)
                        {
                            try { Marshal.ReleaseComObject((object)cell); } catch {}
                        }
                    }

                    done++;
                    if (tracker.Due)
                    {
                        // No "a/b" pattern here: the progress tests look for "1/2" to find a file's slice
                        tracker.Report(ApplyStart + ApplySpan * done / Math.Max(1, total),
                            string.Format("원본에 반영 중 · {0:N0}칸 / 총 {1:N0}칸", done, total), false);
                    }
                }
                result.NewRows = newRowIds.Count;

                foreach (SheetState st in states.Values)
                {
                    if (st != null && (object)st.Ws != null)
                    {
                        try { Marshal.ReleaseComObject((object)st.Ws); } catch {}
                        st.Ws = null;
                    }
                }

                if (logger != null)
                {
                    logger(string.Format("  └ {0}개 칸 반영 완료{1}", result.AppliedCells,
                        result.NewRows > 0 ? string.Format(" (새 행 {0}개 추가)", result.NewRows) : ""));
                }

                // ===== STEP 4: save =====
                tracker.Report(WorkStart + WorkSpan, "취합 결과를 원본에 저장하는 중…", true);
                try { excel.DisplayAlerts = false; } catch {}

                string tempCopyPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(basePath), "temp_agg_save_" + Guid.NewGuid().ToString("N") + ".xlsx");
                bool savedViaCopy = false;

                if (result.AppliedCells == 0)
                {
                    if (logger != null) logger("[STEP 3] 반영된 칸이 없어 원본을 저장하지 않습니다.");
                    isFileSuccessfullySaved = true;
                }
                else
                {
                    if (logger != null) logger("[STEP 3] 모든 취합 결과 원본 저장 처리 중...");
                    try
                    {
                        if (logger != null) logger("  [STEP 3-1] 원본 직접 저장(wbBase.Save) 시도...");
                        wbBase.Save();
                        isFileSuccessfullySaved = true;
                        if (logger != null) logger("  └ 원본 직접 저장(Save) 성공!");
                    }
                    catch (Exception exDirect)
                    {
                        if (logger != null) logger("  ⚠️ 원본 직접 저장 실패(" + exDirect.Message + "), SaveCopyAs 우회 저장 시도 중...");
                        try
                        {
                            wbBase.SaveCopyAs(tempCopyPath);
                            savedViaCopy = true;
                            isFileSuccessfullySaved = true;
                            if (logger != null) logger("  └ SaveCopyAs 임시 파일 저장 성공!");
                        }
                        catch (Exception exCopy)
                        {
                            if (logger != null) logger("  ❌ SaveCopyAs 우회 저장 실패: " + exCopy.Message);
                        }
                    }
                }

                if (logger != null) logger("[STEP 4] 원본 워크북(wbBase) 닫기 시도...");
                try
                {
                    wbBase.Close(false);
                    if (logger != null) logger("  └ wbBase.Close 완료.");
                }
                catch (Exception exCloseBase)
                {
                    if (logger != null) logger("  ⚠️ wbBase.Close 실패: " + exCloseBase.Message);
                }
                finally
                {
                    if ((object)wbBase != null) { try { Marshal.ReleaseComObject((object)wbBase); } catch {} wbBase = null; }
                }

                // ===== STEP 5: change report (a failed report never fails the merge) =====
                if (opt.CreateReport)
                {
                    tracker.Report(WorkStart + WorkSpan + 1, "변경 내역 보고서를 만드는 중…", true);
                    try
                    {
                        result.ReportPath = WriteReport(excel, basePath, opt, result, logger);
                    }
                    catch (Exception exReport)
                    {
                        if (logger != null) logger("  ⚠️ 변경 내역 보고서 생성 실패: " + exReport.Message);
                        DebugLog.Error("보고서 생성 실패", exReport);
                    }
                }

                if (logger != null) logger("[STEP 5] Excel 백그라운드 애플리케이션 정리 중...");
                if ((object)excel != null)
                {
                    try { excel.Quit(); } catch {}
                    try { Marshal.ReleaseComObject((object)excel); } catch {}
                    excel = null;
                }

                if (savedViaCopy && System.IO.File.Exists(tempCopyPath))
                {
                    if (logger != null) logger("[STEP 6] 임시 파일 원본 덮어쓰기(File.Copy) 시도...");
                    try
                    {
                        System.IO.File.Copy(tempCopyPath, basePath, true);
                        System.IO.File.Delete(tempCopyPath);
                        if (logger != null) logger("  └ 임시 파일로 원본 덮어쓰기 최종 반영 성공!");
                    }
                    catch (Exception exOver)
                    {
                        if (logger != null) logger("  ❌ 원본 파일 덮어쓰기 실패: " + exOver.Message);
                        throw new Exception("취합 결과는 임시 파일(" + tempCopyPath + ")에 보존되었습니다.\n원본 파일 덮어쓰기 실패: " + exOver.Message);
                    }
                }

                if (!isFileSuccessfullySaved)
                {
                    throw new Exception("원본 파일 저장에 실패했습니다. 로그를 확인해 주세요.");
                }

                result.Completed = true;
                if (logger != null) logger("🎉 모든 엑셀 취합 및 저장 정상 완료!");
                tracker.Report(100, "취합 완료", true);
            }
            catch (Exception exMain)
            {
                if (isFileSuccessfullySaved && !exMain.Message.StartsWith("취합 결과는 임시 파일"))
                {
                    if (logger != null) logger("  ℹ️ 저장 후 정리 단계에서 예외 발생(" + exMain.Message + "), 하지만 파일 취합 및 디스크 저장은 이미 성공했습니다.");
                    DebugLog.Error("저장 후 정리 단계 예외", exMain);
                    result.Completed = true;
                    tracker.Report(100, "취합 완료", true);
                    return;
                }
                throw;
            }
            finally
            {
                foreach (SheetState st in states.Values)
                {
                    if (st != null && (object)st.Ws != null) { try { Marshal.ReleaseComObject((object)st.Ws); } catch {} st.Ws = null; }
                }

                try
                {
                    if ((object)wbBase != null) { wbBase.Close(false); }
                }
                catch { }
                try
                {
                    if ((object)wbBase != null) { Marshal.ReleaseComObject((object)wbBase); }
                }
                catch { }
                wbBase = null;

                try
                {
                    if ((object)excel != null) { excel.Quit(); }
                }
                catch { }

                try
                {
                    if ((object)excel != null) { Marshal.ReleaseComObject((object)excel); }
                }
                catch { }
                excel = null;
            }
        }

        // Reads a base sheet once per run (values, formula mask, header and key index). Null = sheet missing or unusable.
        private static SheetState GetSheetState(Dictionary<string, SheetState> states, dynamic wbBase, string sName, MergeOptions opt, LogHandler logger)
        {
            SheetState st;
            if (states.TryGetValue(sName, out st)) return st;

            dynamic ws = null;
            try
            {
                ws = wbBase.Worksheets[sName];
            }
            catch (Exception exSheet)
            {
                if (logger != null) logger(string.Format("    ⚠️ 원본에 시트 [{0}]가 존재하지 않아 스킵합니다.", sName));
                DebugLog.Error(string.Format("원본 시트 접근 실패 [{0}]", sName), exSheet);
                states[sName] = null;
                return null;
            }

            try
            {
                st = new SheetState(sName);
                int rMax, cMax;
                GetSheetBounds(ws, out rMax, out cMax);
                st.BaseRows = rMax;
                st.BaseCols = cMax;
                st.BaseArr = GetRangeValueArray(ws, rMax, cMax);
                st.NextNewRow = rMax + 1;

                if (opt.ProtectFormulas)
                {
                    st.Formula = GetFormulaMask(ws, rMax, cMax, out st.FormulaCount);
                    if (st.FormulaCount > 0 && logger != null)
                    {
                        logger(string.Format("    🧮 시트 [{0}]: 수식 칸 {1}개 보호", sName, st.FormulaCount));
                    }
                }

                if (opt.UseKeyMatching)
                {
                    st.HeaderByCol = ReadHeaderRow(st.BaseArr, opt.HeaderRow, cMax);
                    st.BaseHeader = HeaderMap(st.HeaderByCol);
                    st.BaseKeyCols = ResolveKeyColumns(opt.KeyColumns, st.BaseHeader);
                    if (st.BaseKeyCols == null)
                    {
                        st.KeyError = true;
                        if (logger != null) logger(string.Format("    ❌ 원본 시트 [{0}]의 {1}행에서 키 열 '{2}'을(를) 찾지 못해 이 시트는 건너뜁니다.", sName, opt.HeaderRow, opt.KeyColumns.Trim()));
                    }
                    else
                    {
                        BuildKeyIndex(st, opt, logger);
                    }
                }

                if (opt.ExcludeColumns != null && opt.ExcludeColumns.Trim().Length > 0)
                {
                    Dictionary<string, int> header = opt.UseKeyMatching ? st.BaseHeader : HeaderMap(ReadHeaderRow(st.BaseArr, opt.HeaderRow, cMax));
                    List<string> unknown = new List<string>();
                    st.ExcludedCols = CompareRules.ResolveColumns(opt.ExcludeColumns, header, unknown);
                    if (logger != null)
                    {
                        if (st.ExcludedCols.Count > 0)
                        {
                            List<string> names = new List<string>();
                            List<int> sorted = new List<int>(st.ExcludedCols);
                            sorted.Sort();
                            foreach (int c in sorted)
                            {
                                string h = (c <= cMax) ? SheetState.Text(SheetState.Get(st.BaseArr, opt.HeaderRow, c)) : "";
                                names.Add(h.Length > 0 ? ChangeItem.ColumnLetter(c) + "(" + h + ")" : ChangeItem.ColumnLetter(c));
                            }
                            logger(string.Format("    🚫 시트 [{0}]: 비교 제외 열 {1}", sName, string.Join(", ", names.ToArray())));
                        }
                        if (unknown.Count > 0)
                        {
                            logger(string.Format("    ⚠️ 시트 [{0}]의 {1}행에서 제외 열 [{2}]을(를) 찾지 못해 무시합니다.", sName, opt.HeaderRow, string.Join(", ", unknown.ToArray())));
                        }
                    }
                    DebugLog.Write("DEBUG", string.Format("제외 열 [{0}] '{1}' → {2}개, 못 찾음 {3}개", sName, opt.ExcludeColumns, st.ExcludedCols.Count, unknown.Count));
                }

                states[sName] = st;
                return st;
            }
            finally
            {
                if ((object)ws != null) { try { Marshal.ReleaseComObject((object)ws); } catch {} }
            }
        }

        private static bool[,] GetFormulaMask(dynamic ws, int rMax, int cMax, out int count)
        {
            count = 0;
            bool[,] mask = new bool[rMax + 1, cMax + 1];
            if (rMax < 1 || cMax < 1) return mask;

            dynamic rng = null;
            object formulas = null;
            try
            {
                rng = ws.Range[ws.Cells[1, 1], ws.Cells[rMax, cMax]];
                formulas = rng.Formula;
            }
            catch (Exception ex)
            {
                DebugLog.Error("수식 정보 읽기 실패 (수식 보호 없이 진행)", ex);
                return mask;
            }
            finally
            {
                if ((object)rng != null) { try { Marshal.ReleaseComObject((object)rng); } catch {} }
            }

            object[,] arr = formulas as object[,];
            if (arr == null)
            {
                string single = formulas as string;
                if (single != null && single.StartsWith("="))
                {
                    mask[1, 1] = true;
                    count = 1;
                }
                return mask;
            }

            int lb0 = arr.GetLowerBound(0), lb1 = arr.GetLowerBound(1);
            for (int i = lb0; i <= arr.GetUpperBound(0); i++)
            {
                for (int j = lb1; j <= arr.GetUpperBound(1); j++)
                {
                    string s = arr[i, j] as string;
                    if (s != null && s.StartsWith("="))
                    {
                        int r = i - lb0 + 1, c = j - lb1 + 1;
                        if (r <= rMax && c <= cMax)
                        {
                            mask[r, c] = true;
                            count++;
                        }
                    }
                }
            }
            return mask;
        }

        private static string[] ReadHeaderRow(object[,] arr, int headerRow, int cols)
        {
            string[] headers = new string[cols + 1];
            for (int c = 1; c <= cols; c++)
            {
                headers[c] = SheetState.Text(SheetState.Get(arr, headerRow, c));
            }
            return headers;
        }

        private static Dictionary<string, int> HeaderMap(string[] headers)
        {
            Dictionary<string, int> map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c < headers.Length; c++)
            {
                string h = headers[c];
                if (!string.IsNullOrEmpty(h) && !map.ContainsKey(h)) map[h] = c;
            }
            return map;
        }

        // "사번" / "A" / "시군구,읍면동" → column numbers. Header names win over column letters. Null when any token is unknown.
        internal static int[] ResolveKeyColumns(string spec, Dictionary<string, int> headerMap)
        {
            if (spec == null) return null;
            string[] tokens = spec.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<int> cols = new List<int>();
            foreach (string raw in tokens)
            {
                string t = raw.Trim();
                if (t.Length == 0) continue;
                int col;
                if (headerMap.TryGetValue(t, out col))
                {
                    cols.Add(col);
                    continue;
                }
                col = ChangeItem.ColumnIndex(t);
                if (col <= 0) return null;
                cols.Add(col);
            }
            return cols.Count > 0 ? cols.ToArray() : null;
        }

        private static string BuildRowKey(object[,] arr, int row, int[] keyCols, MergeOptions opt)
        {
            bool any = false;
            string[] parts = new string[keyCols.Length];
            for (int i = 0; i < keyCols.Length; i++)
            {
                parts[i] = CompareRules.KeyText(SheetState.Text(SheetState.Get(arr, row, keyCols[i])), opt);
                if (parts[i].Length > 0) any = true;
            }
            return any ? string.Join("\u001F", parts) : "";
        }

        private static string DisplayKey(string key)
        {
            return key.Replace("\u001F", " / ");
        }

        private static void BuildKeyIndex(SheetState st, MergeOptions opt, LogHandler logger)
        {
            int headerRow = opt.HeaderRow;
            int duplicates = 0;
            string firstDuplicate = null;
            for (int r = headerRow + 1; r <= st.BaseRows; r++)
            {
                string key = BuildRowKey(st.BaseArr, r, st.BaseKeyCols, opt);
                if (key.Length == 0) continue;
                if (st.KeyIndex.ContainsKey(key))
                {
                    duplicates++;
                    if (firstDuplicate == null) firstDuplicate = DisplayKey(key);
                    continue;
                }
                st.KeyIndex[key] = r;
            }
            if (duplicates > 0 && logger != null)
            {
                logger(string.Format("    ⚠️ 원본 시트 [{0}]에 같은 키가 {1}번 더 나옵니다 (예: '{2}'). 처음 나온 행만 사용합니다.", st.Name, duplicates, firstDuplicate));
            }
        }

        // Compares one modified sheet with the current in-memory state of the base sheet and records ChangeItems
        private static int AnalyzeSheet(
            SheetState st,
            object[,] modArr,
            string modFileName,
            MergeOptions opt,
            List<ChangeItem> items,
            ref int seq,
            ref ConflictResolutionChoice? userRememberedChoice,
            LogHandler logger,
            ProgressTracker tracker,
            double sheetStart,
            double sheetSpan,
            string sheetLabel)
        {
            if (st.KeyError) return 0;

            int mRows = modArr.GetUpperBound(0);
            int mCols = modArr.GetUpperBound(1);
            int[] colMap = new int[mCols + 1];
            int[] modKeyCols = null;
            int firstRow = 1;
            bool keyMode = opt.UseKeyMatching;

            if (keyMode)
            {
                firstRow = opt.HeaderRow + 1;
                string[] modHeaders = ReadHeaderRow(modArr, opt.HeaderRow, mCols);
                Dictionary<string, int> modHeaderMap = HeaderMap(modHeaders);
                modKeyCols = ResolveKeyColumns(opt.KeyColumns, modHeaderMap);
                if (modKeyCols == null)
                {
                    if (logger != null) logger(string.Format("    ❌ 수정본({0}) 시트 [{1}]의 {2}행에서 키 열 '{3}'을(를) 찾지 못해 건너뜁니다.", modFileName, st.Name, opt.HeaderRow, opt.KeyColumns.Trim()));
                    return 0;
                }

                List<string> unmatched = new List<string>();
                for (int c = 1; c <= mCols; c++)
                {
                    string h = modHeaders[c];
                    int baseCol;
                    if (string.IsNullOrEmpty(h)) colMap[c] = c;                       // 머리글 없는 열은 같은 위치
                    else if (st.BaseHeader.TryGetValue(h, out baseCol)) colMap[c] = baseCol;
                    else { colMap[c] = 0; unmatched.Add(h); }
                }
                if (unmatched.Count > 0 && logger != null)
                {
                    logger(string.Format("    ⚠️ 원본에 없는 열 [{0}]은(는) 비교하지 않습니다.", string.Join(", ", unmatched.ToArray())));
                }
            }
            else
            {
                for (int c = 1; c <= mCols; c++) colMap[c] = c;
            }

            int planned = 0;
            int emptyKeyRows = 0;
            HashSet<string> seenKeys = new HashSet<string>(StringComparer.Ordinal);

            for (int r = firstRow; r <= mRows; r++)
            {
                ReportRow(tracker, sheetStart, sheetSpan, r, mRows, sheetLabel);

                int baseRow = r;
                string rowKey = "";
                bool isNewRow = false;

                if (keyMode)
                {
                    string key = BuildRowKey(modArr, r, modKeyCols, opt);
                    if (key.Length == 0)
                    {
                        // A row with no key value at all is blank or a note row; count only rows that hold data
                        for (int c = 1; c <= mCols; c++)
                        {
                            if (SheetState.Text(SheetState.Get(modArr, r, c)).Length > 0) { emptyKeyRows++; break; }
                        }
                        continue;
                    }
                    seenKeys.Add(key);
                    rowKey = DisplayKey(key);
                    if (!st.KeyIndex.TryGetValue(key, out baseRow))
                    {
                        baseRow = st.NextNewRow++;
                        st.KeyIndex[key] = baseRow;
                        st.NewRows.Add(baseRow);
                        DebugLog.Write("DEBUG", string.Format("새 행 배정 [{0}] 키='{1}' → {2}행 (파일 {3})", st.Name, rowKey, baseRow, modFileName));
                    }
                    isNewRow = st.NewRows.Contains(baseRow);
                }

                for (int c = 1; c <= mCols; c++)
                {
                    int bc = colMap[c];
                    if (bc <= 0) continue;

                    object oMod = SheetState.Get(modArr, r, c);
                    string vMod = SheetState.Text(oMod);
                    if (vMod.Length == 0) continue;           // 빈 칸은 반영하지 않음 (기존 동작)

                    string vCur = st.CurrentText(baseRow, bc);
                    if (vMod == vCur) continue;

                    if (st.ExcludedCols.Contains(bc))
                    {
                        st.ExcludedSkipped++;
                        continue;
                    }
                    if (CompareRules.Same(vMod, vCur, opt))
                    {
                        st.RuleSame++;
                        DebugLog.Write("DEBUG", string.Format("비교 규칙으로 같은 값 [{0}!{1}{2}] 파일={3} 기존='{4}' 수정='{5}'",
                            st.Name, ChangeItem.ColumnLetter(bc), baseRow, modFileName, vCur, vMod));
                        continue;
                    }

                    ChangeItem item = new ChangeItem();
                    item.Seq = ++seq;
                    item.FileName = modFileName;
                    item.SheetName = st.Name;
                    item.Row = baseRow;
                    item.Col = bc;
                    item.RowKey = rowKey;
                    item.ColumnHeader = keyMode ? st.HeaderText(bc) : "";
                    item.Kind = isNewRow ? ChangeKind.NewRow : ChangeKind.Update;
                    item.OldText = vCur;
                    item.NewText = vMod;
                    item.NewValue = oMod;

                    if (opt.ProtectFormulas && st.IsFormula(baseRow, bc))
                    {
                        item.Status = ChangeStatus.FormulaProtected;
                        item.Note = "원본 칸에 수식이 있어 바꾸지 않음";
                        items.Add(item);
                        DebugLog.Write("DEBUG", string.Format("수식 보호 [{0}!{1}] 파일={2} 수정값='{3}' 무시", st.Name, item.Address, modFileName, vMod));
                        continue;
                    }

                    long cellId = SheetState.CellId(baseRow, bc);
                    string prevFile;
                    if (st.TouchedBy.TryGetValue(cellId, out prevFile))
                    {
                        string cellKey = item.CellKey;
                        DebugLog.Write("DEBUG", string.Format("충돌 감지 [{0}] 정책={1} 파일={2} 기존='{3}' 수정='{4}' 먼저 반영한 파일={5}",
                            cellKey, opt.ConflictPolicy, modFileName, vCur, vMod, prevFile));

                        if (opt.ConflictPolicy == ConflictPolicy.FirstWins)
                        {
                            item.Status = ChangeStatus.ConflictKept;
                            item.Note = string.Format("첫 번째 파일 우선 - {0} 값 유지", prevFile);
                            items.Add(item);
                            if (logger != null) logger(string.Format("    ℹ️ [충돌 스킵] [{0}!{1}] 기존 값({2}) 보존 (수정본: {3})", st.Name, item.Address, vCur, vMod));
                            continue;
                        }
                        else if (opt.ConflictPolicy == ConflictPolicy.PromptUser)
                        {
                            ConflictResolutionChoice choiceToApply;
                            if (userRememberedChoice.HasValue)
                            {
                                choiceToApply = userRememberedChoice.Value;
                            }
                            else
                            {
                                tracker.Report(tracker.Current, sheetLabel + " · 충돌 선택 대기 중…", true);
                                using (ConflictDialog dlg = new ConflictDialog(st.Name, baseRow, bc, modFileName, vCur, vMod))
                                {
                                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                    {
                                        choiceToApply = dlg.Choice;
                                        if (dlg.ApplyToAll) userRememberedChoice = dlg.Choice;
                                    }
                                    else
                                    {
                                        // Closing the dialog (X / Esc) must not silently overwrite data: keep the existing value
                                        choiceToApply = ConflictResolutionChoice.KeepExistingValue;
                                        DebugLog.Write("DEBUG", string.Format("충돌 대화상자 닫힘 [{0}] → 기존 값 유지", cellKey));
                                    }
                                }
                            }
                            DebugLog.Write("DEBUG", string.Format("충돌 처리 [{0}] 사용자 선택={1} (전체적용 기억={2})", cellKey, choiceToApply, userRememberedChoice.HasValue));

                            if (choiceToApply == ConflictResolutionChoice.KeepExistingValue)
                            {
                                item.Status = ChangeStatus.ConflictKept;
                                item.Note = string.Format("사용자 선택 - {0} 값 유지", prevFile);
                                items.Add(item);
                                if (logger != null) logger(string.Format("    ℹ️ [사용자 선택] [{0}!{1}] 기존 값({2}) 보존", st.Name, item.Address, vCur));
                                continue;
                            }
                            item.Note = string.Format("충돌 - 사용자 선택으로 {0} 값을 덮어씀", prevFile);
                        }
                        else
                        {
                            item.Note = string.Format("충돌 - 마지막 파일 우선으로 {0} 값을 덮어씀", prevFile);
                            DebugLog.Write("DEBUG", string.Format("충돌 처리 [{0}] 마지막 파일 우선 → '{1}' 로 덮어씀", cellKey, vMod));
                        }
                    }

                    item.Status = ChangeStatus.Planned;
                    st.Overlay[cellId] = oMod;
                    st.TouchedBy[cellId] = modFileName;
                    items.Add(item);
                    planned++;
                }
            }

            if (keyMode)
            {
                if (emptyKeyRows > 0 && logger != null)
                {
                    logger(string.Format("    ⚠️ 키 열이 비어 있는 데이터 행 {0}개는 비교하지 않았습니다.", emptyKeyRows));
                }

                // Rows in the base that this file does not contain: information only, nothing is deleted
                int missing = 0;
                foreach (KeyValuePair<string, int> kv in st.KeyIndex)
                {
                    if (st.NewRows.Contains(kv.Value) || seenKeys.Contains(kv.Key)) continue;
                    ChangeItem info = new ChangeItem();
                    info.Seq = ++seq;
                    info.FileName = modFileName;
                    info.SheetName = st.Name;
                    info.Row = kv.Value;
                    info.Col = 0;
                    info.RowKey = DisplayKey(kv.Key);
                    info.Status = ChangeStatus.MissingInMod;
                    info.Note = "원본에는 있지만 이 수정본에는 없는 행 (삭제되지 않음)";
                    items.Add(info);
                    missing++;
                }
                if (missing > 0 && logger != null)
                {
                    logger(string.Format("    ℹ️ 원본에는 있지만 수정본에 없는 행 {0}개 (삭제하지 않음, 보고서에서 확인)", missing));
                }
            }

            return planned;
        }

        private static void LogPlanSummary(MergeResult result, LogHandler logger)
        {
            int planned = 0, newRowCells = 0, conflictKept = 0, formula = 0, missing = 0;
            HashSet<string> newRows = new HashSet<string>();
            foreach (ChangeItem it in result.Items)
            {
                switch (it.Status)
                {
                    case ChangeStatus.Planned:
                        planned++;
                        if (it.Kind == ChangeKind.NewRow) { newRowCells++; newRows.Add(it.SheetName + "!" + it.Row); }
                        break;
                    case ChangeStatus.ConflictKept: conflictKept++; break;
                    case ChangeStatus.FormulaProtected: formula++; break;
                    case ChangeStatus.MissingInMod: missing++; break;
                }
            }
            string msg = string.Format("📋 분석 결과: 변경 예정 {0}칸 (새 행 {1}개) · 충돌로 기존 값 유지 {2}칸 · 수식 보호 {3}칸",
                planned, newRows.Count, conflictKept, formula);
            if (missing > 0) msg += string.Format(" · 수정본에 없는 행 {0}건", missing);
            if (logger != null) logger(msg);
            DebugLog.Write("DEBUG", msg);
        }

        private static string WriteReport(dynamic excel, string basePath, MergeOptions opt, MergeResult result, LogHandler logger)
        {
            string dir = System.IO.Path.GetDirectoryName(basePath);
            string name = System.IO.Path.GetFileNameWithoutExtension(basePath);
            string path = System.IO.Path.Combine(dir, string.Format("{0}_취합내역_{1}.xlsx", name, DateTime.Now.ToString("yyyyMMdd_HHmmss")));

            dynamic wb = null;
            dynamic wsDetail = null;
            dynamic wsSummary = null;
            try
            {
                wb = excel.Workbooks.Add();
                wsDetail = wb.Worksheets[1];
                wsDetail.Name = "변경내역";
                wsSummary = wb.Worksheets.Add(wsDetail);   // inserted before → becomes the first sheet
                wsSummary.Name = "요약";

                List<object[]> summary = new List<object[]>();
                summary.Add(new object[] { "항목", "내용" });
                summary.Add(new object[] { "원본 파일", basePath });
                summary.Add(new object[] { "실행 시각", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") });
                string[] mods = opt.ModPaths ?? new string[0];
                for (int i = 0; i < mods.Length; i++)
                {
                    summary.Add(new object[] { string.Format("수정본 {0}", i + 1), mods[i] });
                }
                summary.Add(new object[] { "비교 방식", opt.UseKeyMatching ? string.Format("키 열 [{0}] · 머리글 {1}행", opt.KeyColumns.Trim(), opt.HeaderRow) : "같은 위치(행·열)" });
                summary.Add(new object[] { "충돌 처리 규칙", PolicyText(opt.ConflictPolicy) });
                summary.Add(new object[] { "수식 칸 보호", opt.ProtectFormulas ? "사용" : "사용 안 함" });
                summary.Add(new object[] { "비교 규칙", CompareRules.Describe(opt) });
                summary.Add(new object[] { "비교 규칙으로 같은 값 처리", result.RuleSameCells.ToString() });
                summary.Add(new object[] { "제외 열이라 건너뜀", result.ExcludedColumnCells.ToString() });
                summary.Add(new object[] { "반영된 칸", result.AppliedCells.ToString() });
                summary.Add(new object[] { "새로 추가된 행", result.NewRows.ToString() });
                summary.Add(new object[] { "미리보기에서 제외", result.Count(ChangeStatus.Excluded).ToString() });
                summary.Add(new object[] { "나중 파일 값으로 대체", result.Count(ChangeStatus.Superseded).ToString() });
                summary.Add(new object[] { "충돌로 기존 값 유지", result.Count(ChangeStatus.ConflictKept).ToString() });
                summary.Add(new object[] { "수식 보호로 건너뜀", result.Count(ChangeStatus.FormulaProtected).ToString() });
                summary.Add(new object[] { "수정본에 없는 행", result.Count(ChangeStatus.MissingInMod).ToString() });
                summary.Add(new object[] { "반영 실패", result.Count(ChangeStatus.Failed).ToString() });
                summary.Add(new object[] { "백업 파일", result.BackupPath ?? "(백업 안 함)" });
                WriteTable(wsSummary, summary);

                List<object[]> detail = new List<object[]>();
                detail.Add(new object[] { "순번", "수정본 파일", "시트", "셀", "행 키", "열", "구분", "기존 값", "새 값", "처리 결과", "비고" });
                foreach (ChangeItem it in result.Items)
                {
                    detail.Add(new object[] {
                        it.Seq.ToString(), it.FileName, it.SheetName, it.Address, it.RowKey, it.ColumnLabel,
                        it.Col > 0 ? it.KindText : "", it.OldText, it.NewText, it.StatusText, it.Note
                    });
                }
                WriteTable(wsDetail, detail);

                wb.SaveAs(path, 51);
                if (logger != null) logger("📄 변경 내역 보고서 저장: " + System.IO.Path.GetFileName(path));
                DebugLog.Write("DEBUG", "보고서 저장: " + path + " (항목 " + result.Items.Count + "건)");
                return path;
            }
            finally
            {
                if ((object)wsSummary != null) { try { Marshal.ReleaseComObject((object)wsSummary); } catch {} }
                if ((object)wsDetail != null) { try { Marshal.ReleaseComObject((object)wsDetail); } catch {} }
                if ((object)wb != null)
                {
                    try { wb.Close(false); } catch {}
                    try { Marshal.ReleaseComObject((object)wb); } catch {}
                }
            }
        }

        // Writes rows as text in one call (text format keeps values like "001" intact), bold header, sensible widths
        private static void WriteTable(dynamic ws, List<object[]> rows)
        {
            int rowCount = rows.Count;
            int colCount = rows[0].Length;
            object[,] data = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    data[r, c] = c < rows[r].Length ? rows[r][c] : null;
                }
            }

            dynamic rng = ws.Range[ws.Cells[1, 1], ws.Cells[rowCount, colCount]];
            dynamic header = ws.Range[ws.Cells[1, 1], ws.Cells[1, colCount]];
            try
            {
                rng.NumberFormat = "@";
                rng.Value2 = data;
                header.Font.Bold = true;
                header.Interior.Color = 15921906; // #F2F2F2
                try { rng.AutoFilter(1); } catch {}
                rng.Columns.AutoFit();
                for (int c = 1; c <= colCount; c++)
                {
                    dynamic col = ws.Columns[c];
                    try
                    {
                        double w = Convert.ToDouble(col.ColumnWidth);
                        if (w > 60) col.ColumnWidth = 60;
                    }
                    catch {}
                    finally { try { Marshal.ReleaseComObject((object)col); } catch {} }
                }
            }
            finally
            {
                try { Marshal.ReleaseComObject((object)header); } catch {}
                try { Marshal.ReleaseComObject((object)rng); } catch {}
            }
        }

        public static int ResetHighlightsAndComments(
            string basePath,
            string sheetName,
            LogHandler logger,
            ProgressHandler progress = null
        )
        {
            int totalResetCount = 0;
            Exception threadEx = null;
            ProgressTracker tracker = new ProgressTracker(progress, ProgressIntervalMs, null);

            System.Threading.Thread staThread = new System.Threading.Thread(() =>
            {
                try
                {
                    totalResetCount = ResetHighlightsAndCommentsInternal(basePath, sheetName, logger, tracker);
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            });

            staThread.SetApartmentState(System.Threading.ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            if (threadEx != null)
            {
                DebugLog.Error("취합 스레드 예외", threadEx);
                string errStr = threadEx.ToString();
                if (errStr.Contains("0x80010114") || errStr.Contains("RPC_E_DISCONNECTED") || errStr.Contains("요청한 개체가 없습니다"))
                {
                    if (logger != null) logger("  ℹ️ COM 정리 중 무해한 예외가 발생했으나, 파일의 메모 및 음영 초기화는 100% 완료되었습니다.");
                    tracker.Report(100, "초기화 완료", true);
                    return totalResetCount;
                }
                throw threadEx;
            }
            return totalResetCount;
        }

        private static int ResetHighlightsAndCommentsInternal(
            string basePath,
            string sheetName,
            LogHandler logger,
            ProgressTracker tracker
        )
        {
            basePath = System.IO.Path.GetFullPath(basePath);
            if (!System.IO.File.Exists(basePath))
            {
                throw new System.IO.FileNotFoundException("원본 파일이 존재하지 않습니다: " + basePath);
            }

            if (IsFileLocked(basePath))
            {
                throw new Exception(string.Format("원본 엑셀 파일({0})이 현재 Excel이나 다른 프로그램에서 열려 있습니다.\n\n해결 방법: 엑셀 프로그램에서 원본 파일을 닫으신 후 다시 실행해 주세요.", System.IO.Path.GetFileName(basePath)));
            }

            tracker.Report(0, "준비 중…", true);
            if (logger != null) logger("엑셀 애플리케이션 시작 중 (메모 및 음영 초기화)...");
            tracker.Report(1, "엑셀을 시작하는 중…", true);
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new Exception("Microsoft Excel이 이 컴퓨터에 설치되어 있지 않습니다.");
            }

            dynamic excel = Activator.CreateInstance(excelType);
            try { excel.Visible = false; } catch {}
            try { excel.DisplayAlerts = false; } catch {}
            try { excel.ScreenUpdating = false; } catch {}
            try { excel.EnableEvents = false; } catch {}
            try { excel.AskToUpdateLinks = false; } catch {}
            try { excel.AutomationSecurity = 3; } catch {}
            try { excel.FeatureInstall = 0; } catch {}

            int resetSheetsCount = 0;
            dynamic wbBase = null;
            bool isFileSaved = false;

            try
            {
                if (logger != null) logger("원본 파일 열기: " + System.IO.Path.GetFileName(basePath));
                tracker.Report(3, "원본 파일을 여는 중…", true);
                wbBase = OpenWorkbook(excel, basePath, false);

                bool isAutoAllSheets = string.IsNullOrEmpty(sheetName) || 
                                      sheetName == "(전체 시트 자동 인식)" || 
                                      sheetName == "(전체 시트)" || 
                                      sheetName == "(모든 시트)";
                bool isFirstSheetOnly = sheetName == "(첫번째 시트)";

                List<string> targetSheetNames = new List<string>();

                if (isAutoAllSheets)
                {
                    dynamic sheets = wbBase.Worksheets;
                    int count = sheets.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic ws = sheets[i];
                        targetSheetNames.Add(ws.Name);
                        try { Marshal.ReleaseComObject((object)ws); } catch {}
                    }
                    try { Marshal.ReleaseComObject((object)sheets); } catch {}
                }
                else if (isFirstSheetOnly)
                {
                    dynamic ws1 = wbBase.Worksheets[1];
                    targetSheetNames.Add(ws1.Name);
                    try { Marshal.ReleaseComObject((object)ws1); } catch {}
                }
                else
                {
                    targetSheetNames.Add(sheetName);
                }

                int sheetNo = 0;
                foreach (string sName in targetSheetNames)
                {
                    // Equal slice of the bar per sheet: 90% scanning cells, 10% removing comments
                    double sheetStart = WorkStart + WorkSpan * sheetNo / targetSheetNames.Count;
                    double sheetSpan = WorkSpan / targetSheetNames.Count;
                    string sheetLabel = string.Format("시트 [{0}] ({1}/{2})", sName, sheetNo + 1, targetSheetNames.Count);
                    sheetNo++;

                    dynamic wsBase = null;
                    try
                    {
                        wsBase = wbBase.Worksheets[sName];
                    }
                    catch
                    {
                        continue;
                    }

                    try { wsBase.Unprotect(); } catch {}
                    tracker.Report(sheetStart, sheetLabel + " · 음영 확인 중…", true);

                    dynamic usedRange = null;
                    try
                    {
                        usedRange = wsBase.UsedRange;
                        if ((object)usedRange != null)
                        {
                            HashSet<long> targetPastelColors = new HashSet<long>
                            {
                                13431551L,
                                14348258L,
                                16247517L,
                                14083324L,
                                15786728L
                            };
                            int removedHighlightCount = 0;
                            int removedCommentCount = 0;

                            try
                            {
                                int startRow = usedRange.Row;
                                int startCol = usedRange.Column;
                                int rMax = startRow + usedRange.Rows.Count - 1;
                                int cMax = startCol + usedRange.Columns.Count - 1;

                                for (int r = startRow; r <= rMax; r++)
                                {
                                    ReportRow(tracker, sheetStart, sheetSpan * 0.9, r - startRow + 1, rMax - startRow + 1, sheetLabel);

                                    for (int c = startCol; c <= cMax; c++)
                                    {
                                        dynamic cell = null;
                                        try
                                        {
                                            cell = wsBase.Cells[r, c];
                                            dynamic interior = cell.Interior;
                                            object colorIdxObj = interior.ColorIndex;
                                            int colorIdx = 0;
                                            if (colorIdxObj != null && int.TryParse(colorIdxObj.ToString(), out colorIdx) && colorIdx != -4142)
                                            {
                                                object colorObj = interior.Color;
                                                if (colorObj != null)
                                                {
                                                    long colorVal = Convert.ToInt64(colorObj);
                                                    if (targetPastelColors.Contains(colorVal))
                                                    {
                                                        interior.ColorIndex = -4142;
                                                        removedHighlightCount++;
                                                    }
                                                }
                                            }
                                            try { Marshal.ReleaseComObject((object)interior); } catch { }
                                        }
                                        catch { }
                                        finally
                                        {
                                            if ((object)cell != null) { try { Marshal.ReleaseComObject((object)cell); } catch { } }
                                        }
                                    }
                                }
                            }
                            catch (Exception exFind)
                            {
                                if (logger != null) logger(string.Format("  ⚠️ 파스텔 음영 제거 예외: {0}", exFind.Message));
                            }

                            tracker.Report(sheetStart + sheetSpan * 0.9, sheetLabel + " · 메모 정리 중…", true);
                            try
                            {
                                dynamic comments = wsBase.Comments;
                                if ((object)comments != null)
                                {
                                    int commentCount = comments.Count;
                                    for (int i = commentCount; i >= 1; i--)
                                    {
                                        dynamic comment = null;
                                        try
                                        {
                                            comment = comments.Item(i);
                                            string cText = comment.Text();
                                            if (cText != null && (cText.Contains("변경(수정값):") || cText.Contains("기존 값:")))
                                            {
                                                comment.Delete();
                                                removedCommentCount++;
                                            }
                                        }
                                        catch { }
                                        finally
                                        {
                                            if ((object)comment != null) { try { Marshal.ReleaseComObject((object)comment); } catch { } }
                                        }
                                    }
                                    try { Marshal.ReleaseComObject((object)comments); } catch { }
                                }
                            }
                            catch { }

                            resetSheetsCount++;
                            tracker.Report(sheetStart + sheetSpan, sheetLabel + " · 완료", true);
                            if (logger != null) logger(string.Format("  ├ 시트 [{0}]: 취합 음영({1}개) 및 메모({2}개) 선택 제거 완료 (원본 서식 유지)", sName, removedHighlightCount, removedCommentCount));
                        }
                    }
                    catch (Exception ex)
                    {
                        if (logger != null) logger(string.Format("  ⚠️ 시트 [{0}] 초기화 중 예외: {1}", sName, ex.Message));
                    }
                    finally
                    {
                        if ((object)usedRange != null) { try { Marshal.ReleaseComObject((object)usedRange); } catch {} }
                        if ((object)wsBase != null) { try { Marshal.ReleaseComObject((object)wsBase); } catch {} }
                    }
                }

                if (logger != null) logger("초기화 완료된 파일 저장 중...");
                tracker.Report(WorkStart + WorkSpan, "변경 내용을 원본에 저장하는 중…", true);
                try
                {
                    wbBase.Save();
                    isFileSaved = true;
                    if (logger != null) logger("  └ 원본 파일 저장 완료!");
                }
                catch (Exception exSave)
                {
                    if (logger != null) logger("  ❌ 원본 파일 저장 실패: " + exSave.Message);
                }

                try { wbBase.Close(false); } catch {}
                try { if ((object)wbBase != null) Marshal.ReleaseComObject((object)wbBase); } catch {}
                wbBase = null;

                try { excel.Quit(); } catch {}
                try { if ((object)excel != null) Marshal.ReleaseComObject((object)excel); } catch {}
                excel = null;

                if (logger != null) logger("🎉 모든 시트의 메모 및 음영이 성공적으로 제거되었습니다!");
                tracker.Report(100, "초기화 완료", true);
                return resetSheetsCount;
            }
            catch (Exception exMain)
            {
                if (isFileSaved)
                {
                    if (logger != null) logger("  ℹ️ 저장 후 정리 단계 예외, 하지만 초기화 저장은 완료되었습니다.");
                    tracker.Report(100, "초기화 완료", true);
                    return resetSheetsCount;
                }
                throw exMain;
            }
            finally
            {
                if ((object)wbBase != null) { try { Marshal.ReleaseComObject((object)wbBase); } catch {} wbBase = null; }
                if ((object)excel != null) { try { excel.Quit(); } catch {} try { Marshal.ReleaseComObject((object)excel); } catch {} excel = null; }
                try { GC.Collect(); } catch {}
                try { GC.WaitForPendingFinalizers(); } catch {}
            }
        }
    }
}
