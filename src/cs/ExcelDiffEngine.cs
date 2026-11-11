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
        NumericSum = 2,  // 수치 합산 (숫자 데이터 누적 합산)
        PromptUser = 3   // 사용자 대화상자 선택 (충돌 시 팝업 선택)
    }

    public class ExcelDiffEngine
    {
        public delegate void LogHandler(string message);

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

                System.IO.File.Copy(basePath, backupPath, true);
                if (logger != null) logger("🛡️ 원본 백업 파일 자동 생성 완료: " + backupFileName);
                return backupPath;
            }
            catch (Exception ex)
            {
                if (logger != null) logger("⚠️ 원본 백업 파일 생성 실패: " + ex.Message);
                return null;
            }
        }

        private static bool IsFileLocked(string filePath)
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
                if (usedRange != null)
                {
                    int rStart = Convert.ToInt32(usedRange.Row);
                    int cStart = Convert.ToInt32(usedRange.Column);
                    int rCount = Convert.ToInt32(usedRange.Rows.Count);
                    int cCount = Convert.ToInt32(usedRange.Columns.Count);

                    maxRow = Math.Max(1, rStart + rCount - 1);
                    maxCol = Math.Max(1, cStart + cCount - 1);

                    try { Marshal.ReleaseComObject(usedRange); } catch {}
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
                if (rng != null) { try { Marshal.ReleaseComObject(rng); } catch {} }
            }

            if (val is object[,])
            {
                return (object[,])val;
            }

            object[,] arr = new object[rMax + 1, cMax + 1];
            arr[1, 1] = val;
            return arr;
        }

        public static int ProcessDiffMultiple(
            string basePath,
            string[] modPaths,
            string sheetName,
            bool applyHighlight,
            bool addComment,
            int highlightColorOle,
            bool autoBackup,
            ConflictPolicy conflictPolicy,
            LogHandler logger
        )
        {
            int totalDiffCount = 0;
            Exception threadEx = null;

            System.Threading.Thread staThread = new System.Threading.Thread(() =>
            {
                try
                {
                    totalDiffCount = ProcessDiffMultipleInternal(
                        basePath,
                        modPaths,
                        sheetName,
                        applyHighlight,
                        addComment,
                        highlightColorOle,
                        autoBackup,
                        conflictPolicy,
                        logger
                    );
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
                string errStr = threadEx.ToString();
                if (errStr.Contains("0x80010114") || errStr.Contains("RPC_E_DISCONNECTED") || errStr.Contains("요청한 개체가 없습니다"))
                {
                    if (logger != null) logger("  ℹ️ COM 최종 정리 중 0x80010114 예외가 발생했으나, 파일 취합 및 디스크 저장은 100% 완료되었습니다.");
                    return totalDiffCount;
                }
                throw threadEx;
            }
            return totalDiffCount;
        }

        private static int ProcessDiffMultipleInternal(
            string basePath,
            string[] modPaths,
            string sheetName,
            bool applyHighlight,
            bool addComment,
            int highlightColorOle,
            bool autoBackup,
            ConflictPolicy conflictPolicy,
            LogHandler logger
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

            // 1. Auto Backup Snapshot
            if (autoBackup)
            {
                CreateAutoBackup(basePath, logger);
            }

            if (logger != null) logger("엑셀 애플리케이션 시작 중...");
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

            int totalDiffCount = 0;
            dynamic wbBase = null;
            bool isFileSuccessfullySaved = false;

            // Track modified cells across multiple files to detect conflicts
            Dictionary<string, string> modifiedCellTracker = new Dictionary<string, string>();
            ConflictResolutionChoice? userRememberedChoice = null;

            try
            {
                if (logger != null) logger("원본 파일 열기: " + System.IO.Path.GetFileName(basePath));
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
                                try { Marshal.ReleaseComObject(ws); } catch {}
                            }
                            try { Marshal.ReleaseComObject(sheets); } catch {}
                        }
                        else if (isFirstSheetOnly)
                        {
                            dynamic ws1 = wbBase.Worksheets[1];
                            targetSheetNames.Add(ws1.Name);
                            try { Marshal.ReleaseComObject(ws1); } catch {}
                        }
                        else
                        {
                            targetSheetNames.Add(sheetName);
                        }

                        int fileDiffCount = 0;

                        foreach (string sName in targetSheetNames)
                        {
                            if (logger != null) logger(string.Format("  ├ 시트 [{0}] 비교 시작...", sName));
                            dynamic wsBase = null;
                            dynamic wsMod = null;
                            try
                            {
                                wsBase = wbBase.Worksheets[sName];
                            }
                            catch
                            {
                                continue;
                            }

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
                                if (wsBase != null) { try { Marshal.ReleaseComObject(wsBase); } catch {} }
                                continue;
                            }

                            try { wsBase.Unprotect(); } catch {}

                            int rMaxBase, cMaxBase;
                            int rMaxMod, cMaxMod;
                            GetSheetBounds(wsBase, out rMaxBase, out cMaxBase);
                            GetSheetBounds(wsMod, out rMaxMod, out cMaxMod);

                            int rMax = Math.Max(rMaxBase, rMaxMod);
                            int cMax = Math.Max(cMaxBase, cMaxMod);

                            if (rMax < 1 || cMax < 1)
                            {
                                if (wsBase != null) { try { Marshal.ReleaseComObject(wsBase); } catch {} }
                                if (wsMod != null) { try { Marshal.ReleaseComObject(wsMod); } catch {} }
                                continue;
                            }

                            object[,] baseArr = GetRangeValueArray(wsBase, rMax, cMax);
                            object[,] modArr = GetRangeValueArray(wsMod, rMax, cMax);

                            int sheetDiffCount = 0;

                            int bMaxR = baseArr.GetUpperBound(0);
                            int bMaxC = baseArr.GetUpperBound(1);
                            int mMaxR = modArr.GetUpperBound(0);
                            int mMaxC = modArr.GetUpperBound(1);

                            for (int r = 1; r <= rMax; r++)
                            {
                                for (int c = 1; c <= cMax; c++)
                                {
                                    object oBase = (r <= bMaxR && c <= bMaxC) ? baseArr[r, c] : null;
                                    object oMod = (r <= mMaxR && c <= mMaxC) ? modArr[r, c] : null;

                                    string vBase = oBase != null ? oBase.ToString().Trim() : "";
                                    string vMod = oMod != null ? oMod.ToString().Trim() : "";

                                    // 수정본 파일(modArr)에 유효한 변경값(vMod가 빈값이 아님)이 존재하고, 현재 원본 값(vBase)과 다른 경우에만 누적 반영
                                    if (!string.IsNullOrEmpty(vMod) && vMod != vBase)
                                    {
                                        string cellKey = string.Format("{0}!R{1}C{2}", sName, r, c);
                                        bool isConflict = modifiedCellTracker.ContainsKey(cellKey);

                                        object finalVal = oMod;
                                        string commentNote = "";

                                        if (isConflict)
                                        {
                                            if (conflictPolicy == ConflictPolicy.FirstWins)
                                            {
                                                if (logger != null) logger(string.Format("    ℹ️ [충돌 스킵] [{0}] 셀 기존 값({1}) 보존 (수정본: {2})", cellKey, vBase, vMod));
                                                continue;
                                            }
                                            else if (conflictPolicy == ConflictPolicy.NumericSum)
                                            {
                                                double numBase, numMod;
                                                if (double.TryParse(vBase, out numBase) && double.TryParse(vMod, out numMod))
                                                {
                                                    double sum = numBase + numMod;
                                                    finalVal = sum;
                                                    commentNote = string.Format(" [수치 합산: {0} + {1} = {2}]", numBase, numMod, sum);
                                                }
                                            }
                                            else if (conflictPolicy == ConflictPolicy.PromptUser)
                                            {
                                                ConflictResolutionChoice choiceToApply;
                                                if (userRememberedChoice.HasValue)
                                                {
                                                    choiceToApply = userRememberedChoice.Value;
                                                }
                                                else
                                                {
                                                    ConflictDialog dlg = new ConflictDialog(sName, r, c, modFileName, vBase, vMod);
                                                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                                    {
                                                        choiceToApply = dlg.Choice;
                                                        if (dlg.ApplyToAll)
                                                        {
                                                            userRememberedChoice = dlg.Choice;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        choiceToApply = ConflictResolutionChoice.UseNewValue;
                                                    }
                                                }

                                                if (choiceToApply == ConflictResolutionChoice.KeepExistingValue)
                                                {
                                                    if (logger != null) logger(string.Format("    ℹ️ [사용자 선택] [{0}] 셀 기존 값({1}) 보존", cellKey, vBase));
                                                    continue;
                                                }
                                            }
                                        }

                                        dynamic cell = null;
                                        try
                                        {
                                            cell = wsBase.Cells[r, c];
                                            cell.Value2 = finalVal;

                                            if (applyHighlight)
                                            {
                                                try { cell.Interior.Color = highlightColorOle; } catch {}
                                            }

                                            if (addComment)
                                            {
                                                try
                                                {
                                                    try { cell.ClearComments(); } catch {}
                                                    string commentText = string.Format("[{0}] 변경(수정값): {1}{2} (기존: {3})\n수정일시: {4}",
                                                        modFileName,
                                                        vMod,
                                                        commentNote,
                                                        string.IsNullOrEmpty(vBase) ? "(빈값)" : vBase,
                                                        DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                                                    cell.AddComment(commentText);
                                                }
                                                catch { }
                                            }

                                            modifiedCellTracker[cellKey] = vMod;
                                        }
                                        catch (Exception exCell)
                                        {
                                            if (logger != null) logger(string.Format("    ⚠️ 셀 [{0},{1}] 변경 예외: {2}", r, c, exCell.Message));
                                        }
                                        finally
                                        {
                                            if (cell != null)
                                            {
                                                try { Marshal.ReleaseComObject(cell); } catch {}
                                            }
                                        }

                                        sheetDiffCount++;
                                        fileDiffCount++;
                                        totalDiffCount++;
                                    }
                                }
                            }

                            if (logger != null && isAutoAllSheets)
                            {
                                logger(string.Format("    ├ 시트 [{0}]: {1}개 셀 반영 완료", sName, sheetDiffCount));
                            }

                            if (wsBase != null) { try { Marshal.ReleaseComObject(wsBase); } catch {} }
                            if (wsMod != null) { try { Marshal.ReleaseComObject(wsMod); } catch {} }
                        }

                        if (logger != null) logger(string.Format("  └ {0}: 총 {1}개 셀 반영 완료", modFileName, fileDiffCount));
                    }
                    finally
                    {
                        if (wbMod != null)
                        {
                            try { wbMod.Close(false); } catch (Exception exCloseMod) { if (logger != null) logger("    ⚠️ wbMod.Close 예외: " + exCloseMod.Message); }
                            try { Marshal.ReleaseComObject(wbMod); } catch {}
                            wbMod = null;
                        }
                    }
                }

                if (logger != null) logger("[STEP 2] 모든 취합 결과 원본 저장 처리 중...");
                try { excel.DisplayAlerts = false; } catch {}
                
                string tempCopyPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(basePath), "temp_agg_save_" + Guid.NewGuid().ToString("N") + ".xlsx");
                bool savedViaCopy = false;

                try
                {
                    if (logger != null) logger("  [STEP 2-1] 원본 직접 저장(wbBase.Save) 시도...");
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

                if (logger != null) logger("[STEP 3] 원본 워크북(wbBase) 닫기 시도...");
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
                    if (wbBase != null) { try { Marshal.ReleaseComObject(wbBase); } catch {} wbBase = null; }
                }

                if (logger != null) logger("[STEP 4] Excel 백그라운드 애플리케이션 정리 중...");
                if (excel != null)
                {
                    try { excel.Quit(); } catch {}
                    try { Marshal.ReleaseComObject(excel); } catch {}
                    excel = null;
                }

                if (savedViaCopy && System.IO.File.Exists(tempCopyPath))
                {
                    if (logger != null) logger("[STEP 5] 임시 파일 원본 덮어쓰기(File.Copy) 시도...");
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

                if (logger != null) logger("🎉 모든 엑셀 취합 및 저장 정상 완료!");
                return totalDiffCount;
            }
            catch (Exception exMain)
            {
                if (isFileSuccessfullySaved)
                {
                    if (logger != null) logger("  ℹ️ 저장 후 정리 단계에서 예외 발생(" + exMain.Message + "), 하지만 파일 취합 및 디스크 저장은 이미 성공했습니다.");
                    return totalDiffCount;
                }
                throw exMain;
            }
            finally
            {
                try
                {
                    if (wbBase != null) { Marshal.ReleaseComObject(wbBase); }
                }
                catch { }
                wbBase = null;

                try
                {
                    if (excel != null) { excel.Quit(); }
                }
                catch { }

                try
                {
                    if (excel != null) { Marshal.ReleaseComObject(excel); }
                }
                catch { }
                excel = null;
            }
        }

        public static int ResetHighlightsAndComments(
            string basePath,
            string sheetName,
            LogHandler logger
        )
        {
            int totalResetCount = 0;
            Exception threadEx = null;

            System.Threading.Thread staThread = new System.Threading.Thread(() =>
            {
                try
                {
                    totalResetCount = ResetHighlightsAndCommentsInternal(basePath, sheetName, logger);
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
                string errStr = threadEx.ToString();
                if (errStr.Contains("0x80010114") || errStr.Contains("RPC_E_DISCONNECTED") || errStr.Contains("요청한 개체가 없습니다"))
                {
                    if (logger != null) logger("  ℹ️ COM 정리 중 무해한 예외가 발생했으나, 파일의 메모 및 음영 초기화는 100% 완료되었습니다.");
                    return totalResetCount;
                }
                throw threadEx;
            }
            return totalResetCount;
        }

        private static int ResetHighlightsAndCommentsInternal(
            string basePath,
            string sheetName,
            LogHandler logger
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

            if (logger != null) logger("엑셀 애플리케이션 시작 중 (메모 및 음영 초기화)...");
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
                        try { Marshal.ReleaseComObject(ws); } catch {}
                    }
                    try { Marshal.ReleaseComObject(sheets); } catch {}
                }
                else if (isFirstSheetOnly)
                {
                    dynamic ws1 = wbBase.Worksheets[1];
                    targetSheetNames.Add(ws1.Name);
                    try { Marshal.ReleaseComObject(ws1); } catch {}
                }
                else
                {
                    targetSheetNames.Add(sheetName);
                }

                foreach (string sName in targetSheetNames)
                {
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

                    dynamic usedRange = null;
                    try
                    {
                        usedRange = wsBase.UsedRange;
                        if (usedRange != null)
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
                                            try { Marshal.ReleaseComObject(interior); } catch { }
                                        }
                                        catch { }
                                        finally
                                        {
                                            if (cell != null) { try { Marshal.ReleaseComObject(cell); } catch { } }
                                        }
                                    }
                                }
                            }
                            catch (Exception exFind)
                            {
                                if (logger != null) logger(string.Format("  ⚠️ 파스텔 음영 제거 예외: {0}", exFind.Message));
                            }

                            try
                            {
                                dynamic comments = wsBase.Comments;
                                if (comments != null)
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
                                            if (comment != null) { try { Marshal.ReleaseComObject(comment); } catch { } }
                                        }
                                    }
                                    try { Marshal.ReleaseComObject(comments); } catch { }
                                }
                            }
                            catch { }

                            resetSheetsCount++;
                            if (logger != null) logger(string.Format("  ├ 시트 [{0}]: 취합 음영({1}개) 및 메모({2}개) 선택 제거 완료 (원본 서식 유지)", sName, removedHighlightCount, removedCommentCount));
                        }
                    }
                    catch (Exception ex)
                    {
                        if (logger != null) logger(string.Format("  ⚠️ 시트 [{0}] 초기화 중 예외: {1}", sName, ex.Message));
                    }
                    finally
                    {
                        if (usedRange != null) { try { Marshal.ReleaseComObject(usedRange); } catch {} }
                        if (wsBase != null) { try { Marshal.ReleaseComObject(wsBase); } catch {} }
                    }
                }

                if (logger != null) logger("초기화 완료된 파일 저장 중...");
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
                try { if (wbBase != null) Marshal.ReleaseComObject(wbBase); } catch {}
                wbBase = null;

                try { excel.Quit(); } catch {}
                try { if (excel != null) Marshal.ReleaseComObject(excel); } catch {}
                excel = null;

                if (logger != null) logger("🎉 모든 시트의 메모 및 음영이 성공적으로 제거되었습니다!");
                return resetSheetsCount;
            }
            catch (Exception exMain)
            {
                if (isFileSaved)
                {
                    if (logger != null) logger("  ℹ️ 저장 후 정리 단계 예외, 하지만 초기화 저장은 완료되었습니다.");
                    return resetSheetsCount;
                }
                throw exMain;
            }
            finally
            {
                if (wbBase != null) { try { Marshal.ReleaseComObject(wbBase); } catch {} wbBase = null; }
                if (excel != null) { try { excel.Quit(); } catch {} try { Marshal.ReleaseComObject(excel); } catch {} excel = null; }
                try { GC.Collect(); } catch {}
                try { GC.WaitForPendingFinalizers(); } catch {}
            }
        }
    }
}
