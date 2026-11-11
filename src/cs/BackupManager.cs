using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelDiffAggregator
{
    // Finds, prunes and restores the automatic backups named 원본이름_backup_yyyyMMdd_HHmmss[_n].ext
    public static class BackupManager
    {
        public class BackupEntry
        {
            public string FullPath;
            public DateTime Time;
            public long Size;

            public string FileName
            {
                get { return Path.GetFileName(FullPath); }
            }

            public override string ToString()
            {
                return string.Format("{0}    ·    {1:N0} KB    ·    {2}", Time.ToString("yyyy-MM-dd HH:mm:ss"), Math.Max(1, Size / 1024), FileName);
            }
        }

        public static List<BackupEntry> List(string basePath)
        {
            List<BackupEntry> list = new List<BackupEntry>();
            if (string.IsNullOrEmpty(basePath)) return list;
            basePath = Path.GetFullPath(basePath);
            string dir = Path.GetDirectoryName(basePath);
            if (!Directory.Exists(dir)) return list;

            string name = Path.GetFileNameWithoutExtension(basePath);
            string ext = Path.GetExtension(basePath);
            Regex pattern = new Regex("^" + Regex.Escape(name) + @"_backup_(\d{8}_\d{6})(_\d+)?" + Regex.Escape(ext) + "$", RegexOptions.IgnoreCase);

            foreach (string file in Directory.GetFiles(dir, name + "_backup_*" + ext))
            {
                Match m = pattern.Match(Path.GetFileName(file));
                if (!m.Success) continue;
                DateTime time;
                if (!DateTime.TryParseExact(m.Groups[1].Value, "yyyyMMdd_HHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out time)) continue;
                BackupEntry e = new BackupEntry();
                e.FullPath = file;
                e.Time = time;
                try { e.Size = new FileInfo(file).Length; } catch { }
                list.Add(e);
            }

            // Newest first; same second → higher _n suffix (made later) first
            list.Sort((x, y) =>
            {
                int byTime = y.Time.CompareTo(x.Time);
                return byTime != 0 ? byTime : string.Compare(y.FileName, x.FileName, StringComparison.OrdinalIgnoreCase);
            });
            return list;
        }

        // Keeps the newest `keep` backups and deletes the rest. Returns how many were deleted.
        public static int Prune(string basePath, int keep, ExcelDiffEngine.LogHandler logger)
        {
            if (keep <= 0) return 0;
            List<BackupEntry> list = List(basePath);
            int deleted = 0;
            for (int i = keep; i < list.Count; i++)
            {
                try
                {
                    File.Delete(list[i].FullPath);
                    deleted++;
                    DebugLog.Write("DEBUG", "오래된 백업 삭제: " + list[i].FullPath);
                }
                catch (Exception ex)
                {
                    DebugLog.Error("오래된 백업 삭제 실패: " + list[i].FullPath, ex);
                }
            }
            if (deleted > 0 && logger != null)
            {
                logger(string.Format("🧹 오래된 백업 {0}개 정리 (최근 {1}개만 보관)", deleted, keep));
            }
            return deleted;
        }

        // Puts a backup back in place of the base file. The current base is backed up first, so a restore can be undone.
        // Returns the path of that safety backup.
        public static string Restore(string basePath, string backupPath, ExcelDiffEngine.LogHandler logger)
        {
            basePath = Path.GetFullPath(basePath);
            backupPath = Path.GetFullPath(backupPath);
            if (!File.Exists(backupPath)) throw new FileNotFoundException("백업 파일이 없습니다: " + backupPath);
            if (ExcelDiffEngine.IsFileLocked(basePath))
            {
                throw new Exception(string.Format("원본 엑셀 파일({0})이 Excel이나 다른 프로그램에서 열려 있습니다.\n파일을 닫은 뒤 다시 시도해 주세요.", Path.GetFileName(basePath)));
            }

            // Copy the chosen backup aside first: the safety backup below must never be able to touch it
            string temp = Path.Combine(Path.GetDirectoryName(basePath), "temp_restore_" + Guid.NewGuid().ToString("N") + Path.GetExtension(basePath));
            File.Copy(backupPath, temp, false);
            try
            {
                string safety = null;
                if (File.Exists(basePath))
                {
                    safety = ExcelDiffEngine.CreateAutoBackup(basePath, logger);
                    if (safety == null) throw new Exception("복원 전에 현재 원본을 백업하지 못해 복원을 중단했습니다.");
                }
                File.Copy(temp, basePath, true);
                if (logger != null) logger("♻️ 백업으로 복원 완료: " + Path.GetFileName(backupPath) + (safety != null ? " (복원 전 상태는 " + Path.GetFileName(safety) + " 로 보관)" : ""));
                DebugLog.Write("DEBUG", string.Format("복원: {0} → {1} (안전 백업 {2})", backupPath, basePath, safety));
                return safety;
            }
            finally
            {
                try { File.Delete(temp); } catch { }
            }
        }
    }
}
