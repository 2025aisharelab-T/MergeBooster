using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDiffAggregator.Tests
{
    // Tests for the progress reporting added to the merge / reset flows.
    // Unit tests cover ProgressTracker alone; integration tests drive the real engine through Excel.
    static class ProgressTests
    {
        public static int Failures;

        private class ProgressEvent
        {
            public int Percent;
            public string Status;
        }

        private static void Check(bool ok, string name)
        {
            if (ok)
            {
                Console.WriteLine("  PASS: " + name);
            }
            else
            {
                Failures++;
                Console.WriteLine("  FAIL: " + name);
            }
        }

        private static string Join(List<string> items)
        {
            return string.Join("|", items.ToArray());
        }

        private static string Describe(List<ProgressEvent> events)
        {
            List<string> parts = new List<string>();
            foreach (ProgressEvent e in events) parts.Add(e.Percent + ":" + e.Status);
            return Join(parts);
        }

        // ---------------------------------------------------------------- unit tests

        public static void RunUnit()
        {
            Console.WriteLine("ProgressTracker unit tests");

            // Breaks caught: no lower clamp (-5 shown as -5), no upper clamp (150), a lower value moving the bar
            // backwards (5 after 10), rounding 99.9 up to 100 so the bar claims "done" before the work is.
            {
                List<string> seen = new List<string>();
                ProgressTracker t = new ProgressTracker((p, s) => seen.Add(p + ":" + s), 100, () => 0);
                t.Report(-5, "start", true);
                t.Report(10, "a", true);
                t.Report(5, "b", true);
                t.Report(99.9, "c", true);
                t.Report(150, "d", true);
                Check(Join(seen) == "0:start|10:a|10:b|99:c|100:d", "percent is clamped to 0..100, never goes back, fractions floor (got " + Join(seen) + ")");
                Check(t.Current == 100, "Current returns the highest percent reported");
            }

            // Breaks caught: no throttling (b is emitted), force ignored (d is dropped inside the interval).
            {
                List<string> seen = new List<string>();
                long now = 0;
                ProgressTracker t = new ProgressTracker((p, s) => seen.Add(p + ":" + s), 100, () => now);
                t.Report(10, "a", false);
                now = 50; t.Report(20, "b", false);
                now = 120; t.Report(30, "c", false);
                now = 121; t.Report(40, "d", true);
                Check(Join(seen) == "10:a|30:c|40:d", "non-forced reports are throttled, forced ones always go out (got " + Join(seen) + ")");
            }

            // Break caught: an identical update re-sent to the UI on every call.
            {
                List<string> seen = new List<string>();
                ProgressTracker t = new ProgressTracker((p, s) => seen.Add(p + ":" + s), 100, () => 0);
                t.Report(10, "a", true);
                t.Report(10, "a", true);
                Check(Join(seen) == "10:a", "an identical update is not re-sent (got " + Join(seen) + ")");
            }

            // Break caught: Due always true, so callers format a status string for every row of a huge sheet.
            {
                long now = 1000;
                ProgressTracker t = new ProgressTracker((p, s) => { }, 100, () => now);
                bool before = t.Due;
                t.Report(1, "a", false);
                bool justAfter = t.Due;
                now = 1099;
                bool almost = t.Due;
                now = 1100;
                bool elapsed = t.Due;
                Check(before && !justAfter && !almost && elapsed, "Due is true only once the interval has elapsed");
            }

            // Break caught: NullReferenceException for callers that pass no progress handler, or work spent on them.
            {
                ProgressTracker t = new ProgressTracker(null, 100, () => 0);
                bool threw = false;
                try { t.Report(50, "a", true); } catch (Exception) { threw = true; }
                Check(!threw && !t.Due, "a tracker without a handler is a silent no-op");
            }
        }

        // ------------------------------------------------------- integration tests (real Excel)

        private static ProgressHandler Collect(List<ProgressEvent> sink)
        {
            return (p, s) => sink.Add(new ProgressEvent { Percent = p, Status = s });
        }

        private static void WriteWorkbook(dynamic excel, string path, string sheetName, string[,] cells)
        {
            dynamic wb = excel.Workbooks.Add();
            dynamic ws = wb.Worksheets[1];
            ws.Name = sheetName;
            for (int r = 0; r < cells.GetLength(0); r++)
            {
                for (int c = 0; c < cells.GetLength(1); c++)
                {
                    ws.Cells[r + 1, c + 1].Value2 = cells[r, c];
                }
            }
            wb.SaveAs(path);
            wb.Close(false);
            Marshal.ReleaseComObject((object)wb);
        }

        private static bool IsMonotonicAndInRange(List<ProgressEvent> events)
        {
            int previous = 0;
            foreach (ProgressEvent e in events)
            {
                if (e.Percent < previous || e.Percent < 0 || e.Percent > 100) return false;
                previous = e.Percent;
            }
            return true;
        }

        public static void RunIntegration(Type excelType, string testDir)
        {
            string dir = Path.Combine(testDir, "progress");
            Directory.CreateDirectory(dir);
            string basePath = Path.Combine(dir, "base.xlsx");
            string mod1 = Path.Combine(dir, "mod1.xlsx");
            string mod2 = Path.Combine(dir, "mod2.xlsx");

            dynamic excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                WriteWorkbook(excel, basePath, "Sheet1", new string[,] { { "id", "name" }, { "1", "kim" }, { "2", "lee" } });
                // The engine compares each file against the base as already updated by earlier files, so mod2 keeps
                // mod1's edit (KIM); that leaves exactly one new change per file: B2 from mod1, B3 from mod2.
                WriteWorkbook(excel, mod1, "Sheet1", new string[,] { { "id", "name" }, { "1", "KIM" }, { "2", "lee" } });
                WriteWorkbook(excel, mod2, "Sheet1", new string[,] { { "id", "name" }, { "1", "KIM" }, { "2", "LEE" } });
            }
            finally
            {
                try { excel.Quit(); } catch { }
                try { Marshal.ReleaseComObject((object)excel); } catch { }
            }

            Console.WriteLine("Merge progress (2 files, 1 sheet)");
            List<ProgressEvent> merge = new List<ProgressEvent>();
            int diff = ExcelDiffEngine.ProcessDiffMultiple(
                basePath, new string[] { mod1, mod2 }, "(전체 시트 자동 인식)",
                false, false, 13431551, false, ConflictPolicy.LastWins,
                msg => { }, Collect(merge));
            string trace = Describe(merge);

            Check(diff == 2, "merge result is unchanged by progress reporting (2 cells, got " + diff + ")");
            Check(merge.Count > 0, "merge reports progress at all");
            Check(IsMonotonicAndInRange(merge), "merge progress stays within 0..100 and never goes back (got " + trace + ")");
            Check(merge.Count > 0 && merge[merge.Count - 1].Percent == 100,
                "merge ends at 100 even when Excel cleanup throws after the save (got " + trace + ")");

            // Files are analysed inside 5..80 (80..95 is writing approved cells), so file 2 of 2 starts at 5 + 75 * 1/2 = 42.5.
            // Breaks caught: progress not advancing per file, wrong span offset or denominator.
            int firstOfFile2 = -1;
            int highestOfFile1 = -1;
            foreach (ProgressEvent e in merge)
            {
                if (e.Status.Contains("2/2") && firstOfFile2 < 0) firstOfFile2 = e.Percent;
                if (e.Status.Contains("1/2") && e.Percent > highestOfFile1) highestOfFile1 = e.Percent;
            }
            Check(highestOfFile1 > 5 && highestOfFile1 <= 42, "file 1 of 2 reports progress inside 5..42 (highest " + highestOfFile1 + ")");
            Check(firstOfFile2 >= 42 && firstOfFile2 < 80, "file 2 of 2 starts at or after the 42% mark (first " + firstOfFile2 + ")");

            bool namesSheet = false;
            foreach (ProgressEvent e in merge)
            {
                if (e.Percent > 5 && e.Percent < 95 && e.Status.Contains("Sheet1")) namesSheet = true;
            }
            Check(namesSheet, "a mid-run status names the sheet being processed");

            Console.WriteLine("Reset progress (1 sheet)");
            List<ProgressEvent> reset = new List<ProgressEvent>();
            ExcelDiffEngine.ResetHighlightsAndComments(basePath, "(전체 시트 자동 인식)", msg => { }, Collect(reset));
            string resetTrace = Describe(reset);

            Check(reset.Count > 0, "reset reports progress at all");
            Check(IsMonotonicAndInRange(reset), "reset progress stays within 0..100 and never goes back (got " + resetTrace + ")");
            Check(reset.Count > 0 && reset[reset.Count - 1].Percent == 100, "reset ends at 100 (got " + resetTrace + ")");
            bool resetNamesSheet = false;
            foreach (ProgressEvent e in reset)
            {
                if (e.Percent > 5 && e.Percent < 95 && e.Status.Contains("Sheet1")) resetNamesSheet = true;
            }
            Check(resetNamesSheet, "a mid-run reset status names the sheet being processed");

            // Break caught: 100% reported from a finally block, so a failed merge looks finished.
            Console.WriteLine("Merge progress on failure");
            List<ProgressEvent> failed = new List<ProgressEvent>();
            bool threw = false;
            try
            {
                ExcelDiffEngine.ProcessDiffMultiple(
                    Path.Combine(dir, "does_not_exist.xlsx"), new string[] { mod1 }, "(전체 시트 자동 인식)",
                    false, false, 13431551, false, ConflictPolicy.LastWins,
                    msg => { }, Collect(failed));
            }
            catch (Exception)
            {
                threw = true;
            }
            bool reached100 = false;
            foreach (ProgressEvent e in failed) if (e.Percent == 100) reached100 = true;
            Check(threw, "a missing base file still raises the error");
            Check(!reached100, "a failed merge never reports 100% (got " + Describe(failed) + ")");
        }
    }
}
