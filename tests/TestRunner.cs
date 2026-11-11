using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDiffAggregator.Tests
{
    class TestRunner
    {
        static void Main(string[] args)
        {
            string testDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_temp");
            if (Directory.Exists(testDir))
            {
                try { Directory.Delete(testDir, true); } catch {}
            }
            Directory.CreateDirectory(testDir);

            string basePath = Path.Combine(testDir, "test_base.xlsx");
            string modPath = Path.Combine(testDir, "test_mod.xlsx");

            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                Console.WriteLine("FAIL: Excel not installed.");
                return;
            }

            dynamic excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // 1. Create Base Workbook
                dynamic wbBase = excel.Workbooks.Add();
                dynamic wsBase = wbBase.Worksheets[1];
                wsBase.Name = "직원목록";

                wsBase.Cells[1, 1].Value2 = "사번";
                wsBase.Cells[1, 2].Value2 = "성명";
                wsBase.Cells[1, 3].Value2 = "부서";

                // Green fill on header A1
                wsBase.Cells[1, 1].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0, 255, 0));

                wsBase.Cells[2, 1].Value2 = 1001;
                wsBase.Cells[2, 2].Value2 = "홍길동";
                wsBase.Cells[2, 3].Value2 = "영업1팀";
                // Original comment on B2
                wsBase.Cells[2, 2].AddComment("기존 중요 관리 메모");

                wsBase.Cells[3, 1].Value2 = 1002;
                wsBase.Cells[3, 2].Value2 = "김철수";
                wsBase.Cells[3, 3].Value2 = "개발1팀";

                wbBase.SaveAs(basePath);
                wbBase.Close(false);
                Marshal.ReleaseComObject(wbBase);

                // 2. Create Mod Workbook
                dynamic wbMod = excel.Workbooks.Add();
                dynamic wsMod = wbMod.Worksheets[1];
                wsMod.Name = "직원목록";

                wsMod.Cells[1, 1].Value2 = "사번";
                wsMod.Cells[1, 2].Value2 = "성명";
                wsMod.Cells[1, 3].Value2 = "부서";

                wsMod.Cells[2, 1].Value2 = 1001;
                wsMod.Cells[2, 2].Value2 = "홍길동";
                wsMod.Cells[2, 3].Value2 = "영업1팀";

                wsMod.Cells[3, 1].Value2 = 1002;
                wsMod.Cells[3, 2].Value2 = "김철수";
                wsMod.Cells[3, 3].Value2 = "AI개발팀"; // Changed!

                wbMod.SaveAs(modPath);
                wbMod.Close(false);
                Marshal.ReleaseComObject(wbMod);
            }
            finally
            {
                try { excel.Quit(); } catch {}
                try { Marshal.ReleaseComObject(excel); } catch {}
            }

            Console.WriteLine("Step 1: Base and Mod Excel files created.");

            // 3. Process Diff (Aggregation)
            Console.WriteLine("Step 2: Running ProcessDiffMultiple (Aggregation)...");
            int diffCount = ExcelDiffEngine.ProcessDiffMultiple(
                basePath,
                new string[] { modPath },
                "(전체 시트 자동 인식)",
                true,
                true,
                13431551,
                msg => Console.WriteLine("  [LOG] " + msg)
            );
            Console.WriteLine("Aggregation diff count: " + diffCount);

            // 4. Run Reset
            Console.WriteLine("Step 3: Running ResetHighlightsAndComments...");
            int resetCount = ExcelDiffEngine.ResetHighlightsAndComments(
                basePath,
                "(전체 시트 자동 인식)",
                msg => Console.WriteLine("  [LOG] " + msg)
            );
            Console.WriteLine("Reset count: " + resetCount);

            // 5. Verification
            excel = Activator.CreateInstance(excelType);
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                dynamic wbVer = excel.Workbooks.Open(basePath);
                dynamic wsVer = wbVer.Worksheets["직원목록"];

                // Check A1 fill (Must still be Green: RGB(0, 255, 0) = OLE 65280)
                object a1ColorObj = wsVer.Cells[1, 1].Interior.Color;
                long a1Color = Convert.ToInt64(a1ColorObj);
                long expectedGreen = ColorTranslator.ToOle(Color.FromArgb(0, 255, 0));
                Console.WriteLine("A1 Interior.Color: " + a1Color + " (Expected Green: " + expectedGreen + ")");

                // Check B2 Comment (Must still be "기존 중요 관리 메모")
                dynamic b2Comment = wsVer.Cells[2, 2].Comment;
                string b2CommentText = b2Comment != null ? b2Comment.Text() : null;
                Console.WriteLine("B2 Comment: " + (b2CommentText ?? "NULL"));

                // Check C3 Fill (Must be ColorIndex -4142 / xlNone)
                object c3ColorIdxObj = wsVer.Cells[3, 3].Interior.ColorIndex;
                int c3ColorIdx = Convert.ToInt32(c3ColorIdxObj);
                Console.WriteLine("C3 ColorIndex: " + c3ColorIdx + " (Expected: -4142)");

                // Check C3 Comment (Must be NULL)
                dynamic c3Comment = wsVer.Cells[3, 3].Comment;
                string c3CommentText = c3Comment != null ? c3Comment.Text() : null;
                Console.WriteLine("C3 Comment: " + (c3CommentText ?? "NULL"));

                bool pass = true;
                if (a1Color != expectedGreen)
                {
                    Console.WriteLine("FAIL: A1 Green fill was lost or altered! Got: " + a1Color);
                    pass = false;
                }
                if (b2CommentText != "기존 중요 관리 메모")
                {
                    Console.WriteLine("FAIL: B2 original comment was lost or altered!");
                    pass = false;
                }
                if (c3ColorIdx != -4142)
                {
                    Console.WriteLine("FAIL: C3 aggregation yellow fill was not reset! ColorIndex: " + c3ColorIdx);
                    pass = false;
                }
                if (c3CommentText != null)
                {
                    Console.WriteLine("FAIL: C3 aggregation comment was not deleted!");
                    pass = false;
                }

                if (pass)
                {
                    Console.WriteLine("\n🎉 TEST PASSED! Selective reset correctly preserves original fill and comments, while removing only aggregated fill and comments.");
                }
                else
                {
                    Console.WriteLine("\n❌ TEST FAILED!");
                }

                wbVer.Close(false);
                Marshal.ReleaseComObject(wbVer);
            }
            finally
            {
                try { excel.Quit(); } catch {}
                try { Marshal.ReleaseComObject(excel); } catch {}
            }
        }
    }
}
