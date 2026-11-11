using System;
using System.Collections.Generic;

namespace ExcelDiffAggregator
{
    public enum ChangeKind
    {
        Update,  // 원본에 있던 행의 칸 수정
        NewRow   // 키 열 매칭에서 원본에 없던 키 → 맨 아래에 새 행으로 추가
    }

    public enum ChangeStatus
    {
        Planned,          // 분석 결과 반영 예정 (미리보기 대상)
        Applied,          // 원본에 반영됨
        Excluded,         // 미리보기에서 제외 / 취소
        Superseded,       // 같은 칸에 나중 파일 값이 반영되어 대체됨
        ConflictKept,     // 충돌 규칙(첫 파일 우선, 사용자 선택)으로 기존 값 유지
        FormulaProtected, // 원본 칸에 수식이 있어 건너뜀
        MissingInMod,     // 키 열 매칭: 원본에는 있지만 수정본에 없는 행 (정보용)
        Failed            // 반영 중 오류
    }

    // One proposed change from one modified file to one cell of the base workbook
    public class ChangeItem
    {
        public int Seq;
        public string FileName = "";
        public string SheetName = "";
        public int Row;                 // 원본 기준 행 번호 (1부터)
        public int Col;                 // 원본 기준 열 번호 (1부터, 0 = 행 전체 정보)
        public string RowKey = "";      // 키 열 매칭일 때 행 키 (여러 열은 " / " 로 연결)
        public string ColumnHeader = ""; // 키 열 매칭일 때 원본 머리글
        public ChangeKind Kind = ChangeKind.Update;
        public string OldText = "";     // 이 파일 반영 직전 값 (앞 파일 반영분 포함)
        public string NewText = "";
        public object NewValue;
        public ChangeStatus Status = ChangeStatus.Planned;
        public string Note = "";
        public bool Approved = true;    // 미리보기 체크 상태
        public PreviewSheet Sheet;      // 미리보기 표 보기용 원본 시트 값 (없으면 변경 항목만으로 표를 그림)

        public string Address
        {
            get { return Col > 0 ? ColumnLetter(Col) + Row : Row + "행"; }
        }

        public string CellKey
        {
            get { return SheetName + "!R" + Row + "C" + Col; }
        }

        public string ColumnLabel
        {
            get
            {
                if (Col <= 0) return "";
                return string.IsNullOrEmpty(ColumnHeader) ? ColumnLetter(Col) : ColumnLetter(Col) + " (" + ColumnHeader + ")";
            }
        }

        public string KindText
        {
            get { return Kind == ChangeKind.NewRow ? "새 행" : "수정"; }
        }

        public string StatusText
        {
            get { return DescribeStatus(Status); }
        }

        public static string DescribeStatus(ChangeStatus status)
        {
            switch (status)
            {
                case ChangeStatus.Planned: return "반영 예정";
                case ChangeStatus.Applied: return "반영됨";
                case ChangeStatus.Excluded: return "제외됨";
                case ChangeStatus.Superseded: return "나중 파일 값으로 대체";
                case ChangeStatus.ConflictKept: return "충돌 - 기존 값 유지";
                case ChangeStatus.FormulaProtected: return "수식 보호로 건너뜀";
                case ChangeStatus.MissingInMod: return "수정본에 없는 행";
                case ChangeStatus.Failed: return "반영 실패";
            }
            return status.ToString();
        }

        public static string ColumnLetter(int col)
        {
            if (col <= 0) return "";
            string s = "";
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = (char)('A' + m) + s;
                col = (col - 1) / 26;
            }
            return s;
        }

        // "A" → 1, "AB" → 28; returns 0 when the text is not a column letter
        public static int ColumnIndex(string letters)
        {
            if (string.IsNullOrEmpty(letters) || letters.Length > 3) return 0;
            int n = 0;
            foreach (char ch in letters.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                n = n * 26 + (ch - 'A' + 1);
            }
            return n <= 16384 ? n : 0;
        }
    }

    // Read-only snapshot of one base sheet, shared by every ChangeItem of that sheet (used by the preview grid)
    public class PreviewSheet
    {
        public string Name = "";
        public object[,] Values;    // 원본 값 (Excel 범위 배열, 1부터)
        public int Rows;
        public int Cols;
        public int HeaderRow;       // 키 열 매칭일 때 머리글 행 (0 = 없음)

        public string OriginalText(int r, int c)
        {
            return SheetState.Text(SheetState.Get(Values, r, c));
        }
    }

    // Called on the merge thread with the planned items; return false to cancel the whole merge
    public delegate bool PreviewHandler(List<ChangeItem> plannedItems);

    public class MergeOptions
    {
        public string BasePath;
        public string[] ModPaths = new string[0];
        public string SheetName = "(전체 시트 자동 인식)";
        public bool ApplyHighlight = true;
        public bool AddComment = true;
        public int HighlightColorOle = 13431551;
        public bool AutoBackup = true;
        public int BackupKeepCount = 10;    // 자동 백업 보관 개수 (0 이하 = 정리 안 함)
        public ConflictPolicy ConflictPolicy = ConflictPolicy.LastWins;

        // 1단계: 키 열 매칭 / 수식 보호
        public string KeyColumns = "";      // 비우면 같은 위치(행·열)끼리 비교
        public int HeaderRow = 1;
        public bool ProtectFormulas = true;

        // 2단계: 미리보기 / 보고서
        public PreviewHandler Preview;      // null 이면 미리보기 없이 바로 반영
        public bool CreateReport = false;

        // 6번: 비교 규칙 (같은 값으로 볼지 판단할 때만 쓰고, 반영할 때는 수정본 값을 그대로 씀)
        public bool IgnoreWhitespace = false;   // "홍 길동" = "홍길동", 줄바꿈·NBSP 포함
        public bool IgnoreCase = false;         // "abc" = "ABC"
        public bool CompareNumbers = false;     // "1,000" = 1000, "10%" = 0.1, 오차 허용
        public double NumberTolerance = 0;      // |a-b| ≤ 오차면 같은 값 (0 = 형식 차이·부동소수 오차만 무시)
        public string ExcludeColumns = "";      // 비교·반영하지 않을 열: 머리글 이름 또는 열 문자, 쉼표로 구분

        public bool UseKeyMatching
        {
            get { return KeyColumns != null && KeyColumns.Trim().Length > 0; }
        }

        public bool HasCompareRules
        {
            get { return IgnoreWhitespace || IgnoreCase || CompareNumbers || (ExcludeColumns != null && ExcludeColumns.Trim().Length > 0); }
        }
    }

    public class MergeResult
    {
        public List<ChangeItem> Items = new List<ChangeItem>();
        public int AppliedCells;
        public int NewRows;
        public bool Cancelled;
        public bool Completed;      // 저장(또는 저장 불필요 판단)까지 끝남
        public string BackupPath;
        public string ReportPath;
        public int RuleSameCells;       // 비교 규칙 덕분에 '같은 값'으로 보고 건너뛴 칸
        public int ExcludedColumnCells; // 비교 제외 열이라 건너뛴 칸 (값이 달랐던 칸만)

        public int Count(ChangeStatus status)
        {
            int n = 0;
            foreach (ChangeItem it in Items) if (it.Status == status) n++;
            return n;
        }
    }

    // Per-sheet working state for one merge run: base values read once, plus in-memory edits from earlier files
    internal class SheetState
    {
        public string Name;
        public object[,] BaseArr;
        public int BaseRows;
        public int BaseCols;
        public bool[,] Formula;                     // null 이면 수식 보호 안 함
        public int FormulaCount;
        public int NextNewRow;
        public bool KeyError;
        public string[] HeaderByCol = new string[0];
        public Dictionary<string, int> BaseHeader = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public int[] BaseKeyCols;
        public Dictionary<string, int> KeyIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        public HashSet<int> NewRows = new HashSet<int>();
        public HashSet<int> ExcludedCols = new HashSet<int>();     // 원본 기준 비교 제외 열
        public int RuleSame;
        public int ExcludedSkipped;
        public Dictionary<long, object> Overlay = new Dictionary<long, object>();
        public Dictionary<long, string> TouchedBy = new Dictionary<long, string>();
        public dynamic Ws;                          // 반영 단계에서만 열어 둠

        public SheetState(string name)
        {
            Name = name;
        }

        public static long CellId(int row, int col)
        {
            return (long)row * 1048576L + col;   // 열은 최대 16,384라 2^20 안에 들어감
        }

        public static object Get(object[,] arr, int r, int c)
        {
            if (arr == null) return null;
            if (r < arr.GetLowerBound(0) || r > arr.GetUpperBound(0)) return null;
            if (c < arr.GetLowerBound(1) || c > arr.GetUpperBound(1)) return null;
            return arr[r, c];
        }

        public static string Text(object o)
        {
            return o != null ? o.ToString().Trim() : "";
        }

        public string OriginalText(int r, int c)
        {
            return Text(Get(BaseArr, r, c));
        }

        public string CurrentText(int r, int c)
        {
            object v;
            if (Overlay.TryGetValue(CellId(r, c), out v)) return Text(v);
            return OriginalText(r, c);
        }

        public bool IsFormula(int r, int c)
        {
            if (Formula == null) return false;
            if (r < 0 || r > Formula.GetUpperBound(0) || c < 0 || c > Formula.GetUpperBound(1)) return false;
            return Formula[r, c];
        }

        public string HeaderText(int c)
        {
            return (c >= 0 && c < HeaderByCol.Length && HeaderByCol[c] != null) ? HeaderByCol[c] : "";
        }
    }
}
