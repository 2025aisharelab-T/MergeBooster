using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelDiffAggregator
{
    // Decides whether a modified value really differs from the base value under the user's comparison rules.
    // Only the comparison is relaxed; when a cell does count as changed, the modified file's value is written as-is.
    public static class CompareRules
    {
        // Floating-point noise floor relative to the magnitude (0.1 + 0.2 vs 0.3 etc.)
        private const double RelativeEpsilon = 1e-9;

        public static bool Same(string modText, string baseText, MergeOptions opt)
        {
            if (modText == baseText) return true;
            if (opt == null) return false;

            if (opt.CompareNumbers)
            {
                double a, b;
                if (TryNumber(modText, out a) && TryNumber(baseText, out b))
                {
                    double diff = Math.Abs(a - b);
                    double floor = RelativeEpsilon * Math.Max(Math.Abs(a), Math.Abs(b));
                    return diff <= Math.Max(Math.Max(opt.NumberTolerance, 0), floor);
                }
            }

            if (opt.IgnoreWhitespace || opt.IgnoreCase)
            {
                return string.Equals(Normalize(modText, opt), Normalize(baseText, opt), StringComparison.Ordinal);
            }
            return false;
        }

        // Key values follow the whitespace/case rules so "홍 길동" and "홍길동" find the same row
        // (number rules are not applied to keys: codes like "007" and "7" must stay different)
        public static string KeyText(string s, MergeOptions opt)
        {
            if (opt == null || (!opt.IgnoreWhitespace && !opt.IgnoreCase)) return s ?? "";
            return Normalize(s, opt);
        }

        private static string Normalize(string s, MergeOptions opt)
        {
            if (s == null) return "";
            if (opt.IgnoreWhitespace)
            {
                StringBuilder sb = new StringBuilder(s.Length);
                foreach (char ch in s)
                {
                    // char.IsWhiteSpace covers spaces, tabs, line breaks, NBSP (U+00A0) and full-width space (U+3000)
                    if (!char.IsWhiteSpace(ch) && ch != '​' && ch != '﻿') sb.Append(ch);
                }
                s = sb.ToString();
            }
            if (opt.IgnoreCase) s = s.ToUpperInvariant();
            return s;
        }

        // Accepts "1000", "1,000", " 1 000 ", "-3.5", "1.2E3", "10%" (= 0.1); rejects empty and non-numeric text
        public static bool TryNumber(string s, out double value)
        {
            value = 0;
            if (string.IsNullOrEmpty(s)) return false;
            string t = s.Trim().Replace(" ", "").Replace(" ", "");
            bool percent = false;
            if (t.EndsWith("%"))
            {
                percent = true;
                t = t.Substring(0, t.Length - 1);
            }
            if (t.Length == 0) return false;
            // Thousands separators only in a valid position ("1,234,567"); "1,5" is not treated as a number
            if (t.IndexOf(',') >= 0)
            {
                if (!ValidThousands(t)) return false;
                t = t.Replace(",", "");
            }
            if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
            if (double.IsNaN(value) || double.IsInfinity(value)) return false;
            if (percent) value /= 100.0;
            return true;
        }

        private static bool ValidThousands(string t)
        {
            string body = t;
            if (body.StartsWith("-") || body.StartsWith("+")) body = body.Substring(1);
            int dot = body.IndexOf('.');
            string intPart = dot >= 0 ? body.Substring(0, dot) : body;
            if (dot >= 0 && body.IndexOf(',', dot) >= 0) return false;   // no commas after the decimal point
            string[] groups = intPart.Split(',');
            if (groups[0].Length < 1 || groups[0].Length > 3) return false;
            for (int i = 1; i < groups.Length; i++) if (groups[i].Length != 3) return false;
            return true;
        }

        // Resolves "비고, 담당자, H" against the base header row; header names win over column letters.
        // Unknown names are returned in 'unknown' so the caller can warn (they are simply ignored).
        public static HashSet<int> ResolveColumns(string spec, Dictionary<string, int> header, List<string> unknown)
        {
            HashSet<int> cols = new HashSet<int>();
            if (string.IsNullOrEmpty(spec)) return cols;
            foreach (string raw in spec.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string token = raw.Trim();
                if (token.Length == 0) continue;
                int col;
                if (header != null && header.TryGetValue(token, out col)) { cols.Add(col); continue; }
                col = ChangeItem.ColumnIndex(token);
                if (col > 0) { cols.Add(col); continue; }
                if (unknown != null) unknown.Add(token);
            }
            return cols;
        }

        public static string Describe(MergeOptions opt)
        {
            List<string> parts = new List<string>();
            if (opt.IgnoreWhitespace) parts.Add("공백 차이 무시");
            if (opt.IgnoreCase) parts.Add("대소문자 무시");
            if (opt.CompareNumbers)
            {
                parts.Add(opt.NumberTolerance > 0
                    ? "숫자 값 비교(오차 ±" + opt.NumberTolerance.ToString(CultureInfo.InvariantCulture) + ")"
                    : "숫자 값 비교");
            }
            if (opt.ExcludeColumns != null && opt.ExcludeColumns.Trim().Length > 0) parts.Add("제외 열 [" + opt.ExcludeColumns.Trim() + "]");
            return parts.Count > 0 ? string.Join(", ", parts.ToArray()) : "정확히 같은 값만 같음";
        }
    }
}
