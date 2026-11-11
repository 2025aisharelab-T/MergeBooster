using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelDiffAggregator
{
    // Lists every planned change before anything is written; the user approves or excludes items one by one or in groups
    public class PreviewDialog : Form
    {
        private class GroupOption
        {
            public string Label;
            public Predicate<ChangeItem> Match;
            public override string ToString() { return Label; }
        }

        // One base sheet shown in the Excel-style grid
        private class SheetView
        {
            public string Name;
            public PreviewSheet Data;   // null → 원본 값은 변경 항목의 기존 값으로 대신 표시
            public int Rows;
            public int Cols;
            public readonly Dictionary<long, List<ChangeItem>> Cells = new Dictionary<long, List<ChangeItem>>();
            public readonly List<long> Order = new List<long>();   // 변경 칸 (행 → 열 순)
            public override string ToString() { return string.Format("{0}  (변경 {1}칸)", Name, Cells.Count); }
        }

        private const int MaxGridCols = 300;   // 아주 넓은 시트는 변경 칸이 있는 열까지만 표시

        private static readonly Color CellChanged = ColorTranslator.FromHtml("#FFF2CC");
        private static readonly Color CellNewRow = ColorTranslator.FromHtml("#E2EFDA");
        private static readonly Color CellConflict = ColorTranslator.FromHtml("#FCE4D6");
        private static readonly Color CellExcluded = ColorTranslator.FromHtml("#EDEDED");
        private static readonly Color CellHeaderRow = ColorTranslator.FromHtml("#F2F2F2");

        private readonly List<ChangeItem> items;
        private readonly List<GroupOption> groups = new List<GroupOption>();
        private readonly List<SheetView> sheetViews = new List<SheetView>();
        private SheetView current;
        private TabControl tabs;
        private DataGridView grid;
        private ComboBox cboSheet;
        private Label lblCellInfo;
        private Font gridFont;
        private Font gridBoldFont;
        private Font gridStrikeFont;
        private ListView list;
        private Label lblSummary;
        private ComboBox cboGroup;
        private SoftButton btnApply;
        private bool suppressCheckEvents;

        public PreviewDialog(List<ChangeItem> plannedItems)
        {
            items = plannedItems;
            BuildGroups();
            BuildSheetViews();
            InitializeComponent();
            FillList();
            UpdateSummary();
            if (sheetViews.Count > 0) cboSheet.SelectedIndex = 0;
            this.Shown += (s, e) => ShowFromTop();
        }

        // Value the grid shows for a cell once the current check marks are applied (row/col are 1-based)
        public string DisplayText(string sheetName, int row, int col)
        {
            foreach (SheetView v in sheetViews)
            {
                if (v.Name == sheetName) return CellText(v, row, col);
            }
            return "";
        }

        // Same as double-clicking a changed cell in the grid; returns false when the cell has no change
        public bool ToggleCell(string sheetName, int row, int col)
        {
            foreach (SheetView v in sheetViews)
            {
                List<ChangeItem> cell;
                if (v.Name == sheetName && v.Cells.TryGetValue(SheetState.CellId(row, col), out cell))
                {
                    ToggleCells(new List<List<ChangeItem>> { cell });
                    return true;
                }
            }
            return false;
        }

        public void ShowGridTab(bool gridTab)
        {
            tabs.SelectedIndex = gridTab ? 0 : 1;
        }

        public int ApprovedCount
        {
            get
            {
                int n = 0;
                foreach (ChangeItem it in items) if (it.Approved) n++;
                return n;
            }
        }

        public List<string> GroupLabels
        {
            get
            {
                List<string> labels = new List<string>();
                foreach (GroupOption g in groups) labels.Add(g.Label);
                return labels;
            }
        }

        // Sets Approved for every item in the named group ("파일: a.xlsx", "시트: Sheet1", "열: C (부서)", "구분: 새 행", "전체")
        public int SetGroup(string label, bool approved)
        {
            GroupOption group = null;
            foreach (GroupOption g in groups) if (g.Label == label) group = g;
            if (group == null) throw new ArgumentException("알 수 없는 묶음: " + label);

            int changed = 0;
            foreach (ChangeItem it in items)
            {
                if (group.Match(it) && it.Approved != approved)
                {
                    it.Approved = approved;
                    changed++;
                }
            }
            SyncChecks();
            UpdateSummary();
            if (grid != null) grid.Invalidate();
            UpdateCellInfo();
            DebugLog.Write("DEBUG", string.Format("미리보기 묶음 {0}: [{1}] {2}건", approved ? "적용" : "제외", label, changed));
            return changed;
        }

        private void BuildGroups()
        {
            groups.Add(new GroupOption { Label = "전체", Match = it => true });

            List<string> files = new List<string>(), sheets = new List<string>(), columns = new List<string>();
            bool hasNewRow = false, hasUpdate = false, hasConflict = false;
            foreach (ChangeItem it in items)
            {
                if (!files.Contains(it.FileName)) files.Add(it.FileName);
                if (!sheets.Contains(it.SheetName)) sheets.Add(it.SheetName);
                if (!columns.Contains(it.ColumnLabel)) columns.Add(it.ColumnLabel);
                if (it.Kind == ChangeKind.NewRow) hasNewRow = true; else hasUpdate = true;
                if (it.Note.StartsWith("충돌")) hasConflict = true;
            }

            foreach (string f in files)
            {
                string name = f;
                groups.Add(new GroupOption { Label = "파일: " + name, Match = it => it.FileName == name });
            }
            if (sheets.Count > 1)
            {
                foreach (string s in sheets)
                {
                    string name = s;
                    groups.Add(new GroupOption { Label = "시트: " + name, Match = it => it.SheetName == name });
                }
            }
            foreach (string c in columns)
            {
                string label = c;
                groups.Add(new GroupOption { Label = "열: " + label, Match = it => it.ColumnLabel == label });
            }
            if (hasNewRow && hasUpdate)
            {
                groups.Add(new GroupOption { Label = "구분: 새 행", Match = it => it.Kind == ChangeKind.NewRow });
                groups.Add(new GroupOption { Label = "구분: 수정", Match = it => it.Kind == ChangeKind.Update });
            }
            if (hasConflict)
            {
                groups.Add(new GroupOption { Label = "구분: 충돌로 덮어쓰는 항목", Match = it => it.Note.StartsWith("충돌") });
            }
        }

        private void InitializeComponent()
        {
            Color colBg = ColorTranslator.FromHtml("#F8FAFC");
            Color colPrimary = ColorTranslator.FromHtml("#6366F1");
            Color colTextDark = ColorTranslator.FromHtml("#0F172A");
            Color colTextMuted = ColorTranslator.FromHtml("#64748B");
            Color colSecondaryBg = ColorTranslator.FromHtml("#F1F5F9");
            Color colSecondaryHover = ColorTranslator.FromHtml("#E2E8F0");
            Color colSecondaryBorder = ColorTranslator.FromHtml("#CBD5E1");

            this.SuspendLayout();
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Text = "반영 전 미리보기 - 머지부스터";
            this.ClientSize = new Size(1100, 720);
            this.MinimumSize = new Size(760, 520);
            this.Padding = new Padding(18, 10, 18, 12);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            // Shown from the merge's background STA thread (no owner window), so force it in front of the main form
            this.TopMost = true;
            this.ShowInTaskbar = true;
            this.BackColor = colBg;
            this.Font = new Font("맑은 고딕", 9F, FontStyle.Regular);

            Label lblTitle = new Label
            {
                Text = "아직 원본은 바뀌지 않았습니다. 반영할 항목만 체크한 뒤 [반영]을 눌러 주세요.",
                Font = new Font("맑은 고딕", 10.5F, FontStyle.Bold),
                ForeColor = colTextDark,
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            lblSummary = new Label
            {
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                ForeColor = colTextMuted,
                Dock = DockStyle.Top,
                Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Padding = new Point(14, 5)
            };
            TabPage pageGrid = new TabPage("엑셀 표 보기") { BackColor = Color.White, Padding = new Padding(6) };
            TabPage pageList = new TabPage("목록 보기") { BackColor = Color.White, Padding = new Padding(6) };
            tabs.TabPages.Add(pageGrid);
            tabs.TabPages.Add(pageList);
            BuildGridPage(pageGrid, colTextMuted);

            list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                GridLines = true,
                HideSelection = false,
                Font = new Font("맑은 고딕", 9F)
            };
            list.Columns.Add("수정본 파일", 150);
            list.Columns.Add("시트", 90);
            list.Columns.Add("셀", 60);
            list.Columns.Add("행 키", 110);
            list.Columns.Add("열", 110);
            list.Columns.Add("구분", 52);
            list.Columns.Add("기존 값", 120);
            list.Columns.Add("새 값", 120);
            list.Columns.Add("비고", 200);
            list.ItemChecked += (s, e) =>
            {
                if (suppressCheckEvents) return;
                ChangeItem it = (ChangeItem)e.Item.Tag;
                it.Approved = e.Item.Checked;
                UpdateSummary();
                if (grid != null) grid.Invalidate();
            };
            // Double-click a row in the list → jump to that cell in the grid
            list.DoubleClick += (s, e) =>
            {
                if (list.SelectedItems.Count == 0) return;
                ChangeItem it = (ChangeItem)list.SelectedItems[0].Tag;
                for (int i = 0; i < sheetViews.Count; i++)
                {
                    if (sheetViews[i].Name != it.SheetName) continue;
                    tabs.SelectedIndex = 0;
                    if (cboSheet.SelectedIndex != i) cboSheet.SelectedIndex = i;
                    SelectCell(it.Row, it.Col);
                    break;
                }
            };
            pageList.Controls.Add(list);

            Label lblGroup = new Label
            {
                Text = "묶음",
                ForeColor = colTextMuted,
                AutoSize = true,
                Margin = new Padding(0, 13, 6, 0)
            };

            cboGroup = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(190, 24),
                DropDownWidth = 320,
                Margin = new Padding(0, 9, 8, 0),
                Font = new Font("맑은 고딕", 9F)
            };
            foreach (GroupOption g in groups) cboGroup.Items.Add(g);
            cboGroup.SelectedIndex = 0;

            Func<string, int, int, SoftButton> secondary = (text, x, width) => new SoftButton
            {
                Text = text,
                Size = new Size(width, 34),
                Margin = new Padding(0, 4, 6, 0),
                BorderRadius = 10,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colTextDark,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };

            SoftButton btnGroupOn = secondary("묶음 적용", 314, 96);
            btnGroupOn.Click += (s, e) => SetGroup(((GroupOption)cboGroup.SelectedItem).Label, true);
            SoftButton btnGroupOff = secondary("묶음 제외", 416, 96);
            btnGroupOff.Click += (s, e) => SetGroup(((GroupOption)cboGroup.SelectedItem).Label, false);

            Label lblHint = new Label
            {
                Text = "표에서 칸을 더블클릭(또는 스페이스)하면 반영/제외가 바뀝니다. 제외한 칸은 원래 값이 유지되고, 같은 칸을 여러 파일이 바꾸면 체크된 것 중 가장 나중 파일 값이 반영됩니다.",
                ForeColor = colTextMuted,
                Font = new Font("맑은 고딕", 8.5F),
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 4, 0, 0),
                AutoEllipsis = true
            };

            btnApply = new SoftButton
            {
                Text = "반영",
                Size = new Size(150, 40),
                Margin = new Padding(12, 0, 0, 0),
                BorderRadius = 12,
                NormalColor = colPrimary,
                HoverColor = ColorTranslator.FromHtml("#4F46E5"),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnApply.Click += (s, e) =>
            {
                DebugLog.Write("DEBUG", string.Format("미리보기 반영: {0}/{1}건 승인", ApprovedCount, items.Count));
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            SoftButton btnCancel = new SoftButton
            {
                Text = "취소",
                Size = new Size(100, 40),
                Margin = new Padding(8, 0, 0, 0),
                BorderRadius = 12,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colTextDark,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnCancel.Click += (s, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };

            // Bottom bar: the action buttons are docked right so they stay visible at any window size / DPI;
            // the group controls and the hint share whatever width is left (the hint wraps or is cut with …)
            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 26, 0, 0),
                BackColor = colBg
            };
            actions.Controls.Add(btnApply);
            actions.Controls.Add(btnCancel);

            FlowLayoutPanel groupRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 42,
                WrapContents = false,
                BackColor = colBg
            };
            groupRow.Controls.AddRange(new Control[] { lblGroup, cboGroup, btnGroupOn, btnGroupOff });

            Panel left = new Panel { Dock = DockStyle.Fill, BackColor = colBg };
            left.Controls.Add(lblHint);
            left.Controls.Add(groupRow);

            Panel bottom = new Panel { Dock = DockStyle.Bottom, Height = 84, Padding = new Padding(0, 6, 0, 0), BackColor = colBg };
            bottom.Controls.Add(left);
            bottom.Controls.Add(actions);

            // Docking runs from the last-added control inward, so the Fill control (tabs) goes in first
            this.Controls.Add(tabs);
            this.Controls.Add(bottom);
            this.Controls.Add(lblSummary);
            this.Controls.Add(lblTitle);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        // ===================== Excel-style grid =====================

        private void BuildSheetViews()
        {
            Dictionary<string, SheetView> byName = new Dictionary<string, SheetView>();
            foreach (ChangeItem it in items)
            {
                if (it.Row <= 0 || it.Col <= 0) continue;
                SheetView v;
                if (!byName.TryGetValue(it.SheetName, out v))
                {
                    v = new SheetView { Name = it.SheetName, Data = it.Sheet };
                    if (it.Sheet != null)
                    {
                        v.Rows = it.Sheet.Rows;
                        v.Cols = it.Sheet.Cols;
                    }
                    byName[it.SheetName] = v;
                    sheetViews.Add(v);
                }
                if (v.Data == null && it.Sheet != null) v.Data = it.Sheet;

                long id = SheetState.CellId(it.Row, it.Col);
                List<ChangeItem> cell;
                if (!v.Cells.TryGetValue(id, out cell))
                {
                    cell = new List<ChangeItem>();
                    v.Cells[id] = cell;
                    v.Order.Add(id);
                }
                cell.Add(it);
            }

            foreach (SheetView v in sheetViews)
            {
                v.Order.Sort();
                int maxItemRow = 0, maxItemCol = 0;
                foreach (long id in v.Order)
                {
                    maxItemRow = Math.Max(maxItemRow, (int)(id / 1048576L));
                    maxItemCol = Math.Max(maxItemCol, (int)(id % 1048576L));
                }
                v.Rows = Math.Max(v.Rows, maxItemRow);
                v.Cols = Math.Max(Math.Min(v.Cols, MaxGridCols), maxItemCol);
            }
        }

        private void BuildGridPage(TabPage page, Color colTextMuted)
        {
            gridFont = new Font("맑은 고딕", 9F, FontStyle.Regular);
            gridBoldFont = new Font("맑은 고딕", 9F, FontStyle.Bold);
            gridStrikeFont = new Font("맑은 고딕", 9F, FontStyle.Strikeout);

            FlowLayoutPanel bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                BackColor = Color.White,
                Padding = new Padding(0, 0, 0, 4)
            };
            Label lblSheet = new Label { Text = "시트", AutoSize = true, Margin = new Padding(0, 8, 4, 0), ForeColor = colTextMuted, Font = gridFont };
            cboSheet = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(230, 24),
                Margin = new Padding(0, 4, 10, 0),
                Font = gridFont
            };
            foreach (SheetView v in sheetViews) cboSheet.Items.Add(v);
            cboSheet.SelectedIndexChanged += (s, e) => ShowSheet(cboSheet.SelectedItem as SheetView);

            Button btnPrev = new Button { Text = "◀ 이전 변경", AutoSize = true, MinimumSize = new Size(92, 26), Margin = new Padding(0, 3, 4, 0), Font = gridBoldFont, FlatStyle = FlatStyle.System };
            btnPrev.Click += (s, e) => JumpToChange(-1);
            Button btnNext = new Button { Text = "다음 변경 ▶", AutoSize = true, MinimumSize = new Size(92, 26), Margin = new Padding(0, 3, 14, 0), Font = gridBoldFont, FlatStyle = FlatStyle.System };
            btnNext.Click += (s, e) => JumpToChange(1);
            bar.Controls.AddRange(new Control[] { lblSheet, cboSheet, btnPrev, btnNext });

            AddLegend(bar, CellChanged, "반영 예정");
            AddLegend(bar, CellNewRow, "새 행");
            AddLegend(bar, CellConflict, "충돌로 덮어씀");
            AddLegend(bar, CellExcluded, "제외 (원래 값 유지)");

            lblCellInfo = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = ColorTranslator.FromHtml("#0F172A"),
                BackColor = ColorTranslator.FromHtml("#F8FAFC"),
                Font = gridFont,
                AutoEllipsis = true
            };

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                VirtualMode = true,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = true,
                ShowCellToolTips = true,
                EnableHeadersVisualStyles = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = ColorTranslator.FromHtml("#D4D4D4"),
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 24,
                RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing,
                RowHeadersWidth = 56,
                Font = gridFont
            };
            grid.RowTemplate.Height = 22;
            DataGridViewCellStyle header = new DataGridViewCellStyle
            {
                BackColor = ColorTranslator.FromHtml("#F3F3F3"),
                ForeColor = ColorTranslator.FromHtml("#444444"),
                SelectionBackColor = ColorTranslator.FromHtml("#D2D2D2"),
                SelectionForeColor = Color.Black,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Font = gridFont
            };
            grid.ColumnHeadersDefaultCellStyle = header;
            grid.RowHeadersDefaultCellStyle = header.Clone();
            grid.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#BDD7EE");
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            // Smoother scrolling on large sheets
            typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .SetValue(grid, true, null);

            grid.CellValueNeeded += (s, e) =>
            {
                if (current != null) e.Value = CellText(current, e.RowIndex + 1, e.ColumnIndex + 1);
            };
            grid.CellFormatting += Grid_CellFormatting;
            grid.CellToolTipTextNeeded += (s, e) =>
            {
                if (current == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;
                List<ChangeItem> cell;
                if (current.Cells.TryGetValue(SheetState.CellId(e.RowIndex + 1, e.ColumnIndex + 1), out cell))
                {
                    e.ToolTipText = DescribeCell(cell, "\n");
                }
            };
            grid.RowPostPaint += (s, e) =>
            {
                // Excel-style row numbers (drawn instead of stored, so huge sheets stay virtual)
                Rectangle r = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth - 1, e.RowBounds.Height - 1);
                bool currentRow = grid.CurrentCell != null && grid.CurrentCell.RowIndex == e.RowIndex;
                using (SolidBrush b = new SolidBrush(ColorTranslator.FromHtml(currentRow ? "#D2D2D2" : "#F3F3F3")))
                {
                    e.Graphics.FillRectangle(b, r);   // covers the default ▶ row indicator
                }
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), gridFont, r, ColorTranslator.FromHtml("#444444"),
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            };
            grid.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0) ToggleSelected();
            };
            grid.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Space) { ToggleSelected(); e.Handled = true; }
                else if (e.KeyCode == Keys.F3) { JumpToChange(e.Shift ? -1 : 1); e.Handled = true; }
            };
            int lastRow = -1;
            grid.SelectionChanged += (s, e) =>
            {
                int nowRow = grid.CurrentCell != null ? grid.CurrentCell.RowIndex : -1;
                if (nowRow != lastRow)
                {
                    if (lastRow >= 0 && lastRow < grid.RowCount) grid.InvalidateRow(lastRow);
                    if (nowRow >= 0) grid.InvalidateRow(nowRow);
                    lastRow = nowRow;
                }
                UpdateCellInfo();
            };

            page.Controls.Add(grid);
            page.Controls.Add(lblCellInfo);
            page.Controls.Add(bar);
        }

        // Swatch + text kept together so a wrapping toolbar never separates them
        private static void AddLegend(FlowLayoutPanel bar, Color color, string text)
        {
            FlowLayoutPanel pair = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0, 6, 12, 0),
                BackColor = Color.White
            };
            pair.Controls.Add(new Label
            {
                Size = new Size(14, 14),
                Margin = new Padding(0, 3, 4, 0),
                BackColor = color,
                BorderStyle = BorderStyle.FixedSingle
            });
            pair.Controls.Add(new Label
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(0, 1, 0, 0),
                ForeColor = ColorTranslator.FromHtml("#475569"),
                Font = new Font("맑은 고딕", 8.5F)
            });
            bar.Controls.Add(pair);
        }

        private void ShowSheet(SheetView view)
        {
            current = view;
            grid.SuspendLayout();
            grid.RowCount = 0;
            grid.Columns.Clear();
            if (view != null)
            {
                DataGridViewColumn[] cols = new DataGridViewColumn[view.Cols];
                for (int c = 1; c <= view.Cols; c++)
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn
                    {
                        HeaderText = ChangeItem.ColumnLetter(c),
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        FillWeight = 1,   // DataGridView caps the FillWeight total at 65535
                        Width = GuessWidth(view, c)
                    };
                    cols[c - 1] = col;
                }
                grid.Columns.AddRange(cols);
                grid.RowCount = view.Rows;
                // Row numbers need room on big sheets
                grid.RowHeadersWidth = Math.Max(48, TextRenderer.MeasureText(view.Rows.ToString(), gridFont).Width + 24);
            }
            grid.ResumeLayout();
            if (grid.IsHandleCreated) ShowFromTop();
            UpdateCellInfo();
        }

        // Always open a sheet at A1 like Excel, with row 1 (up to the header row) frozen at the top
        private void ShowFromTop()
        {
            if (current == null || grid.RowCount == 0 || grid.ColumnCount == 0) return;
            int frozen = 1;
            if (current.Data != null && current.Data.HeaderRow > 1) frozen = Math.Min(current.Data.HeaderRow, 10);
            if (frozen >= grid.RowCount) frozen = 0;   // nothing left to scroll; freezing would only hide rows
            for (int i = 0; i < frozen; i++) grid.Rows[i].Frozen = true;
            try
            {
                grid.ClearSelection();
                grid.CurrentCell = grid[0, 0];
                grid.CurrentCell.Selected = true;
                grid.HorizontalScrollingOffset = 0;
                if (frozen < grid.RowCount) grid.FirstDisplayedScrollingRowIndex = frozen;
            }
            catch (InvalidOperationException) { }
            UpdateCellInfo();
        }

        private int GuessWidth(SheetView view, int col)
        {
            int widest = 6;
            int sampleRows = Math.Min(view.Rows, 200);
            for (int r = 1; r <= sampleRows; r++) widest = Math.Max(widest, TextUnits(CellText(view, r, col)));
            foreach (long id in view.Order)
            {
                if ((int)(id % 1048576L) != col) continue;
                foreach (ChangeItem it in view.Cells[id]) widest = Math.Max(widest, TextUnits(it.NewText));
            }
            return Math.Max(56, Math.Min(240, widest * 7 + 18));
        }

        // Rough width in half-width characters (Korean counts as two)
        private static int TextUnits(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int n = 0;
            foreach (char ch in s) n += ch > 0x2E80 ? 2 : 1;
            return n;
        }

        private static string OriginalText(SheetView view, int row, int col, List<ChangeItem> cell)
        {
            if (view.Data != null) return view.Data.OriginalText(row, col);
            return cell != null && cell.Count > 0 ? cell[0].OldText : "";
        }

        // The last checked item wins, exactly as the merge applies it; with none checked the original stays
        private static string CellText(SheetView view, int row, int col)
        {
            List<ChangeItem> cell;
            if (!view.Cells.TryGetValue(SheetState.CellId(row, col), out cell)) return OriginalText(view, row, col, null);
            for (int i = cell.Count - 1; i >= 0; i--)
            {
                if (cell[i].Approved) return cell[i].NewText;
            }
            return OriginalText(view, row, col, cell);
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (current == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            int row = e.RowIndex + 1;
            string text = e.Value as string;
            double num;
            if (!string.IsNullOrEmpty(text) && double.TryParse(text, out num))
            {
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;   // numbers right-aligned like Excel
            }

            List<ChangeItem> cell;
            if (current.Cells.TryGetValue(SheetState.CellId(row, e.ColumnIndex + 1), out cell))
            {
                bool anyApproved = false, conflict = false, newRow = false;
                foreach (ChangeItem it in cell)
                {
                    if (!it.Approved) continue;
                    anyApproved = true;
                    if (it.Note.StartsWith("충돌")) conflict = true;
                    if (it.Kind == ChangeKind.NewRow) newRow = true;
                }
                if (!anyApproved)
                {
                    e.CellStyle.BackColor = CellExcluded;
                    e.CellStyle.ForeColor = ColorTranslator.FromHtml("#7F7F7F");
                }
                else
                {
                    e.CellStyle.BackColor = conflict ? CellConflict : (newRow ? CellNewRow : CellChanged);
                    e.CellStyle.Font = gridBoldFont;
                    e.CellStyle.ForeColor = ColorTranslator.FromHtml("#1F2937");
                }
                e.CellStyle.SelectionBackColor = Blend(e.CellStyle.BackColor, grid.DefaultCellStyle.SelectionBackColor);
                e.CellStyle.SelectionForeColor = e.CellStyle.ForeColor;
            }
            else if (current.Data != null && row == current.Data.HeaderRow)
            {
                e.CellStyle.BackColor = CellHeaderRow;
                e.CellStyle.Font = gridBoldFont;
            }
        }

        private static Color Blend(Color a, Color b)
        {
            return Color.FromArgb((a.R + b.R) / 2, (a.G + b.G) / 2, (a.B + b.B) / 2);
        }

        private static string DescribeCell(List<ChangeItem> cell, string separator)
        {
            List<string> lines = new List<string>();
            foreach (ChangeItem it in cell)
            {
                string line = string.Format("[{0}] {1} → {2}  ({3})",
                    it.FileName, string.IsNullOrEmpty(it.OldText) ? "(빈값)" : it.OldText,
                    string.IsNullOrEmpty(it.NewText) ? "(빈값)" : it.NewText, it.Approved ? "반영" : "제외");
                if (!string.IsNullOrEmpty(it.Note)) line += " · " + it.Note;
                lines.Add(line);
            }
            return string.Join(separator, lines.ToArray());
        }

        private void UpdateCellInfo()
        {
            if (lblCellInfo == null) return;
            if (current == null || grid.CurrentCell == null)
            {
                lblCellInfo.Text = "  바뀌는 칸은 색으로 표시됩니다. 칸에 마우스를 올리면 기존 값과 새 값을 볼 수 있습니다.";
                return;
            }
            int row = grid.CurrentCell.RowIndex + 1, col = grid.CurrentCell.ColumnIndex + 1;
            string address = ChangeItem.ColumnLetter(col) + row;
            List<ChangeItem> cell;
            if (current.Cells.TryGetValue(SheetState.CellId(row, col), out cell))
            {
                lblCellInfo.Text = "  " + address + "   " + DescribeCell(cell, "   |   ");
            }
            else
            {
                string v = CellText(current, row, col);
                lblCellInfo.Text = "  " + address + "   " + (v.Length > 0 ? v : "(빈칸)") + "   · 바뀌지 않는 칸";
            }
        }

        private void SelectCell(int row, int col)
        {
            if (current == null || row < 1 || col < 1 || row > grid.RowCount || col > grid.ColumnCount) return;
            try
            {
                grid.ClearSelection();
                grid.CurrentCell = grid[col - 1, row - 1];
                grid.CurrentCell.Selected = true;
            }
            catch (InvalidOperationException) { }
        }

        // direction 0 = first change, 1 = next, -1 = previous (wraps around)
        private void JumpToChange(int direction)
        {
            if (current == null || current.Order.Count == 0) return;
            long here = grid.CurrentCell != null ? SheetState.CellId(grid.CurrentCell.RowIndex + 1, grid.CurrentCell.ColumnIndex + 1) : -1;
            long target = current.Order[0];
            if (direction > 0)
            {
                target = current.Order[0];
                foreach (long id in current.Order) if (id > here) { target = id; break; }
            }
            else if (direction < 0)
            {
                target = current.Order[current.Order.Count - 1];
                for (int i = current.Order.Count - 1; i >= 0; i--) if (current.Order[i] < here) { target = current.Order[i]; break; }
            }
            SelectCell((int)(target / 1048576L), (int)(target % 1048576L));
        }

        private void ToggleSelected()
        {
            if (current == null) return;
            List<List<ChangeItem>> targets = new List<List<ChangeItem>>();
            foreach (DataGridViewCell c in grid.SelectedCells)
            {
                List<ChangeItem> cell;
                if (current.Cells.TryGetValue(SheetState.CellId(c.RowIndex + 1, c.ColumnIndex + 1), out cell)) targets.Add(cell);
            }
            ToggleCells(targets);
        }

        // Any checked → exclude all of them; none checked → include all of them
        private void ToggleCells(List<List<ChangeItem>> targets)
        {
            if (targets.Count == 0) return;
            bool anyApproved = false;
            foreach (List<ChangeItem> cell in targets)
                foreach (ChangeItem it in cell) if (it.Approved) anyApproved = true;
            int changed = 0;
            foreach (List<ChangeItem> cell in targets)
            {
                foreach (ChangeItem it in cell)
                {
                    if (it.Approved == anyApproved) { it.Approved = !anyApproved; changed++; }
                }
            }
            SyncChecks();
            UpdateSummary();
            if (grid != null) grid.Invalidate();
            UpdateCellInfo();
            DebugLog.Write("DEBUG", string.Format("미리보기 표에서 {0}칸 {1} ({2}건)", targets.Count, anyApproved ? "제외" : "반영", changed));
        }

        private void FillList()
        {
            Color newRowColor = ColorTranslator.FromHtml("#ECFDF5");
            Color conflictColor = ColorTranslator.FromHtml("#FFF7ED");
            suppressCheckEvents = true;
            list.BeginUpdate();
            try
            {
                foreach (ChangeItem it in items)
                {
                    ListViewItem row = new ListViewItem(new string[] {
                        it.FileName, it.SheetName, it.Address, it.RowKey, it.ColumnLabel, it.KindText,
                        string.IsNullOrEmpty(it.OldText) ? "(빈값)" : it.OldText, it.NewText, it.Note
                    });
                    row.Tag = it;
                    row.Checked = it.Approved;
                    if (it.Note.StartsWith("충돌")) row.BackColor = conflictColor;
                    else if (it.Kind == ChangeKind.NewRow) row.BackColor = newRowColor;
                    list.Items.Add(row);
                }
            }
            finally
            {
                list.EndUpdate();
                suppressCheckEvents = false;
            }
        }

        private void SyncChecks()
        {
            suppressCheckEvents = true;
            list.BeginUpdate();
            try
            {
                foreach (ListViewItem row in list.Items)
                {
                    bool want = ((ChangeItem)row.Tag).Approved;
                    if (row.Checked != want) row.Checked = want;
                }
            }
            finally
            {
                list.EndUpdate();
                suppressCheckEvents = false;
            }
        }

        private void UpdateSummary()
        {
            int approved = ApprovedCount;
            HashSet<string> newRows = new HashSet<string>();
            int conflicts = 0;
            foreach (ChangeItem it in items)
            {
                if (it.Kind == ChangeKind.NewRow) newRows.Add(it.SheetName + "!" + it.Row);
                if (it.Note.StartsWith("충돌")) conflicts++;
            }
            lblSummary.Text = string.Format("변경 예정 {0:N0}건 중 {1:N0}건 선택 · 새 행 {2}개 · 충돌로 덮어쓰는 항목 {3}건  (초록 = 새 행, 주황 = 충돌)",
                items.Count, approved, newRows.Count, conflicts);
            if (btnApply != null) btnApply.Text = string.Format("{0:N0}건 반영", approved);
        }
    }

    // Lets the user pick one of the automatic backups to put back in place of the base file
    public class BackupRestoreDialog : Form
    {
        private ListBox lstBackups;
        public BackupManager.BackupEntry Selected { get; private set; }

        public BackupRestoreDialog(string baseFileName, List<BackupManager.BackupEntry> backups)
        {
            Color colBg = ColorTranslator.FromHtml("#F8FAFC");
            Color colTextDark = ColorTranslator.FromHtml("#0F172A");
            Color colTextMuted = ColorTranslator.FromHtml("#64748B");

            this.SuspendLayout();
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Text = "백업으로 되돌리기";
            this.ClientSize = new Size(620, 380);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = colBg;
            this.Font = new Font("맑은 고딕", 9F);

            Label lblTitle = new Label
            {
                Text = "되돌릴 백업을 고르세요 (위쪽이 가장 최근)",
                Font = new Font("맑은 고딕", 10.5F, FontStyle.Bold),
                ForeColor = colTextDark,
                Location = new Point(18, 14),
                AutoSize = true
            };
            Label lblDesc = new Label
            {
                Text = "대상: " + baseFileName + "\n지금 원본은 복원 직전에 새 백업으로 보관되므로, 복원한 뒤에도 다시 되돌릴 수 있습니다.",
                ForeColor = colTextMuted,
                Location = new Point(18, 42),
                Size = new Size(584, 36)
            };

            lstBackups = new ListBox
            {
                Location = new Point(18, 84),
                Size = new Size(584, 230),
                Font = new Font("맑은 고딕", 9F),
                IntegralHeight = false
            };
            foreach (BackupManager.BackupEntry b in backups) lstBackups.Items.Add(b);
            if (lstBackups.Items.Count > 0) lstBackups.SelectedIndex = 0;

            SoftButton btnRestore = new SoftButton
            {
                Text = "선택한 백업으로 복원",
                Location = new Point(338, 328),
                Size = new Size(160, 38),
                BorderRadius = 12,
                NormalColor = ColorTranslator.FromHtml("#6366F1"),
                HoverColor = ColorTranslator.FromHtml("#4F46E5"),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnRestore.Click += (s, e) =>
            {
                Selected = lstBackups.SelectedItem as BackupManager.BackupEntry;
                if (Selected == null) return;
                DialogResult answer = MessageBox.Show(this,
                    string.Format("원본을 아래 백업 상태로 되돌립니다.\n\n{0}\n\n계속하시겠습니까?", Selected),
                    "복원 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes) return;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            SoftButton btnCancel = new SoftButton
            {
                Text = "닫기",
                Location = new Point(504, 328),
                Size = new Size(98, 38),
                BorderRadius = 12,
                NormalColor = ColorTranslator.FromHtml("#F1F5F9"),
                HoverColor = ColorTranslator.FromHtml("#E2E8F0"),
                BorderColor = ColorTranslator.FromHtml("#CBD5E1"),
                BorderThickness = 1,
                ForeColor = colTextDark,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            this.Controls.AddRange(new Control[] { lblTitle, lblDesc, lstBackups, btnRestore, btnCancel });
            this.ResumeLayout(false);
        }
    }
}
