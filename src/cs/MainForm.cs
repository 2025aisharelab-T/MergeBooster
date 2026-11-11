using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelDiffAggregator
{
    // Soft & Emotional Aesthetic Button (Inherits Control directly to eliminate all Win32 black corner artifacts)
    public class SoftButton : Control
    {
        public int BorderRadius { get; set; }
        public Color NormalColor { get; set; }
        public Color HoverColor { get; set; }
        public Color BorderColor { get; set; }
        public int BorderThickness { get; set; }

        public Image ButtonImage { get; set; }

        private bool isHovered = false;
        private bool isPressed = false;

        public SoftButton()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.SupportsTransparentBackColor,
                true
            );

            this.BorderRadius = 12;
            this.NormalColor = ColorTranslator.FromHtml("#6366F1"); // Soft Indigo (#6366F1)
            this.HoverColor = ColorTranslator.FromHtml("#4F46E5");
            this.BorderColor = Color.Transparent;
            this.BorderThickness = 0;
            this.BackColor = Color.Transparent;
            this.ForeColor = Color.White;
            this.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            this.Cursor = Cursors.Hand;
            this.DoubleBuffered = true;

            this.MouseEnter += (s, e) => { isHovered = true; this.Invalidate(); };
            this.MouseLeave += (s, e) => { isHovered = false; isPressed = false; this.Invalidate(); };
            this.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { isPressed = true; this.Invalidate(); } };
            this.MouseUp += (s, e) => { isPressed = false; this.Invalidate(); };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            // Paint parent background for 100% seamless transparent corners
            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            int offset = BorderThickness > 0 ? BorderThickness : 1;
            Rectangle rect = new Rectangle(offset, offset, this.Width - (offset * 2), this.Height - (offset * 2));

            using (GraphicsPath path = GetRoundedPath(rect, BorderRadius))
            {
                Color fill = isPressed ? HoverColor : (isHovered ? HoverColor : NormalColor);
                using (SolidBrush brush = new SolidBrush(fill))
                {
                    g.FillPath(brush, path);
                }

                if (BorderThickness > 0 && BorderColor != Color.Transparent)
                {
                    using (Pen pen = new Pen(BorderColor, BorderThickness))
                    {
                        g.DrawPath(pen, path);
                    }
                }

                if (ButtonImage != null)
                {
                    Size textSize = TextRenderer.MeasureText(this.Text, this.Font);
                    int totalW = 24 + 6 + textSize.Width;
                    int startX = Math.Max(10, (this.Width - totalW) / 2);
                    int imgY = (this.Height - 24) / 2;

                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(ButtonImage, new Rectangle(startX, imgY, 24, 24));

                    Rectangle textRect = new Rectangle(startX + 28, 0, this.Width - (startX + 28), this.Height);
                    TextRenderer.DrawText(
                        g,
                        this.Text,
                        this.Font,
                        textRect,
                        this.ForeColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter
                    );
                }
                else
                {
                    TextRenderer.DrawText(
                        g,
                        this.Text,
                        this.Font,
                        this.ClientRectangle,
                        this.ForeColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                    );
                }
            }
        }

        private GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = Math.Max(1, radius * 2);
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    // Soft & Emotional Aesthetic Card Panel
    public class SoftPanel : Panel
    {
        public int BorderRadius { get; set; }
        public Color BorderColor { get; set; }

        public SoftPanel()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.SupportsTransparentBackColor,
                true
            );

            this.BorderRadius = 16;
            this.BorderColor = ColorTranslator.FromHtml("#E2E8F0");
            this.BackColor = Color.White;
            this.DoubleBuffered = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            Rectangle rect = new Rectangle(1, 1, this.Width - 3, this.Height - 3);
            using (GraphicsPath path = GetRoundedPath(rect, BorderRadius))
            {
                using (SolidBrush b = new SolidBrush(this.BackColor))
                {
                    g.FillPath(b, path);
                }

                using (Pen pen = new Pen(BorderColor, 1f))
                {
                    g.DrawPath(pen, path);
                }
            }
        }

        private GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = Math.Max(1, radius * 2);
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    // Soft macOS Style Terminal Panel
    public class SoftTerminalPanel : Panel
    {
        public int BorderRadius { get; set; }

        public SoftTerminalPanel()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.OptimizedDoubleBuffer,
                true
            );

            this.BorderRadius = 14;
            this.BackColor = ColorTranslator.FromHtml("#0F172A");
            this.DoubleBuffered = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            Rectangle rect = new Rectangle(1, 1, this.Width - 3, this.Height - 3);
            using (GraphicsPath path = GetRoundedPath(rect, BorderRadius))
            {
                using (SolidBrush mainBrush = new SolidBrush(ColorTranslator.FromHtml("#0F172A")))
                {
                    g.FillPath(mainBrush, path);
                }

                // Title bar background
                Rectangle titleRect = new Rectangle(1, 1, this.Width - 3, 34);
                using (SolidBrush titleBrush = new SolidBrush(ColorTranslator.FromHtml("#1E293B")))
                {
                    g.FillRectangle(titleBrush, titleRect);
                }

                // macOS Traffic Lights
                using (SolidBrush bRed = new SolidBrush(ColorTranslator.FromHtml("#EF4444")))
                    g.FillEllipse(bRed, 14, 11, 12, 12);
                using (SolidBrush bYellow = new SolidBrush(ColorTranslator.FromHtml("#F59E0B")))
                    g.FillEllipse(bYellow, 32, 11, 12, 12);
                using (SolidBrush bGreen = new SolidBrush(ColorTranslator.FromHtml("#10B981")))
                    g.FillEllipse(bGreen, 50, 11, 12, 12);

                TextRenderer.DrawText(
                    g,
                    "Terminal — Execution & Conflict Logs",
                    new Font("Consolas", 8.5F, FontStyle.Bold),
                    new Rectangle(0, 0, this.Width, 34),
                    ColorTranslator.FromHtml("#94A3B8"),
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                );

                using (Pen pen = new Pen(ColorTranslator.FromHtml("#334155"), 1f))
                {
                    g.DrawPath(pen, path);
                }
            }
        }

        private GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = Math.Max(1, radius * 2);
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public class ColorChipButton : Control
    {
        public Color FillColor { get; set; }
        public int OleColor { get; set; }
        public string ColorName { get; set; }
        private bool isSelected = false;

        public bool IsSelected
        {
            get { return isSelected; }
            set { isSelected = value; this.Invalidate(); }
        }

        public ColorChipButton(Color fill, int ole, string name)
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.SupportsTransparentBackColor,
                true
            );

            this.FillColor = fill;
            this.OleColor = ole;
            this.ColorName = name;
            this.Size = new Size(24, 24);
            this.Cursor = Cursors.Hand;
            this.DoubleBuffered = true;

            ToolTip tt = new ToolTip();
            tt.SetToolTip(this, name);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            Rectangle rect = new Rectangle(2, 2, 20, 20);
            using (SolidBrush b = new SolidBrush(FillColor))
            {
                g.FillEllipse(b, rect);
            }

            if (isSelected)
            {
                using (Pen pen = new Pen(ColorTranslator.FromHtml("#6366F1"), 2f))
                {
                    g.DrawEllipse(pen, rect);
                }
                TextRenderer.DrawText(g, "✓", new Font("맑은 고딕", 8F, FontStyle.Bold), new Rectangle(0, 0, this.Width, this.Height), Color.FromArgb(15, 23, 42), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            else
            {
                using (Pen pen = new Pen(Color.FromArgb(203, 213, 225), 1f))
                {
                    g.DrawEllipse(pen, rect);
                }
            }
        }
    }

    // Render User's Exact 3D Color Rocket Emoji PNG Control
    public class ColorRocketIcon : Control
    {
        private Image rocketImg;

        public ColorRocketIcon()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.DoubleBuffer |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.SupportsTransparentBackColor,
                true
            );
            this.Size = new Size(28, 28);
            this.BackColor = Color.Transparent;
            this.DoubleBuffered = true;

            string p = "rocket_hd.png";
            if (!System.IO.File.Exists(p)) p = "rocket.png";
            if (System.IO.File.Exists(p))
            {
                try { rocketImg = Image.FromFile(p); } catch {}
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            if (rocketImg != null)
            {
                g.DrawImage(rocketImg, new Rectangle(0, 0, 28, 28));
            }
        }
    }

    public class MainForm : Form
    {
        private TextBox txtBasePath;
        private CheckedListBox lstModFiles;
        private TextBox txtSheetName;
        private CheckBox chkHighlight;
        private CheckBox chkComment;
        private CheckBox chkAutoBackup;
        private ComboBox cboConflictPolicy;
        private List<ColorChipButton> colorChips = new List<ColorChipButton>();
        private int selectedColorOle = 13431551;
        private SoftButton btnRun;
        private SoftButton btnReset;
        private SoftButton btnClearBoth;
        private TextBox txtLog;
        private BackgroundWorker worker;
        private BackgroundWorker resetWorker;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Soft Rounded Emotional Design Tokens (from design.md)
            Color colBg = ColorTranslator.FromHtml("#F8FAFC");              // Light Slate Background
            Color colSurface = ColorTranslator.FromHtml("#FFFFFF");         // Pure White Card Surface
            Color colPrimary = ColorTranslator.FromHtml("#6366F1");         // Soft Indigo (#6366F1)
            Color colPrimaryHover = ColorTranslator.FromHtml("#4F46E5");
            Color colDangerBg = ColorTranslator.FromHtml("#FEF2F2");        // Soft Pink Rose-50 (#FEF2F2)
            Color colDangerHover = ColorTranslator.FromHtml("#FEE2E2");
            Color colDangerBorder = ColorTranslator.FromHtml("#FCA5A5");   // Rose-300
            Color colDangerText = ColorTranslator.FromHtml("#DC2626");     // Rose-600
            Color colSecondaryBg = ColorTranslator.FromHtml("#F1F5F9");      // Soft Slate-100
            Color colSecondaryHover = ColorTranslator.FromHtml("#E2E8F0");
            Color colSecondaryBorder = ColorTranslator.FromHtml("#CBD5E1"); // Slate-300
            Color colSecondaryText = ColorTranslator.FromHtml("#334155");   // Slate-700
            Color colTextDark = ColorTranslator.FromHtml("#0F172A");        // Deep Navy Black
            Color colTextMuted = ColorTranslator.FromHtml("#64748B");       // Muted Gray
            Color colBorder = ColorTranslator.FromHtml("#E2E8F0");          // Subtle Light Border
            Color colConsoleBg = ColorTranslator.FromHtml("#0F172A");       // Slate-900 Dark Terminal
            Color colConsoleFg = ColorTranslator.FromHtml("#38BDF8");       // Crisp Sky Blue

            this.Text = "머지부스터 (Merge Booster)";
            this.Size = new Size(760, 810);
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // 창 크기 고정 원복!
            this.MaximizeBox = false;                       // 최대화 버튼 비활성화 원복!
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("맑은 고딕", 9F, FontStyle.Regular);
            this.BackColor = colBg;

            // Load HD Rocket Icon for Window Header & Windows Taskbar
            if (System.IO.File.Exists("app_icon.ico"))
            {
                try {
                    this.Icon = new Icon("app_icon.ico");
                    this.ShowIcon = true;
                } catch {}
            }

            // Header Container (Soft 16px Rounded Card)
            SoftPanel pnlHeader = new SoftPanel
            {
                Location = new Point(24, 16),
                Size = new Size(696, 76),
                BorderRadius = 16,
                BorderColor = colBorder,
                BackColor = colSurface
            };

            // User's Exact 3D Color Rocket Emoji PNG Control
            ColorRocketIcon rocketImg = new ColorRocketIcon
            {
                Location = new Point(18, 11)
            };
            pnlHeader.Controls.Add(rocketImg);

            Label lblTitle = new Label
            {
                Text = "머지부스터 (Merge Booster)",
                Font = new Font("맑은 고딕", 14.5F, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#1E1B4B"),
                Location = new Point(50, 12),
                AutoSize = true
            };
            pnlHeader.Controls.Add(lblTitle);

            SoftButton badgePill = new SoftButton
            {
                Text = "v3.5 Merge Engine",
                Location = new Point(540, 14),
                Size = new Size(136, 28),
                BorderRadius = 14,
                NormalColor = ColorTranslator.FromHtml("#EEF2FF"),
                HoverColor = ColorTranslator.FromHtml("#E0E7FF"),
                BorderColor = ColorTranslator.FromHtml("#818CF8"),
                BorderThickness = 1,
                ForeColor = ColorTranslator.FromHtml("#312E81"),
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            pnlHeader.Controls.Add(badgePill);

            Label lblSubTitle = new Label
            {
                Text = "퇴근 시간을 앞당기는 강력한 엔진! 취합은 머지부스터에 맡기고 여유롭게 커피 한잔 어때요?",
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#3730A3"),
                Location = new Point(22, 44),
                AutoSize = true
            };
            pnlHeader.Controls.Add(lblSubTitle);
            this.Controls.Add(pnlHeader);

            // Group 1: Base Reference File Card
            SoftPanel grpBase = new SoftPanel
            {
                Location = new Point(24, 104),
                Size = new Size(696, 82),
                BorderRadius = 16,
                BorderColor = colBorder,
                BackColor = colSurface
            };

            Label lblGrp1Title = new Label
            {
                Text = "1. 원본(기준) 엑셀 파일 지정",
                Location = new Point(20, 12),
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                ForeColor = colTextDark,
                AutoSize = true
            };

            Label lblBase = new Label { Text = "원본 파일:", Location = new Point(20, 44), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextMuted };
            
            SoftPanel pnlTxtBase = new SoftPanel
            {
                Location = new Point(88, 34),
                Size = new Size(346, 36),
                BorderRadius = 10,
                BorderColor = colSecondaryBorder,
                BackColor = ColorTranslator.FromHtml("#F8FAFC")
            };
            txtBasePath = new TextBox
            {
                Location = new Point(10, 9),
                Size = new Size(326, 18),
                ReadOnly = true,
                Font = new Font("맑은 고딕", 9F),
                BackColor = ColorTranslator.FromHtml("#F8FAFC"),
                BorderStyle = BorderStyle.None
            };
            pnlTxtBase.Controls.Add(txtBasePath);

            SoftButton btnSelectBase = new SoftButton
            {
                Text = "찾아보기...",
                Location = new Point(444, 34),
                Size = new Size(116, 36),
                BorderRadius = 10,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colSecondaryText,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnSelectBase.Click += (s, e) => {
                OpenFileDialog dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm;*.xls" };
                if (dlg.ShowDialog() == DialogResult.OK) txtBasePath.Text = dlg.FileName;
            };

            SoftButton btnOpenBase = new SoftButton
            {
                Text = "엑셀 열기",
                Location = new Point(568, 34),
                Size = new Size(110, 36),
                BorderRadius = 10,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colPrimary,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnOpenBase.Click += (s, e) => {
                OpenExcelFile(txtBasePath.Text);
            };

            grpBase.Controls.AddRange(new Control[] { lblGrp1Title, lblBase, pnlTxtBase, btnSelectBase, btnOpenBase });
            this.Controls.Add(grpBase);

            // Group 2: Target Modified Files Card
            SoftPanel grpMod = new SoftPanel
            {
                Location = new Point(24, 198),
                Size = new Size(696, 198),
                BorderRadius = 16,
                BorderColor = colBorder,
                BackColor = colSurface
            };

            Label lblGrp2Title = new Label
            {
                Text = "2. 취합 대상 수정본 엑셀 파일 목록 (⚠️ 원본 기준 셀 데이터 변경 시 취합오류 발생)",
                Location = new Point(20, 12),
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                ForeColor = colTextDark,
                AutoSize = true
            };

            SoftPanel pnlLstMod = new SoftPanel
            {
                Location = new Point(20, 38),
                Size = new Size(510, 144),
                BorderRadius = 12,
                BorderColor = colSecondaryBorder,
                BackColor = colSurface
            };

            lstModFiles = new CheckedListBox
            {
                Location = new Point(8, 8),
                Size = new Size(494, 128),
                CheckOnClick = true,
                Font = new Font("맑은 고딕", 9F),
                BorderStyle = BorderStyle.None,
                BackColor = colSurface
            };
            pnlLstMod.Controls.Add(lstModFiles);

            SoftButton btnAddMod = new SoftButton
            {
                Text = "파일 추가",
                Location = new Point(542, 38),
                Size = new Size(136, 34),
                BorderRadius = 10,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colTextDark,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnAddMod.Click += (s, e) => {
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xlsm;*.xls",
                    Multiselect = true,
                    Title = "수정본 엑셀 파일(들) 선택"
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in dlg.FileNames)
                    {
                        if (!lstModFiles.Items.Contains(file))
                        {
                            lstModFiles.Items.Add(file, true);
                        }
                    }
                }
            };

            SoftButton btnOpenMod = new SoftButton
            {
                Text = "엑셀 열기",
                Location = new Point(542, 93),
                Size = new Size(136, 34),
                BorderRadius = 10,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colPrimary,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnOpenMod.Click += (s, e) => {
                if (lstModFiles.SelectedItems.Count == 0)
                {
                    MessageBox.Show("열고자 하는 수정본 엑셀 파일을 목록에서 선택해 주세요.", "선택 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                foreach (object item in lstModFiles.SelectedItems)
                {
                    OpenExcelFile(item.ToString());
                }
            };

            SoftButton btnRemoveMod = new SoftButton
            {
                Text = "선택 삭제",
                Location = new Point(542, 148),
                Size = new Size(136, 34),
                BorderRadius = 10,
                NormalColor = colDangerBg,
                HoverColor = colDangerHover,
                BorderColor = colDangerBorder,
                BorderThickness = 1,
                ForeColor = colDangerText,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnRemoveMod.Click += (s, e) => {
                List<object> toRemove = new List<object>();
                foreach (object item in lstModFiles.SelectedItems)
                {
                    toRemove.Add(item);
                }
                foreach (object item in toRemove)
                {
                    lstModFiles.Items.Remove(item);
                }
            };

            grpMod.Controls.AddRange(new Control[] { lblGrp2Title, pnlLstMod, btnAddMod, btnOpenMod, btnRemoveMod });
            this.Controls.Add(grpMod);

            // Group 3: Options Card
            SoftPanel grpOpts = new SoftPanel
            {
                Location = new Point(24, 408),
                Size = new Size(696, 114),
                BorderRadius = 16,
                BorderColor = colBorder,
                BackColor = colSurface
            };

            Label lblGrp3Title = new Label
            {
                Text = "3. 취합 및 비교 세부 옵션",
                Location = new Point(20, 10),
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                ForeColor = colTextDark,
                AutoSize = true
            };

            // Row 1 (Y = 36)
            Label lblSheet = new Label { Text = "비교 시트:", Location = new Point(20, 39), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextMuted };

            SoftPanel pnlTxtSheet = new SoftPanel
            {
                Location = new Point(84, 34),
                Size = new Size(140, 30),
                BorderRadius = 8,
                BorderColor = colSecondaryBorder,
                BackColor = ColorTranslator.FromHtml("#F8FAFC")
            };
            txtSheetName = new TextBox
            {
                Text = "(전체 시트 자동 인식)",
                Location = new Point(4, 6),
                Size = new Size(132, 18),
                TextAlign = HorizontalAlignment.Center, // 상하좌우 완전 가운데 정렬 적용!
                Font = new Font("맑은 고딕", 8.5F),
                BackColor = ColorTranslator.FromHtml("#F8FAFC"),
                BorderStyle = BorderStyle.None
            };
            pnlTxtSheet.Controls.Add(txtSheetName);

            Panel div1 = new Panel { Location = new Point(234, 36), Size = new Size(1, 26), BackColor = colBorder };

            chkHighlight = new CheckBox { Text = "음영 표시", Checked = true, Location = new Point(246, 38), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextDark };
            
            colorChips = new List<ColorChipButton>();
            ColorChipButton yellowChip = new ColorChipButton(Color.FromArgb(255, 242, 204), 13431551, "파스텔 노랑 (#FFF2CC)") { Location = new Point(326, 36) };
            ColorChipButton mintChip   = new ColorChipButton(Color.FromArgb(226, 239, 218), 14348258, "파스텔 민트 (#E2EFDA)") { Location = new Point(354, 36) };
            ColorChipButton blueChip   = new ColorChipButton(Color.FromArgb(221, 235, 247), 16247517, "파스텔 블루 (#DDEBF7)") { Location = new Point(382, 36) };
            ColorChipButton pinkChip   = new ColorChipButton(Color.FromArgb(252, 228, 214), 14083324, "파스텔 핑크 (#FCE4D6)") { Location = new Point(410, 36) };
            ColorChipButton purpleChip = new ColorChipButton(Color.FromArgb(232, 226, 240), 15786728, "파스텔 퍼플 (#E8E2F0)") { Location = new Point(438, 36) };

            colorChips.Add(yellowChip);
            colorChips.Add(mintChip);
            colorChips.Add(blueChip);
            colorChips.Add(pinkChip);
            colorChips.Add(purpleChip);

            foreach (var chip in colorChips)
            {
                ColorChipButton btn = chip;
                btn.Click += (s, e) => SelectColorChip(btn);
            }
            SelectColorChip(yellowChip);

            Panel div2 = new Panel { Location = new Point(474, 36), Size = new Size(1, 26), BackColor = colBorder };

            chkComment = new CheckBox { Text = "변경 메모 표시", Checked = true, Location = new Point(486, 38), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextDark };

            // Row 2 (Y = 74) - Auto Backup & Conflict Policy
            chkAutoBackup = new CheckBox
            {
                Text = "취합 전 원본 자동 백업 (base_backup_날짜.xlsx)",
                Checked = true,
                Location = new Point(20, 77),
                AutoSize = true,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#15803D")
            };

            Panel div3 = new Panel { Location = new Point(340, 74), Size = new Size(1, 26), BackColor = colBorder };

            Label lblConflict = new Label
            {
                Text = "충돌 처리 규칙:",
                Location = new Point(350, 78),
                AutoSize = true,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                ForeColor = colTextDark
            };

            cboConflictPolicy = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(446, 75),
                Size = new Size(232, 24),
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular),
                BackColor = colSurface
            };
            cboConflictPolicy.Items.Add("마지막 파일 우선 (덮어쓰기)");
            cboConflictPolicy.Items.Add("첫 번째 파일 우선 (최초값 보존)");
            cboConflictPolicy.Items.Add("수치 합산 (숫자 데이터 누적)");
            cboConflictPolicy.Items.Add("사용자 선택 대화상자 (팝업)");
            cboConflictPolicy.SelectedIndex = 0;

            grpOpts.Controls.AddRange(new Control[] {
                lblGrp3Title, lblSheet, pnlTxtSheet, div1, chkHighlight,
                yellowChip, mintChip, blueChip, pinkChip, purpleChip, div2, chkComment,
                chkAutoBackup, div3, lblConflict, cboConflictPolicy
            });
            this.Controls.Add(grpOpts);

            // Primary Soft Action Button (Soft Indigo #6366F1, 42px height, 21px pill rounding)
            btnRun = new SoftButton
            {
                Text = "선택한 수정본 파일(들) 순차 취합 실행",
                ButtonImage = null,
                Location = new Point(24, 534),
                Size = new Size(350, 42),
                BorderRadius = 21,
                NormalColor = colPrimary,
                HoverColor = colPrimaryHover,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold)
            };
            btnRun.Click += BtnRun_Click;
            this.Controls.Add(btnRun);

            // Danger Action Button (Soft Pink #FEF2F2, 42px height, 21px pill rounding)
            btnReset = new SoftButton
            {
                Text = "원본 메모 & 음영 일괄 제거",
                ButtonImage = null,
                Location = new Point(382, 534),
                Size = new Size(210, 42),
                BorderRadius = 21,
                NormalColor = colDangerBg,
                HoverColor = colDangerHover,
                BorderColor = colDangerBorder,
                BorderThickness = 1,
                ForeColor = colDangerText,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnReset.Click += BtnReset_Click;
            this.Controls.Add(btnReset);

            // Secondary Clear Both Button
            btnClearBoth = new SoftButton
            {
                Text = "전체 비우기",
                ButtonImage = null,
                Location = new Point(600, 534),
                Size = new Size(120, 42),
                BorderRadius = 21,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colTextMuted,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
            };
            btnClearBoth.Click += (s, e) => {
                txtBasePath.Text = "";
                lstModFiles.Items.Clear();
                AppendLog("원본 파일 경로 및 수정본 파일 목록이 모두 비워졌습니다.");
            };
            this.Controls.Add(btnClearBoth);

            // macOS Terminal Panel (#0F172A Slate-900 Dark Terminal)
            SoftTerminalPanel pnlConsole = new SoftTerminalPanel
            {
                Location = new Point(24, 588),
                Size = new Size(696, 160),
                BorderRadius = 14,
                BackColor = colConsoleBg
            };

            txtLog = new TextBox
            {
                Location = new Point(14, 40),
                Size = new Size(668, 108),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = colConsoleBg,
                ForeColor = colConsoleFg,
                Font = new Font("Consolas", 9.5F, FontStyle.Regular),
                BorderStyle = BorderStyle.None
            };
            pnlConsole.Controls.Add(txtLog);
            this.Controls.Add(pnlConsole);

            worker = new BackgroundWorker();
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

            resetWorker = new BackgroundWorker();
            resetWorker.DoWork += ResetWorker_DoWork;
            resetWorker.RunWorkerCompleted += ResetWorker_RunWorkerCompleted;
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

            List<string> checkedFiles = new List<string>();
            foreach (object item in lstModFiles.CheckedItems)
            {
                checkedFiles.Add(item.ToString());
            }

            if (checkedFiles.Count == 0 && lstModFiles.Items.Count > 0)
            {
                for (int i = 0; i < lstModFiles.Items.Count; i++)
                {
                    lstModFiles.SetItemChecked(i, true);
                    checkedFiles.Add(lstModFiles.Items[i].ToString());
                }
            }

            if (checkedFiles.Count == 0)
            {
                MessageBox.Show("취합할 수정본 파일을 최소 1개 이상 목록에 추가해 주세요.", "선택 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnRun.Enabled = false;
            btnReset.Enabled = false;
            btnClearBoth.Enabled = false;
            txtLog.Clear();
            AppendLog("다중 파일 취합 작업 시작 (" + checkedFiles.Count + "개 파일 선택됨)...");

            string basePath = txtBasePath.Text;
            string[] modPaths = checkedFiles.ToArray();
            string sheetName = txtSheetName.Text.Trim();
            bool highlight = chkHighlight.Checked;
            bool comment = chkComment.Checked;
            int highlightColor = selectedColorOle;
            bool autoBackup = chkAutoBackup.Checked;
            ConflictPolicy conflictPolicy = (ConflictPolicy)cboConflictPolicy.SelectedIndex;

            AppendLog("⚙️ 적용 충돌 정책: " + cboConflictPolicy.SelectedItem.ToString());

            worker.RunWorkerAsync(new object[] { basePath, modPaths, sheetName, highlight, comment, highlightColor, autoBackup, conflictPolicy });
        }

        private void SelectColorChip(ColorChipButton targetChip)
        {
            if (targetChip == null) return;
            foreach (var chip in colorChips)
            {
                chip.IsSelected = (chip == targetChip);
            }
            selectedColorOle = targetChip.OleColor;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] args = (object[])e.Argument;
            int totalCount = ExcelDiffEngine.ProcessDiffMultiple(
                (string)args[0],
                (string[])args[1],
                (string)args[2],
                (bool)args[3],
                (bool)args[4],
                (int)args[5],
                (bool)args[6],
                (ConflictPolicy)args[7],
                AppendLog
            );
            e.Result = totalCount;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnRun.Enabled = true;
            btnReset.Enabled = true;
            btnClearBoth.Enabled = true;
            if (e.Error != null)
            {
                AppendLog("❌ 오류 발생: " + e.Error.Message);
                MessageBox.Show("취합 중 오류가 발생했습니다:\n" + e.Error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                int count = (int)e.Result;
                AppendLog("🎉 취합 완료! 총 " + count + "개 셀 변경사항 누적 반영됨.");
                MessageBox.Show("성공적으로 취합이 완료되었습니다!\n총 " + count + "개 셀이 누적 반영 및 하이라이트 되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBasePath.Text))
            {
                MessageBox.Show("초기화할 원본 파일 경로를 먼저 선택해 주세요.", "경로 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(
                "선택한 원본 파일의 시트에서 취합으로 반영된 메모(주석)와 하이라이트 음영만 선택 초기화하시겠습니까?\n(원본 고유 서식 및 사용자 개별 주석은 그대로 유지됩니다)\n\n대상 파일: " + System.IO.Path.GetFileName(txtBasePath.Text),
                "메모 및 음영 선택 초기화 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dr != DialogResult.Yes) return;

            btnRun.Enabled = false;
            btnReset.Enabled = false;
            btnClearBoth.Enabled = false;
            txtLog.Clear();
            AppendLog("원본 파일의 메모 및 음영 제거 작업 시작...");

            string basePath = txtBasePath.Text;
            string sheetName = txtSheetName.Text.Trim();

            resetWorker.RunWorkerAsync(new object[] { basePath, sheetName });
        }

        private void ResetWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] args = (object[])e.Argument;
            int count = ExcelDiffEngine.ResetHighlightsAndComments(
                (string)args[0],
                (string)args[1],
                AppendLog
            );
            e.Result = count;
        }

        private void ResetWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnRun.Enabled = true;
            btnReset.Enabled = true;
            btnClearBoth.Enabled = true;
            if (e.Error != null)
            {
                AppendLog("❌ 초기화 중 오류 발생: " + e.Error.Message);
                MessageBox.Show("초기화 중 오류가 발생했습니다:\n" + e.Error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                int count = (int)e.Result;
                AppendLog("🎉 초기화 완료! 총 " + count + "개 시트의 메모 및 음영이 깨끗하게 제거되었습니다.");
                MessageBox.Show("성공적으로 원본 파일의 메모 및 음영 배경색이 제거되었습니다!", "초기화 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void OpenExcelFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("엑셀 파일 경로가 지정되지 않았습니다.", "경로 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                string fullPath = System.IO.Path.GetFullPath(filePath);
                if (!System.IO.File.Exists(fullPath))
                {
                    MessageBox.Show("해당 엑셀 파일이 존재하지 않습니다:\n" + fullPath, "파일 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                System.Diagnostics.Process.Start(fullPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("엑셀 파일 실행 중 오류가 발생했습니다:\n" + ex.Message, "실행 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
