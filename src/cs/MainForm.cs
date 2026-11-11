using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDiffAggregator
{
    internal static class SoftGraphics
    {
        private static readonly float DpiScale = ReadDpiScale();

        private static float ReadDpiScale()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                return g.DpiX / 96f;
            }
        }

        // Converts a size designed at 96 DPI to device pixels
        public static int Px(int logical)
        {
            return (int)Math.Round(logical * DpiScale);
        }

        public static float Stroke(float logical)
        {
            return logical * DpiScale;
        }

        public static GraphicsPath RoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = Math.Max(1, Math.Min(radius * 2, Math.Min(rect.Width, rect.Height)));
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

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

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            isHovered = false;
            isPressed = false;
            this.Invalidate();
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            this.Invalidate();
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            this.Invalidate();
        }

        protected override void OnChangeUICues(UICuesEventArgs e)
        {
            base.OnChangeUICues(e);
            this.Invalidate();
        }

        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Enter || keyData == Keys.Space || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                e.Handled = true;
                this.OnClick(EventArgs.Empty);
            }
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

            int offset = SoftGraphics.Px(BorderThickness > 0 ? BorderThickness : 1);
            Rectangle rect = new Rectangle(offset, offset, this.Width - (offset * 2), this.Height - (offset * 2));

            using (GraphicsPath path = SoftGraphics.RoundedPath(rect, SoftGraphics.Px(BorderRadius)))
            {
                Color fill = !this.Enabled
                    ? ColorTranslator.FromHtml("#E2E8F0")
                    : (isPressed || isHovered ? HoverColor : NormalColor);
                Color textColor = this.Enabled ? this.ForeColor : ColorTranslator.FromHtml("#94A3B8");
                using (SolidBrush brush = new SolidBrush(fill))
                {
                    g.FillPath(brush, path);
                }

                if (this.Enabled && BorderThickness > 0 && BorderColor != Color.Transparent)
                {
                    using (Pen pen = new Pen(BorderColor, SoftGraphics.Stroke(BorderThickness)))
                    {
                        g.DrawPath(pen, path);
                    }
                }

                if (this.Focused && this.ShowFocusCues)
                {
                    Rectangle ringRect = Rectangle.Inflate(rect, -SoftGraphics.Px(2), -SoftGraphics.Px(2));
                    using (GraphicsPath ring = SoftGraphics.RoundedPath(ringRect, SoftGraphics.Px(Math.Max(1, BorderRadius - 2))))
                    using (Pen ringPen = new Pen(ColorTranslator.FromHtml("#818CF8"), SoftGraphics.Stroke(1.5f)))
                    {
                        g.DrawPath(ringPen, ring);
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
                        textColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix
                    );
                }
                else
                {
                    TextRenderer.DrawText(
                        g,
                        this.Text,
                        this.Font,
                        this.ClientRectangle,
                        textColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix
                    );
                }
            }
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

            // Fill before enabling anti-aliasing so edge pixels are fully covered
            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            g.SmoothingMode = SmoothingMode.AntiAlias;
            int inset = SoftGraphics.Px(1);
            Rectangle rect = new Rectangle(inset, inset, this.Width - inset * 2 - 1, this.Height - inset * 2 - 1);
            using (GraphicsPath path = SoftGraphics.RoundedPath(rect, SoftGraphics.Px(BorderRadius)))
            {
                using (SolidBrush b = new SolidBrush(this.BackColor))
                {
                    g.FillPath(b, path);
                }

                using (Pen pen = new Pen(BorderColor, SoftGraphics.Stroke(1f)))
                {
                    g.DrawPath(pen, path);
                }
            }
        }
    }

    // Read-only path display that keeps the file name visible by eliding the middle of long paths
    public class PathField : Control
    {
        public string Placeholder { get; set; }
        public Color PlaceholderColor { get; set; }

        private readonly ToolTip tip = new ToolTip();

        public PathField()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer,
                true
            );
            this.SetStyle(ControlStyles.Selectable, false);
            this.TabStop = false;
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            tip.SetToolTip(this, this.Text);
            this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(this.BackColor);
            bool empty = string.IsNullOrEmpty(this.Text);
            TextRenderer.DrawText(
                e.Graphics,
                empty ? Placeholder : this.Text,
                this.Font,
                this.ClientRectangle,
                empty ? PlaceholderColor : this.ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                TextFormatFlags.PathEllipsis | TextFormatFlags.NoPrefix
            );
        }
    }

    // Drop-down list drawn to match the rounded input fields
    public class SoftComboBox : ComboBox
    {
        public int BorderRadius { get; set; }
        public Color BorderColor { get; set; }
        public Color FocusBorderColor { get; set; }
        public Color ArrowColor { get; set; }
        public Color ItemHighlightColor { get; set; }

        public SoftComboBox()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer,
                true
            );
            this.DropDownStyle = ComboBoxStyle.DropDownList;
            this.DrawMode = DrawMode.OwnerDrawFixed;
            this.ItemHeight = SoftGraphics.Px(28);
            this.BorderRadius = 8;
        }

        protected override void OnSelectedIndexChanged(EventArgs e)
        {
            base.OnSelectedIndexChanged(e);
            this.Invalidate();
        }

        protected override void OnDropDownClosed(EventArgs e)
        {
            base.OnDropDownClosed(e);
            this.Invalidate();
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            this.Invalidate();
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;

            // Fill before enabling anti-aliasing so edge pixels are fully covered
            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            g.SmoothingMode = SmoothingMode.AntiAlias;
            int inset = SoftGraphics.Px(1);
            Rectangle rect = new Rectangle(inset, inset, this.Width - inset * 2 - 1, this.Height - inset * 2 - 1);
            using (GraphicsPath path = SoftGraphics.RoundedPath(rect, SoftGraphics.Px(BorderRadius)))
            {
                using (SolidBrush fill = new SolidBrush(this.BackColor))
                {
                    g.FillPath(fill, path);
                }
                using (Pen border = new Pen(this.Focused ? FocusBorderColor : BorderColor, SoftGraphics.Stroke(1f)))
                {
                    g.DrawPath(border, path);
                }
            }

            string text = this.SelectedItem == null ? "" : this.SelectedItem.ToString();
            TextRenderer.DrawText(
                g,
                text,
                this.Font,
                new Rectangle(SoftGraphics.Px(10), 0, this.Width - SoftGraphics.Px(38), this.Height),
                this.ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix
            );

            int cx = this.Width - SoftGraphics.Px(18);
            int cy = this.Height / 2;
            int arrowHalfWidth = SoftGraphics.Px(4);
            int arrowHalfHeight = SoftGraphics.Px(2);
            using (Pen arrow = new Pen(ArrowColor, SoftGraphics.Stroke(1.5f)))
            {
                g.DrawLines(arrow, new[] {
                    new Point(cx - arrowHalfWidth, cy - arrowHalfHeight),
                    new Point(cx, cy + arrowHalfHeight),
                    new Point(cx + arrowHalfWidth, cy - arrowHalfHeight)
                });
            }
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            bool highlighted = (e.State & DrawItemState.Selected) != 0;
            using (SolidBrush fill = new SolidBrush(highlighted ? ItemHighlightColor : this.BackColor))
            {
                e.Graphics.FillRectangle(fill, e.Bounds);
            }

            Rectangle textRect = new Rectangle(e.Bounds.X + SoftGraphics.Px(10), e.Bounds.Y, e.Bounds.Width - SoftGraphics.Px(12), e.Bounds.Height);
            TextRenderer.DrawText(
                e.Graphics,
                this.Items[e.Index].ToString(),
                this.Font,
                textRect,
                this.ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix
            );
        }
    }

    // Soft macOS Style Terminal Panel
    public class SoftTerminalPanel : Panel
    {
        private static readonly Font TitleFont = new Font("Consolas", 8.5F, FontStyle.Bold);

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

            // Fill before enabling anti-aliasing so edge pixels are fully covered
            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            g.SmoothingMode = SmoothingMode.AntiAlias;
            int inset = SoftGraphics.Px(1);
            Rectangle rect = new Rectangle(inset, inset, this.Width - inset * 2 - 1, this.Height - inset * 2 - 1);
            using (GraphicsPath path = SoftGraphics.RoundedPath(rect, SoftGraphics.Px(BorderRadius)))
            {
                using (SolidBrush mainBrush = new SolidBrush(ColorTranslator.FromHtml("#0F172A")))
                {
                    g.FillPath(mainBrush, path);
                }

                // Title bar background, following the panel's rounded top corners
                int d = Math.Max(1, SoftGraphics.Px(BorderRadius) * 2);
                int titleHeight = SoftGraphics.Px(34);
                using (GraphicsPath titlePath = new GraphicsPath())
                using (SolidBrush titleBrush = new SolidBrush(ColorTranslator.FromHtml("#1E293B")))
                {
                    titlePath.AddArc(rect.X, rect.Y, d, d, 180, 90);
                    titlePath.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                    titlePath.AddLine(rect.Right, rect.Y + titleHeight, rect.X, rect.Y + titleHeight);
                    titlePath.CloseFigure();
                    g.FillPath(titleBrush, titlePath);
                }

                // macOS Traffic Lights
                int lightY = SoftGraphics.Px(11);
                int lightSize = SoftGraphics.Px(12);
                using (SolidBrush bRed = new SolidBrush(ColorTranslator.FromHtml("#EF4444")))
                    g.FillEllipse(bRed, SoftGraphics.Px(14), lightY, lightSize, lightSize);
                using (SolidBrush bYellow = new SolidBrush(ColorTranslator.FromHtml("#F59E0B")))
                    g.FillEllipse(bYellow, SoftGraphics.Px(32), lightY, lightSize, lightSize);
                using (SolidBrush bGreen = new SolidBrush(ColorTranslator.FromHtml("#10B981")))
                    g.FillEllipse(bGreen, SoftGraphics.Px(50), lightY, lightSize, lightSize);

                TextRenderer.DrawText(
                    g,
                    "Terminal — Execution & Conflict Logs",
                    TitleFont,
                    new Rectangle(0, 0, this.Width, titleHeight),
                    ColorTranslator.FromHtml("#94A3B8"),
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix
                );

                using (Pen pen = new Pen(ColorTranslator.FromHtml("#334155"), SoftGraphics.Stroke(1f)))
                {
                    g.DrawPath(pen, path);
                }
            }
        }
    }

    public enum ProgressState
    {
        Idle,
        Running,
        Done,
        Error
    }

    // Status line (what is happening + percent) above a pill-shaped progress track
    public class SoftProgressBar : Control
    {
        private static readonly Font PercentFont = new Font("맑은 고딕", 9.5F, FontStyle.Bold);

        private int percent;
        private string status = "";
        private ProgressState state = ProgressState.Idle;

        public int Percent { get { return percent; } }
        public ProgressState State { get { return state; } }

        public SoftProgressBar()
        {
            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true
            );
            this.DoubleBuffered = true;
            this.TabStop = false;
        }

        public void SetProgress(ProgressState newState, int newPercent, string newStatus)
        {
            newPercent = Math.Max(0, Math.Min(100, newPercent));
            newStatus = newStatus ?? "";
            if (newState == state && newPercent == percent && newStatus == status) return;

            state = newState;
            percent = newPercent;
            status = newStatus;
            this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            if (this.Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(this.Parent.BackColor))
                {
                    g.FillRectangle(parentBrush, this.ClientRectangle);
                }
            }

            Color fillColor;
            Color statusColor;
            Color percentColor;
            switch (state)
            {
                case ProgressState.Running:
                    fillColor = ColorTranslator.FromHtml("#6366F1");     // primary
                    statusColor = ColorTranslator.FromHtml("#334155");
                    percentColor = ColorTranslator.FromHtml("#4F46E5");
                    break;
                case ProgressState.Done:
                    fillColor = ColorTranslator.FromHtml("#10B981");     // emerald
                    statusColor = ColorTranslator.FromHtml("#334155");
                    percentColor = ColorTranslator.FromHtml("#059669");
                    break;
                case ProgressState.Error:
                    fillColor = ColorTranslator.FromHtml("#DC2626");     // rose
                    statusColor = ColorTranslator.FromHtml("#DC2626");
                    percentColor = ColorTranslator.FromHtml("#DC2626");
                    break;
                default:
                    fillColor = ColorTranslator.FromHtml("#CBD5E1");
                    statusColor = ColorTranslator.FromHtml("#94A3B8");
                    percentColor = ColorTranslator.FromHtml("#94A3B8");
                    break;
            }

            int trackHeight = SoftGraphics.Px(8);
            int trackBottomGap = SoftGraphics.Px(1);
            int trackY = this.Height - trackHeight - trackBottomGap;
            int textHeight = Math.Max(1, trackY - SoftGraphics.Px(1));
            int percentWidth = SoftGraphics.Px(48);

            TextRenderer.DrawText(
                g,
                status,
                this.Font,
                new Rectangle(0, 0, Math.Max(1, this.Width - percentWidth), textHeight),
                statusColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix
            );
            TextRenderer.DrawText(
                g,
                percent + "%",
                PercentFont,
                new Rectangle(this.Width - percentWidth, 0, percentWidth, textHeight),
                percentColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix
            );

            Rectangle track = new Rectangle(0, trackY, this.Width - 1, trackHeight);
            using (GraphicsPath trackPath = SoftGraphics.RoundedPath(track, trackHeight / 2))
            using (SolidBrush trackBrush = new SolidBrush(ColorTranslator.FromHtml("#E2E8F0")))
            {
                g.FillPath(trackBrush, trackPath);
            }

            if (percent > 0 || state == ProgressState.Done)
            {
                // Never narrower than the pill's own height, otherwise a low percent draws as a sliver
                int fillWidth = Math.Max(trackHeight, (int)Math.Round(track.Width * percent / 100.0));
                Rectangle fill = new Rectangle(track.X, track.Y, Math.Min(track.Width, fillWidth), track.Height);
                using (GraphicsPath fillPath = SoftGraphics.RoundedPath(fill, trackHeight / 2))
                using (SolidBrush fillBrush = new SolidBrush(fillColor))
                {
                    g.FillPath(fillBrush, fillPath);
                }
            }
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
            this.TabStop = false;
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

            Rectangle rect = new Rectangle(SoftGraphics.Px(2), SoftGraphics.Px(2), SoftGraphics.Px(20), SoftGraphics.Px(20));
            using (SolidBrush b = new SolidBrush(FillColor))
            {
                g.FillEllipse(b, rect);
            }

            if (isSelected)
            {
                using (Pen pen = new Pen(ColorTranslator.FromHtml("#6366F1"), SoftGraphics.Stroke(2f)))
                {
                    g.DrawEllipse(pen, rect);
                }
                TextRenderer.DrawText(g, "✓", new Font("맑은 고딕", 8F, FontStyle.Bold), new Rectangle(0, 0, this.Width, this.Height), Color.FromArgb(15, 23, 42), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            else
            {
                using (Pen pen = new Pen(Color.FromArgb(203, 213, 225), SoftGraphics.Stroke(1f)))
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
            this.TabStop = false;
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
                g.DrawImage(rocketImg, this.ClientRectangle);
            }
        }
    }

    // Modified-file list entry: shows the file name and folder, keeps the full path for processing
    internal sealed class ModFileEntry
    {
        public readonly string FullPath;

        public ModFileEntry(string fullPath)
        {
            FullPath = fullPath;
        }

        public override string ToString()
        {
            string directory = System.IO.Path.GetDirectoryName(FullPath);
            string folder = System.IO.Path.GetFileName(directory);
            if (string.IsNullOrEmpty(folder)) folder = directory;
            return System.IO.Path.GetFileName(FullPath) + "   ·   " + folder;
        }
    }

    public class MainForm : Form
    {
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string appName, string idList);

        private PathField txtBasePath;
        private CheckedListBox lstModFiles;
        private Label lblModEmpty;
        private TextBox txtSheetName;
        private CheckBox chkHighlight;
        private CheckBox chkComment;
        private CheckBox chkAutoBackup;
        private SoftComboBox cboConflictPolicy;
        private List<ColorChipButton> colorChips = new List<ColorChipButton>();
        private int selectedColorOle = 13431551;
        private SoftButton btnRun;
        private SoftButton btnReset;
        private SoftButton btnClearBoth;
        private TextBox txtLog;
        private SoftProgressBar progressBar;
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

            this.SuspendLayout();
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Text = "머지부스터 (Merge Booster)";
            this.ClientSize = new Size(744, 794);
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

            Func<int, int, SoftPanel> createCard = (top, height) => new SoftPanel
            {
                Location = new Point(24, top),
                Size = new Size(696, height),
                BorderRadius = 16,
                BorderColor = colBorder,
                BackColor = colSurface
            };

            Func<string, Label> createCardTitle = text => new Label
            {
                Text = text,
                Location = new Point(20, 14),
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold),
                ForeColor = colTextDark,
                AutoSize = true
            };

            Func<string, Point, Size, SoftButton> createSecondaryButton = (text, location, size) => new SoftButton
            {
                Text = text,
                Location = location,
                Size = size,
                BorderRadius = 10,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colSecondaryText,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
            };

            // Header
            Panel pnlHeader = new Panel
            {
                Location = new Point(24, 14),
                Size = new Size(696, 60),
                BackColor = colBg
            };

            // User's Exact 3D Color Rocket Emoji PNG Control
            ColorRocketIcon rocketImg = new ColorRocketIcon
            {
                Location = new Point(2, 3)
            };
            pnlHeader.Controls.Add(rocketImg);

            Label lblTitle = new Label
            {
                Text = "머지부스터 (Merge Booster)",
                Font = new Font("맑은 고딕", 14.5F, FontStyle.Bold),
                ForeColor = colTextDark,
                Location = new Point(34, 2),
                AutoSize = true
            };
            pnlHeader.Controls.Add(lblTitle);

            Label lblVersion = new Label
            {
                Text = "v3.5 Merge Engine",
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular),
                ForeColor = colTextMuted,
                Location = new Point(496, 8),
                Size = new Size(200, 20),
                TextAlign = ContentAlignment.MiddleRight
            };
            pnlHeader.Controls.Add(lblVersion);

            Label lblSubTitle = new Label
            {
                Text = "퇴근 시간을 앞당기는 강력한 엔진! 취합은 머지부스터에 맡기고 여유롭게 커피 한잔 어때요?",
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                ForeColor = colTextMuted,
                Location = new Point(36, 36),
                AutoSize = true
            };
            pnlHeader.Controls.Add(lblSubTitle);
            this.Controls.Add(pnlHeader);

            // Group 1: Base Reference File Card
            SoftPanel grpBase = createCard(82, 90);
            Label lblGrp1Title = createCardTitle("1. 원본(기준) 엑셀 파일 지정");

            SoftPanel pnlTxtBase = new SoftPanel
            {
                Location = new Point(20, 42),
                Size = new Size(412, 36),
                BorderRadius = 10,
                BorderColor = colSecondaryBorder,
                BackColor = colBg
            };
            txtBasePath = new PathField
            {
                Location = new Point(12, 6),
                Size = new Size(388, 24),
                Font = new Font("맑은 고딕", 9F),
                BackColor = colBg,
                ForeColor = colTextDark,
                Placeholder = "원본 엑셀 파일을 선택하세요",
                PlaceholderColor = colTextMuted
            };
            pnlTxtBase.Controls.Add(txtBasePath);

            SoftButton btnSelectBase = createSecondaryButton("찾아보기...", new Point(440, 42), new Size(116, 36));
            btnSelectBase.Click += (s, e) => {
                OpenFileDialog dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm;*.xls" };
                if (dlg.ShowDialog() == DialogResult.OK) txtBasePath.Text = dlg.FileName;
            };

            SoftButton btnOpenBase = createSecondaryButton("엑셀 열기", new Point(564, 42), new Size(112, 36));
            btnOpenBase.Click += (s, e) => {
                OpenExcelFile(txtBasePath.Text);
            };

            grpBase.Controls.AddRange(new Control[] { lblGrp1Title, pnlTxtBase, btnSelectBase, btnOpenBase });
            this.Controls.Add(grpBase);

            // Group 2: Target Modified Files Card
            SoftPanel grpMod = createCard(184, 182);
            Label lblGrp2Title = createCardTitle("2. 취합 대상 수정본 엑셀 파일 목록");

            Label lblGrp2Warning = new Label
            {
                Text = "원본 기준 셀 데이터 변경 시 취합오류 발생",
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular),
                ForeColor = ColorTranslator.FromHtml("#B45309"),
                Location = new Point(376, 15),
                Size = new Size(300, 18),
                TextAlign = ContentAlignment.MiddleRight
            };

            SoftPanel pnlLstMod = new SoftPanel
            {
                Location = new Point(20, 42),
                Size = new Size(510, 124),
                BorderRadius = 12,
                BorderColor = colSecondaryBorder,
                BackColor = colSurface
            };

            lstModFiles = new CheckedListBox
            {
                Location = new Point(8, 8),
                Size = new Size(494, 108),
                CheckOnClick = true,
                Font = new Font("맑은 고딕", 9F),
                BorderStyle = BorderStyle.None,
                BackColor = colSurface
            };
            pnlLstMod.Controls.Add(lstModFiles);

            ToolTip modFileTip = new ToolTip();
            int lastTipIndex = -1;
            lstModFiles.MouseMove += (s, e) => {
                int index = lstModFiles.IndexFromPoint(e.Location);
                if (index == lastTipIndex) return;
                lastTipIndex = index;
                modFileTip.SetToolTip(lstModFiles, index >= 0 ? ((ModFileEntry)lstModFiles.Items[index]).FullPath : "");
            };

            lblModEmpty = new Label
            {
                Text = "[파일 추가]를 눌러 수정본 엑셀 파일을 추가하세요",
                Location = new Point(8, 8),
                Size = new Size(494, 108),
                Font = new Font("맑은 고딕", 9F),
                ForeColor = ColorTranslator.FromHtml("#94A3B8"),
                BackColor = colSurface,
                TextAlign = ContentAlignment.MiddleCenter
            };
            pnlLstMod.Controls.Add(lblModEmpty);
            lblModEmpty.BringToFront();

            SoftButton btnAddMod = createSecondaryButton("파일 추가", new Point(542, 42), new Size(134, 36));
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
                        AddModFile(file);
                    }
                }
            };

            SoftButton btnOpenMod = createSecondaryButton("엑셀 열기", new Point(542, 86), new Size(134, 36));
            btnOpenMod.Click += (s, e) => {
                if (lstModFiles.SelectedItems.Count == 0)
                {
                    MessageBox.Show("열고자 하는 수정본 엑셀 파일을 목록에서 선택해 주세요.", "선택 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                foreach (ModFileEntry item in lstModFiles.SelectedItems)
                {
                    OpenExcelFile(item.FullPath);
                }
            };

            SoftButton btnRemoveMod = new SoftButton
            {
                Text = "선택 삭제",
                Location = new Point(542, 130),
                Size = new Size(134, 36),
                BorderRadius = 10,
                NormalColor = colDangerBg,
                HoverColor = colDangerHover,
                BorderColor = colDangerBorder,
                BorderThickness = 1,
                ForeColor = colDangerText,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
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
                UpdateModEmptyState();
            };

            grpMod.Controls.AddRange(new Control[] { lblGrp2Title, lblGrp2Warning, pnlLstMod, btnAddMod, btnOpenMod, btnRemoveMod });
            this.Controls.Add(grpMod);

            // Group 3: Options Card
            SoftPanel grpOpts = createCard(378, 126);
            Label lblGrp3Title = createCardTitle("3. 취합 및 비교 세부 옵션");

            // Row 1: sheet name and conflict policy
            Label lblSheet = new Label { Text = "비교 시트", Location = new Point(20, 50), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextMuted };

            SoftPanel pnlTxtSheet = new SoftPanel
            {
                Location = new Point(88, 42),
                Size = new Size(252, 34),
                BorderRadius = 8,
                BorderColor = colSecondaryBorder,
                BackColor = colBg
            };
            txtSheetName = new TextBox
            {
                Text = "(전체 시트 자동 인식)",
                Location = new Point(10, 9),
                Size = new Size(232, 18),
                TextAlign = HorizontalAlignment.Center, // 상하좌우 완전 가운데 정렬 적용!
                Font = new Font("맑은 고딕", 9F),
                BackColor = colBg,
                BorderStyle = BorderStyle.None
            };
            pnlTxtSheet.Controls.Add(txtSheetName);

            Label lblConflict = new Label
            {
                Text = "충돌 처리 규칙",
                Location = new Point(356, 50),
                AutoSize = true,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                ForeColor = colTextMuted
            };

            cboConflictPolicy = new SoftComboBox
            {
                Location = new Point(446, 42),
                Size = new Size(230, 34),
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                BackColor = colBg,
                ForeColor = colTextDark,
                BorderColor = colSecondaryBorder,
                FocusBorderColor = colPrimary,
                ArrowColor = colTextMuted,
                ItemHighlightColor = ColorTranslator.FromHtml("#EEF2FF")
            };
            cboConflictPolicy.Items.Add("마지막 파일 우선 (덮어쓰기)");
            cboConflictPolicy.Items.Add("첫 번째 파일 우선 (최초값 보존)");
            cboConflictPolicy.Items.Add("수치 합산 (숫자 데이터 누적)");
            cboConflictPolicy.Items.Add("사용자 선택 대화상자 (팝업)");
            cboConflictPolicy.SelectedIndex = 0;

            // Row 2: highlight colors, comment and backup toggles
            chkHighlight = new CheckBox { Text = "음영 표시", Checked = true, Location = new Point(22, 90), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextDark };

            colorChips = new List<ColorChipButton>();
            ColorChipButton yellowChip = new ColorChipButton(Color.FromArgb(255, 242, 204), 13431551, "파스텔 노랑 (#FFF2CC)") { Location = new Point(104, 88) };
            ColorChipButton mintChip   = new ColorChipButton(Color.FromArgb(226, 239, 218), 14348258, "파스텔 민트 (#E2EFDA)") { Location = new Point(132, 88) };
            ColorChipButton blueChip   = new ColorChipButton(Color.FromArgb(221, 235, 247), 16247517, "파스텔 블루 (#DDEBF7)") { Location = new Point(160, 88) };
            ColorChipButton pinkChip   = new ColorChipButton(Color.FromArgb(252, 228, 214), 14083324, "파스텔 핑크 (#FCE4D6)") { Location = new Point(188, 88) };
            ColorChipButton purpleChip = new ColorChipButton(Color.FromArgb(232, 226, 240), 15786728, "파스텔 퍼플 (#E8E2F0)") { Location = new Point(216, 88) };

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

            chkComment = new CheckBox { Text = "변경 메모 표시", Checked = true, Location = new Point(358, 90), AutoSize = true, Font = new Font("맑은 고딕", 9F, FontStyle.Regular), ForeColor = colTextDark };

            chkAutoBackup = new CheckBox
            {
                Text = "취합 전 원본 자동 백업",
                Checked = true,
                Location = new Point(490, 90),
                AutoSize = true,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                ForeColor = colTextDark
            };
            new ToolTip().SetToolTip(chkAutoBackup, "취합 전에 원본을 같은 폴더에 원본이름_backup_날짜_시간.xlsx 로 복사합니다.");

            grpOpts.Controls.AddRange(new Control[] {
                lblGrp3Title, lblSheet, pnlTxtSheet, lblConflict, cboConflictPolicy,
                chkHighlight, yellowChip, mintChip, blueChip, pinkChip, purpleChip,
                chkComment, chkAutoBackup
            });
            this.Controls.Add(grpOpts);

            // Primary Soft Action Button
            btnRun = new SoftButton
            {
                Text = "선택한 수정본 파일(들) 순차 취합 실행",
                ButtonImage = null,
                Location = new Point(24, 516),
                Size = new Size(360, 44),
                BorderRadius = 22,
                NormalColor = colPrimary,
                HoverColor = colPrimaryHover,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold)
            };
            btnRun.Click += BtnRun_Click;
            this.Controls.Add(btnRun);

            // Danger Action Button
            btnReset = new SoftButton
            {
                Text = "원본 메모 & 음영 일괄 제거",
                ButtonImage = null,
                Location = new Point(392, 516),
                Size = new Size(212, 44),
                BorderRadius = 22,
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
                Location = new Point(612, 516),
                Size = new Size(108, 44),
                BorderRadius = 22,
                NormalColor = colSecondaryBg,
                HoverColor = colSecondaryHover,
                BorderColor = colSecondaryBorder,
                BorderThickness = 1,
                ForeColor = colSecondaryText,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
            };
            btnClearBoth.Click += (s, e) => {
                txtBasePath.Text = "";
                lstModFiles.Items.Clear();
                UpdateModEmptyState();
                progressBar.SetProgress(ProgressState.Idle, 0, "대기 중");
                AppendLog("원본 파일 경로 및 수정본 파일 목록이 모두 비워졌습니다.");
            };
            this.Controls.Add(btnClearBoth);

            // Progress: always visible (grey while idle) so the layout does not jump when a job starts
            progressBar = new SoftProgressBar
            {
                Location = new Point(24, 572),
                Size = new Size(696, 28),
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
            };
            progressBar.SetProgress(ProgressState.Idle, 0, "대기 중");
            this.Controls.Add(progressBar);

            // macOS Terminal Panel (#0F172A Slate-900 Dark Terminal)
            SoftTerminalPanel pnlConsole = new SoftTerminalPanel
            {
                Location = new Point(24, 610),
                Size = new Size(696, 160),
                BorderRadius = 14,
                BackColor = colConsoleBg
            };

            txtLog = new TextBox
            {
                Location = new Point(14, 42),
                Size = new Size(668, 106),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = colConsoleBg,
                ForeColor = colConsoleFg,
                Font = new Font("Consolas", 9.5F, FontStyle.Regular),
                BorderStyle = BorderStyle.None
            };
            txtLog.HandleCreated += (s, e) => SetWindowTheme(txtLog.Handle, "DarkMode_Explorer", null);
            pnlConsole.Controls.Add(txtLog);
            this.Controls.Add(pnlConsole);

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

            resetWorker = new BackgroundWorker();
            resetWorker.WorkerReportsProgress = true;
            resetWorker.DoWork += ResetWorker_DoWork;
            resetWorker.ProgressChanged += Worker_ProgressChanged;
            resetWorker.RunWorkerCompleted += ResetWorker_RunWorkerCompleted;

            this.ResumeLayout(false);
        }

        // Closing mid-run is what people do when they cannot tell whether the program is still working
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            bool busy = (worker != null && worker.IsBusy) || (resetWorker != null && resetWorker.IsBusy);
            if (busy && e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult answer = MessageBox.Show(
                    this,
                    "작업이 아직 진행 중입니다. 위쪽 진행률 바가 움직이고 있다면 정상 작동 중입니다.\n\n" +
                    "지금 닫으면 작업이 중간에 끊겨 원본 파일이 저장되지 않거나, 엑셀이 백그라운드에 남을 수 있습니다.\n\n" +
                    "그래도 닫으시겠습니까?",
                    "작업 진행 중",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2
                );
                if (answer != DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
            }
            base.OnFormClosing(e);
        }

        private void AddModFile(string path)
        {
            foreach (ModFileEntry existing in lstModFiles.Items)
            {
                if (string.Equals(existing.FullPath, path, StringComparison.OrdinalIgnoreCase)) return;
            }
            lstModFiles.Items.Add(new ModFileEntry(path), true);
            UpdateModEmptyState();
        }

        private void UpdateModEmptyState()
        {
            lblModEmpty.Visible = lstModFiles.Items.Count == 0;
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
            foreach (ModFileEntry item in lstModFiles.CheckedItems)
            {
                checkedFiles.Add(item.FullPath);
            }

            if (checkedFiles.Count == 0 && lstModFiles.Items.Count > 0)
            {
                for (int i = 0; i < lstModFiles.Items.Count; i++)
                {
                    lstModFiles.SetItemChecked(i, true);
                    checkedFiles.Add(((ModFileEntry)lstModFiles.Items[i]).FullPath);
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
            progressBar.SetProgress(ProgressState.Running, 0, "취합 작업을 시작합니다…");
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
            BackgroundWorker source = (BackgroundWorker)sender;
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
                AppendLog,
                (percent, status) => source.ReportProgress(percent, status)
            );
            e.Result = totalCount;
        }

        // Shared by the merge and reset workers; runs on the UI thread
        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.SetProgress(ProgressState.Running, e.ProgressPercentage, e.UserState as string);
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnRun.Enabled = true;
            btnReset.Enabled = true;
            btnClearBoth.Enabled = true;
            if (e.Error != null)
            {
                progressBar.SetProgress(ProgressState.Error, progressBar.Percent, "오류로 중단되었습니다 — 아래 로그를 확인해 주세요");
                AppendLog("❌ 오류 발생: " + e.Error.Message);
                MessageBox.Show("취합 중 오류가 발생했습니다:\n" + e.Error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                int count = (int)e.Result;
                progressBar.SetProgress(ProgressState.Done, 100, "취합 완료 · 총 " + count + "개 셀 반영");
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
            progressBar.SetProgress(ProgressState.Running, 0, "메모 & 음영 제거를 시작합니다…");
            AppendLog("원본 파일의 메모 및 음영 제거 작업 시작...");

            string basePath = txtBasePath.Text;
            string sheetName = txtSheetName.Text.Trim();

            resetWorker.RunWorkerAsync(new object[] { basePath, sheetName });
        }

        private void ResetWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker source = (BackgroundWorker)sender;
            object[] args = (object[])e.Argument;
            int count = ExcelDiffEngine.ResetHighlightsAndComments(
                (string)args[0],
                (string)args[1],
                AppendLog,
                (percent, status) => source.ReportProgress(percent, status)
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
                progressBar.SetProgress(ProgressState.Error, progressBar.Percent, "오류로 중단되었습니다 — 아래 로그를 확인해 주세요");
                AppendLog("❌ 초기화 중 오류 발생: " + e.Error.Message);
                MessageBox.Show("초기화 중 오류가 발생했습니다:\n" + e.Error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                int count = (int)e.Result;
                progressBar.SetProgress(ProgressState.Done, 100, "초기화 완료 · " + count + "개 시트 정리됨");
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
