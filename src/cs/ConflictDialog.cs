using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelDiffAggregator
{
    public enum ConflictResolutionChoice
    {
        UseNewValue,      // 수정본 값 적용
        KeepExistingValue // 기존 원본/이전 취합값 유지
    }

    public class ConflictDialog : Form
    {
        public ConflictResolutionChoice Choice { get; private set; }
        public bool ApplyToAll { get; private set; }

        public ConflictDialog(
            string sheetName,
            int row,
            int col,
            string modFileName,
            string existingValue,
            string newValue
        )
        {
            InitializeComponent(sheetName, row, col, modFileName, existingValue, newValue);
        }

        private void InitializeComponent(
            string sheetName,
            int row,
            int col,
            string modFileName,
            string existingValue,
            string newValue
        )
        {
            Color colBg = ColorTranslator.FromHtml("#F8FAFC");
            Color colSurface = ColorTranslator.FromHtml("#FFFFFF");
            Color colPrimary = ColorTranslator.FromHtml("#6366F1"); // Soft Indigo
            Color colBorder = ColorTranslator.FromHtml("#E2E8F0");
            Color colTextDark = ColorTranslator.FromHtml("#0F172A");
            Color colTextMuted = ColorTranslator.FromHtml("#64748B");

            this.SuspendLayout();
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Text = "셀 취합 충돌 발생 (Conflict Resolution)";
            this.ClientSize = new Size(524, 352);
            this.ShowIcon = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            // Shown from the merge's background STA thread (no owner window), so force it in front of the main form
            this.TopMost = true;
            this.ShowInTaskbar = true;
            this.BackColor = colBg;
            this.Font = new Font("맑은 고딕", 9F, FontStyle.Regular);

            // Title Bar Section
            Label lblTitle = new Label
            {
                Text = string.Format("[{0}] 시트 셀 [{1},{2}] 충돌 탐지", sheetName, row, col),
                Font = new Font("맑은 고딕", 11F, FontStyle.Bold),
                ForeColor = colTextDark,
                Location = new Point(24, 16),
                Size = new Size(476, 26),
                AutoEllipsis = true,
                UseMnemonic = false
            };
            this.Controls.Add(lblTitle);

            Label lblDesc = new Label
            {
                Text = string.Format("이전 취합으로 반영된 기존 값과 수정본 파일({0})의 변경 값이 서로 다릅니다.", modFileName),
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                ForeColor = colTextMuted,
                Location = new Point(24, 44),
                Size = new Size(476, 38),
                AutoEllipsis = true,
                UseMnemonic = false
            };
            this.Controls.Add(lblDesc);

            // Card Panel for Value Comparison (SoftPanel)
            SoftPanel pnlCard = new SoftPanel
            {
                Location = new Point(24, 90),
                Size = new Size(476, 160),
                BorderRadius = 14,
                BorderColor = colBorder,
                BackColor = colSurface
            };

            // Existing Value Section
            Label lblExistHeader = new Label
            {
                Text = "기존 반영된 값 (Existing Value):",
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#475569"),
                Location = new Point(18, 16),
                AutoSize = true
            };
            Label lblExistVal = new Label
            {
                Text = string.IsNullOrEmpty(existingValue) ? "(빈 값)" : existingValue,
                Font = new Font("맑은 고딕", 11.5F, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#D97706"),
                Location = new Point(18, 38),
                Size = new Size(440, 26),
                AutoEllipsis = true,
                UseMnemonic = false
            };
            pnlCard.Controls.AddRange(new Control[] { lblExistHeader, lblExistVal });

            Panel pnlDiv = new Panel
            {
                Location = new Point(18, 78),
                Size = new Size(440, 1),
                BackColor = colBorder
            };
            pnlCard.Controls.Add(pnlDiv);

            Label lblNewHeader = new Label
            {
                Text = string.Format("수정본({0})의 신규 입력 값:", modFileName),
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#475569"),
                Location = new Point(18, 90),
                Size = new Size(440, 18),
                AutoEllipsis = true,
                UseMnemonic = false
            };
            Label lblNewVal = new Label
            {
                Text = string.IsNullOrEmpty(newValue) ? "(빈 값)" : newValue,
                Font = new Font("맑은 고딕", 11.5F, FontStyle.Bold),
                ForeColor = colPrimary,
                Location = new Point(18, 112),
                Size = new Size(440, 26),
                AutoEllipsis = true,
                UseMnemonic = false
            };
            pnlCard.Controls.AddRange(new Control[] { lblNewHeader, lblNewVal });

            this.Controls.Add(pnlCard);

            CheckBox chkApplyAll = new CheckBox
            {
                Text = "이후 발생하는 모든 충돌에 동일한 선택 적용 (다시 묻지 않음)",
                Location = new Point(26, 262),
                AutoSize = true,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular),
                ForeColor = colTextMuted
            };
            this.Controls.Add(chkApplyAll);

            SoftButton btnUseNew = new SoftButton
            {
                Text = "수정본 값 적용",
                Location = new Point(234, 296),
                Size = new Size(130, 36),
                BorderRadius = 10,
                NormalColor = colPrimary,
                HoverColor = ColorTranslator.FromHtml("#4F46E5"),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnUseNew.Click += (s, e) => {
                this.Choice = ConflictResolutionChoice.UseNewValue;
                this.ApplyToAll = chkApplyAll.Checked;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            SoftButton btnKeepExist = new SoftButton
            {
                Text = "기존 값 유지",
                Location = new Point(370, 296),
                Size = new Size(130, 36),
                BorderRadius = 10,
                NormalColor = ColorTranslator.FromHtml("#F1F5F9"),
                HoverColor = ColorTranslator.FromHtml("#E2E8F0"),
                BorderColor = colBorder,
                BorderThickness = 1,
                ForeColor = colTextDark,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            btnKeepExist.Click += (s, e) => {
                this.Choice = ConflictResolutionChoice.KeepExistingValue;
                this.ApplyToAll = chkApplyAll.Checked;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.AddRange(new Control[] { btnUseNew, btnKeepExist });
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
