namespace GhiIssue
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            btnToggleCategory = new Button();
            btnTheme = new Button();
            btnSyncToken = new Button();
            btnBulkEdit = new Button();
            btnCreateTicket = new Button();
            dgvCreateTickets = new DataGridView();
            tabPage2 = new TabPage();
            btnSearch = new Button();
            btnViewOpen = new Button();
            dgvTickets = new DataGridView();
            btnExecute = new Button();
            cboEmployees = new ComboBox();
            statusStrip1 = new StatusStrip();
            lblStatusCount = new ToolStripStatusLabel();
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvCreateTickets).BeginInit();
            tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvTickets).BeginInit();
            statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            tabControl1.Location = new Point(12, 12);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1224, 632);
            tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(btnToggleCategory);
            tabPage1.Controls.Add(btnTheme);
            tabPage1.Controls.Add(btnSyncToken);
            tabPage1.Controls.Add(btnBulkEdit);
            tabPage1.Controls.Add(btnCreateTicket);
            tabPage1.Controls.Add(dgvCreateTickets);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1216, 604);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "TẠO PHIẾU";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnToggleCategory
            // 
            btnToggleCategory.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnToggleCategory.BackColor = Color.Cornsilk;
            btnToggleCategory.Cursor = Cursors.Hand;
            btnToggleCategory.FlatAppearance.BorderSize = 0;
            btnToggleCategory.FlatStyle = FlatStyle.Flat;
            btnToggleCategory.Location = new Point(130, 475);
            btnToggleCategory.Name = "btnToggleCategory";
            btnToggleCategory.Size = new Size(120, 120);
            btnToggleCategory.TabIndex = 8;
            btnToggleCategory.Text = "👁️ Bật/Tắt Phân loại";
            btnToggleCategory.UseVisualStyleBackColor = false;
            // 
            // btnTheme
            // 
            btnTheme.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnTheme.BackColor = Color.Gainsboro;
            btnTheme.Cursor = Cursors.Hand;
            btnTheme.FlatAppearance.BorderSize = 0;
            btnTheme.FlatStyle = FlatStyle.Flat;
            btnTheme.Location = new Point(6, 475);
            btnTheme.Name = "btnTheme";
            btnTheme.Size = new Size(118, 57);
            btnTheme.TabIndex = 7;
            btnTheme.Text = "🎨 Giao diện";
            btnTheme.UseVisualStyleBackColor = false;
            // 
            // btnSyncToken
            // 
            btnSyncToken.Anchor = AnchorStyles.Bottom;
            btnSyncToken.BackColor = Color.CornflowerBlue;
            btnSyncToken.Cursor = Cursors.Hand;
            btnSyncToken.FlatAppearance.BorderSize = 0;
            btnSyncToken.FlatStyle = FlatStyle.Flat;
            btnSyncToken.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnSyncToken.Location = new Point(531, 475);
            btnSyncToken.Name = "btnSyncToken";
            btnSyncToken.Size = new Size(120, 120);
            btnSyncToken.TabIndex = 6;
            btnSyncToken.Text = "🔄 ĐỒNG BỘ";
            btnSyncToken.UseVisualStyleBackColor = false;
            // 
            // btnBulkEdit
            // 
            btnBulkEdit.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnBulkEdit.BackColor = Color.Tan;
            btnBulkEdit.Cursor = Cursors.Hand;
            btnBulkEdit.FlatAppearance.BorderSize = 0;
            btnBulkEdit.FlatStyle = FlatStyle.Flat;
            btnBulkEdit.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnBulkEdit.Location = new Point(7, 538);
            btnBulkEdit.Name = "btnBulkEdit";
            btnBulkEdit.Size = new Size(117, 57);
            btnBulkEdit.TabIndex = 3;
            btnBulkEdit.Text = "Sửa hàng loạt";
            btnBulkEdit.UseVisualStyleBackColor = false;
            // 
            // btnCreateTicket
            // 
            btnCreateTicket.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCreateTicket.BackColor = Color.Gold;
            btnCreateTicket.Cursor = Cursors.Hand;
            btnCreateTicket.FlatAppearance.BorderSize = 0;
            btnCreateTicket.FlatStyle = FlatStyle.Flat;
            btnCreateTicket.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnCreateTicket.ForeColor = SystemColors.ControlText;
            btnCreateTicket.Location = new Point(1090, 473);
            btnCreateTicket.Name = "btnCreateTicket";
            btnCreateTicket.Size = new Size(120, 120);
            btnCreateTicket.TabIndex = 2;
            btnCreateTicket.Text = "TẠO PHIẾU";
            btnCreateTicket.UseVisualStyleBackColor = false;
            // 
            // dgvCreateTickets
            // 
            dataGridViewCellStyle1.BackColor = Color.AliceBlue;
            dgvCreateTickets.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dgvCreateTickets.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvCreateTickets.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = SystemColors.Window;
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            dataGridViewCellStyle2.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;
            dgvCreateTickets.DefaultCellStyle = dataGridViewCellStyle2;
            dgvCreateTickets.Location = new Point(6, 5);
            dgvCreateTickets.Name = "dgvCreateTickets";
            dgvCreateTickets.RowTemplate.Height = 32;
            dgvCreateTickets.Size = new Size(1204, 464);
            dgvCreateTickets.TabIndex = 0;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(btnSearch);
            tabPage2.Controls.Add(btnViewOpen);
            tabPage2.Controls.Add(dgvTickets);
            tabPage2.Controls.Add(btnExecute);
            tabPage2.Controls.Add(cboEmployees);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1216, 604);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "ĐÓNG ISSUE";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnSearch
            // 
            btnSearch.Location = new Point(6, 4);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(32, 26);
            btnSearch.TabIndex = 4;
            btnSearch.Text = "🔍";
            btnSearch.UseVisualStyleBackColor = true;
            btnSearch.Click += btnSearch_Click;
            // 
            // btnViewOpen
            // 
            btnViewOpen.Location = new Point(328, 7);
            btnViewOpen.Name = "btnViewOpen";
            btnViewOpen.Size = new Size(139, 23);
            btnViewOpen.TabIndex = 3;
            btnViewOpen.Text = "Xem phiếu đang mở";
            btnViewOpen.UseVisualStyleBackColor = true;
            // 
            // dgvTickets
            // 
            dgvTickets.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvTickets.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvTickets.Location = new Point(6, 36);
            dgvTickets.Name = "dgvTickets";
            dgvTickets.RowTemplate.Height = 32;
            dgvTickets.Size = new Size(1204, 517);
            dgvTickets.TabIndex = 2;
            // 
            // btnExecute
            // 
            btnExecute.Location = new Point(247, 7);
            btnExecute.Name = "btnExecute";
            btnExecute.Size = new Size(75, 23);
            btnExecute.TabIndex = 1;
            btnExecute.Text = "ĐÓNG ISSUE";
            btnExecute.UseVisualStyleBackColor = true;
            btnExecute.Click += btnExecute_Click;
            // 
            // cboEmployees
            // 
            cboEmployees.FormattingEnabled = true;
            cboEmployees.Location = new Point(44, 5);
            cboEmployees.Name = "cboEmployees";
            cboEmployees.Size = new Size(197, 23);
            cboEmployees.TabIndex = 0;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { lblStatusCount });
            statusStrip1.Location = new Point(0, 634);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(1248, 22);
            statusStrip1.TabIndex = 7;
            statusStrip1.Text = "statusStrip1";
            // 
            // lblStatusCount
            // 
            lblStatusCount.Name = "lblStatusCount";
            lblStatusCount.Size = new Size(63, 17);
            lblStatusCount.Text = "Đếm dòng";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1248, 656);
            Controls.Add(statusStrip1);
            Controls.Add(tabControl1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Ghi Issue - v3.2";
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvCreateTickets).EndInit();
            tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvTickets).EndInit();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TabControl tabControl1;
        private TabPage tabPage2;
        private DataGridView dgvTickets;
        private Button btnExecute;
        private ComboBox cboEmployees;
        private Button btnViewOpen;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel lblStatusCount;
        private Button btnSearch;
        private TabPage tabPage1;
        private Button btnToggleCategory;
        private Button btnTheme;
        private Button btnSyncToken;
        private Button btnBulkEdit;
        private Button btnCreateTicket;
        private DataGridView dgvCreateTickets;
    }
}
