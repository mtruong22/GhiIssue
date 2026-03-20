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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            btnCheckRecent = new Button();
            btnRetry = new Button();
            btnBulkEdit = new Button();
            btnCreateTicket = new Button();
            btnAddRow = new Button();
            dgvCreateTickets = new DataGridView();
            tabPage2 = new TabPage();
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
            tabControl1.Size = new Size(1026, 591);
            tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(btnCheckRecent);
            tabPage1.Controls.Add(btnRetry);
            tabPage1.Controls.Add(btnBulkEdit);
            tabPage1.Controls.Add(btnCreateTicket);
            tabPage1.Controls.Add(btnAddRow);
            tabPage1.Controls.Add(dgvCreateTickets);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1018, 563);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "TẠO PHIẾU";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnCheckRecent
            // 
            btnCheckRecent.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCheckRecent.BackColor = Color.LightGray;
            btnCheckRecent.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnCheckRecent.Location = new Point(692, 495);
            btnCheckRecent.Name = "btnCheckRecent";
            btnCheckRecent.Size = new Size(136, 35);
            btnCheckRecent.TabIndex = 5;
            btnCheckRecent.Text = "Kiểm tra phiếu vừa tạo";
            btnCheckRecent.UseVisualStyleBackColor = false;
            // 
            // btnRetry
            // 
            btnRetry.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnRetry.BackColor = Color.Salmon;
            btnRetry.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnRetry.Location = new Point(692, 434);
            btnRetry.Name = "btnRetry";
            btnRetry.Size = new Size(136, 35);
            btnRetry.TabIndex = 4;
            btnRetry.Text = "Gửi lại các phiếu lỗi";
            btnRetry.UseVisualStyleBackColor = false;
            // 
            // btnBulkEdit
            // 
            btnBulkEdit.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnBulkEdit.BackColor = Color.Tan;
            btnBulkEdit.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnBulkEdit.Location = new Point(190, 467);
            btnBulkEdit.Name = "btnBulkEdit";
            btnBulkEdit.Size = new Size(136, 35);
            btnBulkEdit.TabIndex = 3;
            btnBulkEdit.Text = "Sửa hàng loạt";
            btnBulkEdit.UseVisualStyleBackColor = false;
            // 
            // btnCreateTicket
            // 
            btnCreateTicket.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCreateTicket.BackColor = Color.Gold;
            btnCreateTicket.Font = new Font("Segoe UI", 15F, FontStyle.Bold);
            btnCreateTicket.ForeColor = SystemColors.ControlText;
            btnCreateTicket.Location = new Point(834, 434);
            btnCreateTicket.Name = "btnCreateTicket";
            btnCreateTicket.Size = new Size(178, 96);
            btnCreateTicket.TabIndex = 2;
            btnCreateTicket.Text = "TẠO PHIẾU";
            btnCreateTicket.UseVisualStyleBackColor = false;
            // 
            // btnAddRow
            // 
            btnAddRow.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnAddRow.BackColor = Color.Honeydew;
            btnAddRow.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            btnAddRow.Location = new Point(6, 434);
            btnAddRow.Name = "btnAddRow";
            btnAddRow.Size = new Size(178, 96);
            btnAddRow.TabIndex = 1;
            btnAddRow.Text = "➕ Thêm dòng";
            btnAddRow.UseVisualStyleBackColor = false;
            // 
            // dgvCreateTickets
            // 
            dgvCreateTickets.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvCreateTickets.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvCreateTickets.Location = new Point(6, 5);
            dgvCreateTickets.Name = "dgvCreateTickets";
            dgvCreateTickets.Size = new Size(1006, 423);
            dgvCreateTickets.TabIndex = 0;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(btnViewOpen);
            tabPage2.Controls.Add(dgvTickets);
            tabPage2.Controls.Add(btnExecute);
            tabPage2.Controls.Add(cboEmployees);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1018, 563);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "ĐÓNG ISSUE";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnViewOpen
            // 
            btnViewOpen.Location = new Point(290, 5);
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
            dgvTickets.Size = new Size(1043, 517);
            dgvTickets.TabIndex = 2;
            // 
            // btnExecute
            // 
            btnExecute.Location = new Point(209, 6);
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
            cboEmployees.Location = new Point(6, 7);
            cboEmployees.Name = "cboEmployees";
            cboEmployees.Size = new Size(197, 23);
            cboEmployees.TabIndex = 0;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { lblStatusCount });
            statusStrip1.Location = new Point(0, 593);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(1050, 22);
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
            ClientSize = new Size(1050, 615);
            Controls.Add(statusStrip1);
            Controls.Add(tabControl1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Ghi Issue - v2.0";
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
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Button btnCreateTicket;
        private Button btnAddRow;
        private DataGridView dgvCreateTickets;
        private DataGridView dgvTickets;
        private Button btnExecute;
        private ComboBox cboEmployees;
        private Button btnCheckRecent;
        private Button btnRetry;
        private Button btnBulkEdit;
        private Button btnViewOpen;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel lblStatusCount;
    }
}
