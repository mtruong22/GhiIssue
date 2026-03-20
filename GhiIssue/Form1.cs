using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace GhiIssue
{
    public partial class Form1 : Form
    {
        // ================== CẤU HÌNH & FILE HỆ THỐNG ==================
        private string OMICRM_TOKEN = "";
        private string tokenFilePath = Path.Combine(Application.StartupPath, "token.txt");
        private string draftFilePath = Path.Combine(Application.StartupPath, "draft.json");
        private string loginEmailCached = "";

        private List<Employee> employees;
        private List<CategoryItem> categoryList = new List<CategoryItem>();
        private List<TagItem> tagList = new List<TagItem>();

        private BindingList<PredefinedTitle> defaultTitles;

        private string defaultAssigneeId = "";
        private HashSet<string> recentlyClosedTicketIds = new HashSet<string>();
        private System.Windows.Forms.Timer autoSaveTimer;

        // Biến để liên kết với thanh StatusStrip trên giao diện
        private ToolStripStatusLabel myStatusLabel;

        public Form1()
        {
            this.Text = "Ghi Issue v2.0";

            if (!Directory.Exists(Application.StartupPath)) Directory.CreateDirectory(Application.StartupPath);

            InitializeComponent();
            this.Load += Form1_Load;

            // KẾT NỐI CÁC NÚT CƠ BẢN
            if (btnCreateTicket != null)
            {
                btnCreateTicket.Click -= btnCreateTicket_Click;
                btnCreateTicket.Click += btnCreateTicket_Click;
            }
            if (btnAddRow != null)
            {
                btnAddRow.Click -= btnAddRow_Click;
                btnAddRow.Click += btnAddRow_Click;
            }
            if (btnExecute != null)
            {
                btnExecute.Click -= btnExecute_Click;
                btnExecute.Click += btnExecute_Click;
            }

            // KẾT NỐI CÁC NÚT TÍNH NĂNG MỞ RỘNG
            if (this.Controls.Find("btnRetry", true).FirstOrDefault() is Button btnRetry)
            {
                btnRetry.Click -= BtnRetry_Click;
                btnRetry.Click += BtnRetry_Click;
            }
            if (this.Controls.Find("btnBulkEdit", true).FirstOrDefault() is Button btnBulkEdit)
            {
                btnBulkEdit.Click -= BtnBulkEdit_Click;
                btnBulkEdit.Click += BtnBulkEdit_Click;
            }
            if (this.Controls.Find("btnCheckRecent", true).FirstOrDefault() is Button btnCheckRecent)
            {
                btnCheckRecent.Click -= BtnCheckRecent_Click;
                btnCheckRecent.Click += BtnCheckRecent_Click;
            }
            if (this.Controls.Find("btnViewOpen", true).FirstOrDefault() is Button btnViewOpen)
            {
                btnViewOpen.Click -= BtnViewOpen_Click;
                btnViewOpen.Click += BtnViewOpen_Click;
            }

            if (dgvCreateTickets != null)
            {
                dgvCreateTickets.DataError += DgvCreateTickets_DataError;
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                dgvCreateTickets.CurrentCellDirtyStateChanged += DgvCreateTickets_CurrentCellDirtyStateChanged;
                dgvCreateTickets.EditingControlShowing += DgvCreateTickets_EditingControlShowing;
                dgvCreateTickets.CellValidating += DgvCreateTickets_CellValidating;
                dgvCreateTickets.CellEndEdit += (s, e) => SaveDraft();
                dgvCreateTickets.RowsRemoved += (s, e) => SaveDraft();

                // Kích hoạt đếm dòng
                dgvCreateTickets.SelectionChanged += (s, e) => UpdateStatusCount();
                dgvCreateTickets.RowsAdded += (s, e) => UpdateStatusCount();
                dgvCreateTickets.RowsRemoved += (s, e) => UpdateStatusCount();
            }

            if (dgvTickets != null)
            {
                // FIX LỖI: Bật mỏ neo để ép bảng Đóng Issue tự co giãn bám sát viền Form
                dgvTickets.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

                // Kích hoạt đếm dòng
                dgvTickets.SelectionChanged += (s, e) => UpdateStatusCount();
                dgvTickets.DataBindingComplete += (s, e) => UpdateStatusCount();
            }

            this.FormClosing += Form1_FormClosing;

            autoSaveTimer = new System.Windows.Forms.Timer();
            autoSaveTimer.Interval = 5000;
            autoSaveTimer.Tick += (s, e) => SaveDraft();
            autoSaveTimer.Start();
        }

        // =========================================================================
        // HÀM TỰ ĐỘNG TÌM THANH TRẠNG THÁI VÀ BƠM DỮ LIỆU
        // =========================================================================
        private void HookStatusStrip()
        {
            // Thuật toán tìm kiếm tất cả StatusStrip trên form (kể cả bị giấu trong TabControl)
            IEnumerable<StatusStrip> FindStatusStrips(Control parent)
            {
                foreach (Control c in parent.Controls)
                {
                    if (c is StatusStrip ss) yield return ss;
                    foreach (var child in FindStatusStrips(c)) yield return child;
                }
            }

            var strip = FindStatusStrips(this).FirstOrDefault();

            if (strip != null)
            {
                // FIX LỖI: Ép thanh StatusStrip thoát khỏi TabControl để hiển thị chung cho toàn Form
                if (strip.Parent != this)
                {
                    strip.Parent = this;
                    strip.Dock = DockStyle.Bottom;
                    strip.BringToFront();
                }

                // Tìm bằng Tên (nếu bạn đã đặt tên lblStatusCount)
                var foundItems = strip.Items.Find("lblStatusCount", true);
                if (foundItems.Length > 0 && foundItems[0] is ToolStripStatusLabel)
                {
                    myStatusLabel = (ToolStripStatusLabel)foundItems[0];
                }
                // Dự phòng: Nếu quên đặt tên, cứ bốc đại cái chữ đầu tiên trên thanh đó
                else if (strip.Items.Count > 0 && strip.Items[0] is ToolStripStatusLabel fallbackLbl)
                {
                    myStatusLabel = fallbackLbl;
                }
            }

            // Mông má lại cho đẹp
            if (myStatusLabel != null)
            {
                myStatusLabel.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                myStatusLabel.ForeColor = Color.DarkBlue;
                UpdateStatusCount();
            }
        }

        private void UpdateStatusCount()
        {
            try
            {
                if (myStatusLabel == null) return;

                int createTotal = 0, createSelected = 0;
                int ticketTotal = 0, ticketSelected = 0;

                if (dgvCreateTickets != null)
                {
                    createTotal = dgvCreateTickets.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
                    if (dgvCreateTickets.SelectedCells.Count > 0)
                    {
                        createSelected = dgvCreateTickets.SelectedCells.Cast<DataGridViewCell>().Select(c => c.RowIndex).Distinct().Count();
                    }
                }

                if (dgvTickets != null && dgvTickets.Rows.Count > 0)
                {
                    ticketTotal = dgvTickets.Rows.Count;
                    if (dgvTickets.SelectedCells.Count > 0)
                    {
                        ticketSelected = dgvTickets.SelectedCells.Cast<DataGridViewCell>().Select(c => c.RowIndex).Distinct().Count();
                    }
                }

                myStatusLabel.Text = $"📝 Tạo Phiếu: {createTotal} phiếu (Đang chọn: {createSelected})   |   🗂️ Đóng Issue: {ticketTotal} phiếu (Đang chọn: {ticketSelected})";
            }
            catch { }
        }

        // =========================================================================
        // TÍNH NĂNG MỞ RỘNG: RETRY, XEM PHIẾU (TẠO/ĐÓNG), SỬA HÀNG LOẠT
        // =========================================================================

        private void BtnRetry_Click(object sender, EventArgs e)
        {
            int retryCount = 0;
            foreach (DataGridViewRow row in dgvCreateTickets.Rows)
            {
                if (row.IsNewRow) continue;
                string currentResult = row.Cells["colResult"].Value?.ToString() ?? "";

                if (currentResult.Contains("❌"))
                {
                    row.Cells["colResult"].Value = ""; // Xóa lỗi để kích hoạt gửi lại
                    retryCount++;
                }
            }

            if (retryCount == 0)
            {
                MessageBox.Show("Không có phiếu nào bị lỗi để gửi lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            btnCreateTicket_Click(null, null);
        }

        private async void BtnCheckRecent_Click(object sender, EventArgs e)
        {
            Employee selectedEmp = (Employee)cboEmployees.SelectedItem;
            if (selectedEmp == null) return;

            var searchBody = new
            {
                status_filters = new[] { "active_state" },
                assignee_contact_ids = new[] { selectedEmp.Id },
                additional_layout = new[] { "object_association" },
                has_notify_report = true,
                page = "1",
                size = "1000",
                current_status = new[] { 0, 1, 2, 3 }
            };

            var searchContent = new StringContent(JsonSerializer.Serialize(searchBody), Encoding.UTF8, "application/json");
            Cursor.Current = Cursors.WaitCursor;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

                try
                {
                    HttpResponseMessage searchResponse = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/search?lng=vi&utm_source=web", searchContent);
                    string resStr = await searchResponse.Content.ReadAsStringAsync();

                    if (searchResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized || resStr.Contains("Unauthorized"))
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Phiên đăng nhập đã hết hạn!\nVui lòng đăng nhập lại.", "Cảnh Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
                        OMICRM_TOKEN = "";
                        await PerformLoginSequenceAsync();
                        return;
                    }

                    if (searchResponse.IsSuccessStatusCode)
                    {
                        OmiResponse omiData = JsonSerializer.Deserialize<OmiResponse>(resStr);
                        if (omiData?.payload?.items != null && omiData.payload.items.Count > 0)
                        {
                            dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                            foreach (var t in omiData.payload.items)
                            {
                                int idx = dgvCreateTickets.Rows.Add();
                                var row = dgvCreateTickets.Rows[idx];

                                row.Cells["colTitle"].Value = t.name;
                                row.Cells["colAssignee"].Value = selectedEmp.Id;
                                row.Cells["colResult"].Value = "☁️ Đã có trên hệ thống";

                                row.DefaultCellStyle.BackColor = Color.LightGreen;
                                row.ReadOnly = true;
                            }

                            dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                            Cursor.Current = Cursors.Default;

                            UpdateStatusCount();

                            MessageBox.Show($"Đã kéo về {omiData.payload.items.Count} phiếu đang hoạt động trên hệ thống.\nHãy kéo xuống cuối bảng để kiểm tra!", "Tải Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            Cursor.Current = Cursors.Default;
                            MessageBox.Show("Không tìm thấy phiếu nào của bạn trên hệ thống!", "Trống", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show($"Có lỗi xảy ra: {ex.Message}", "Lỗi Code");
                }
            }
        }

        private async void BtnViewOpen_Click(object sender, EventArgs e)
        {
            Employee selectedEmp = (Employee)cboEmployees.SelectedItem;
            if (selectedEmp == null) return;

            dgvTickets.DataSource = null;

            var searchBody = new
            {
                status_filters = new[] { "active_state" },
                assignee_contact_ids = new[] { selectedEmp.Id },
                additional_layout = new[] { "object_association" },
                has_notify_report = true,
                page = "1",
                size = "1000",
                current_status = new[] { 0, 1, 2, 3 }
            };

            var searchContent = new StringContent(JsonSerializer.Serialize(searchBody), Encoding.UTF8, "application/json");
            Cursor.Current = Cursors.WaitCursor;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

                try
                {
                    HttpResponseMessage searchResponse = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/search?lng=vi&utm_source=web", searchContent);
                    string resStr = await searchResponse.Content.ReadAsStringAsync();

                    if (searchResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized || resStr.Contains("Unauthorized"))
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Phiên đăng nhập đã hết hạn!\nVui lòng đăng nhập lại.", "Cảnh Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
                        OMICRM_TOKEN = "";
                        await PerformLoginSequenceAsync();
                        return;
                    }

                    if (!searchResponse.IsSuccessStatusCode)
                    {
                        MessageBox.Show($"Lỗi API Tìm kiếm: {searchResponse.StatusCode}", "Lỗi");
                        Cursor.Current = Cursors.Default;
                        return;
                    }

                    OmiResponse omiData = JsonSerializer.Deserialize<OmiResponse>(resStr);
                    List<TicketItem> allTickets = new List<TicketItem>();
                    if (omiData?.payload?.items != null) allTickets = omiData.payload.items;

                    if (allTickets.Count == 0)
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Sạch sẽ! Không có phiếu nào tồn đọng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (var ticket in allTickets) ticket.NguoiNhan = selectedEmp.Name;

                    Cursor.Current = Cursors.Default;
                    dgvTickets.DataSource = allTickets;

                    if (dgvTickets.Columns.Contains("_id")) dgvTickets.Columns["_id"].Visible = false;
                    if (dgvTickets.Columns.Contains("id")) dgvTickets.Columns["id"].Visible = false;
                    if (dgvTickets.Columns.Contains("current_status")) dgvTickets.Columns["current_status"].Visible = false;

                    if (dgvTickets.Columns.Contains("unique_id")) { dgvTickets.Columns["unique_id"].HeaderText = "ID Phiếu"; dgvTickets.Columns["unique_id"].DisplayIndex = 0; dgvTickets.Columns["unique_id"].Width = 80; }
                    if (dgvTickets.Columns.Contains("name")) { dgvTickets.Columns["name"].HeaderText = "Tên Phiếu"; dgvTickets.Columns["name"].DisplayIndex = 1; }
                    if (dgvTickets.Columns.Contains("TrangThaiHienThi")) { dgvTickets.Columns["TrangThaiHienThi"].HeaderText = "Trạng Thái"; dgvTickets.Columns["TrangThaiHienThi"].DisplayIndex = 2; dgvTickets.Columns["TrangThaiHienThi"].Width = 140; }
                    if (dgvTickets.Columns.Contains("NguoiNhan")) { dgvTickets.Columns["NguoiNhan"].HeaderText = "Người Nhận"; dgvTickets.Columns["NguoiNhan"].DisplayIndex = 3; }

                    dgvTickets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    UpdateStatusCount();

                    MessageBox.Show($"Đang hiển thị {allTickets.Count} phiếu đang mở. \n(Dữ liệu chỉ hiển thị, chưa thực hiện lệnh Đóng)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show($"Có lỗi xảy ra: {ex.Message}", "Lỗi Code");
                }
            }
        }

        private void BtnBulkEdit_Click(object sender, EventArgs e)
        {
            if (dgvCreateTickets.SelectedRows.Count < 2)
            {
                MessageBox.Show("Vui lòng bôi đen (chọn) từ 2 dòng trở lên ở cột ngoài cùng bên trái để sử dụng tính năng sửa hàng loạt!", "Hướng dẫn", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Form popup = new Form()
            {
                Width = 350,
                Height = 220,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Sửa dữ liệu hàng loạt",
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false
            };

            Label lbl1 = new Label() { Left = 20, Top = 20, Text = "Chọn Cột cần đổi:", AutoSize = true };
            ComboBox cbColumn = new ComboBox() { Left = 20, Top = 40, Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
            cbColumn.Items.AddRange(new string[] { "Tiêu đề", "Chủ đề", "Phân loại", "Tag", "Người xử lý" });

            Label lbl2 = new Label() { Left = 20, Top = 80, Text = "Chọn Giá trị mới:", AutoSize = true };
            ComboBox cbValue = new ComboBox() { Left = 20, Top = 100, Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };

            cbColumn.SelectedIndexChanged += (s, ev) =>
            {
                cbValue.DataSource = null;
                cbValue.Items.Clear();
                string selCol = cbColumn.SelectedItem.ToString();

                if (selCol == "Tiêu đề") { cbValue.DataSource = defaultTitles; cbValue.DisplayMember = "Title"; cbValue.ValueMember = "Title"; }
                else if (selCol == "Chủ đề") { cbValue.DataSource = categoryList; cbValue.DisplayMember = "name"; cbValue.ValueMember = "_id"; }
                else if (selCol == "Tag") { cbValue.DataSource = tagList; cbValue.DisplayMember = "name"; cbValue.ValueMember = "id"; }
                else if (selCol == "Người xử lý") { cbValue.DataSource = employees; cbValue.DisplayMember = "Name"; cbValue.ValueMember = "Id"; }
                else if (selCol == "Phân loại")
                {
                    string catId = dgvCreateTickets.SelectedRows[0].Cells["colCategory"].Value?.ToString();
                    var cat = categoryList.FirstOrDefault(c => c._id == catId);
                    if (cat != null && cat.types != null) { cbValue.DataSource = cat.types.ToList(); cbValue.DisplayMember = "name"; cbValue.ValueMember = "uuid"; }
                    else { MessageBox.Show("Dòng đầu tiên chưa chọn Chủ đề, không thể load Phân loại!"); }
                }
            };
            cbColumn.SelectedIndex = 0;

            Button btnApply = new Button() { Text = "Áp dụng", Left = 190, Top = 140, Width = 120, DialogResult = DialogResult.OK };
            popup.Controls.Add(lbl1); popup.Controls.Add(cbColumn);
            popup.Controls.Add(lbl2); popup.Controls.Add(cbValue);
            popup.Controls.Add(btnApply); popup.AcceptButton = btnApply;

            if (popup.ShowDialog() == DialogResult.OK)
            {
                string selCol = cbColumn.SelectedItem.ToString();
                object newVal = cbValue.SelectedValue;
                if (newVal != null)
                {
                    dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                    foreach (DataGridViewRow row in dgvCreateTickets.SelectedRows)
                    {
                        if (row.IsNewRow) continue;
                        if (selCol == "Tiêu đề")
                        {
                            row.Cells["colTitle"].Value = newVal;
                            var match = defaultTitles.FirstOrDefault(t => t.Title == newVal.ToString());
                            if (match != null) row.Cells["colGroup"].Value = match.Group;
                        }
                        else if (selCol == "Chủ đề") { row.Cells["colCategory"].Value = newVal; row.Cells["colSubCategory"].Value = null; }
                        else if (selCol == "Phân loại") { row.Cells["colSubCategory"].Value = newVal; }
                        else if (selCol == "Tag") { row.Cells["colTag"].Value = newVal; }
                        else if (selCol == "Người xử lý") { row.Cells["colAssignee"].Value = newVal; }
                    }
                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                    MessageBox.Show($"Đã thay đổi hàng loạt cho {dgvCreateTickets.SelectedRows.Count} dòng!", "Thành công");
                    UpdateStatusCount();
                }
            }
        }

        // =========================================================================

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show(
                    "Bạn có chắc chắn muốn thoát Tool Ghi Issue không?\n(Dữ liệu đang nhập dở sẽ được tự động lưu nháp an toàn)",
                    "Xác nhận thoát",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }

            if (dgvCreateTickets != null && dgvCreateTickets.IsCurrentCellInEditMode) dgvCreateTickets.EndEdit();
            SaveDraft();
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            LoadDefaultTitles();

            employees = new List<Employee>()
            {
                new Employee { Name = "THUẬN, Nguyễn (CC | HCM)", Id = "65780d66afb20447c7b6ca00" },
                new Employee { Name = "ĐỐNG, Lê (CC | HCM)", Id = "65780d63afb20447c7b6c9c4" },
                new Employee { Name = "THÁI, Đặng (CC | HCM)", Id = "6942a99f64787205142cae90" },
                new Employee { Name = "TRƯỜNG, Mai (CC | HCM)", Id = "6943a4ca9090931741267c7d" },
                new Employee { Name = "THÀNH, Đặng (CC | HCM)", Id = "65780f62ab74fd73cc5f0556" },
                new Employee { Name = "HOAN, Đinh (CC | HCM)", Id = "654207d44c09ca32ee60e80f" },
                new Employee { Name = "HUY, Phan (CC | HCM)", Id = "6942a9faf81afd12f7bd19c2" },
                new Employee { Name = "CHƯƠNG, Lê (CC | HCM)", Id = "6942a9a064787205142caeb8" },
                new Employee { Name = "DANH, Nguyễn (CC | HCM)", Id = "65780d65afb20447c7b6c9d8" },
                new Employee { Name = "Đại, Lê (CC | HCM)", Id = "6942a9a064787205142caea4" },
                new Employee { Name = "BÌNH, Tô (CC | HCM)", Id = "68142e2cc400946977b39882" },
                new Employee { Name = "THƯ, Đinh (CC | HCM)", Id = "6943a7551fdc73044ad8a2ce" },
                new Employee { Name = "Hoàng AN, Nguyễn (CC | HCM)", Id = "65780d65afb20447c7b6c9ec" },
                new Employee { Name = "DŨNG, Bùi (CC | HCM)", Id = "654207f91dcb3048bfecc2f9" },
                new Employee { Name = "KHANG, Chiêm (CC | HCM)", Id = "65780f61ab74fd73cc5f0542" },
                new Employee { Name = "TRUNG, Nguyễn (CC | HCM)", Id = "65420805dd376b76c2f9b118" }
            };

            if (cboEmployees != null)
            {
                cboEmployees.DataSource = employees;
                cboEmployees.AutoCompleteMode = AutoCompleteMode.None;
                cboEmployees.DropDownStyle = ComboBoxStyle.DropDown;
                cboEmployees.TextUpdate -= cboEmployees_TextUpdate;
                cboEmployees.TextUpdate += cboEmployees_TextUpdate;
            }

            SetupCreateTicketGrid();
            CheckAndDownloadUpdateAsync();
            await PerformLoginSequenceAsync();

            AutoAssignFromEmail(loginEmailCached);

            // GỌI HÀM HOOK THANH TRẠNG THÁI (Đã vá lỗi ở đây)
            HookStatusStrip();

            LoadDraft();
        }

        private void LoadDefaultTitles()
        {
            defaultTitles = new BindingList<PredefinedTitle>
            {
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Phần mềm POS không vào được do services không chạy" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Phần mềm POS bị treo do hệ thống" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Phần mềm POS không vào được do hết hạn license" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Phần mềm POS không in được do hạ tầng mạng" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Cài lại phần mềm POS cho khách hàng" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Phần mềm POS ngắt kết nối do bật 2 phần mềm Doshcash" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Không Đồng bộ master data POS" },
                new PredefinedTitle { Group = "Phần mềm POS", Title = "Phần mềm POS treo tạm tính không in được do bug hệ thống" },

                new PredefinedTitle { Group = "RK7 MANAGER", Title = "Phần mềm Rk7 Manager không vào được do Services không chạy" },
                new PredefinedTitle { Group = "RK7 MANAGER", Title = "Phần mềm Rk7 Manager không vào được do hệ thống máy chủ" },
                new PredefinedTitle { Group = "RK7 MANAGER", Title = "Phần mềm Rk7 Manager không vào xem báo cáo được do Services không chạy" },
                new PredefinedTitle { Group = "RK7 MANAGER", Title = "Phần mềm Rk7 Manager không vào được do hết hạn license" },
                new PredefinedTitle { Group = "RK7 MANAGER", Title = "Cài lại phần mềm RK7 Manager cho khách hàng" },

                new PredefinedTitle { Group = "KDS", Title = "Phần mềm KDS bị treo cần phải khởi động lại KDS Server" },
                new PredefinedTitle { Group = "KDS", Title = "Phần mềm KDS không vào được do Services không chạy" },
                new PredefinedTitle { Group = "KDS", Title = "Phần mềm KDS không lên món do hệ thống mạng của khách hàng" },
                new PredefinedTitle { Group = "KDS", Title = "Phần mềm KDS không vào được do hết hạn license" },
                new PredefinedTitle { Group = "KDS", Title = "Phần mềm KDS không vào được do sai cấu hình" },
                new PredefinedTitle { Group = "KDS", Title = "Cài lại phần mềm KDS cho khách hàng" },
                new PredefinedTitle { Group = "KDS", Title = "Thay đổi cấu hình phần mềm KDS theo yêu cầu của khách hàng" },

                new PredefinedTitle { Group = "SkyTab", Title = "Phần mềm SkyTab bị treo cần phải khởi động lại App Server" },
                new PredefinedTitle { Group = "SkyTab", Title = "Phần mềm SkyTab không vào được do Services không chạy" },
                new PredefinedTitle { Group = "SkyTab", Title = "Phần mềm SkyTab không vào được do hệ thống mạng của khách hàng" },
                new PredefinedTitle { Group = "SkyTab", Title = "Phần mềm SkyTab không vào được do hết hạn license XML" },
                new PredefinedTitle { Group = "SkyTab", Title = "Cài lại phần mềm SkyTab cho khách hàng" },
                new PredefinedTitle { Group = "SkyTab", Title = "Thay đổi cấu hình phần mềm SkyTab theo yêu cầu của khách hàng" },

                new PredefinedTitle { Group = "SkyMenu", Title = "Phần mềm SkyMenu bị treo cần phải khởi động lại App Server" },
                new PredefinedTitle { Group = "SkyMenu", Title = "Phần mềm SkyMenu không vào được do Services không chạy" },
                new PredefinedTitle { Group = "SkyMenu", Title = "Phần mềm SkyMenu không vào được do hệ thống mạng của khách hàng" },
                new PredefinedTitle { Group = "SkyMenu", Title = "Phần mềm SkyMenu không vào được do hết hạn license XML" },
                new PredefinedTitle { Group = "SkyMenu", Title = "Cài lại phần mềm SkyMenu cho khách hàng" },
                new PredefinedTitle { Group = "SkyMenu", Title = "Thay đổi cấu hình phần mềm SkyMenu theo yêu cầu của khách hàng" },

                new PredefinedTitle { Group = "SkyOrder", Title = "Phần mềm SkyOrder bị treo cần phải khởi động lại App Server" },
                new PredefinedTitle { Group = "SkyOrder", Title = "Phần mềm SkyOrder không vào được do Services không chạy" },
                new PredefinedTitle { Group = "SkyOrder", Title = "Phần mềm SkyOrder không vào được do hệ thống mạng của khách hàng" },
                new PredefinedTitle { Group = "SkyOrder", Title = "Phần mềm SkyOrder không vào được do hết hạn license XML" },
                new PredefinedTitle { Group = "SkyOrder", Title = "Cài lại phần mềm SkyOrder cho khách hàng" },
                new PredefinedTitle { Group = "SkyOrder", Title = "Thay đổi cấu hình phần mềm SkyOrder theo yêu cầu của khách hàng" },

                new PredefinedTitle { Group = "SkyGuestScreen", Title = "Màn hình 02 không hiển thị món do Services không chạy" },
                new PredefinedTitle { Group = "SkyGuestScreen", Title = "Màn hình 02 bị treo cần khởi động lại Services" },
                new PredefinedTitle { Group = "SkyGuestScreen", Title = "Màn hình 02 hiện sai hình ảnh" },
                new PredefinedTitle { Group = "SkyGuestScreen", Title = "Thay đổi cấu hình phần mềm SkyGuestScreen theo yêu cầu của khách hàng" },

                new PredefinedTitle { Group = "OM", Title = "OM bị mất kết nối với máy chủ" },
                new PredefinedTitle { Group = "OM", Title = "OM không xác nhận được đơn do đã có đơn tồn tại trên POS" },
                new PredefinedTitle { Group = "OM", Title = "OM hủy đơn do Agent bị treo" },
                new PredefinedTitle { Group = "OM", Title = "OM hủy đơn do Agent không chạy" },
                new PredefinedTitle { Group = "OM", Title = "OM bị hủy đơn do áp dụng sai coupon, sai giá tiền" },
                new PredefinedTitle { Group = "OM", Title = "OM hủy đơn do masterdata sai" },

                new PredefinedTitle { Group = "Payment Hub", Title = "Hệ thống thanh toán Momo bị treo" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hệ thống thanh toán Napas bị treo" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hệ thống thanh toán ZaloPay bị treo" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hệ thống thanh toán VietQR bị treo" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hệ thống thanh toán Payoo bị treo" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hỗ trợ thanh toán VietQR thủ công" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hỗ trợ thanh toán Payoo thủ công" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hỗ trợ thanh toán Zalo thủ công" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hỗ trợ thanh toán Napas thủ công" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Hỗ trợ thanh toán Momo thủ công" },
                new PredefinedTitle { Group = "Payment Hub", Title = "Server Payment Hub gặp sự cố" },

                new PredefinedTitle { Group = "CRM", Title = "CRM không tích điểm được do Services Farcard bị treo" },
                new PredefinedTitle { Group = "CRM", Title = "CRM không tích điểm được do Services Farcard bị stop" },
                new PredefinedTitle { Group = "CRM", Title = "CRM bị treo không giảm giá được do máy chủ" },
                new PredefinedTitle { Group = "CRM", Title = "CRM không sử dụng được hết hạn license" },

                new PredefinedTitle { Group = "SkyInvoice", Title = "SkyInvoice bị thiếu bill từ POS đẩy lên" },
                new PredefinedTitle { Group = "SkyInvoice", Title = "SkyInvoice bị treo do hệ thống máy chủ." },
                new PredefinedTitle { Group = "SkyInvoice", Title = "SkyInvoice bị treo do hệ thống mạng của khách hàng." },
                new PredefinedTitle { Group = "SkyInvoice", Title = "Chỉnh lại mẫu hóa đơn trên phần mềm SkyInvoice" },

                new PredefinedTitle { Group = "Phần mềm Kho", Title = "Phần mềm kho không vào được" },
                new PredefinedTitle { Group = "Phần mềm Kho", Title = "Phần mềm kho trừ hàng sai" },

                new PredefinedTitle { Group = "BÁO CÁO", Title = "Bổ sung báo cáo mới theo yêu cầu của khách hàng trên IR REPORT" },
                new PredefinedTitle { Group = "BÁO CÁO", Title = "Bổ sung thêm trường dữ liệu trên báo cáo IR REPORT cho khách hàng" },
                new PredefinedTitle { Group = "BÁO CÁO", Title = "Hỗ trợ báo cáo bị sai lệch dữ liệu trên IR Report" },
                new PredefinedTitle { Group = "BÁO CÁO", Title = "Báo cáo IR bị thiếu dữ liệu từ nhà hàng đổ về" },
                new PredefinedTitle { Group = "BÁO CÁO", Title = "Bổ sung báo cáo mới theo yêu cầu của khách hàng trên POS" },
                new PredefinedTitle { Group = "BÁO CÁO", Title = "Bổ sung thêm trường dữ liệu trên báo cáo POS cho khách hàng" },
                new PredefinedTitle { Group = "BÁO CÁO", Title = "Hỗ trợ báo cáo bị sai lệch dữ liệu trên POS" },

                new PredefinedTitle { Group = "MASTER DATA", Title = "Setup Master Data khách hàng gửi" },
                new PredefinedTitle { Group = "MASTER DATA", Title = "Hỗ trợ đồng bộ Master Data" },
                new PredefinedTitle { Group = "MASTER DATA", Title = "Hỗ trợ Master Data bị setup sai, thiếu" },

                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Máy POS không lên do hư nguồn" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Máy POS bị hư màn hình cảm ứng" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Máy POS hư ổ cứng" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Máy in không lên do hư nguồn" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Máy in bị lỗi dao cắt" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Máy in thanh nhiệt bị mờ mực in" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Két tiền bị hư chìa khóa" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Két tiền bị mất chìa khóa" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Két tiền bị kẹt giấy" },
                new PredefinedTitle { Group = "PHẦN CỨNG", Title = "Két tiền không tự động bật ra khi tính tiền" },

                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn khách hàng thao tác trên hệ thống POS" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn khách hàng thao tác trên phần mềm Kho" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn khách hàng thao tác trên phần mềm RK7 Manager" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn nhà hàng thanh toán thủ công" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn nhà hàng đóng ca chung (DayEnd)" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn nhà hàng cấu hình đăng nhập Skytab" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hướng dẫn nhà hàng mở màn hình 02" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hỗ trợ hủy bill" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hỗ trợ phục hồi bill" },
                new PredefinedTitle { Group = "Thao tác Vận Hành", Title = "Hỗ trợ hủy tạm tính" }
            };
        }

        private void AutoAssignFromEmail(string email)
        {
            if (string.IsNullOrEmpty(email)) return;
            email = email.ToLower();
            string empId = "";

            if (email.Contains("truongmh")) empId = "6943a4ca9090931741267c7d";
            else if (email.Contains("chuongld")) empId = "6942a9a064787205142caeb8";
            else if (email.Contains("thuan")) empId = "65780d66afb20447c7b6ca00";
            else if (email.Contains("dongl")) empId = "65780d63afb20447c7b6c9c4";
            else if (email.Contains("thaid")) empId = "6942a99f64787205142cae90";
            else if (email.Contains("thanhd")) empId = "65780f62ab74fd73cc5f0556";
            else if (email.Contains("hoand")) empId = "654207d44c09ca32ee60e80f";
            else if (email.Contains("huyp")) empId = "6942a9faf81afd12f7bd19c2";
            else if (email.Contains("danhn")) empId = "65780d65afb20447c7b6c9d8";
            else if (email.Contains("dail")) empId = "6942a9a064787205142caea4";
            else if (email.Contains("binht")) empId = "68142e2cc400946977b39882";
            else if (email.Contains("thud")) empId = "6943a7551fdc73044ad8a2ce";
            else if (email.Contains("hoangan")) empId = "65780d65afb20447c7b6c9ec";
            else if (email.Contains("dungb")) empId = "654207f91dcb3048bfecc2f9";
            else if (email.Contains("khangc")) empId = "65780f61ab74fd73cc5f0542";
            else if (email.Contains("trungn")) empId = "65420805dd376b76c2f9b118";

            if (!string.IsNullOrEmpty(empId))
            {
                defaultAssigneeId = empId;
                var me = employees.FirstOrDefault(e => e.Id == empId);
                if (me != null && cboEmployees != null)
                {
                    this.Invoke(new Action(() => { cboEmployees.SelectedItem = me; }));
                }
            }
        }

        private void cboEmployees_TextUpdate(object sender, EventArgs e)
        {
            string keyword = cboEmployees.Text;
            string searchKeyword = ConvertToUnSign(keyword);

            cboEmployees.TextUpdate -= cboEmployees_TextUpdate;

            cboEmployees.DataSource = string.IsNullOrEmpty(searchKeyword)
                ? employees
                : employees.Where(x => ConvertToUnSign(x.Name).Contains(searchKeyword)).ToList();

            cboEmployees.Text = keyword;
            cboEmployees.SelectionStart = keyword.Length;
            cboEmployees.DroppedDown = true;

            cboEmployees.TextUpdate += cboEmployees_TextUpdate;
            Cursor.Current = Cursors.Default;
        }

        private void SaveDraft()
        {
            try
            {
                var drafts = new List<DraftTicket>();
                foreach (DataGridViewRow row in dgvCreateTickets.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (row.Cells["colResult"].Value != null && row.Cells["colResult"].Value.ToString().Contains("☁️")) continue;

                    if (row.Cells["colTitle"].Value != null || row.Cells["colCategory"].Value != null)
                    {
                        drafts.Add(new DraftTicket
                        {
                            Title = row.Cells["colTitle"].Value?.ToString() ?? "",
                            Desc = row.Cells["colDesc"].Value?.ToString() ?? "",
                            CatId = row.Cells["colCategory"].Value?.ToString() ?? "",
                            SubCatId = row.Cells["colSubCategory"].Value?.ToString() ?? "",
                            TagId = row.Cells["colTag"].Value?.ToString() ?? "",
                            EmpId = row.Cells["colAssignee"].Value?.ToString() ?? ""
                        });
                    }
                }

                if (drafts.Count > 0)
                {
                    File.WriteAllText(draftFilePath, JsonSerializer.Serialize(drafts));
                }
                else if (File.Exists(draftFilePath))
                {
                    File.Delete(draftFilePath);
                }
            }
            catch { }
        }

        private void LoadDraft()
        {
            if (File.Exists(draftFilePath))
            {
                try
                {
                    string json = File.ReadAllText(draftFilePath);
                    var drafts = JsonSerializer.Deserialize<List<DraftTicket>>(json);

                    if (drafts != null && drafts.Count > 0)
                    {
                        dgvCreateTickets.Rows.Clear();
                        dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                        foreach (var draft in drafts)
                        {
                            int idx = dgvCreateTickets.Rows.Add();
                            var row = dgvCreateTickets.Rows[idx];

                            if (!string.IsNullOrEmpty(draft.Title) && !defaultTitles.Any(t => t.Title.Equals(draft.Title, StringComparison.OrdinalIgnoreCase)))
                            {
                                defaultTitles.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = draft.Title });
                            }

                            row.Cells["colTitle"].Value = draft.Title;
                            var matchedGroup = defaultTitles.FirstOrDefault(t => t.Title == draft.Title);
                            if (matchedGroup != null) row.Cells["colGroup"].Value = matchedGroup.Group;

                            row.Cells["colDesc"].Value = draft.Desc;
                            row.Cells["colCategory"].Value = draft.CatId;

                            if (!string.IsNullOrEmpty(draft.CatId))
                            {
                                var cat = categoryList.FirstOrDefault(c => c._id == draft.CatId);
                                var cellSubCat = (DataGridViewComboBoxCell)row.Cells["colSubCategory"];
                                if (cat != null && cat.types != null)
                                {
                                    cellSubCat.DataSource = cat.types;
                                    cellSubCat.DisplayMember = "name";
                                    cellSubCat.ValueMember = "uuid";
                                }
                                row.Cells["colSubCategory"].Value = draft.SubCatId;
                            }

                            row.Cells["colTag"].Value = draft.TagId;
                            row.Cells["colAssignee"].Value = draft.EmpId;
                        }

                        dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                    }
                }
                catch { }
            }
        }

        private async void CheckAndDownloadUpdateAsync()
        {
            string currentVersion = "2.0";
            string versionUrl = "https://raw.githubusercontent.com/mtruong22/GhiIssue/master/version.txt";
            string exeUrl = "https://github.com/mtruong22/GhiIssue/releases/latest/download/GhiIssue.exe";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string onlineVersion = await client.GetStringAsync(versionUrl);
                    onlineVersion = onlineVersion.Trim();

                    if (onlineVersion != currentVersion)
                    {
                        DialogResult dialogResult = MessageBox.Show(
                            $"Có bản cập nhật mới (v{onlineVersion})!\nBạn có muốn tải và cài đặt tự động ngay không?",
                            "Cập nhật Tool OmiCRM",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                        if (dialogResult == DialogResult.Yes)
                        {
                            Form progressForm = new Form()
                            {
                                Text = "Đang tải bản cập nhật...",
                                Size = new Size(400, 120),
                                FormBorderStyle = FormBorderStyle.FixedDialog,
                                StartPosition = FormStartPosition.CenterScreen,
                                ControlBox = false
                            };
                            ProgressBar pb = new ProgressBar() { Left = 20, Top = 20, Width = 340, Height = 25 };
                            Label lbl = new Label() { Left = 20, Top = 50, Width = 340, Text = "Đang kết nối đến GitHub..." };
                            progressForm.Controls.Add(pb);
                            progressForm.Controls.Add(lbl);
                            progressForm.Show();

                            string tempExeName = "GhiIssue_Update.exe";

                            try
                            {
                                using (HttpResponseMessage res = await client.GetAsync(exeUrl, HttpCompletionOption.ResponseHeadersRead))
                                {
                                    res.EnsureSuccessStatusCode();
                                    long? totalBytes = res.Content.Headers.ContentLength;

                                    using (Stream contentStream = await res.Content.ReadAsStreamAsync())
                                    using (FileStream fileStream = new FileStream(tempExeName, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                                    {
                                        var buffer = new byte[8192];
                                        bool isMoreToRead = true;
                                        long totalRead = 0;

                                        while (isMoreToRead)
                                        {
                                            int read = await contentStream.ReadAsync(buffer, 0, buffer.Length);
                                            if (read == 0) isMoreToRead = false;
                                            else
                                            {
                                                await fileStream.WriteAsync(buffer, 0, read);
                                                totalRead += read;
                                                if (totalBytes.HasValue)
                                                {
                                                    int percent = (int)((double)totalRead / totalBytes.Value * 100);
                                                    pb.Value = percent;
                                                    lbl.Text = $"Đang tải: {totalRead / 1048576} MB / {totalBytes.Value / 1048576} MB ({percent}%)";
                                                    Application.DoEvents();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception downloadEx)
                            {
                                progressForm.Close();
                                MessageBox.Show("Lỗi tải file: " + downloadEx.Message, "Lỗi mạng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            progressForm.Close();

                            string batCode =
                                "@echo off\n" +
                                "timeout /t 2 /nobreak > NUL\n" +
                                "del \"GhiIssue.exe\"\n" +
                                "ren \"GhiIssue_Update.exe\" \"GhiIssue.exe\"\n" +
                                "start \"\" \"GhiIssue.exe\"\n" +
                                "del \"%~f0\"";

                            File.WriteAllText("updater.bat", batCode);

                            ProcessStartInfo psi = new ProcessStartInfo()
                            {
                                FileName = "updater.bat",
                                CreateNoWindow = true,
                                WindowStyle = ProcessWindowStyle.Hidden
                            };
                            Process.Start(psi);
                            Application.Exit();
                        }
                    }
                }
            }
            catch { }
        }

        private async Task PerformLoginSequenceAsync()
        {
            bool isDataLoaded = false;

            if (File.Exists(tokenFilePath))
            {
                OMICRM_TOKEN = File.ReadAllText(tokenFilePath).Trim();
                isDataLoaded = await LoadDataWithTokenAsync();

                string credFile = Path.Combine(Application.StartupPath, "user.dat");
                if (File.Exists(credFile))
                {
                    try { loginEmailCached = Encoding.UTF8.GetString(Convert.FromBase64String(File.ReadAllText(credFile))).Split('|')[0]; } catch { }
                }
            }

            while (!isDataLoaded)
            {
                string password = "";
                if (ShowLoginDialog(out loginEmailCached, out password))
                {
                    Cursor.Current = Cursors.WaitCursor;
                    bool loginSuccess = await LoginAndGetTokenAsync(loginEmailCached, password);
                    Cursor.Current = Cursors.Default;

                    if (loginSuccess)
                    {
                        isDataLoaded = await LoadDataWithTokenAsync();
                        if (isDataLoaded) break;
                    }
                    else
                    {
                        MessageBox.Show("Đăng nhập thất bại! Kiểm tra lại Email/Mật khẩu", "Lỗi");
                    }
                }
                else
                {
                    MessageBox.Show("Bạn phải đăng nhập để sử dụng Tool!", "Thông báo");
                    Application.Exit();
                    return;
                }
            }
        }

        private bool ShowLoginDialog(out string email, out string password)
        {
            Form prompt = new Form()
            {
                Width = 350,
                Height = 250,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Đăng nhập OmiCRM",
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false
            };
            Label textLabel = new Label() { Left = 20, Top = 20, Text = "Email:" };
            TextBox inputBox = new TextBox() { Left = 20, Top = 40, Width = 290 };
            Label passLabel = new Label() { Left = 20, Top = 70, Text = "Mật khẩu:" };
            TextBox passBox = new TextBox() { Left = 20, Top = 90, Width = 290, PasswordChar = '*' };

            CheckBox chkRemember = new CheckBox() { Left = 20, Top = 125, Text = "Ghi nhớ mật khẩu", Width = 150 };
            Button confirmation = new Button() { Text = "Đăng nhập", Left = 210, Width = 100, Top = 150, DialogResult = DialogResult.OK };

            string credFile = Path.Combine(Application.StartupPath, "user.dat");
            if (File.Exists(credFile))
            {
                try
                {
                    string decoded = Encoding.UTF8.GetString(Convert.FromBase64String(File.ReadAllText(credFile)));
                    string[] parts = decoded.Split('|');
                    if (parts.Length == 2)
                    {
                        inputBox.Text = parts[0];
                        passBox.Text = parts[1];
                        chkRemember.Checked = true;
                    }
                }
                catch { }
            }

            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(inputBox);
            prompt.Controls.Add(passLabel);
            prompt.Controls.Add(passBox);
            prompt.Controls.Add(chkRemember);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            email = ""; password = "";
            if (prompt.ShowDialog() == DialogResult.OK)
            {
                email = inputBox.Text.Trim();
                password = passBox.Text.Trim();

                if (chkRemember.Checked)
                {
                    string encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(email + "|" + password));
                    File.WriteAllText(credFile, encoded);
                }
                else
                {
                    if (File.Exists(credFile)) File.Delete(credFile);
                }
                return true;
            }
            return false;
        }

        private async Task<bool> LoginAndGetTokenAsync(string email, string password)
        {
            using (HttpClient client = new HttpClient())
            {
                var loginBody = new
                {
                    kind = "internal",
                    identify_info = email,
                    password = password,
                    tenant_id = "6541fc4753504c21f1db821c" // Dcorp
                };

                var content = new StringContent(JsonSerializer.Serialize(loginBody), Encoding.UTF8, "application/json");

                try
                {
                    HttpResponseMessage response = await client.PostAsync("https://auth-v1-stg.omicrm.com/auth_2fa/login?lng=vi&utm_source=web", content);
                    if (response.IsSuccessStatusCode)
                    {
                        string resString = await response.Content.ReadAsStringAsync();
                        using (JsonDocument doc = JsonDocument.Parse(resString))
                        {
                            JsonElement root = doc.RootElement;
                            string rawToken = root.GetProperty("payload").GetProperty("access_token").GetString();
                            OMICRM_TOKEN = "Bearer " + rawToken;
                            File.WriteAllText(tokenFilePath, OMICRM_TOKEN);
                            return true;
                        }
                    }
                }
                catch { }
            }
            return false;
        }

        private async Task<bool> LoadDataWithTokenAsync()
        {
            if (string.IsNullOrEmpty(OMICRM_TOKEN)) return false;
            try
            {
                await Task.WhenAll(LoadCategoriesAsync(), LoadTagsAsync());
                if (categoryList.Count > 0) return true;
            }
            catch { }
            return false;
        }

        private void DgvCreateTickets_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (dgvCreateTickets.Columns[e.ColumnIndex].Name == "colTitle")
            {
                string typedText = e.FormattedValue?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(typedText))
                {
                    if (!defaultTitles.Any(t => t.Title.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        defaultTitles.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = typedText });
                    }
                }
            }
        }

        private void DgvCreateTickets_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox cb)
            {
                cb.DropDownStyle = ComboBoxStyle.DropDown;
                cb.AutoCompleteMode = AutoCompleteMode.None;

                BeginInvoke(new Action(() => {
                    if (cb.Focused)
                    {
                        cb.SelectionLength = 0;
                        cb.SelectionStart = cb.Text.Length;
                    }
                }));

                string colName = dgvCreateTickets.CurrentCell.OwningColumn.Name;

                if (colName == "colTitle")
                {
                    string currentVal = dgvCreateTickets.CurrentRow.Cells["colTitle"].Value?.ToString() ?? "";
                    var cleanList = defaultTitles.Where(t => t.Group != "Khác (Nhập tay)" || t.Title == currentVal).ToList();

                    cb.DataSource = cleanList;
                    cb.DisplayMember = "Title";
                    cb.ValueMember = "Title";
                }
                else if (colName == "colCategory") cb.DataSource = categoryList.ToList();
                else if (colName == "colTag") cb.DataSource = tagList.ToList();
                else if (colName == "colAssignee") cb.DataSource = employees.ToList();
                else if (colName == "colSubCategory")
                {
                    string catId = dgvCreateTickets.CurrentRow.Cells["colCategory"].Value?.ToString();
                    var cat = categoryList.FirstOrDefault(c => c._id == catId);
                    cb.DataSource = (cat != null && cat.types != null) ? cat.types.ToList() : new List<SubCategoryItem>();
                }

                cb.TextUpdate -= Cb_TextUpdate;
                cb.TextUpdate += Cb_TextUpdate;
                cb.KeyDown -= Cb_KeyDown;
                cb.KeyDown += Cb_KeyDown;
            }
        }

        private void Cb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SendKeys.Send("{TAB}");
            }
        }

        private string ConvertToUnSign(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("\\p{IsCombiningDiacriticalMarks}+");
            string temp = s.Normalize(NormalizationForm.FormD);
            return regex.Replace(temp, String.Empty).Replace('\u0111', 'd').Replace('\u0110', 'D').ToLower();
        }

        private void Cb_TextUpdate(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb == null) return;

            string keyword = cb.Text;
            string searchKeyword = ConvertToUnSign(keyword);

            cb.TextUpdate -= Cb_TextUpdate;
            string colName = dgvCreateTickets.CurrentCell.OwningColumn.Name;

            if (colName == "colTitle")
            {
                var cleanList = defaultTitles.Where(t => t.Group != "Khác (Nhập tay)").ToList();
                cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? cleanList : cleanList.Where(x => ConvertToUnSign(x.Title).Contains(searchKeyword)).ToList();
            }
            else if (colName == "colCategory") cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? categoryList.ToList() : categoryList.Where(x => ConvertToUnSign(x.name).Contains(searchKeyword)).ToList();
            else if (colName == "colTag") cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? tagList.ToList() : tagList.Where(x => ConvertToUnSign(x.name).Contains(searchKeyword)).ToList();
            else if (colName == "colAssignee") cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? employees.ToList() : employees.Where(x => ConvertToUnSign(x.Name).Contains(searchKeyword)).ToList();
            else if (colName == "colSubCategory")
            {
                string catId = dgvCreateTickets.CurrentRow.Cells["colCategory"].Value?.ToString();
                var cat = categoryList.FirstOrDefault(c => c._id == catId);
                if (cat != null && cat.types != null) cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? cat.types.ToList() : cat.types.Where(x => ConvertToUnSign(x.name).Contains(searchKeyword)).ToList();
            }

            cb.Text = keyword;
            cb.SelectionStart = keyword.Length;
            cb.DroppedDown = true;
            cb.TextUpdate += Cb_TextUpdate;
            Cursor.Current = Cursors.Default;
        }

        private void DgvCreateTickets_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvCreateTickets.IsCurrentCellDirty)
            {
                string colName = dgvCreateTickets.CurrentCell?.OwningColumn?.Name;
                if (colName == "colCategory" || colName == "colSubCategory" || colName == "colAssignee" || colName == "colTag")
                {
                    dgvCreateTickets.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void DgvCreateTickets_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string colName = dgvCreateTickets.Columns[e.ColumnIndex].Name;
            var currentRow = dgvCreateTickets.Rows[e.RowIndex];

            if (colName == "colTitle")
            {
                string selectedTitle = currentRow.Cells["colTitle"].Value?.ToString();
                var matched = defaultTitles.FirstOrDefault(t => t.Title == selectedTitle);

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                if (matched != null && matched.Group != "Khác (Nhập tay)")
                {
                    currentRow.Cells["colGroup"].Value = matched.Group;
                }
                else
                {
                    currentRow.Cells["colGroup"].Value = "Khác (Nhập tay)";
                }
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
            }

            if (colName == "colCategory")
            {
                string selectedCatId = currentRow.Cells["colCategory"].Value?.ToString();
                var cellSubCat = (DataGridViewComboBoxCell)currentRow.Cells["colSubCategory"];
                var cat = categoryList.FirstOrDefault(c => c._id == selectedCatId);
                if (cat != null && cat.types != null)
                {
                    cellSubCat.DataSource = cat.types.ToList();
                    cellSubCat.DisplayMember = "name";
                    cellSubCat.ValueMember = "uuid";
                }
                else { cellSubCat.DataSource = new List<SubCategoryItem>(); }
            }

            if (colName == "colCategory" || colName == "colSubCategory" || colName == "colAssignee")
            {
                var cellValue = currentRow.Cells[colName].Value;
                if (cellValue == null) return;

                string valueToCopy = cellValue.ToString();
                string currentCatId = currentRow.Cells["colCategory"].Value?.ToString();

                if (colName == "colAssignee") defaultAssigneeId = valueToCopy;

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                try
                {
                    for (int i = e.RowIndex + 1; i < dgvCreateTickets.Rows.Count; i++)
                    {
                        var targetRow = dgvCreateTickets.Rows[i];
                        if (targetRow.IsNewRow) continue;

                        if (colName == "colCategory")
                        {
                            targetRow.Cells["colCategory"].Value = valueToCopy;

                            var targetCellSubCat = (DataGridViewComboBoxCell)targetRow.Cells["colSubCategory"];
                            var cat = categoryList.FirstOrDefault(c => c._id == valueToCopy);
                            if (cat != null && cat.types != null)
                            {
                                targetCellSubCat.DataSource = cat.types.ToList();
                                targetCellSubCat.DisplayMember = "name";
                                targetCellSubCat.ValueMember = "uuid";
                            }
                            else { targetCellSubCat.DataSource = new List<SubCategoryItem>(); }

                            targetRow.Cells["colSubCategory"].Value = null;
                        }
                        else if (colName == "colSubCategory")
                        {
                            if (targetRow.Cells["colCategory"].Value?.ToString() == currentCatId)
                            {
                                targetRow.Cells["colSubCategory"].Value = valueToCopy;
                            }
                        }
                        else if (colName == "colAssignee")
                        {
                            targetRow.Cells["colAssignee"].Value = valueToCopy;
                        }
                    }
                }
                finally
                {
                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                }
            }
        }

        private void DgvCreateTickets_DataError(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; }

        private async Task LoadCategoriesAsync()
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);
                try
                {
                    HttpResponseMessage response = await client.GetAsync("https://ticket-v2-stg.omicrm.com/ticket/category/get-all?lng=vi&utm_source=web");
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        CategoryResponse catData = JsonSerializer.Deserialize<CategoryResponse>(json);
                        if (catData?.payload != null)
                        {
                            categoryList = catData.payload.Where(c => !string.IsNullOrEmpty(c.name) && c.types != null && c.types.Count > 0).ToList();
                            if (dgvCreateTickets != null && dgvCreateTickets.Columns.Contains("colCategory"))
                            {
                                var col = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colCategory"];
                                col.DataSource = categoryList;
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private async Task LoadTagsAsync()
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);
                try
                {
                    tagList = new List<TagItem>();
                    int currentPage = 1;
                    int pageSize = 1000;
                    bool hasMoreData = true;

                    while (hasMoreData)
                    {
                        var body = new { size = pageSize, page = currentPage };
                        var content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

                        HttpResponseMessage response = await client.PostAsync("https://tenant-config-stg.omicrm.com/tag/search-all?lng=vi&utm_source=web", content);

                        if (!response.IsSuccessStatusCode)
                            response = await client.GetAsync($"https://tenant-config-stg.omicrm.com/tag/search-all?lng=vi&utm_source=web&size={pageSize}&page={currentPage}");

                        if (response.IsSuccessStatusCode)
                        {
                            string json = await response.Content.ReadAsStringAsync();
                            TagResponse tagData = JsonSerializer.Deserialize<TagResponse>(json);

                            if (tagData?.payload?.items != null && tagData.payload.items.Count > 0)
                            {
                                tagList.AddRange(tagData.payload.items);
                                currentPage++;
                            }
                            else hasMoreData = false;
                        }
                        else hasMoreData = false;
                    }

                    if (dgvCreateTickets != null && dgvCreateTickets.Columns.Contains("colTag"))
                    {
                        var col = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colTag"];
                        col.DataSource = tagList;
                    }
                }
                catch { }
            }
        }

        private void SetupCreateTicketGrid()
        {
            if (dgvCreateTickets == null) return;
            dgvCreateTickets.Columns.Clear();
            dgvCreateTickets.AutoGenerateColumns = false;
            dgvCreateTickets.AllowUserToAddRows = true;

            dgvCreateTickets.EditMode = DataGridViewEditMode.EditOnEnter;
            dgvCreateTickets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvCreateTickets.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            dgvCreateTickets.MultiSelect = true;
            dgvCreateTickets.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            dgvCreateTickets.KeyDown -= DgvCreateTickets_KeyDown;
            dgvCreateTickets.KeyDown += DgvCreateTickets_KeyDown;

            ContextMenuStrip gridMenu = new ContextMenuStrip();
            gridMenu.Items.Add("📋 Copy (Ctrl+C)", null, (s, e) => CopyToClipboard());
            gridMenu.Items.Add("📋 Paste (Ctrl+V)", null, (s, e) => PasteFromClipboard());
            gridMenu.Items.Add("✂️ Cut (Ctrl+X)", null, (s, e) => CutToClipboard());
            gridMenu.Items.Add("-");
            gridMenu.Items.Add("❌ Delete", null, (s, e) => SmartDelete());
            dgvCreateTickets.ContextMenuStrip = gridMenu;

            DataGridViewTextBoxColumn colGroup = new DataGridViewTextBoxColumn();
            colGroup.HeaderText = "Nhóm";
            colGroup.Name = "colGroup";
            colGroup.ReadOnly = true;
            colGroup.DefaultCellStyle.BackColor = Color.LightGray;
            colGroup.FillWeight = 10;
            dgvCreateTickets.Columns.Add(colGroup);

            DataGridViewComboBoxColumn colTitle = new DataGridViewComboBoxColumn();
            colTitle.HeaderText = "Tiêu đề phiếu (mặc định)";
            colTitle.Name = "colTitle";
            colTitle.DataSource = defaultTitles;
            colTitle.DisplayMember = "Title";
            colTitle.ValueMember = "Title";
            colTitle.FillWeight = 25;
            dgvCreateTickets.Columns.Add(colTitle);

            dgvCreateTickets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Mô tả chi tiết", Name = "colDesc", FillWeight = 15 });

            DataGridViewComboBoxColumn colCategory = new DataGridViewComboBoxColumn();
            colCategory.HeaderText = "Chủ đề";
            colCategory.Name = "colCategory";
            colCategory.DisplayMember = "name";
            colCategory.ValueMember = "_id";
            colCategory.FillWeight = 10;
            dgvCreateTickets.Columns.Add(colCategory);

            DataGridViewComboBoxColumn colSubCategory = new DataGridViewComboBoxColumn();
            colSubCategory.HeaderText = "Phân loại";
            colSubCategory.Name = "colSubCategory";
            colSubCategory.DisplayMember = "name";
            colSubCategory.ValueMember = "uuid";
            colSubCategory.FillWeight = 15;
            dgvCreateTickets.Columns.Add(colSubCategory);

            DataGridViewComboBoxColumn colTag = new DataGridViewComboBoxColumn();
            colTag.HeaderText = "Tag";
            colTag.Name = "colTag";
            colTag.DisplayMember = "name";
            colTag.ValueMember = "id";
            colTag.FillWeight = 15;
            dgvCreateTickets.Columns.Add(colTag);

            DataGridViewComboBoxColumn colAssignee = new DataGridViewComboBoxColumn();
            colAssignee.HeaderText = "Người xử lý";
            colAssignee.Name = "colAssignee";
            colAssignee.DataSource = employees;
            colAssignee.DisplayMember = "Name";
            colAssignee.ValueMember = "Id";
            colAssignee.FillWeight = 15;
            dgvCreateTickets.Columns.Add(colAssignee);

            DataGridViewTextBoxColumn colResult = new DataGridViewTextBoxColumn();
            colResult.HeaderText = "Kết quả";
            colResult.Name = "colResult";
            colResult.ReadOnly = true;
            colResult.FillWeight = 10;
            dgvCreateTickets.Columns.Add(colResult);
        }

        private void DgvCreateTickets_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C) { CopyToClipboard(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.V) { PasteFromClipboard(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.X) { CutToClipboard(); e.Handled = true; }
            else if (e.KeyCode == Keys.Delete)
            {
                SmartDelete();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SendKeys.Send("{TAB}");
            }
        }

        private void CopyToClipboard()
        {
            try { DataObject d = dgvCreateTickets.GetClipboardContent(); if (d != null) Clipboard.SetDataObject(d); } catch { }
        }

        private void CutToClipboard()
        {
            try
            {
                CopyToClipboard();
                DeleteSelectedCells();
            }
            catch { }
        }

        private void SmartDelete()
        {
            try
            {
                if (dgvCreateTickets.SelectedRows.Count > 0)
                {
                    foreach (DataGridViewRow row in dgvCreateTickets.SelectedRows.Cast<DataGridViewRow>().ToList())
                    {
                        if (!row.IsNewRow) dgvCreateTickets.Rows.Remove(row);
                    }
                }
                else
                {
                    DeleteSelectedCells();
                }
            }
            catch { }
        }

        private void DeleteSelectedCells()
        {
            try
            {
                foreach (DataGridViewCell cell in dgvCreateTickets.SelectedCells)
                {
                    if (!cell.ReadOnly) cell.Value = null;
                }
            }
            catch { }
        }

        private void PasteFromClipboard()
        {
            try
            {
                string copiedText = Clipboard.GetText();
                if (string.IsNullOrEmpty(copiedText)) return;

                string[] lines = copiedText.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return;

                int startRow = dgvCreateTickets.CurrentCell.RowIndex;
                int startCol = dgvCreateTickets.CurrentCell.ColumnIndex;

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                foreach (string line in lines)
                {
                    if (startRow >= dgvCreateTickets.Rows.Count) dgvCreateTickets.Rows.Add();
                    if (dgvCreateTickets.Rows[startRow].IsNewRow) continue;

                    string[] cells = line.Split('\t');
                    for (int i = 0; i < cells.Length; i++)
                    {
                        int targetColIndex = startCol + i;
                        if (targetColIndex >= dgvCreateTickets.Columns.Count) break;

                        var targetCell = dgvCreateTickets[targetColIndex, startRow];

                        if (targetCell.ReadOnly) continue;

                        string cellText = cells[i].Trim();
                        string colName = targetCell.OwningColumn.Name;

                        if (colName == "colTitle")
                        {
                            var match = defaultTitles.FirstOrDefault(t => t.Title.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                            if (match != null)
                            {
                                targetCell.Value = match.Title;
                                dgvCreateTickets.Rows[startRow].Cells["colGroup"].Value = match.Group;
                            }
                            else
                            {
                                if (!defaultTitles.Any(t => t.Title.Equals(cellText, StringComparison.OrdinalIgnoreCase)))
                                {
                                    defaultTitles.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = cellText });
                                }
                                targetCell.Value = cellText;
                                dgvCreateTickets.Rows[startRow].Cells["colGroup"].Value = "Khác (Nhập tay)";
                            }
                        }
                        else if (colName == "colDesc") targetCell.Value = cellText;
                        else if (colName == "colCategory")
                        {
                            var match = categoryList.FirstOrDefault(c => c.name.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                            if (match != null)
                            {
                                targetCell.Value = match._id;
                                var cellSubCat = (DataGridViewComboBoxCell)dgvCreateTickets.Rows[startRow].Cells["colSubCategory"];
                                if (match.types != null)
                                {
                                    cellSubCat.DataSource = match.types.ToList();
                                    cellSubCat.DisplayMember = "name";
                                    cellSubCat.ValueMember = "uuid";
                                }
                            }
                        }
                        else if (colName == "colSubCategory")
                        {
                            string currentCatId = dgvCreateTickets.Rows[startRow].Cells["colCategory"].Value?.ToString();
                            var parentCat = categoryList.FirstOrDefault(c => c._id == currentCatId);
                            if (parentCat != null && parentCat.types != null)
                            {
                                var match = parentCat.types.FirstOrDefault(t => t.name.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                                if (match != null) targetCell.Value = match.uuid;
                            }
                        }
                        else if (colName == "colTag")
                        {
                            var match = tagList.FirstOrDefault(t => t.name.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                            if (match != null) targetCell.Value = match.id;
                        }
                        else if (colName == "colAssignee")
                        {
                            var match = employees.FirstOrDefault(e => e.Name.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                            if (match != null) targetCell.Value = match.Id;
                        }
                    }
                    startRow++;
                }
            }
            catch { }
            finally
            {
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
            }
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            if (dgvCreateTickets != null)
            {
                int oldRowCount = dgvCreateTickets.Rows.Count;
                dgvCreateTickets.Rows.Add(5);
                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                try
                {
                    int lastValidRowIndex = oldRowCount - 2;
                    string lastCat = "";
                    string lastSubCat = "";

                    if (lastValidRowIndex >= 0)
                    {
                        var sourceRow = dgvCreateTickets.Rows[lastValidRowIndex];
                        lastCat = sourceRow.Cells["colCategory"].Value?.ToString() ?? "";
                        lastSubCat = sourceRow.Cells["colSubCategory"].Value?.ToString() ?? "";
                    }

                    for (int i = oldRowCount - 1; i < dgvCreateTickets.Rows.Count; i++)
                    {
                        var targetRow = dgvCreateTickets.Rows[i];
                        if (targetRow.IsNewRow) continue;

                        if (!string.IsNullOrEmpty(defaultAssigneeId)) targetRow.Cells["colAssignee"].Value = defaultAssigneeId;

                        if (!string.IsNullOrEmpty(lastCat))
                        {
                            targetRow.Cells["colCategory"].Value = lastCat;
                            var targetCellSubCat = (DataGridViewComboBoxCell)targetRow.Cells["colSubCategory"];
                            var cat = categoryList.FirstOrDefault(c => c._id == lastCat);
                            if (cat != null && cat.types != null)
                            {
                                targetCellSubCat.DataSource = cat.types.ToList();
                                targetCellSubCat.DisplayMember = "name";
                                targetCellSubCat.ValueMember = "uuid";
                            }
                            else { targetCellSubCat.DataSource = new List<SubCategoryItem>(); }

                            if (!string.IsNullOrEmpty(lastSubCat)) targetRow.Cells["colSubCategory"].Value = lastSubCat;
                        }
                    }
                }
                finally
                {
                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                }
            }
        }

        private async void btnCreateTicket_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            int successCount = 0;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
                client.DefaultRequestHeaders.Add("Accept", "application/json, text/plain, */*");
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

                foreach (DataGridViewRow row in dgvCreateTickets.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (row.Cells["colResult"].Value != null && row.Cells["colResult"].Value.ToString().Contains("☁️")) continue;

                    try
                    {
                        string title = row.Cells["colTitle"].Value?.ToString() ?? "";
                        string desc = row.Cells["colDesc"].Value?.ToString() ?? "";
                        string catId = row.Cells["colCategory"].Value?.ToString() ?? "";
                        string subCatId = row.Cells["colSubCategory"].Value?.ToString() ?? "";
                        string tagId = row.Cells["colTag"].Value?.ToString() ?? "";
                        string empId = row.Cells["colAssignee"].Value?.ToString() ?? "";

                        if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(catId)) continue;

                        row.Cells["colResult"].Value = "⏳ Đang gửi...";

                        var catInfo = categoryList.FirstOrDefault(c => c._id == catId);
                        var subCatInfo = catInfo?.types?.FirstOrDefault(t => t.uuid == subCatId);
                        int typeIndex = subCatInfo != null ? subCatInfo.index : 0;

                        string finalTag = string.IsNullOrEmpty(tagId) ? null : tagId;
                        string finalEmp = string.IsNullOrEmpty(empId) ? null : empId;

                        var createBody = new
                        {
                            name = title,
                            description = string.IsNullOrEmpty(desc) ? "" : $"<div style=\"font-size: 15px;\">{desc}</div>",
                            category_id = catId,
                            current_type = typeIndex,
                            assignee_contact_ids = string.IsNullOrEmpty(finalEmp) ? new string[] { } : new[] { finalEmp },
                            tags = string.IsNullOrEmpty(finalTag) ? new string[] { } : new[] { finalTag },
                            source = "crud",
                            priority = "medium",
                            embed_files = new string[] { },
                            reporter_contact_ids = new string[] { },
                            work_list = new string[] { },
                            attribute_structures = new object[] { }
                        };

                        var content = new StringContent(JsonSerializer.Serialize(createBody), Encoding.UTF8, "application/json");
                        HttpResponseMessage response = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/add?lng=vi&utm_source=web", content);
                        string resStr = await response.Content.ReadAsStringAsync();

                        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized || resStr.Contains("Unauthorized") || resStr.Contains("\"status_code\": 401"))
                        {
                            Cursor.Current = Cursors.Default;
                            MessageBox.Show("Phiên đăng nhập OmiCRM đã hết hạn!\nTool đã DỪNG GỬI để bảo vệ các phiếu còn lại.\n\nVui lòng đăng nhập lại để tiếp tục.", "Cảnh Báo Mất Phiếu", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
                            OMICRM_TOKEN = "";
                            await PerformLoginSequenceAsync();
                            return;
                        }

                        if (response.IsSuccessStatusCode)
                        {
                            row.Cells["colResult"].Value = "✅ Thành công";
                            successCount++;
                        }
                        else
                        {
                            if (resStr.StartsWith("<")) row.Cells["colResult"].Value = "❌ Tường lửa chặn / Lỗi Token";
                            else row.Cells["colResult"].Value = "❌ " + (resStr.Length > 80 ? resStr.Substring(0, 80) : resStr);
                        }
                    }
                    catch (Exception ex) { row.Cells["colResult"].Value = "❌ Lỗi: " + ex.Message; }
                }
            }

            Cursor.Current = Cursors.Default;
            MessageBox.Show($"Xong! Đã tạo thành công {successCount} phiếu.", "Kết quả");

            if (successCount > 0)
            {
                for (int i = dgvCreateTickets.Rows.Count - 1; i >= 0; i--)
                {
                    DataGridViewRow currentRow = dgvCreateTickets.Rows[i];
                    if (currentRow.IsNewRow) continue;
                    if (currentRow.Cells["colResult"].Value != null && currentRow.Cells["colResult"].Value.ToString().Contains("✅"))
                    {
                        dgvCreateTickets.Rows.RemoveAt(i);
                    }
                }
                if (dgvCreateTickets.Rows.Count <= 1) btnAddRow_Click(null, null);

                SaveDraft();
            }
        }

        private async void btnExecute_Click(object sender, EventArgs e)
        {
            Employee selectedEmp = (Employee)cboEmployees.SelectedItem;
            if (selectedEmp == null) return;
            string targetId = selectedEmp.Id;

            dgvTickets.DataSource = null;

            var searchBody = new
            {
                status_filters = new[] { "active_state" },
                assignee_contact_ids = new[] { targetId },
                additional_layout = new[] { "object_association" },
                has_notify_report = true,
                page = "1",
                size = "1000",
                current_status = new[] { 0, 1, 2, 3 }
            };

            var searchContent = new StringContent(JsonSerializer.Serialize(searchBody), Encoding.UTF8, "application/json");
            Cursor.Current = Cursors.WaitCursor;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

                try
                {
                    HttpResponseMessage searchResponse = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/search?lng=vi&utm_source=web", searchContent);
                    string resStr = await searchResponse.Content.ReadAsStringAsync();

                    if (searchResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized || resStr.Contains("Unauthorized") || resStr.Contains("\"status_code\": 401"))
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Phiên đăng nhập đã hết hạn!\nVui lòng đăng nhập lại.", "Cảnh Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
                        OMICRM_TOKEN = "";
                        await PerformLoginSequenceAsync();
                        return;
                    }

                    if (!searchResponse.IsSuccessStatusCode)
                    {
                        MessageBox.Show($"Lỗi API Tìm kiếm: {searchResponse.StatusCode}", "Lỗi");
                        Cursor.Current = Cursors.Default;
                        return;
                    }

                    OmiResponse omiData = JsonSerializer.Deserialize<OmiResponse>(resStr);

                    List<TicketItem> allTickets = new List<TicketItem>();
                    if (omiData?.payload?.items != null) allTickets = omiData.payload.items;

                    allTickets = allTickets.Where(t => !recentlyClosedTicketIds.Contains(t.id) && !recentlyClosedTicketIds.Contains(t._id)).ToList();

                    if (allTickets.Count == 0)
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Sạch sẽ! Không có phiếu nào tồn đọng (0, 1, 2, 3).", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    var ticketsToClose = allTickets.Where(t => t.current_status == 0).ToList();
                    Cursor.Current = Cursors.Default;

                    if (ticketsToClose.Count > 0)
                    {
                        DialogResult confirm = MessageBox.Show(
                            $"Tìm thấy tổng cộng {allTickets.Count} phiếu.\n" +
                            $"- Có {ticketsToClose.Count} phiếu MỚI (sẽ được tự động ĐÓNG).\n" +
                            $"- Có {allTickets.Count - ticketsToClose.Count} phiếu ĐANG XỬ LÝ/KHÁC (sẽ giữ nguyên).\n\n" +
                            $"Bạn có muốn thực hiện không?", "Xác nhận xử lý", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirm == DialogResult.No) return;
                    }
                    else
                    {
                        MessageBox.Show($"Tìm thấy {allTickets.Count} phiếu đang xử lý/leo thang.\nSẽ hiển thị danh sách để bạn theo dõi.", "Thông tin");
                    }

                    Cursor.Current = Cursors.WaitCursor;
                    int successCount = 0;

                    foreach (var ticket in allTickets)
                    {
                        ticket.NguoiNhan = selectedEmp.Name;

                        if (ticket.current_status == 0)
                        {
                            string realId = !string.IsNullOrEmpty(ticket._id) ? ticket._id : ticket.id;
                            var updateBody = new { id = realId, status = 4, message = "DONE" };
                            var updateContent = new StringContent(JsonSerializer.Serialize(updateBody), Encoding.UTF8, "application/json");

                            try
                            {
                                HttpResponseMessage updateResponse = await client.PutAsync("https://ticket-v2-stg.omicrm.com/ticket/status/update?lng=vi&utm_source=web", updateContent);
                                if (updateResponse.IsSuccessStatusCode)
                                {
                                    successCount++;
                                    ticket.current_status = 4;
                                    recentlyClosedTicketIds.Add(realId);
                                }
                            }
                            catch { }
                        }
                    }

                    Cursor.Current = Cursors.Default;
                    dgvTickets.DataSource = allTickets;

                    if (dgvTickets.Columns.Contains("_id")) dgvTickets.Columns["_id"].Visible = false;
                    if (dgvTickets.Columns.Contains("id")) dgvTickets.Columns["id"].Visible = false;
                    if (dgvTickets.Columns.Contains("current_status")) dgvTickets.Columns["current_status"].Visible = false;

                    if (dgvTickets.Columns.Contains("unique_id")) { dgvTickets.Columns["unique_id"].HeaderText = "ID Phiếu"; dgvTickets.Columns["unique_id"].DisplayIndex = 0; dgvTickets.Columns["unique_id"].Width = 80; }
                    if (dgvTickets.Columns.Contains("name")) { dgvTickets.Columns["name"].HeaderText = "Tên Phiếu"; dgvTickets.Columns["name"].DisplayIndex = 1; }
                    if (dgvTickets.Columns.Contains("TrangThaiHienThi")) { dgvTickets.Columns["TrangThaiHienThi"].HeaderText = "Trạng Thái"; dgvTickets.Columns["TrangThaiHienThi"].DisplayIndex = 2; dgvTickets.Columns["TrangThaiHienThi"].Width = 140; }
                    if (dgvTickets.Columns.Contains("NguoiNhan")) { dgvTickets.Columns["NguoiNhan"].HeaderText = "Người Nhận"; dgvTickets.Columns["NguoiNhan"].DisplayIndex = 3; }

                    dgvTickets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                    if (ticketsToClose.Count > 0)
                    {
                        MessageBox.Show($"Đã xử lý xong!\n- Tự động đóng: {successCount}/{ticketsToClose.Count} phiếu Mới.\n- Giữ nguyên: {allTickets.Count - ticketsToClose.Count} phiếu đang xử lý.", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show($"Có lỗi xảy ra: {ex.Message}", "Lỗi Code");
                }
            }
        }
    }

    public class PredefinedTitle
    {
        public string Group { get; set; }
        public string Title { get; set; }
    }

    public class DraftTicket
    {
        public string Title { get; set; }
        public string Desc { get; set; }
        public string CatId { get; set; }
        public string SubCatId { get; set; }
        public string TagId { get; set; }
        public string EmpId { get; set; }
    }

    public class Employee
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public override string ToString() { return Name; }
    }

    public class OmiResponse
    {
        public OmiPayload payload { get; set; }
    }

    public class OmiPayload
    {
        public List<TicketItem> items { get; set; }
    }

    public class TicketItem
    {
        [System.Text.Json.Serialization.JsonPropertyName("_id")]
        public string _id { get; set; }
        public string id { get; set; }
        public string unique_id { get; set; }
        public string name { get; set; }
        public int current_status { get; set; }

        public string TrangThaiHienThi
        {
            get
            {
                switch (current_status)
                {
                    case 0: return "Mới tiếp nhận (0)";
                    case 1: return "Đang xử lý (1)";
                    case 2: return "Đợi phản hồi (2)";
                    case 3: return "Leo thang (3)";
                    case 4: return "Đã Đóng (4)";
                    default: return "Khác (" + current_status + ")";
                }
            }
        }
        public string NguoiNhan { get; set; }
    }

    public class CategoryResponse
    {
        public List<CategoryItem> payload { get; set; }
    }

    public class CategoryItem
    {
        public string _id { get; set; }
        public string name { get; set; }
        public List<SubCategoryItem> types { get; set; }
        public override string ToString() { return name; }
    }

    public class SubCategoryItem
    {
        public string uuid { get; set; }
        public string name { get; set; }
        public int index { get; set; }
        public override string ToString() { return name; }
    }

    public class TagResponse
    {
        public TagPayload payload { get; set; }
    }

    public class TagPayload
    {
        public List<TagItem> items { get; set; }
    }

    public class TagItem
    {
        public string id { get; set; }
        public string name { get; set; }
        public override string ToString() { return name; }
    }
}