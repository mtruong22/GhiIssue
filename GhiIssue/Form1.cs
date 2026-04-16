using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace GhiIssue
{
    public partial class Form1 : Form
    {
        // ==================================================================================
        // 🌟 TÍNH NĂNG MỚI: LỌC NHANH TAG/NHÓM TRỰC TIẾP TRÊN BẢNG ĐÓNG ISSUE
        // ==================================================================================
        // ==================================================================================
        // 🌟 TÍNH NĂNG MỚI: LỌC NHANH ĐỘC LẬP (NHÓM VÀ TAG) CÓ TÍCH HỢP TÌM KIẾM
        // ==================================================================================
        public static Dictionary<string, string> dictTypeEng = new Dictionary<string, string>();
        public class QuickFilterItem { public string Text { get; set; } public string Id { get; set; } }
        public static bool isStrictTitle = true;
        //public static bool isStrictMode = true;
        public static bool isStrictType = true;
        public static bool allowDelete = false;
        public static bool isMaintenance = false;
        private List<string> adminEmails = new List<string> { "truongmh" };
        private string configUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRGeOsJ71sqpI__BcxboJTp-SvETPWY4-H21rs2mnTQb_K1LRSJJaAkBIP8TyGjgdidcF47gH_lCfUJ/pub?gid=1103587946&single=true&output=tsv";
        private List<TicketItem> _masterTicketList = new List<TicketItem>();
        private List<QuickFilterItem> _baseGroupFilters = new List<QuickFilterItem>();
        private List<QuickFilterItem> _baseTagFilters = new List<QuickFilterItem>();

        private void UpdateQuickFilter(List<TicketItem> loadedTickets)
        {
            _masterTicketList = loadedTickets;

            // Ngắt sự kiện để bơm dữ liệu không bị vướng
            cbQuickGroup.SelectedIndexChanged -= CbQuickGroup_SelectedIndexChanged;
            cbQuickGroup.TextUpdate -= CbQuickGroup_TextUpdate;
            cbQuickTag.SelectedIndexChanged -= CbQuickTag_SelectedIndexChanged;
            cbQuickTag.TextUpdate -= CbQuickTag_TextUpdate;

            cbQuickGroup.DataSource = null;
            cbQuickTag.DataSource = null;

            // 1. BƠM DỮ LIỆU CHO Ô COMBOBOX NHÓM
            List<QuickFilterItem> groupItems = new List<QuickFilterItem>();
            groupItems.Add(new QuickFilterItem { Text = $"🔥 TẤT CẢ NHÓM", Id = "" });
            var allGroups = loadedTickets.Where(t => t.group_name != "Khác").GroupBy(t => t.group_name).OrderByDescending(g => g.Count());
            foreach (var g in allGroups)
            {
                groupItems.Add(new QuickFilterItem { Text = $"{g.Key} ({g.Count()})", Id = g.Key });
            }
            _baseGroupFilters = groupItems;
            cbQuickGroup.DataSource = _baseGroupFilters;
            cbQuickGroup.DisplayMember = "Text";
            cbQuickGroup.SelectedIndex = 0;

            // 2. BƠM DỮ LIỆU CHO Ô COMBOBOX TAG
            List<QuickFilterItem> tagItems = new List<QuickFilterItem>();
            tagItems.Add(new QuickFilterItem { Text = $"🏷️ TẤT CẢ TAG", Id = "" });
            var tagCounts = new Dictionary<string, int>();
            foreach (var t in loadedTickets)
            {
                if (t.tags != null)
                {
                    foreach (var tg in t.tags)
                    {
                        string tid = tg.ValueKind == System.Text.Json.JsonValueKind.Object && tg.TryGetProperty("id", out var idProp) ? idProp.GetString() : tg.ToString();
                        if (tagCounts.ContainsKey(tid)) tagCounts[tid]++; else tagCounts[tid] = 1;
                    }
                }
            }
            var allTags = tagCounts.OrderByDescending(kv => kv.Value);
            foreach (var kv in allTags)
            {
                string tName = tagList.FirstOrDefault(x => x.id == kv.Key)?.name ?? kv.Key;
                tagItems.Add(new QuickFilterItem { Text = $"{tName} ({kv.Value})", Id = kv.Key });
            }
            _baseTagFilters = tagItems;
            cbQuickTag.DataSource = _baseTagFilters;
            cbQuickTag.DisplayMember = "Text";
            cbQuickTag.SelectedIndex = 0;

            // Bật lại sự kiện
            cbQuickGroup.SelectedIndexChanged += CbQuickGroup_SelectedIndexChanged;
            cbQuickGroup.TextUpdate += CbQuickGroup_TextUpdate;
            cbQuickTag.SelectedIndexChanged += CbQuickTag_SelectedIndexChanged;
            cbQuickTag.TextUpdate += CbQuickTag_TextUpdate;
        }

        // --- XỬ LÝ TÌM KIẾM CHO Ô NHÓM ---
        private void CbQuickGroup_TextUpdate(object sender, EventArgs e)
        {
            string keyword = cbQuickGroup.Text;
            string searchKeyword = ConvertToUnSignStatic(keyword);
            int cursorPos = cbQuickGroup.SelectionStart;

            cbQuickGroup.TextUpdate -= CbQuickGroup_TextUpdate;
            cbQuickGroup.SelectedIndexChanged -= CbQuickGroup_SelectedIndexChanged;

            var ds = string.IsNullOrEmpty(searchKeyword) ? _baseGroupFilters : _baseGroupFilters.Where(x => ConvertToUnSignStatic(x.Text).Contains(searchKeyword)).ToList();
            cbQuickGroup.DataSource = null; cbQuickGroup.DataSource = ds; cbQuickGroup.DisplayMember = "Text";
            cbQuickGroup.DroppedDown = ds.Count > 0; cbQuickGroup.Text = keyword; cbQuickGroup.SelectionStart = cursorPos;

            cbQuickGroup.SelectedIndexChanged += CbQuickGroup_SelectedIndexChanged;
            cbQuickGroup.TextUpdate += CbQuickGroup_TextUpdate;
            Cursor.Current = Cursors.Default;
        }

        // --- XỬ LÝ TÌM KIẾM CHO Ô TAG ---
        private void CbQuickTag_TextUpdate(object sender, EventArgs e)
        {
            string keyword = cbQuickTag.Text;
            string searchKeyword = ConvertToUnSignStatic(keyword);
            int cursorPos = cbQuickTag.SelectionStart;

            cbQuickTag.TextUpdate -= CbQuickTag_TextUpdate;
            cbQuickTag.SelectedIndexChanged -= CbQuickTag_SelectedIndexChanged;

            var ds = string.IsNullOrEmpty(searchKeyword) ? _baseTagFilters : _baseTagFilters.Where(x => ConvertToUnSignStatic(x.Text).Contains(searchKeyword)).ToList();
            cbQuickTag.DataSource = null; cbQuickTag.DataSource = ds; cbQuickTag.DisplayMember = "Text";
            cbQuickTag.DroppedDown = ds.Count > 0; cbQuickTag.Text = keyword; cbQuickTag.SelectionStart = cursorPos;

            cbQuickTag.SelectedIndexChanged += CbQuickTag_SelectedIndexChanged;
            cbQuickTag.TextUpdate += CbQuickTag_TextUpdate;
            Cursor.Current = Cursors.Default;
        }

        // --- LOGIC LỌC KẾT HỢP (GIAO NHAU) ---
        private void CbQuickGroup_SelectedIndexChanged(object sender, EventArgs e) { ApplyQuickFilters(); }
        private void CbQuickTag_SelectedIndexChanged(object sender, EventArgs e) { ApplyQuickFilters(); }

        private void ApplyQuickFilters()
        {
            if (cbQuickGroup.SelectedItem == null || cbQuickTag.SelectedItem == null) return;

            var selectedGroup = (QuickFilterItem)cbQuickGroup.SelectedItem;
            var selectedTag = (QuickFilterItem)cbQuickTag.SelectedItem;

            var filtered = _masterTicketList.AsEnumerable();

            // Nếu có chọn Nhóm cụ thể
            if (!string.IsNullOrEmpty(selectedGroup.Id))
            {
                filtered = filtered.Where(t => t.group_name == selectedGroup.Id);
            }

            // Nếu có chọn Tag cụ thể
            if (!string.IsNullOrEmpty(selectedTag.Id))
            {
                filtered = filtered.Where(t => t.tags != null && t.tags.Any(tg => tg.ToString().Contains(selectedTag.Id)));
            }

            dgvTickets.DataSource = null;
            dgvTickets.DataSource = filtered.ToList();
            UpdateStatusCount();
        }
        // ==================================================================================
        // ==================================================================================

        // ================== CẤU HÌNH & FILE HỆ THỐNG ==================
        private string OMICRM_TOKEN = "";
        private string tokenFilePath = Path.Combine(Application.StartupPath, "token.txt");
        private string draftFilePath = Path.Combine(Application.StartupPath, "draft.json");
        private string titlesCachePath = Path.Combine(Application.StartupPath, "titles_cache.json");
        private string catsCachePath = Path.Combine(Application.StartupPath, "cats_cache.json");
        private string tagsCachePath = Path.Combine(Application.StartupPath, "tags_cache.json");
        // --- BỘ NHỚ LƯU PHÂN LOẠI MẶC ĐỊNH ---
        private string catCachePath = Path.Combine(Application.StartupPath, "last_cat.txt");
        private string defaultCat = "FO_1_POS";
        private string defaultSubCat = "POS_Lỗi kết nối";
        private string loginEmailCached = "";

        // BỘ NHỚ LƯU LỊCH SỬ GÕ PHÍM CHO CTRL+Z
        private Stack<string> cellUndoStack = new Stack<string>();
        private bool isUndoing = false;

        private List<Employee> employees;
        private List<CategoryItem> categoryList = new List<CategoryItem>();
        private List<TagItem> tagList = new List<TagItem>();

        private BindingList<PredefinedTitle> defaultTitles = new BindingList<PredefinedTitle>();
        private BindingList<ComboItem> defaultTypeIssues = new BindingList<ComboItem>();
        private BindingList<ComboItem> defaultTechActions = new BindingList<ComboItem>();

        private string defaultAssigneeId = "";
        private HashSet<string> recentlyClosedTicketIds = new HashSet<string>();
        private System.Windows.Forms.Timer autoSaveTimer;

        // Biến để liên kết với thanh StatusStrip trên giao diện
        private ToolStripStatusLabel myStatusLabel;
        // Biến để liên kết với thanh StatusStrip trên giao diện
        // ==== ZaloOA ====
        private WebView2 webViewZalo;
        // ===============================
        private WebView2 webViewOmicall;

        public Form1()
        {
            this.Text = "Ghi Issue v4.6";

            if (!Directory.Exists(Application.StartupPath)) Directory.CreateDirectory(Application.StartupPath);

            InitializeComponent();
            //SetupQuickFilter();
            this.Load += Form1_Load;

            // KẾT NỐI CÁC NÚT CƠ BẢN
            if (btnCreateTicket != null)
            {
                btnCreateTicket.Click -= btnCreateTicket_Click;
                btnCreateTicket.Click += btnCreateTicket_Click;
            }
            //if (btnAddRow != null)
            //{
            //    btnAddRow.Click -= btnAddRow_Click;
            //    btnAddRow.Click += btnAddRow_Click;
            //}
            if (btnExecute != null)
            {
                btnExecute.Click -= btnExecute_Click;
                btnExecute.Click += btnExecute_Click;
            }

            // KẾT NỐI CÁC NÚT TÍNH NĂNG MỞ RỘNG
            //if (this.Controls.Find("btnRetry", true).FirstOrDefault() is Button btnRetry)
            //{
            //    btnRetry.Click -= BtnRetry_Click;
            //    btnRetry.Click += BtnRetry_Click;
            //}
            if (this.Controls.Find("btnTheme", true).FirstOrDefault() is Button btnTheme)
            {
                btnTheme.Click -= btnTheme_Click;
                btnTheme.Click += btnTheme_Click;
            }
            if (this.Controls.Find("btnToggleCategory", true).FirstOrDefault() is Button btnToggle)
            {
                btnToggle.Click -= btnToggleCategory_Click;
                btnToggle.Click += btnToggleCategory_Click;
            }
            if (this.Controls.Find("btnBulkEdit", true).FirstOrDefault() is Button btnBulkEdit)
            {
                btnBulkEdit.Click -= BtnBulkEdit_Click;
                btnBulkEdit.Click += BtnBulkEdit_Click;
            }
            //if (this.Controls.Find("btnCheckRecent", true).FirstOrDefault() is Button btnCheckRecent)
            //{
            //    btnCheckRecent.Click -= BtnCheckRecent_Click;
            //    btnCheckRecent.Click += BtnCheckRecent_Click;
            //}
            if (this.Controls.Find("btnViewOpen", true).FirstOrDefault() is Button btnViewOpen)
            {
                btnViewOpen.Click -= BtnViewOpen_Click;
                btnViewOpen.Click += BtnViewOpen_Click;
            }
            if (this.Controls.Find("btnSyncToken", true).FirstOrDefault() is Button btnSyncToken)
            {
                btnSyncToken.Click -= BtnSyncToken_Click;
                btnSyncToken.Click += BtnSyncToken_Click;
            }
            if (this.Controls.Find("btnSearch", true).FirstOrDefault() is Button btnSearch)
            {
                btnSearch.Click -= btnSearch_Click;
                btnSearch.Click += btnSearch_Click;
            }
            if (this.Controls.Find("btnAdmin", true).FirstOrDefault() is Button btnAdminFound)
            {
                btnAdminFound.Click -= btnAdmin_Click;
                btnAdminFound.Click += btnAdmin_Click;
            }
            if (dgvCreateTickets != null)
            {
                dgvCreateTickets.DataError += DgvCreateTickets_DataError;
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                dgvCreateTickets.CurrentCellDirtyStateChanged += DgvCreateTickets_CurrentCellDirtyStateChanged;
                dgvCreateTickets.EditingControlShowing += DgvCreateTickets_EditingControlShowing;
                dgvCreateTickets.CellValidating += DgvCreateTickets_CellValidating;
                //dgvCreateTickets.CellEndEdit += (s, e) => SaveDraft();
                dgvCreateTickets.CellEndEdit += DgvCreateTickets_CellEndEdit;
                dgvCreateTickets.RowsRemoved += (s, e) => SaveDraft();

                // Kích hoạt đếm dòng
                dgvCreateTickets.SelectionChanged += (s, e) => UpdateStatusCount();
                dgvCreateTickets.RowsAdded += (s, e) => UpdateStatusCount();
                dgvCreateTickets.RowsRemoved += (s, e) => UpdateStatusCount();
            }

            if (dgvTickets != null)
            {
                // Bật mỏ neo để ép bảng Đóng Issue tự co giãn bám sát viền Form
                dgvTickets.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

                // Kích hoạt đếm dòng
                dgvTickets.SelectionChanged += (s, e) => UpdateStatusCount();
                // Cắm điện tính năng tự Việt Hóa và Bấm cột để Sắp xếp
                dgvTickets.DataBindingComplete += DgvTickets_DataBindingComplete;
                dgvTickets.ColumnHeaderMouseClick += DgvTickets_ColumnHeaderMouseClick;
                // --- THÊM DÒNG NÀY ĐỂ BẬT TÍNH NĂNG SỬA PHIẾU ---
                dgvTickets.CellDoubleClick += DgvTickets_CellDoubleClick;
                // --- MENU CHUỘT PHẢI ---
                ContextMenuStrip ticketMenu = new ContextMenuStrip();
                ticketMenu.Items.Add("✏️ Chỉnh sửa phiếu", null, (s, e) =>
                {
                    if (dgvTickets.CurrentRow != null)
                        DgvTickets_CellDoubleClick(dgvTickets, new DataGridViewCellEventArgs(dgvTickets.CurrentCell.ColumnIndex, dgvTickets.CurrentRow.Index));
                });
                ticketMenu.Items.Add("-");
                ticketMenu.Items.Add("🗑️ Xóa phiếu (Trên hệ thống)", null, (s, e) => DeleteSelectedTicketsAsync());
                ticketMenu.Items.Add("-"); // Đường gạch ngang
                ticketMenu.Items.Add("📊 Xuất danh sách ra Excel", null, (s, e) => ExportTicketsToExcel());
                dgvTickets.ContextMenuStrip = ticketMenu;
                // --------------------------------------------------
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
                if (strip.Parent != this)
                {
                    strip.Parent = this;
                    strip.Dock = DockStyle.Bottom;
                    strip.BringToFront();
                }

                var foundItems = strip.Items.Find("lblStatusCount", true);
                if (foundItems.Length > 0 && foundItems[0] is ToolStripStatusLabel)
                {
                    myStatusLabel = (ToolStripStatusLabel)foundItems[0];
                }
                else if (strip.Items.Count > 0 && strip.Items[0] is ToolStripStatusLabel fallbackLbl)
                {
                    myStatusLabel = fallbackLbl;
                }
            }

            if (myStatusLabel != null)
            {
                myStatusLabel.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                myStatusLabel.ForeColor = Color.DarkBlue;
                UpdateStatusCount();
            }
        }

        // =========================================================================
        // HÀM CHUYÊN DỤNG: BÓC VỎ JSON VÀ DỊCH MÃ ID SANG TÊN TIẾNG VIỆT
        // =========================================================================
        private void FormatTicketDisplay(List<TicketItem> list, Employee defaultEmp = null)
        {
            foreach (var ticket in list)
            {
                // 1. Dịch ID Người xử lý
                if (ticket.assignee_contact_ids != null && ticket.assignee_contact_ids.Count > 0)
                {
                    var names = ticket.assignee_contact_ids.Select(je =>
                    {
                        string id = je.ValueKind == JsonValueKind.Object && je.TryGetProperty("id", out var idProp) ? idProp.GetString() : je.ToString();
                        return employees.FirstOrDefault(e => e.Id == id)?.Name ?? id;
                    });
                    ticket.NguoiNhan = string.Join(", ", names);
                }
                else if (defaultEmp != null) ticket.NguoiNhan = defaultEmp.Name;

                // 2. Dịch ID Tag (Bóc vỏ Object nếu OmiCRM gửi dạng {id: "..."})
                if (ticket.tags != null && ticket.tags.Count > 0)
                {
                    var tNames = ticket.tags.Select(je =>
                    {
                        string id = je.ValueKind == JsonValueKind.Object && je.TryGetProperty("id", out var idProp) ? idProp.GetString() : je.ToString();
                        return tagList.FirstOrDefault(t => t.id == id)?.name ?? id;
                    });
                    ticket.TenTag = string.Join(", ", tNames);
                }
                else ticket.TenTag = "";
                // 🌟 TỰ ĐỘNG DÒ TYPE ISSUE DỰA TRÊN TÊN PHIẾU
                // 🌟 TỰ ĐỘNG DÒ NHÓM DỰA TRÊN TÊN PHIẾU
                var match = defaultTitles.FirstOrDefault(t => t.Title.Equals(ticket.name?.Trim(), StringComparison.OrdinalIgnoreCase));

                // ❌ ĐÃ XÓA LỆNH TỰ MAP TYPE ISSUE ĐỂ KHÔNG GHI ĐÈ LÊN DỮ LIỆU THỰC TẾ ĐÃ LƯU

                ticket.group_name = (match != null && !string.IsNullOrEmpty(match.Group) && match.Group != "Khác (Nhập tay)") ? match.Group : "Khác";
                // 🌟 BÓC TÁCH THỜI GIAN TỪ THẺ HTML TÀNG HÌNH
                // 🌟 BÓC TÁCH THỜI GIAN THEO CHUẨN MỚI
                // 🌟 BÓC TÁCH THỜI GIAN LINH HOẠT TỪ MÔ TẢ OMiCRM
                // 🌟 BÓC TÁCH THỜI GIAN VÀ TYPE ISSUE TỪ MÔ TẢ OMiCRM
                // 🌟 BÓC TÁCH THỜI GIAN VÀ TYPE ISSUE TỪ MÔ TẢ OMiCRM (LẤY CÁI MỚI NHẤT)
                if (!string.IsNullOrEmpty(ticket.description))
                {
                    // 1. Bóc Thời gian
                    var regexTG = new System.Text.RegularExpressions.Regex(@"\[TG:\s*(.*?)\s*-\s*(.*?)\]");
                    var matchesTG = regexTG.Matches(ticket.description);
                    if (matchesTG.Count > 0)
                    {
                        var mTG = matchesTG[matchesTG.Count - 1]; // Lấy giá trị mới nhất (Cuối cùng)
                        ticket.ThoiGianNhan = mTG.Groups[1].Value.Trim();
                        ticket.ThoiGianXong = mTG.Groups[2].Value.Trim();
                    }
                    ticket.description = regexTG.Replace(ticket.description, ""); // Xóa SẠCH toàn bộ rác TG cũ

                    // 2. Bóc Type Issue
                    var regexType = new System.Text.RegularExpressions.Regex(@"\[Type:\s*(.*?)\]");
                    var matchesType = regexType.Matches(ticket.description);
                    if (matchesType.Count > 0)
                    {
                        var mType = matchesType[matchesType.Count - 1]; // Lấy giá trị mới nhất
                        ticket.TypeIssue = mType.Groups[1].Value.Trim();
                    }
                    ticket.description = regexType.Replace(ticket.description, ""); // Xóa SẠCH toàn bộ rác Type cũ

                    // Tẩy sạch rác HTML còn sót lại
                    ticket.description = ticket.description.Replace("<br><br>", "").Replace("<br>", "").Trim();
                }
                // 🌟 BÓC TÁCH THỜI GIAN TỪ MÔ TẢ RA CỘT RIÊNG
                //var regex = new System.Text.RegularExpressions.Regex(@"\[TG: (.*?) - (.*?)\]");
                //var m = regex.Match(ticket.MoTa);
                //if (m.Success)
                //{
                //    ticket.ThoiGianNhan = m.Groups[1].Value;
                //    ticket.ThoiGianXong = m.Groups[2].Value;
                //    // Tẩy xóa phần thời gian khỏi Mô tả để người dùng không thấy sự "lén lút" này
                //    //ticket.MoTa = ticket.MoTa.Replace(m.Value, "").Trim();
                //    ticket.description = ticket.description?.Replace(m.Value, "").Trim();
                //}
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

                myStatusLabel.Text = $"📝 Tạo Phiếu: {createTotal} dòng (Đang chọn: {createSelected})   |   🗂️ Đóng Issue: {ticketTotal} phiếu (Đang chọn: {ticketSelected})";
            }
            catch { }
        }

        // =========================================================================
        // TÍNH NĂNG MỞ RỘNG: RETRY, XEM PHIẾU (TẠO/ĐÓNG), SỬA HÀNG LOẠT
        // =========================================================================

        //private void BtnRetry_Click(object sender, EventArgs e)
        //{
        //    int retryCount = 0;
        //    foreach (DataGridViewRow row in dgvCreateTickets.Rows)
        //    {
        //        if (row.IsNewRow) continue;
        //        string currentResult = row.Cells["colResult"].Value?.ToString() ?? "";

        //        if (currentResult.Contains("❌"))
        //        {
        //            row.Cells["colResult"].Value = ""; // Xóa lỗi để kích hoạt gửi lại
        //            retryCount++;
        //        }
        //    }

        //    if (retryCount == 0)
        //    {
        //        MessageBox.Show("Không có phiếu nào bị lỗi để gửi lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    btnCreateTicket_Click(null, null);
        //}

        //private async void BtnCheckRecent_Click(object sender, EventArgs e)
        //{
        //    Employee selectedEmp = (Employee)cboEmployees.SelectedItem;
        //    if (selectedEmp == null) return;

        //    var searchBody = new
        //    {
        //        status_filters = new[] { "active_state" },
        //        assignee_contact_ids = new[] { selectedEmp.Id },
        //        additional_layout = new[] { "object_association" },
        //        has_notify_report = true,
        //        page = "1",
        //        size = "1000",
        //        current_status = new[] { 0, 1, 2, 3 }
        //    };

        //    var searchContent = new StringContent(JsonSerializer.Serialize(searchBody), Encoding.UTF8, "application/json");
        //    Cursor.Current = Cursors.WaitCursor;

        //    using (HttpClient client = new HttpClient())
        //    {
        //        client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

        //        try
        //        {
        //            HttpResponseMessage searchResponse = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/search?lng=vi&utm_source=web", searchContent);
        //            string resStr = await searchResponse.Content.ReadAsStringAsync();

        //            if (searchResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized || resStr.Contains("Unauthorized"))
        //            {
        //                Cursor.Current = Cursors.Default;
        //                MessageBox.Show("Phiên đăng nhập đã hết hạn!\nVui lòng đăng nhập lại.", "Cảnh Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
        //                OMICRM_TOKEN = "";
        //                await PerformLoginSequenceAsync();
        //                return;
        //            }

        //            if (searchResponse.IsSuccessStatusCode)
        //            {
        //                OmiResponse omiData = JsonSerializer.Deserialize<OmiResponse>(resStr);
        //                if (omiData?.payload?.items != null && omiData.payload.items.Count > 0)
        //                {
        //                    dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

        //                    foreach (var t in omiData.payload.items)
        //                    {
        //                        int idx = dgvCreateTickets.Rows.Add();
        //                        var row = dgvCreateTickets.Rows[idx];

        //                        row.Cells["colTitle"].Value = t.name;
        //                        row.Cells["colCategory"].Value = "FO_1_POS";
        //                        row.Cells["colSubCategory"].Value = "POS_Lỗi kết nối";
        //                        row.Cells["colAssignee"].Value = selectedEmp.Id;
        //                        row.Cells["colResult"].Value = "☁️ Đã có trên hệ thống";

        //                        row.DefaultCellStyle.BackColor = Color.LightGreen;
        //                        row.ReadOnly = true;
        //                    }

        //                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
        //                    Cursor.Current = Cursors.Default;

        //                    UpdateStatusCount();

        //                    // Ghi log khi lấy phiếu về
        //                    WriteLog("INFO", $"Kéo phiếu đang xử lý về bảng", $"Số lượng: {omiData.payload.items.Count} phiếu | Của: {selectedEmp.Name}");

        //                    MessageBox.Show($"Đã kéo về {omiData.payload.items.Count} phiếu đang hoạt động trên hệ thống.\nHãy kéo xuống cuối bảng để kiểm tra!", "Tải Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                }
        //                else
        //                {
        //                    Cursor.Current = Cursors.Default;
        //                    MessageBox.Show("Không tìm thấy phiếu nào của bạn trên hệ thống!", "Trống", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Cursor.Current = Cursors.Default;
        //            MessageBox.Show($"Có lỗi xảy ra: {ex.Message}", "Lỗi Code");
        //        }
        //    }
        //}




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

                    // Bật cỗ máy phiên dịch ID sang Tên trước khi nhét vào bảng
                    FormatTicketDisplay(allTickets, selectedEmp);

                    Cursor.Current = Cursors.Default;
                    dgvTickets.DataSource = allTickets;

                    if (dgvTickets.Columns.Contains("_id")) dgvTickets.Columns["_id"].Visible = false;
                    if (dgvTickets.Columns.Contains("id")) dgvTickets.Columns["id"].Visible = false;
                    if (dgvTickets.Columns.Contains("current_status")) dgvTickets.Columns["current_status"].Visible = false;

                    if (dgvTickets.Columns.Contains("unique_id")) { dgvTickets.Columns["unique_id"].HeaderText = "ID Phiếu"; dgvTickets.Columns["unique_id"].DisplayIndex = 0; dgvTickets.Columns["unique_id"].Width = 80; }
                    if (dgvTickets.Columns.Contains("name")) { dgvTickets.Columns["name"].HeaderText = "Tên Phiếu"; dgvTickets.Columns["name"].DisplayIndex = 1; }
                    if (dgvTickets.Columns.Contains("TrangThaiHienThi")) { dgvTickets.Columns["TrangThaiHienThi"].HeaderText = "Trạng Thái"; dgvTickets.Columns["TrangThaiHienThi"].DisplayIndex = 2; dgvTickets.Columns["TrangThaiHienThi"].Width = 140; }
                    if (dgvTickets.Columns.Contains("NguoiNhan")) { dgvTickets.Columns["NguoiNhan"].HeaderText = "Người Nhận"; dgvTickets.Columns["NguoiNhan"].DisplayIndex = 3; }
                    if (dgvTickets.Columns.Contains("TenTag")) { dgvTickets.Columns["TenTag"].HeaderText = "Tags"; dgvTickets.Columns["TenTag"].DisplayIndex = 4; }
                    if (dgvTickets.Columns.Contains("tags")) dgvTickets.Columns["tags"].Visible = false;
                    if (dgvTickets.Columns.Contains("assignee_contact_ids")) dgvTickets.Columns["assignee_contact_ids"].Visible = false;

                    dgvTickets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    UpdateQuickFilter(allTickets);
                    UpdateStatusCount();

                    // Ghi log khi xem danh sách phiếu
                    WriteLog("INFO", $"Xem danh sách phiếu cần Đóng", $"Số lượng: {allTickets.Count} phiếu | Của: {selectedEmp.Name}");

                    MessageBox.Show($"Đang hiển thị {allTickets.Count} phiếu đang mở. \n(Dữ liệu chỉ hiển thị, chưa thực hiện lệnh Đóng)", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show($"Có lỗi xảy ra: {ex.Message}", "Lỗi Code");
                }
            }
        }

        private async void BtnSyncToken_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show(
                "Bạn có muốn đăng nhập lại để lấy Token và làm mới TOÀN BỘ dữ liệu (Tiêu đề Sheet, Tag, Người xử lý...) không?",
                "Đồng bộ Token",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dialogResult == DialogResult.Yes)
            {
                if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
                OMICRM_TOKEN = "";

                // 1. Lấy Token mới và load lại Tag / Category
                await PerformLoginSequenceAsync();

                // 2. ÉP TẢI LẠI GOOGLE SHEET MỚI NHẤT
                await SyncTitlesBackgroundAsync();

                UpdateStatusCount();

                MessageBox.Show("Đã đồng bộ Token và tải lại toàn bộ danh mục MỚI NHẤT!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            // Đã xóa Chủ đề & Phân loại khỏi Bulk Edit vì đã gán cứng
            cbColumn.Items.AddRange(new string[] { "Tiêu đề", "Tag", "Người xử lý" });

            Label lbl2 = new Label() { Left = 20, Top = 80, Text = "Chọn Giá trị mới:", AutoSize = true };
            ComboBox cbValue = new ComboBox() { Left = 20, Top = 100, Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };

            cbColumn.SelectedIndexChanged += (s, ev) =>
            {
                cbValue.DataSource = null;
                cbValue.Items.Clear();
                string selCol = cbColumn.SelectedItem.ToString();

                if (selCol == "Tiêu đề") { cbValue.DataSource = defaultTitles; cbValue.DisplayMember = "Title"; cbValue.ValueMember = "Title"; }
                else if (selCol == "Tag") { cbValue.DataSource = tagList; cbValue.DisplayMember = "name"; cbValue.ValueMember = "id"; }
                else if (selCol == "Người xử lý") { cbValue.DataSource = employees; cbValue.DisplayMember = "Name"; cbValue.ValueMember = "Id"; }
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
                            var match = defaultTitles.FirstOrDefault(t => t.Title.Equals(newVal.ToString(), StringComparison.OrdinalIgnoreCase));
                            if (match != null) row.Cells["colGroup"].Value = match.Group;
                        }
                        else if (selCol == "Tag") { row.Cells["colTag"].Value = newVal; }
                        else if (selCol == "Người xử lý") { row.Cells["colAssignee"].Value = newVal; }
                    }
                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;

                    // Ghi log hành động Sửa hàng loạt
                    string displayValue = cbValue.Text; // Lấy tên hiển thị chữ thay vì ID
                    WriteLog("INFO", $"Sửa hàng loạt {dgvCreateTickets.SelectedRows.Count} dòng", $"Cột: {selCol} -> Giá trị mới: {displayValue}");

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

        //private async void SetupOmicallWebView()
        //{
        //    webViewOmicall = new WebView2();

        //    // TRẢ LẠI 1 DÒNG NÀY (Ép giãn full Tab, không chừa khoảng trắng nữa)
        //    webViewOmicall.Dock = DockStyle.Fill;

        //    tabOmicall.Controls.Add(webViewOmicall);

        //    try
        //    {
        //        string userAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        //        string userDataFolder = Path.Combine(userAppData, "Dcorp_GhiIssue_OmicallData");

        //        var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
        //        await webViewOmicall.EnsureCoreWebView2Async(environment);

        //        webViewOmicall.CoreWebView2.PermissionRequested += (sender, args) =>
        //        {
        //            args.State = CoreWebView2PermissionState.Allow;
        //        };

        //        webViewOmicall.CoreWebView2.Navigate("https://dcorpsupport.omicrm.com/integrated/social/general?lng=vi");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Lỗi tải Omicall: " + ex.Message);
        //    }
        //}

        private async void SetupOmicallWebView()
        {
            webViewOmicall = new WebView2();
            webViewOmicall.Dock = DockStyle.Fill;
            tabOmicall.Controls.Add(webViewOmicall);

            try
            {
                string userAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userDataFolder = Path.Combine(userAppData, "Dcorp_GhiIssue_OmicallData");

                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
                await webViewOmicall.EnsureCoreWebView2Async(environment);

                // ==============================================================================
                // 🌟 1. THÊM 2 DÒNG NÀY ĐỂ BẬT TÍNH NĂNG "LƯU MẬT KHẨU" NHƯ TRÌNH DUYỆT THẬT
                // ==============================================================================
                webViewOmicall.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
                webViewOmicall.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;
                // ==============================================================================

                webViewOmicall.CoreWebView2.PermissionRequested += (sender, args) =>
                {
                    args.State = CoreWebView2PermissionState.Allow;
                };

                // ==============================================================================
                // 🌟 2. BỔ SUNG: CHẶN MỞ TAB MỚI (Giúp trải nghiệm trực chat không bị gián đoạn)
                // Nếu click vào link ảnh/tài liệu, nó sẽ mở đè ngay trong khung Tool luôn
                // ==============================================================================
                webViewOmicall.CoreWebView2.NewWindowRequested += (sender, args) =>
                {
                    args.Handled = true;
                    webViewOmicall.CoreWebView2.Navigate(args.Uri);
                };
                // ==============================================================================

                webViewOmicall.CoreWebView2.Navigate("https://dcorpsupport.omicrm.com/integrated/social/general?lng=vi");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải Omicall: " + ex.Message);
            }
        }

        private async void SetupZaloWebView()
        {
            webViewZalo = new WebView2();
            webViewZalo.Dock = DockStyle.Fill;
            tabZalo.Controls.Add(webViewZalo);

            try
            {
                string userAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userDataFolder = Path.Combine(userAppData, "Dcorp_GhiIssue_ZaloData");

                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);

                await webViewZalo.EnsureCoreWebView2Async(environment);

                // =========================================================================
                // 🌟 BẮT ĐẦU ĐOẠN CODE "BẢO VỆ" TRỊ BỆNH LỖI QR CỦA ZALO 🌟
                // =========================================================================

                // 1. Bắt quả tang Zalo đòi chuyển hướng láo và ép quay xe
                webViewZalo.CoreWebView2.NavigationStarting += (sender, args) =>
                {
                    string url = args.Uri.ToString();
                    if (url.Contains("oa.zalo.me_zoauth") || url.StartsWith("zalo:"))
                    {
                        args.Cancel = true; // Khóa mõm, không cho chuyển trang
                        webViewZalo.CoreWebView2.Navigate("https://oa.zalo.me/chatv2"); // Bế thẳng vào phòng chat
                    }
                };

                // 2. Chặn Zalo mở Tab mới (Ví dụ bấm vào hình ảnh, file nó hay đòi bật tab mới)
                webViewZalo.CoreWebView2.NewWindowRequested += (sender, args) =>
                {
                    args.Handled = true; // Báo là "Tôi tự lo được, không cần mở tab popup"
                    webViewZalo.CoreWebView2.Navigate(args.Uri); // Mở đè ngay trên khung chat này
                };

                // =========================================================================
                // 🌟 KẾT THÚC ĐOẠN CODE BẢO VỆ 🌟
                // =========================================================================

                webViewZalo.CoreWebView2.Navigate("https://oa.zalo.me/chatv2");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải Zalo OA: " + ex.Message);
            }
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            // BƯỚC 1: LOAD DỮ LIỆU OFFLINE TỪ LẦN TRƯỚC
            if (File.Exists(titlesCachePath)) { try { defaultTitles = JsonSerializer.Deserialize<BindingList<PredefinedTitle>>(File.ReadAllText(titlesCachePath)); } catch { } }
            else { LoadDefaultTitlesFallback(); } // Nếu chưa có thì dùng tạm danh sách cứng

            if (File.Exists(catsCachePath)) { try { categoryList = JsonSerializer.Deserialize<List<CategoryItem>>(File.ReadAllText(catsCachePath)); } catch { } }
            if (File.Exists(tagsCachePath)) { try { tagList = JsonSerializer.Deserialize<List<TagItem>>(File.ReadAllText(tagsCachePath)); } catch { } }
            // LÔI TRÍ NHỚ TỪ LẦN TẮT MÁY TRƯỚC RA
            //if (File.Exists(catCachePath))
            //{
            //    try
            //    {
            //        string[] parts = File.ReadAllText(catCachePath).Split('|');
            //        if (parts.Length == 2) { defaultCat = parts[0]; defaultSubCat = parts[1]; }
            //    }
            //    catch { }
            //}

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
                cboEmployees.TextUpdate += cboEmployees_TextUpdate;
            }

            SetupCreateTicketGrid();
            CheckAndDownloadUpdateAsync();
            await PerformLoginSequenceAsync();

            AutoAssignFromEmail(loginEmailCached);
            HookStatusStrip();
            LoadDraft();

            //if (dgvCreateTickets.Rows.Count <= 1) btnAddRow_Click(null, null);
            EnsureSufficientRows(100);
            // BƯỚC 2: CHẠY NGẦM BACKGROUND ĐỂ TẢI SHEET & API TAG
            _ = Task.Run(async () =>
            {
                await SyncTitlesBackgroundAsync();
                if (!string.IsNullOrEmpty(OMICRM_TOKEN))
                {
                    await SyncCategoriesBackgroundAsync();
                    await SyncTagsBackgroundAsync();
                }
            });
            // LOAD GIAO DIỆN MÀU SẮC NGƯỜI DÙNG ĐÃ CUSTOM
            string themePath = Path.Combine(Application.StartupPath, "theme.txt");
            if (File.Exists(themePath))
            {
                try
                {
                    string[] c = File.ReadAllText(themePath).Split('|');
                    if (c.Length == 4) ApplyTheme(ColorTranslator.FromHtml(c[0]), ColorTranslator.FromHtml(c[1]), ColorTranslator.FromHtml(c[2]), ColorTranslator.FromHtml(c[3]));
                    else ApplyTheme(ColorTranslator.FromHtml(c[0]), Color.WhiteSmoke, Color.White, SystemColors.Highlight); // Chống lỗi cho file theme cũ
                }
                catch { }
            }
            // ==== THÊM ĐÚNG 1 DÒNG NÀY VÀO ĐÂY ====
            SetupZaloWebView();
            SetupOmicallWebView();

            if (adminEmails.Any(email => loginEmailCached.ToLower().Contains(email.ToLower())))
            {
                btnAdmin.Visible = true;
            }
            _ = Task.Run(async () => await SyncRemoteConfigAsync());
            // ======================================
        }

        private void LoadDefaultTitlesFallback()
        {
            defaultTitles = new BindingList<PredefinedTitle> {
                new PredefinedTitle { Group = "Offline Backup", Title = "Mất mạng - Hãy kiểm tra lại kết nối" }
            };
        }
        private async Task SyncTitlesBackgroundAsync()
        {
            string sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRGeOsJ71sqpI__BcxboJTp-SvETPWY4-H21rs2mnTQb_K1LRSJJaAkBIP8TyGjgdidcF47gH_lCfUJ/pub?gid=1073183768&single=true&output=tsv";
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string tsvData = await client.GetStringAsync(sheetUrl);
                    string[] lines = tsvData.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    var freshList = new BindingList<PredefinedTitle>();
                    string lastGroup = "Khác (Nhập tay)";

                    HashSet<string> tempTypes = new HashSet<string>();
                    HashSet<string> tempActions = new HashSet<string>();

                    for (int i = 1; i < lines.Length; i++)
                    {
                        string[] cols = lines[i].Split('\t');
                        if (cols.Length >= 2)
                        {
                            string groupName = cols[0].Trim();
                            string titleName = cols[1].Trim();
                            string typeIssueVi = cols.Length >= 3 ? cols[2].Trim() : "";
                            string techAction = cols.Length >= 4 ? cols[3].Trim() : "";

                            // 🌟 HÚT THÊM CỘT E (TIẾNG ANH - INDEX 4)
                            string typeIssueEn = cols.Length >= 5 ? cols[4].Trim() : "";

                            if (!string.IsNullOrEmpty(groupName)) lastGroup = groupName; else groupName = lastGroup;
                            if (!string.IsNullOrEmpty(titleName))
                            {
                                freshList.Add(new PredefinedTitle { Group = groupName, Title = titleName, TypeIssue = typeIssueVi, TechAction = techAction });
                            }

                            if (!string.IsNullOrEmpty(typeIssueVi))
                            {
                                tempTypes.Add(typeIssueVi);
                                // 🌟 Nạp vào từ điển (Nếu Google Sheet chưa dịch kịp thì lấy tiếng Việt xài đỡ)
                                dictTypeEng[typeIssueVi] = !string.IsNullOrEmpty(typeIssueEn) ? typeIssueEn : typeIssueVi;
                            }
                            if (!string.IsNullOrEmpty(techAction)) tempActions.Add(techAction);
                        }
                    }

                    if (freshList.Count > 0)
                    {
                        File.WriteAllText(titlesCachePath, JsonSerializer.Serialize(freshList));
                        this.Invoke(new Action(() =>
                        {
                            defaultTitles = freshList;

                            defaultTypeIssues.Clear(); foreach (var t in tempTypes) defaultTypeIssues.Add(new ComboItem { Text = t });
                            defaultTechActions.Clear(); foreach (var a in tempActions) defaultTechActions.Add(new ComboItem { Text = a });

                            if (dgvCreateTickets != null)
                            {
                                if (dgvCreateTickets.Columns.Contains("colTitle")) ((DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colTitle"]).DataSource = defaultTitles;
                                if (dgvCreateTickets.Columns.Contains("colTypeIssue")) ((DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colTypeIssue"]).DataSource = defaultTypeIssues;
                                if (dgvCreateTickets.Columns.Contains("colDesc")) ((DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colDesc"]).DataSource = defaultTechActions;
                            }
                        }));
                    }

                    if (freshList.Count > 0)
                    {
                        File.WriteAllText(titlesCachePath, JsonSerializer.Serialize(freshList));
                        this.Invoke(new Action(() =>
                        {
                            defaultTitles = freshList;

                            // FIX LỖI BỊ "MÙ" Ô TIÊU ĐỀ: Báo cho cái cột biết là danh sách đã có bản mới!
                            if (dgvCreateTickets != null && dgvCreateTickets.Columns.Contains("colTitle"))
                            {
                                var col = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colTitle"];
                                col.DataSource = defaultTitles;
                            }
                        }));
                    }
                }
            }
            catch { }
        }

        private async Task SyncCategoriesBackgroundAsync()
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
                        if (catData?.payload != null && catData.payload.Count > 0)
                        {
                            var freshList = catData.payload.Where(c => !string.IsNullOrEmpty(c.name) && c.types != null && c.types.Count > 0).ToList();
                            File.WriteAllText(catsCachePath, JsonSerializer.Serialize(freshList));
                            this.Invoke(new Action(() =>
                            {
                                categoryList = freshList;

                                // Bơm danh sách mới vào Não của 2 cột
                                if (dgvCreateTickets != null && dgvCreateTickets.Columns.Contains("colCategory"))
                                {
                                    var colCat = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colCategory"];
                                    colCat.DataSource = categoryList.ToList();

                                    var colSub = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colSubCategory"];
                                    var allSubs = categoryList.Where(c => c.types != null).SelectMany(c => c.types).ToList();
                                    colSub.DataSource = allSubs;
                                }
                            }));
                        }
                    }
                }
                catch { }
            }
        }

        private async Task SyncTagsBackgroundAsync()
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);
                try
                {
                    var freshList = new List<TagItem>();
                    int currentPage = 1; int pageSize = 1000; bool hasMoreData = true;

                    while (hasMoreData)
                    {
                        var body = new { size = pageSize, page = currentPage };
                        var content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
                        HttpResponseMessage response = await client.PostAsync("https://tenant-config-stg.omicrm.com/tag/search-all?lng=vi&utm_source=web", content);

                        if (!response.IsSuccessStatusCode) response = await client.GetAsync($"https://tenant-config-stg.omicrm.com/tag/search-all?lng=vi&utm_source=web&size={pageSize}&page={currentPage}");

                        if (response.IsSuccessStatusCode)
                        {
                            string json = await response.Content.ReadAsStringAsync();
                            TagResponse tagData = JsonSerializer.Deserialize<TagResponse>(json);
                            if (tagData?.payload?.items != null && tagData.payload.items.Count > 0) { freshList.AddRange(tagData.payload.items); currentPage++; }
                            else hasMoreData = false;
                        }
                        else hasMoreData = false;
                    }

                    if (freshList.Count > 0)
                    {
                        File.WriteAllText(tagsCachePath, JsonSerializer.Serialize(freshList));
                        this.Invoke(new Action(() =>
                        {
                            tagList = freshList;
                            if (dgvCreateTickets != null && dgvCreateTickets.Columns.Contains("colTag"))
                                ((DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colTag"]).DataSource = tagList;
                        }));
                    }
                }
                catch { }
            }
        }

        // =========================================================================
        // HÀM TẢI TIÊU ĐỀ TỪ GOOGLE SHEET BẤT TỬ (V4.6 FIX LỖI MERGED CELL)
        // =========================================================================
        private async Task LoadDefaultTitlesAsync()
        {
            defaultTitles = new BindingList<PredefinedTitle>();
            string sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRGeOsJ71sqpI__BcxboJTp-SvETPWY4-H21rs2mnTQb_K1LRSJJaAkBIP8TyGjgdidcF47gH_lCfUJ/pub?gid=1073183768&single=true&output=tsv";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string tsvData = await client.GetStringAsync(sheetUrl);
                    string[] lines = tsvData.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                    string lastGroup = "Khác (Nhập tay)"; // Tên nhóm dự phòng ban đầu

                    for (int i = 1; i < lines.Length; i++)
                    {
                        string[] cols = lines[i].Split('\t');
                        if (cols.Length >= 2)
                        {
                            string groupName = cols[0].Trim();
                            string titleName = cols[1].Trim();

                            // Tự động lấy tên Nhóm của ô phía trên nếu ô hiện tại bị trống/gộp
                            if (!string.IsNullOrEmpty(groupName))
                            {
                                lastGroup = groupName;
                            }
                            else
                            {
                                groupName = lastGroup;
                            }

                            if (!string.IsNullOrEmpty(titleName))
                            {
                                defaultTitles.Add(new PredefinedTitle { Group = groupName, Title = titleName });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể tải Tiêu đề từ Google Sheet. Tạm thời sử dụng danh sách offline.\nLỗi: " + ex.Message, "Cảnh báo Mạng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                defaultTitles = new BindingList<PredefinedTitle>
                {
                    new PredefinedTitle { Group = "Offline Backup", Title = "Mất mạng - Hãy kiểm tra lại kết nối" }
                };
            }
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


        // ================== HỆ THỐNG GHI LOG ==================
        private static readonly object _logLock = new object();

        // Hàm ghi log tổng quát mới
        private void WriteLog(string level, string context, string detail = "")
        {
            try
            {
                lock (_logLock)
                {
                    // Đổi tên file thành GhiIssue.log cho chung chung (lưu cả lỗi lẫn thành công)
                    string logFile = Path.Combine(Application.StartupPath, "GhiIssue.log");
                    if (File.Exists(logFile) && new FileInfo(logFile).Length > 2 * 1024 * 1024)
                    {
                        File.Delete(logFile);
                    }
                    string logMsg = $"[{DateTime.Now:dd/MM/yyyy HH:mm:ss}] [{level}] {context}\n";
                    if (!string.IsNullOrEmpty(detail)) logMsg += $"Chi tiết: {detail}\n";
                    logMsg += new string('-', 40) + "\n";
                    File.AppendAllText(logFile, logMsg);
                }
            }
            catch { /* Bỏ qua nếu lỗi hệ thống file */ }
        }

        // Giữ lại hàm LogError cũ gọi ké vào hàm mới để code cũ không bị báo đỏ
        private void LogError(string context, Exception ex = null)
        {
            WriteLog("ERROR", context, ex?.Message);
        }

        // ================== BẮT PHÍM TẮT & HACK WINFORMS ==================
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // TÍNH NĂNG MỚI: BẤM CTRL + F ĐỂ TÌM KIẾM MỌI LÚC MỌI NƠI
            if (keyData == (Keys.Control | Keys.F))
            {
                SearchGlobalTicket();
                return true; // Chặn phím tắt lại không truyền đi nữa
            }

            // Bắt trọn ổ: Ép WinForms lướt sang ngang thay vì rớt dòng khi gõ tạo phiếu
            if (keyData == Keys.Enter && (dgvCreateTickets.Focused || dgvCreateTickets.EditingControl != null))
            {
                if (dgvCreateTickets.IsCurrentCellInEditMode) dgvCreateTickets.CommitEdit(DataGridViewDataErrorContexts.Commit);
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // =========================================================================
        // BỘ MÁY TÌM KIẾM NÂNG CAO (LỌC NGÀY, TRẠNG THÁI, NGƯỜI XỬ LÝ & BỎ GIỚI HẠN)
        // =========================================================================
        // =========================================================================
        // BỘ MÁY TÌM KIẾM NÂNG CAO (ĐÃ FIX LỖI 50K PHIẾU BẰNG UNIX TIMESTAMP)
        // =========================================================================
        private async void SearchGlobalTicket()
        {
            Form searchForm = new Form() { Width = 450, Height = 630, Text = "Tìm kiếm Nâng cao OmiCRM", StartPosition = FormStartPosition.CenterScreen, MaximizeBox = false, MinimizeBox = false, FormBorderStyle = FormBorderStyle.FixedDialog };
            int currentY = 15;

            Label lbl1 = new Label() { Left = 20, Top = currentY, Text = "Từ khóa (ID Phiếu, Tên, SĐT... - Để trống để lấy tất cả):", AutoSize = true };
            currentY += 20; TextBox txtSearch = new TextBox() { Left = 20, Top = currentY, Width = 390 }; currentY += 35;

            Label lbl2 = new Label() { Left = 20, Top = currentY, Text = "Phạm vi tìm kiếm (Trạng thái):", AutoSize = true };
            currentY += 20; ComboBox cbStatus = new ComboBox() { Left = 20, Top = currentY, Width = 390, DropDownStyle = ComboBoxStyle.DropDownList };
            cbStatus.Items.Add("Tất cả (Cả Mở và Đã Đóng)"); cbStatus.Items.Add("Chỉ các phiếu Đang mở (0, 1, 2, 3)"); cbStatus.Items.Add("Chỉ các phiếu Đã Đóng (4)"); cbStatus.SelectedIndex = 0; currentY += 35;

            Label lbl3 = new Label() { Left = 20, Top = currentY, Text = "Lọc theo Ngày tạo (Tối đa 31 ngày):", AutoSize = true };
            currentY += 20; ComboBox cbTime = new ComboBox() { Left = 20, Top = currentY, Width = 390, DropDownStyle = ComboBoxStyle.DropDownList };
            cbTime.Items.AddRange(new[] { "Hôm nay", "Hôm qua", "7 Ngày qua", "30 Ngày qua", "Tháng này", "Tháng trước", "Tùy chọn ngày..." }); cbTime.SelectedIndex = 0; currentY += 35;

            DateTimePicker dtpStart = new DateTimePicker() { Left = 20, Top = currentY, Width = 180, Format = DateTimePickerFormat.Short, Enabled = false };
            Label lblTo = new Label() { Left = 210, Top = currentY + 3, Text = "đến", AutoSize = true };
            DateTimePicker dtpEnd = new DateTimePicker() { Left = 245, Top = currentY, Width = 165, Format = DateTimePickerFormat.Short, Enabled = false }; currentY += 45;

            cbTime.SelectedIndexChanged += (s, ev) =>
            {
                DateTime now = DateTime.Now; dtpStart.Enabled = dtpEnd.Enabled = (cbTime.Text == "Tùy chọn ngày...");
                if (cbTime.Text == "Hôm nay") { dtpStart.Value = now.Date; dtpEnd.Value = now.Date; }
                else if (cbTime.Text == "Hôm qua") { dtpStart.Value = now.Date.AddDays(-1); dtpEnd.Value = now.Date.AddDays(-1); }
                else if (cbTime.Text == "7 Ngày qua") { dtpStart.Value = now.Date.AddDays(-6); dtpEnd.Value = now.Date; }
                else if (cbTime.Text == "30 Ngày qua") { dtpStart.Value = now.Date.AddDays(-29); dtpEnd.Value = now.Date; }
                else if (cbTime.Text == "Tháng này") { dtpStart.Value = new DateTime(now.Year, now.Month, 1); dtpEnd.Value = now.Date; }
                else if (cbTime.Text == "Tháng trước") { var lastMonth = now.AddMonths(-1); dtpStart.Value = new DateTime(lastMonth.Year, lastMonth.Month, 1); dtpEnd.Value = new DateTime(now.Year, now.Month, 1).AddDays(-1); }
            };

            Label lbl4 = new Label() { Left = 20, Top = currentY, Text = "Người xử lý (KHÔNG TÍCH = TÌM TẤT CẢ):", AutoSize = true };
            currentY += 20; CheckedListBox clbAssignees = new CheckedListBox() { Left = 20, Top = currentY, Width = 390, Height = 90, CheckOnClick = true };
            foreach (var emp in employees) clbAssignees.Items.Add(emp, false); currentY += 100;

            Label lbl5 = new Label() { Left = 20, Top = currentY, Text = "Lọc theo Tag (Bao gồm cả Tag con):", AutoSize = true };
            TextBox txtSearchTag = new TextBox() { Left = 250, Top = currentY - 3, Width = 160, Text = "🔍 Tìm Tag..." };
            txtSearchTag.GotFocus += (s, ev) => { if (txtSearchTag.Text == "🔍 Tìm Tag...") txtSearchTag.Text = ""; }; currentY += 25;

            CheckedListBox clbTags = new CheckedListBox() { Left = 20, Top = currentY, Width = 390, Height = 100, CheckOnClick = true };
            HashSet<string> checkedTagIds = new HashSet<string>();

            Action loadTags = () =>
            {
                clbTags.Items.Clear(); string kw = ConvertToUnSignStatic(txtSearchTag.Text == "🔍 Tìm Tag..." ? "" : txtSearchTag.Text);
                var filtered = string.IsNullOrEmpty(kw) ? tagList.Take(150).ToList() : tagList.Where(x => x.UnsignedName.Contains(kw)).Take(150).ToList();
                var checkedItemsList = tagList.Where(t => checkedTagIds.Contains(t.id)).ToList();
                var finalDisplayList = checkedItemsList.Union(filtered).Distinct().ToList();
                foreach (var tag in finalDisplayList) clbTags.Items.Add(tag, checkedTagIds.Contains(tag.id));
            }; loadTags();

            clbTags.ItemCheck += (s, ev) => { var tag = clbTags.Items[ev.Index] as TagItem; if (ev.NewValue == CheckState.Checked) checkedTagIds.Add(tag.id); else checkedTagIds.Remove(tag.id); };
            txtSearchTag.TextChanged += (s, ev) => loadTags(); currentY += 115;

            // ================== NÚT THỐNG KÊ (ĐÃ FIX LỖI CRASH) ==================
            // ================== NÚT THỐNG KÊ (ĐÃ CÂN GIỮA GIAO DIỆN) ==================
            int btnWidth = 200; // Cho 2 nút bằng chiều rộng nhau cho đẹp

            // Nút Tìm Kiếm (Nằm trên)
            Button btnDoSearch = new Button() { Text = "Tìm", Left = (searchForm.Width - btnWidth) / 2 - 8, Top = currentY, Width = btnWidth, Height = 35, DialogResult = DialogResult.OK, BackColor = Color.LightBlue, FlatStyle = FlatStyle.Flat };
            currentY += 45;

            // Nút Thống Kê (Nằm dưới)
            Button btnTopStats = new Button() { Text = "📊 Xem Top Lỗi / Tag", Left = (searchForm.Width - btnWidth) / 2 - 8, Top = currentY, Width = btnWidth, Height = 35, BackColor = Color.Orange, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            currentY += 50;

            searchForm.Height = currentY + 40; // Tự động kéo dài đáy Form ra một chút để không bị lẹm nút
            // =======================================================================

            btnTopStats.Click += async (s, ev) =>
            {
                if ((dtpEnd.Value.Date - dtpStart.Value.Date).TotalDays > 31) { MessageBox.Show("Khoảng thời gian quá lớn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                long startMs = new DateTimeOffset(dtpStart.Value.Date).ToUnixTimeMilliseconds(); long endMs = new DateTimeOffset(dtpEnd.Value.Date.AddDays(1).AddTicks(-1)).ToUnixTimeMilliseconds();
                Cursor.Current = Cursors.WaitCursor; btnTopStats.Text = "Đang thống kê..."; btnTopStats.Enabled = false;

                var searchBody = new Dictionary<string, object> { { "ranges", new[] { new { field = "created_date", from = startMs, to = endMs } } }, { "additional_layout", new[] { "object_association" } }, { "has_notify_report", true } };
                if (cbStatus.SelectedIndex == 1) { searchBody["status_filters"] = new[] { "active_state" }; searchBody["current_status"] = new[] { 0, 1, 2, 3 }; } else if (cbStatus.SelectedIndex == 2) { searchBody["current_status"] = new[] { 4 }; }

                List<TicketItem> allTickets = new List<TicketItem>(); int currentPage = 1; int pageSize = 1000; bool hasMore = true;

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);
                    try
                    {
                        while (hasMore)
                        {
                            searchBody["page"] = currentPage; searchBody["size"] = pageSize;
                            var searchContent = new StringContent(System.Text.Json.JsonSerializer.Serialize(searchBody), Encoding.UTF8, "application/json");
                            HttpResponseMessage res = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/search?lng=vi&utm_source=web", searchContent);
                            if (res.IsSuccessStatusCode)
                            {
                                OmiResponse omiData = System.Text.Json.JsonSerializer.Deserialize<OmiResponse>(await res.Content.ReadAsStringAsync());
                                if (omiData?.payload?.items != null && omiData.payload.items.Count > 0) { allTickets.AddRange(omiData.payload.items); if (omiData.payload.items.Count < pageSize) hasMore = false; else currentPage++; } else hasMore = false;
                            }
                            else hasMore = false;
                        }

                        if (allTickets.Count > 0)
                        {
                            foreach (var ticket in allTickets)
                            {
                                string tenPhieuOmi = ticket.name != null ? ticket.name.Trim() : ""; var match = defaultTitles.FirstOrDefault(t => t.Title.Equals(tenPhieuOmi, StringComparison.OrdinalIgnoreCase));
                                ticket.group_name = (match != null && !string.IsNullOrEmpty(match.Group) && match.Group != "Khác (Nhập tay)") ? match.Group : "Khác";
                            }
                            var topGroups = allTickets.GroupBy(t => t.group_name).Select(g => new { Nhóm_Lỗi = g.Key, Số_Lượng = g.Count() }).OrderByDescending(x => x.Số_Lượng).ToList();
                            var tagCounts = new Dictionary<string, int>();
                            foreach (var t in allTickets) { if (t.tags != null) { foreach (var tg in t.tags) { string tid = tg.ValueKind == JsonValueKind.Object && tg.TryGetProperty("id", out var idProp) ? idProp.GetString() : tg.ToString(); if (tagCounts.ContainsKey(tid)) tagCounts[tid]++; else tagCounts[tid] = 1; } } }
                            var topTags = tagCounts.Select(kvp => new { ID = kvp.Key, Tên_Tag = tagList.FirstOrDefault(x => x.id == kvp.Key)?.name ?? kvp.Key, Số_Lượng = kvp.Value }).OrderByDescending(x => x.Số_Lượng).ToList();

                            Form lbForm = new Form() { Text = "Bảng Xếp Hạng Top (Nháy đúp chuột để Lọc ngay)", Width = 450, Height = 550, StartPosition = FormStartPosition.CenterParent, ShowIcon = false };
                            TabControl tabLb = new TabControl() { Dock = DockStyle.Fill }; TabPage pageGroup = new TabPage("🔥 TOP NHÓM LỖI"); TabPage pageTag = new TabPage("🏷️ TOP TAG");
                            DataGridView dgvGroups = new DataGridView() { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, BackgroundColor = Color.White };
                            dgvGroups.DataSource = topGroups;

                            DataGridView dgvTags = new DataGridView() { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false, BackgroundColor = Color.White };
                            dgvTags.DataSource = topTags;
                            // SỬA LỖI CRASH Ở ĐÂY BẰNG CÁCH DÙNG SỰ KIỆN DATABINDINGCOMPLETE
                            dgvTags.DataBindingComplete += (sender_bind, e_bind) => { if (dgvTags.Columns.Contains("ID")) dgvTags.Columns["ID"].Visible = false; };

                            pageGroup.Controls.Add(dgvGroups); pageTag.Controls.Add(dgvTags); tabLb.Controls.Add(pageGroup); tabLb.Controls.Add(pageTag); lbForm.Controls.Add(tabLb);
                            dgvGroups.CellDoubleClick += (sender_dgv, e_dgv) =>
                            {
                                if (e_dgv.RowIndex < 0) return; string selGroup = dgvGroups.Rows[e_dgv.RowIndex].Cells["Nhóm_Lỗi"].Value.ToString();
                                var filtered = allTickets.Where(t => t.group_name == selGroup).ToList(); FormatTicketDisplay(filtered, null);
                                dgvTickets.DataSource = null; dgvTickets.DataSource = filtered; UpdateQuickFilter(filtered); UpdateStatusCount(); lbForm.DialogResult = DialogResult.OK; searchForm.DialogResult = DialogResult.Abort;
                            };
                            dgvTags.CellDoubleClick += (sender_dgv, e_dgv) =>
                            {
                                if (e_dgv.RowIndex < 0) return; string selTagId = dgvTags.Rows[e_dgv.RowIndex].Cells["ID"].Value.ToString();
                                var filtered = allTickets.Where(t => t.tags != null && t.tags.Any(tg => tg.ToString().Contains(selTagId))).ToList(); FormatTicketDisplay(filtered, null);
                                dgvTickets.DataSource = null; dgvTickets.DataSource = filtered; UpdateQuickFilter(filtered); UpdateStatusCount(); lbForm.DialogResult = DialogResult.OK; searchForm.DialogResult = DialogResult.Abort;
                            };
                            Cursor.Current = Cursors.Default; lbForm.ShowDialog();
                        }
                        else MessageBox.Show("Không có phiếu nào trong thời gian này để thống kê!", "Trống", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) { MessageBox.Show("Lỗi Thống kê: " + ex.Message); }
                }
                btnTopStats.Text = "📊 Xem Top Lỗi / Tag"; btnTopStats.Enabled = true; Cursor.Current = Cursors.Default;
            };

            searchForm.Controls.Add(lbl1); searchForm.Controls.Add(txtSearch); searchForm.Controls.Add(lbl2); searchForm.Controls.Add(cbStatus); searchForm.Controls.Add(lbl3); searchForm.Controls.Add(cbTime); searchForm.Controls.Add(dtpStart); searchForm.Controls.Add(lblTo); searchForm.Controls.Add(dtpEnd); searchForm.Controls.Add(lbl4); searchForm.Controls.Add(clbAssignees); searchForm.Controls.Add(lbl5); searchForm.Controls.Add(txtSearchTag); searchForm.Controls.Add(clbTags); searchForm.Controls.Add(btnTopStats); searchForm.Controls.Add(btnDoSearch); searchForm.AcceptButton = btnDoSearch;

            if (searchForm.ShowDialog() == DialogResult.OK)
            {
                if ((dtpEnd.Value.Date - dtpStart.Value.Date).TotalDays > 31) { MessageBox.Show("Khoảng thời gian tối đa là 31 ngày để tránh treo hệ thống!", "Giới hạn", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                if (dtpStart.Value.Date > dtpEnd.Value.Date) { MessageBox.Show("Ngày bắt đầu không được lớn hơn ngày kết thúc!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                string keyword = txtSearch.Text.Trim(); int statusIndex = cbStatus.SelectedIndex; var selectedEmpIds = clbAssignees.CheckedItems.Cast<Employee>().Select(e => e.Id).ToArray();
                List<string> finalTagIdsToSearch = new List<string>();
                foreach (string tId in checkedTagIds) { finalTagIdsToSearch.Add(tId); var childTags = tagList.Where(t => t.parent_id == tId).Select(t => t.id); finalTagIdsToSearch.AddRange(childTags); }
                finalTagIdsToSearch = finalTagIdsToSearch.Distinct().ToList();

                long startMs = new DateTimeOffset(dtpStart.Value.Date).ToUnixTimeMilliseconds(); long endMs = new DateTimeOffset(dtpEnd.Value.Date.AddDays(1).AddTicks(-1)).ToUnixTimeMilliseconds();

                Cursor.Current = Cursors.WaitCursor; dgvTickets.DataSource = null;
                var searchBody = new Dictionary<string, object> { { "search", keyword }, { "ranges", new[] { new { field = "created_date", from = startMs, to = endMs } } }, { "additional_layout", new[] { "object_association" } }, { "has_notify_report", true } };
                if (statusIndex == 1) { searchBody["status_filters"] = new[] { "active_state" }; searchBody["current_status"] = new[] { 0, 1, 2, 3 }; } else if (statusIndex == 2) { searchBody["current_status"] = new[] { 4 }; }
                if (selectedEmpIds.Length > 0) searchBody["assignee_contact_ids"] = selectedEmpIds;

                List<TicketItem> allTickets = new List<TicketItem>(); int currentPage = 1; int pageSize = 1000; bool hasMore = true;

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN); client.Timeout = TimeSpan.FromMinutes(3);
                    try
                    {
                        while (hasMore)
                        {
                            searchBody["page"] = currentPage; searchBody["size"] = pageSize;
                            var searchContent = new StringContent(System.Text.Json.JsonSerializer.Serialize(searchBody), Encoding.UTF8, "application/json");
                            HttpResponseMessage res = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/search?lng=vi&utm_source=web", searchContent);
                            if (res.IsSuccessStatusCode)
                            {
                                OmiResponse omiData = System.Text.Json.JsonSerializer.Deserialize<OmiResponse>(await res.Content.ReadAsStringAsync());
                                if (omiData?.payload?.items != null && omiData.payload.items.Count > 0) { allTickets.AddRange(omiData.payload.items); if (omiData.payload.items.Count < pageSize) hasMore = false; else currentPage++; } else hasMore = false;
                            }
                            else hasMore = false;
                        }

                        if (allTickets.Count > 0)
                        {
                            if (finalTagIdsToSearch != null && finalTagIdsToSearch.Count > 0) { allTickets = allTickets.Where(ticket => ticket.tags != null && ticket.tags.Any(t => finalTagIdsToSearch.Contains(t.GetProperty("id").ToString()))).ToList(); }
                            if (allTickets.Count > 0)
                            {
                                foreach (var ticket in allTickets)
                                {
                                    string tenPhieuOmi = ticket.name != null ? ticket.name.Trim() : ""; var match = defaultTitles.FirstOrDefault(t => t.Title.Equals(tenPhieuOmi, StringComparison.OrdinalIgnoreCase));
                                    ticket.group_name = (match != null && !string.IsNullOrEmpty(match.Group) && match.Group != "Khác (Nhập tay)") ? match.Group : "Khác";
                                }

                                FormatTicketDisplay(allTickets, null);
                                dgvTickets.DataSource = null; dgvTickets.DataSource = allTickets;
                                dgvTickets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                                // ⚡ BƠM DỮ LIỆU VÀO Ô LỌC NHANH Ở ĐÂY
                                UpdateQuickFilter(allTickets);
                                UpdateStatusCount();

                                MessageBox.Show($"Đã kéo chuẩn xác {allTickets.Count} phiếu!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else MessageBox.Show("Không có phiếu nào thỏa mãn điều kiện lọc của bạn!", "Trống", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else MessageBox.Show("Không có phiếu nào được tìm thấy từ API!", "Trống", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) { MessageBox.Show("Lỗi Tìm kiếm: " + ex.Message); }
                }
                Cursor.Current = Cursors.Default;
            }
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
                    string jsonContent = JsonSerializer.Serialize(drafts);
                    string tempFilePath = draftFilePath + ".tmp";
                    File.WriteAllText(tempFilePath, jsonContent);

                    // Ghi đè an toàn, có backup nhẹ
                    if (File.Exists(draftFilePath))
                        File.Replace(tempFilePath, draftFilePath, draftFilePath + ".bak");
                    else
                        File.Move(tempFilePath, draftFilePath);
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
                            var matchedGroup = defaultTitles.FirstOrDefault(t => t.Title.Equals(draft.Title, StringComparison.OrdinalIgnoreCase));
                            if (matchedGroup != null) row.Cells["colGroup"].Value = matchedGroup.Group;

                            row.Cells["colDesc"].Value = draft.Desc;

                            // Nạp cứng 2 cột không đổi
                            //row.Cells["colCategory"].Value = "FO_1_POS";
                            //row.Cells["colSubCategory"].Value = "POS_Lỗi kết nối";
                            // Lấy Phân loại của bản nháp, nếu nháp rỗng thì lấy Mặc định
                            row.Cells["colCategory"].Value = !string.IsNullOrEmpty(draft.CatId) ? draft.CatId : defaultCat;
                            row.Cells["colSubCategory"].Value = !string.IsNullOrEmpty(draft.SubCatId) ? draft.SubCatId : defaultSubCat;
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
            string currentVersion = "4.6"; // ĐỔI SỐ VER ĐỂ CẬP NHẬT
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

        // FIX LỖI ENTER BỊ MẤT CHỮ KHI GÕ TAY
        // FIX LỖI "BỊ MÙ" KHÔNG HIỂN THỊ CHỮ KHI CLICK SANG Ô KHÁC
        // FIX LỖI ENTER BỊ MẤT CHỮ KHI GÕ TAY
        // FIX LỖI "BỊ MÙ" KHÔNG HIỂN THỊ CHỮ KHI CLICK CHỌN
        private void DgvCreateTickets_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Bắt lấy ô đang được gõ/chọn
            if (dgvCreateTickets.EditingControl is ComboBox cb)
            {
                string colName = dgvCreateTickets.Columns[e.ColumnIndex].Name;
                string typedText = cb.Text.Trim();

                // NẾU NGƯỜI DÙNG XÓA TRẮNG -> XÓA LUÔN DỮ LIỆU TRONG Ô VÀ THOÁT
                if (string.IsNullOrEmpty(typedText))
                {
                    dgvCreateTickets.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                    return;
                }

                object finalValue = cb.SelectedValue;

                if (colName == "colTitle")
                {
                    if (isStrictTitle && !defaultTitles.Any(t => t.Title.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show("Không tìm thấy nội dung trong template!\nAdmin đang bật chế độ bắt buộc chọn từ danh sách.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cb.Text = ""; finalValue = null;
                    }
                    else
                    {
                        var match = defaultTitles.FirstOrDefault(t => t.Title.Equals(typedText, StringComparison.OrdinalIgnoreCase));
                        if (match != null) finalValue = match.Title;
                        else
                        {
                            // 🌟 TUYỆT CHIÊU CHỮA BỆNH MÙ CHỮ: Âm thầm nạp chữ lạ vào bộ nhớ
                            defaultTitles.RaiseListChangedEvents = false;
                            defaultTitles.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = typedText });
                            defaultTitles.RaiseListChangedEvents = true;
                            finalValue = typedText;
                        }
                    }
                }
                else if (colName == "colTypeIssue")
                {
                    if (isStrictType && !defaultTypeIssues.Any(t => t.Text.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show("Type Issue bắt buộc phải chọn theo Template có sẵn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cb.Text = ""; finalValue = null;
                    }
                    else
                    {
                        var match = defaultTypeIssues.FirstOrDefault(t => t.Text.Equals(typedText, StringComparison.OrdinalIgnoreCase));
                        if (match != null) finalValue = match.Text;
                        else
                        {
                            // 🌟 TUYỆT CHIÊU CHỮA BỆNH MÙ CHỮ: Âm thầm nạp chữ lạ vào bộ nhớ
                            defaultTypeIssues.RaiseListChangedEvents = false;
                            defaultTypeIssues.Insert(0, new ComboItem { Text = typedText });
                            defaultTypeIssues.RaiseListChangedEvents = true;
                            finalValue = typedText;
                        }
                    }
                }
                else if (colName == "colDesc")
                {
                    if (!defaultTechActions.Any(t => t.Text.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                        defaultTechActions.Insert(0, new ComboItem { Text = typedText });
                    finalValue = typedText;
                }
                else if (colName == "colTag")
                {
                    if (!tagList.Any(t => t.name.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        tagList.Insert(0, new TagItem { id = typedText, name = typedText });
                        var col = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colTag"];
                        col.DataSource = null; col.DataSource = tagList;
                        col.DisplayMember = "name"; col.ValueMember = "id";
                        finalValue = typedText;
                    }
                }
                else if (colName == "colAssignee")
                {
                    if (!employees.Any(em => em.Name.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        employees.Insert(0, new Employee { Id = typedText, Name = typedText });
                        var col = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colAssignee"];
                        col.DataSource = null; col.DataSource = employees;
                        col.DisplayMember = "Name"; col.ValueMember = "Id";
                        finalValue = typedText;
                    }
                }

                // CHỐT HẠ: Ép đúng Value vào ô để WinForms hiểu
                if (finalValue != null)
                {
                    dgvCreateTickets.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = finalValue;
                }
            }
        }

        // 1. CHẶN KHÔNG CHO GÕ CHỮ, CHỈ CHO GÕ SỐ
        private void TimeColumn_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Chặn đứng phím đó lại
            }
        }

        // 2. LOGIC TỰ ÉP KIỂU VÀ CHÈN DẤU HAI CHẤM
        private bool isTimeFormatting = false;
        private void TimeColumn_TextChanged(object sender, EventArgs e)
        {
            if (isTimeFormatting) return;
            TextBox txt = sender as TextBox;
            if (txt == null) return;

            isTimeFormatting = true;
            int cursor = txt.SelectionStart;

            // Lột sạch, chỉ giữ lại số
            string raw = new string(txt.Text.Where(char.IsDigit).ToArray());

            // 🌟 CHẶN GÕ LỐ NGAY TỪ TRONG TRỨNG NƯỚC (Cắt cụt nếu > 4 số)
            if (raw.Length > 4) raw = raw.Substring(0, 4);

            if (raw.Length > 0)
            {
                int h = 0, m = 0;
                string newText = "";

                // Xử lý Giờ
                if (raw.Length >= 1)
                {
                    string hourStr = raw.Length >= 2 ? raw.Substring(0, 2) : raw.Substring(0, 1);
                    h = int.Parse(hourStr);

                    if (h > 23) { h = 23; hourStr = "23"; } // Ép ngay lập tức
                    newText = hourStr;

                    // Gõ đủ 2 số -> Tự động chèn ":"
                    if (raw.Length >= 2)
                    {
                        newText += ":";
                        if (cursor == 2 && !txt.Text.Contains(":")) cursor++;
                    }
                }

                // Xử lý Phút
                if (raw.Length >= 3)
                {
                    string minStr = raw.Length >= 4 ? raw.Substring(2, 2) : raw.Substring(2, 1);
                    m = int.Parse(minStr);

                    if (m > 59) { m = 59; minStr = "59"; } // Ép ngay lập tức
                    newText += minStr;
                }

                if (txt.Text != newText)
                {
                    txt.Text = newText;
                    txt.SelectionStart = cursor <= txt.Text.Length ? cursor : txt.Text.Length;
                }
            }
            else txt.Text = "";

            isTimeFormatting = false;
        }

        // FIX LỖI TỰ NHẢY TÊN THUẬN KHI ENTER
        private void DgvCreateTickets_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgvCreateTickets.CurrentCell == null) return;
            string colName = dgvCreateTickets.CurrentCell.OwningColumn.Name;
            object cellValue = dgvCreateTickets.CurrentCell.Value;

            // ==========================================================
            // PHẦN 1: XỬ LÝ CHO CÁC CỘT DẠNG XỔ XUỐNG (COMBOBOX)
            // ==========================================================
            if (e.Control is ComboBox cb)
            {
                cb.DropDownStyle = ComboBoxStyle.DropDown;
                cb.AutoCompleteMode = AutoCompleteMode.None;

                cb.TextUpdate -= Cb_TextUpdate;
                cb.KeyDown -= Cb_KeyDown;
                cb.Validating -= Cb_Validating;

                // XỬ LÝ ĐẶC BIỆT CHO 2 CỘT PHÂN LOẠI (KHÓA GÕ TAY)
                if (colName == "colCategory" || colName == "colSubCategory")
                {
                    cb.DropDownStyle = ComboBoxStyle.DropDownList;

                    if (colName == "colCategory")
                    {
                        cb.DataSource = categoryList.ToList();
                        cb.DisplayMember = "name"; cb.ValueMember = "name";
                    }
                    else
                    {
                        string catName = dgvCreateTickets.CurrentRow.Cells["colCategory"].Value?.ToString() ?? "";
                        var catInfo = categoryList.FirstOrDefault(c => c.name == catName);
                        if (catInfo != null && catInfo.types != null)
                        {
                            cb.DataSource = catInfo.types.ToList();
                            cb.DisplayMember = "name"; cb.ValueMember = "name";
                        }
                        else cb.DataSource = null;
                    }

                    if (cellValue != null) cb.Text = cellValue.ToString(); else cb.SelectedIndex = -1;
                    return; // Dừng hàm tại đây
                }

                if (colName == "colTitle")
                {
                    string currentVal = cellValue?.ToString() ?? "";
                    var cleanList = defaultTitles.Where(t => t.Group != "Khác (Nhập tay)" || t.Title == currentVal).ToList();
                    cb.DataSource = cleanList; cb.DisplayMember = "Title"; cb.ValueMember = "Title";
                    if (!string.IsNullOrEmpty(currentVal)) cb.Text = currentVal; else cb.SelectedIndex = -1;
                }
                // 🌟 FIX LỖI 1: ÉP Ô TYPE ISSUE VÀ MÔ TẢ PHẢI TRẮNG TINH KHI CHƯA CHỌN GÌ
                else if (colName == "colTypeIssue")
                {
                    string currentVal = cellValue?.ToString() ?? "";
                    cb.DataSource = defaultTypeIssues.ToList(); cb.DisplayMember = "Text"; cb.ValueMember = "Text";
                    if (!string.IsNullOrEmpty(currentVal)) cb.Text = currentVal; else cb.SelectedIndex = -1;
                }
                else if (colName == "colDesc")
                {
                    string currentVal = cellValue?.ToString() ?? "";
                    cb.DataSource = defaultTechActions.ToList(); cb.DisplayMember = "Text"; cb.ValueMember = "Text";
                    if (!string.IsNullOrEmpty(currentVal)) cb.Text = currentVal; else cb.SelectedIndex = -1;
                }
                else if (colName == "colTag")
                {
                    var initialTagList = tagList.Take(100).ToList();
                    if (cellValue != null && cellValue.ToString() != "")
                    {
                        if (!initialTagList.Any(t => t.id == cellValue.ToString()))
                        {
                            var missingTag = tagList.FirstOrDefault(t => t.id == cellValue.ToString());
                            if (missingTag != null) initialTagList.Insert(0, missingTag);
                        }
                    }
                    cb.DataSource = null; cb.DataSource = initialTagList; cb.DisplayMember = "name"; cb.ValueMember = "id";
                    if (cellValue != null && cellValue.ToString() != "") cb.SelectedValue = cellValue; else cb.SelectedIndex = -1;
                }
                else if (colName == "colAssignee")
                {
                    cb.DataSource = employees.ToList(); cb.DisplayMember = "Name"; cb.ValueMember = "Id";
                    if (cellValue != null) cb.SelectedValue = cellValue; else cb.SelectedIndex = -1;
                }

                BeginInvoke(new Action(() =>
                {
                    if (cb.Focused && !string.IsNullOrEmpty(cb.Text))
                    {
                        cb.SelectionLength = 0;
                        cb.SelectionStart = cb.Text.Length;
                    }
                }));

                cb.TextUpdate += Cb_TextUpdate;
                cb.KeyDown += Cb_KeyDown;
                cb.Validating += Cb_Validating;
            }

            // ==========================================================
            // PHẦN 2: XỬ LÝ CHO CÁC CỘT DẠNG GÕ CHỮ (TEXTBOX)
            // ==========================================================
            else if (e.Control is TextBox txt)
            {
                if (colName == "colStartTime" || colName == "colEndTime")
                {
                    // 🌟 NÓ NẰM Ở ĐÂY MỚI ĐÚNG NHÉ!
                    txt.MaxLength = 5;

                    txt.KeyPress -= TimeColumn_KeyPress;
                    txt.TextChanged -= TimeColumn_TextChanged;

                    txt.KeyPress += TimeColumn_KeyPress;
                    txt.TextChanged += TimeColumn_TextChanged;
                }
            }
        }

        // HÀM MỚI: Bắt sự kiện khi bấm Enter hoặc click sang ô khác để lưu lại
        private void Cb_Validating(object sender, CancelEventArgs e)
        {
            if (sender is ComboBox cb && dgvCreateTickets.CurrentCell != null)
            {
                string typedText = cb.Text.Trim();
                if (string.IsNullOrEmpty(typedText)) return;

                string colName = dgvCreateTickets.CurrentCell.OwningColumn.Name;

                if (colName == "colTitle")
                {
                    if (!defaultTitles.Any(t => t.Title.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        // CHỈ CHO PHÉP NHÉT VÀO BỘ NHỚ KHI ADMIN ĐÃ TẮT LUẬT
                        if (!isStrictTitle)
                        {
                            defaultTitles.RaiseListChangedEvents = false;
                            defaultTitles.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = typedText });
                            defaultTitles.RaiseListChangedEvents = true;
                        }
                    }
                }
                else if (colName == "colTypeIssue")
                {
                    if (!defaultTypeIssues.Any(t => t.Text.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        // CHỈ CHO PHÉP NHÉT VÀO BỘ NHỚ KHI ADMIN ĐÃ TẮT LUẬT
                        if (!isStrictType)
                        {
                            defaultTypeIssues.RaiseListChangedEvents = false;
                            defaultTypeIssues.Insert(0, new ComboItem { Text = typedText });
                            defaultTypeIssues.RaiseListChangedEvents = true;
                        }
                    }
                }
                else if (colName == "colTag")
                {
                    if (!tagList.Any(t => t.name.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        tagList.Insert(0, new TagItem { id = typedText, name = typedText });
                    }
                }
                else if (colName == "colAssignee")
                {
                    if (!employees.Any(em => em.Name.Equals(typedText, StringComparison.OrdinalIgnoreCase)))
                    {
                        employees.Insert(0, new Employee { Id = typedText, Name = typedText });
                    }
                }
            }
        }

        private void Cb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;

                // Chốt lưu dữ liệu của ô đang gõ ngay lập tức
                dgvCreateTickets.CommitEdit(DataGridViewDataErrorContexts.Commit);

                // Gửi phím Tab ảo để WinForms tự động chuyển ô sang ngang mượt nhất
                SendKeys.Send("{TAB}");
            }
        }

        public static string ConvertToUnSignStatic(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("\\p{IsCombiningDiacriticalMarks}+");
            string temp = s.Normalize(NormalizationForm.FormD);
            return regex.Replace(temp, String.Empty).Replace('\u0111', 'd').Replace('\u0110', 'D').ToLower();
        }

        private string ConvertToUnSign(string s) => ConvertToUnSignStatic(s);

        // FIX LỖI CRASH KHI TÌM KIẾM KHÔNG CÓ KẾT QUẢ
        // FIX LỖI KHÔNG LƯU KHI ENTER VÀ CRASH KHI TÌM KIẾM
        private void Cb_TextUpdate(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb == null) return;

            string keyword = cb.Text;
            int cursorPos = cb.SelectionStart; // <-- LƯU LẠI VỊ TRÍ CON TRỎ CHUỘT HIỆN TẠI

            // LIÊN TỤC LƯU LẠI CHỮ ĐANG GÕ VÀO BỘ NHỚ
            if (!isUndoing)
            {
                if (cellUndoStack.Count == 0 || cellUndoStack.Peek() != keyword)
                    cellUndoStack.Push(keyword);
            }

            string searchKeyword = ConvertToUnSign(keyword);

            cb.TextUpdate -= Cb_TextUpdate;
            string colName = dgvCreateTickets.CurrentCell.OwningColumn.Name;

            System.Collections.IList dataSource = null;

            if (colName == "colTitle")
            {
                var cleanList = defaultTitles.Where(t => t.Group != "Khác (Nhập tay)").ToList();
                var ds = string.IsNullOrEmpty(searchKeyword) ? cleanList : cleanList.Where(x => x.UnsignedTitle.Contains(searchKeyword)).ToList();

                // -------------------------------------------------------------
                // --- TẠM ẨN: KHÔNG TỰ ĐỘNG TẠO ITEM "KHÁC" KHI ĐANG GÕ TÌM ---
                // if (!string.IsNullOrEmpty(keyword) && !ds.Any(x => x.Title.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                //     ds.Insert(0, new PredefinedTitle { Title = keyword, Group = "Khác (Nhập tay)" });
                // else if (ds.Count == 0) ds.Add(new PredefinedTitle { Title = keyword, Group = "Khác (Nhập tay)" });
                // -------------------------------------------------------------

                // 🌟 FIX LỖI CRASH: Nếu gõ sai không ra kết quả nào, ép vào 1 dòng để ComboBox không bị lỗi
                if (ds.Count == 0)
                {
                    ds.Add(new PredefinedTitle { Title = "Không tìm thấy template...", Group = "Khác" });
                }
                cb.DataSource = ds; cb.DisplayMember = "Title"; cb.ValueMember = "Title"; dataSource = ds;
            }
            else if (colName == "colTypeIssue")
            {
                var ds = string.IsNullOrEmpty(searchKeyword) ? defaultTypeIssues.ToList() : defaultTypeIssues.Where(x => x.UnsignedText.Contains(searchKeyword)).ToList();

                // 🌟 ÉP THEO TEMPLATE: Báo lỗi nếu gõ sai
                if (ds.Count == 0) ds.Add(new ComboItem { Text = "Không tìm thấy Type Issue..." });

                cb.DataSource = ds; cb.DisplayMember = "Text"; cb.ValueMember = "Text"; dataSource = ds;
            }
            else if (colName == "colDesc")
            {
                var ds = string.IsNullOrEmpty(searchKeyword) ? defaultTechActions.ToList() : defaultTechActions.Where(x => x.UnsignedText.Contains(searchKeyword)).ToList();
                if (ds.Count == 0) ds.Add(new ComboItem { Text = keyword }); // Rỗng thì ép vào để không crash
                cb.DataSource = ds; cb.DisplayMember = "Text"; cb.ValueMember = "Text"; dataSource = ds;
            }
            else if (colName == "colTag")
            {
                var ds = string.IsNullOrEmpty(searchKeyword) ? tagList.Take(100).ToList() : tagList.Where(x => x.UnsignedName.Contains(searchKeyword)).Take(100).ToList();
                if (!string.IsNullOrEmpty(keyword) && !ds.Any(x => x.name.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                    ds.Insert(0, new TagItem { name = keyword, id = keyword });
                else if (ds.Count == 0) ds.Add(new TagItem { name = keyword, id = keyword });

                cb.DataSource = ds; cb.DisplayMember = "name"; cb.ValueMember = "id"; dataSource = ds;
            }
            else if (colName == "colAssignee")
            {
                var ds = string.IsNullOrEmpty(searchKeyword) ? employees.ToList() : employees.Where(x => x.UnsignedName.Contains(searchKeyword)).ToList();
                if (!string.IsNullOrEmpty(keyword) && !ds.Any(x => x.Name.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                    ds.Insert(0, new Employee { Name = keyword, Id = keyword });
                else if (ds.Count == 0) ds.Add(new Employee { Name = keyword, Id = keyword });

                cb.DataSource = ds; cb.DisplayMember = "Name"; cb.ValueMember = "Id"; dataSource = ds;
            }

            if (dataSource != null && dataSource.Count > 0)
                cb.DroppedDown = dataSource.Count > 1 || (dataSource.Count == 1 && dataSource[0].ToString() != keyword);
            else cb.DroppedDown = false;

            cb.Text = keyword;
            //cb.SelectionStart = cursorPos; // <-- TRẢ LẠI CHÍNH XÁC VỊ TRÍ CON TRỎ CHUỘT
            // FIX LỖI CRASH: Rào lại vị trí con trỏ, không cho vượt quá độ dài chuỗi
            cb.SelectionStart = Math.Min(cursorPos, cb.Text.Length);
            cb.TextUpdate += Cb_TextUpdate;
            Cursor.Current = Cursors.Default;
        }

        // FIX UPDATE NHÓM NGAY LẬP TỨC: Báo cho bảng biết là ô Tiêu đề vừa đổi chữ
        private void DgvCreateTickets_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvCreateTickets.IsCurrentCellDirty)
            {
                string colName = dgvCreateTickets.CurrentCell?.OwningColumn?.Name;

                if (colName == "colAssignee" || colName == "colTag")
                {
                    dgvCreateTickets.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }

        private void DgvCreateTickets_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvCreateTickets.Columns[e.ColumnIndex].Name;

            if (colName == "colStartTime" || colName == "colEndTime")
            {
                var cell = dgvCreateTickets.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string val = cell.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    string raw = new string(val.Where(char.IsDigit).ToArray());

                    if (raw.Length == 3) raw = "0" + raw;

                    if (raw.Length == 4)
                    {
                        int h = int.Parse(raw.Substring(0, 2));
                        int m = int.Parse(raw.Substring(2, 2));
                        if (h > 23) h = 23;
                        if (m > 59) m = 59;
                        cell.Value = $"{h:D2}:{m:D2}";
                    }
                    else
                    {
                        // CHỈ BÁO LỖI KHI CỐ TÌNH GÕ THIẾU SỐ
                        MessageBox.Show("Vui lòng nhập đủ 4 số!\nVí dụ: 0830 (cho 08:30)", "Chưa hoàn thành", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cell.Value = null;
                        BeginInvoke(new Action(() => { dgvCreateTickets.CurrentCell = cell; dgvCreateTickets.BeginEdit(true); }));
                    }
                }
            }
            SaveDraft();
        }

        private void DgvCreateTickets_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string colName = dgvCreateTickets.Columns[e.ColumnIndex].Name;
            var currentRow = dgvCreateTickets.Rows[e.RowIndex];

            // TỰ ĐỘNG RESET PHÂN LOẠI & LƯU VÀO TRÍ NHỚ
            if (colName == "colCategory" || colName == "colSubCategory")
            {
                if (colName == "colCategory")
                {
                    dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                    // LẤY TỰ ĐỘNG PHÂN LOẠI ĐẦU TIÊN CỦA CHỦ ĐỀ MỚI ĐỂ TRÁNH RỖNG
                    string selectedCat = currentRow.Cells["colCategory"].Value?.ToString();
                    var catInfo = categoryList.FirstOrDefault(c => c.name == selectedCat);
                    if (catInfo != null && catInfo.types != null && catInfo.types.Count > 0)
                    {
                        currentRow.Cells["colSubCategory"].Value = catInfo.types[0].name;
                    }
                    else
                    {
                        currentRow.Cells["colSubCategory"].Value = null;
                    }

                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                }

                string curCat = currentRow.Cells["colCategory"].Value?.ToString();
                string curSub = currentRow.Cells["colSubCategory"].Value?.ToString();

                if (!string.IsNullOrEmpty(curCat) && !string.IsNullOrEmpty(curSub))
                {
                    defaultCat = curCat;
                    defaultSubCat = curSub;
                    File.WriteAllText(catCachePath, defaultCat + "|" + defaultSubCat);
                }
            }

            if (colName == "colTitle")
            {
                string selectedTitle = currentRow.Cells["colTitle"].Value?.ToString() ?? "";
                var matched = defaultTitles.FirstOrDefault(t => t.Title.Equals(selectedTitle, StringComparison.OrdinalIgnoreCase));

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                if (matched != null && matched.Group != "Khác (Nhập tay)")
                {
                    currentRow.Cells["colGroup"].Value = matched.Group;
                    //currentRow.Cells["colTypeIssue"].Value = matched.TypeIssue; // 🌟 Tự động mapping
                    //currentRow.Cells["colDesc"].Value = matched.TechAction;     // 🌟 Tự động mapping
                }
                else
                {
                    currentRow.Cells["colGroup"].Value = "Khác (Nhập tay)";
                }
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                EnsureSufficientRows(5);
            }


            if (colName == "colAssignee")
            {
                var cellValue = currentRow.Cells[colName].Value;
                if (cellValue == null) return;

                string valueToCopy = cellValue.ToString();
                defaultAssigneeId = valueToCopy;

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                try
                {
                    for (int i = e.RowIndex + 1; i < dgvCreateTickets.Rows.Count; i++)
                    {
                        var targetRow = dgvCreateTickets.Rows[i];
                        if (targetRow.IsNewRow) continue;
                        targetRow.Cells["colAssignee"].Value = valueToCopy;
                    }
                }
                finally
                {
                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                }
            }
        }

        private void DgvCreateTickets_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
            e.Cancel = false; // Lệnh tối thượng ép Grid không được khóa chết ô
        }
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

                            // Bơm danh sách vào Cột ngay lập tức
                            this.Invoke(new Action(() =>
                            {
                                if (dgvCreateTickets != null && dgvCreateTickets.Columns.Contains("colCategory"))
                                {
                                    var colCat = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colCategory"];
                                    colCat.DataSource = categoryList.ToList();

                                    var colSub = (DataGridViewComboBoxColumn)dgvCreateTickets.Columns["colSubCategory"];
                                    var allSubs = categoryList.Where(c => c.types != null).SelectMany(c => c.types).ToList();
                                    colSub.DataSource = allSubs;
                                }
                            }));
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
            gridMenu.Items.Add("🗑 Xóa (Delete)", null, (s, e) => SmartDelete());
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
            colTitle.FillWeight = 35;
            colTitle.DisplayStyleForCurrentCellOnly = true; // Chỉ hiện mũi tên ở ô đang chọn
            colTitle.FlatStyle = FlatStyle.Flat; // Xóa viền 3D, làm nút phẳng hiện đại
            dgvCreateTickets.Columns.Add(colTitle);

            DataGridViewComboBoxColumn colTypeIssue = new DataGridViewComboBoxColumn();
            colTypeIssue.HeaderText = "Type Issue";
            colTypeIssue.Name = "colTypeIssue";
            colTypeIssue.DataSource = defaultTypeIssues;
            colTypeIssue.DisplayMember = "Text"; colTypeIssue.ValueMember = "Text";
            colTypeIssue.FillWeight = 30;
            colTypeIssue.DisplayStyleForCurrentCellOnly = true; colTypeIssue.FlatStyle = FlatStyle.Flat;
            dgvCreateTickets.Columns.Add(colTypeIssue);

            DataGridViewComboBoxColumn colDesc = new DataGridViewComboBoxColumn();
            colDesc.HeaderText = "Mô tả (Technical Action)";
            colDesc.Name = "colDesc";
            colDesc.DataSource = defaultTechActions;
            colDesc.DisplayMember = "Text"; colDesc.ValueMember = "Text";
            colDesc.FillWeight = 30;
            colDesc.DisplayStyleForCurrentCellOnly = true; colDesc.FlatStyle = FlatStyle.Flat;
            dgvCreateTickets.Columns.Add(colDesc);
            //dgvCreateTickets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Mô tả chi tiết", Name = "colDesc", FillWeight = 15 });

            // CỘT CHỦ ĐỀ (Đã nạp bộ nhớ chống mù chữ)
            DataGridViewComboBoxColumn colCategory = new DataGridViewComboBoxColumn();
            colCategory.HeaderText = "Chủ đề";
            colCategory.Name = "colCategory";
            colCategory.Visible = false;
            colCategory.DisplayStyleForCurrentCellOnly = true;
            colCategory.FlatStyle = FlatStyle.Flat;
            colCategory.FillWeight = 15;
            colCategory.DataSource = categoryList.ToList();
            colCategory.DisplayMember = "name";
            colCategory.ValueMember = "name";
            dgvCreateTickets.Columns.Add(colCategory);

            // CỘT PHÂN LOẠI (Đã nạp bộ nhớ chống mù chữ)
            DataGridViewComboBoxColumn colSubCategory = new DataGridViewComboBoxColumn();
            colSubCategory.HeaderText = "Phân loại";
            colSubCategory.Name = "colSubCategory";
            colSubCategory.Visible = false;
            colSubCategory.DisplayStyleForCurrentCellOnly = true;
            colSubCategory.FlatStyle = FlatStyle.Flat;
            colSubCategory.FillWeight = 15;
            var allSubCats = categoryList.Where(c => c.types != null).SelectMany(c => c.types).ToList();
            colSubCategory.DataSource = allSubCats;
            colSubCategory.DisplayMember = "name";
            colSubCategory.ValueMember = "name";
            dgvCreateTickets.Columns.Add(colSubCategory);

            DataGridViewComboBoxColumn colTag = new DataGridViewComboBoxColumn();
            colTag.HeaderText = "Tag";
            colTag.Name = "colTag";
            colTag.DisplayMember = "name";
            colTag.ValueMember = "id";
            colTag.FillWeight = 20;
            colTag.DisplayStyleForCurrentCellOnly = true;
            colTag.FlatStyle = FlatStyle.Flat;
            dgvCreateTickets.Columns.Add(colTag);

            DataGridViewTextBoxColumn colStartTime = new DataGridViewTextBoxColumn();
            colStartTime.HeaderText = "Time Nhận";
            colStartTime.Name = "colStartTime";
            colStartTime.FillWeight = 7;
            dgvCreateTickets.Columns.Add(colStartTime);

            DataGridViewTextBoxColumn colEndTime = new DataGridViewTextBoxColumn();
            colEndTime.HeaderText = "Time Hoàn Thành";
            colEndTime.Name = "colEndTime";
            colEndTime.FillWeight = 7;
            dgvCreateTickets.Columns.Add(colEndTime);

            DataGridViewComboBoxColumn colAssignee = new DataGridViewComboBoxColumn();
            colAssignee.HeaderText = "Người xử lý";
            colAssignee.Name = "colAssignee";
            colAssignee.DataSource = employees;
            colAssignee.DisplayMember = "Name";
            colAssignee.ValueMember = "Id";
            colAssignee.FillWeight = 15;
            colAssignee.DisplayStyleForCurrentCellOnly = true;
            colAssignee.FlatStyle = FlatStyle.Flat;
            dgvCreateTickets.Columns.Add(colAssignee);

            DataGridViewTextBoxColumn colResult = new DataGridViewTextBoxColumn();
            colResult.HeaderText = "Kết quả";
            colResult.Name = "colResult";
            colResult.ReadOnly = true;
            colResult.FillWeight = 10;
            dgvCreateTickets.Columns.Add(colResult);
            dgvCreateTickets.DefaultValuesNeeded += DgvCreateTickets_DefaultValuesNeeded;
        }

        private void DgvCreateTickets_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            if (e.Row.Index > 0)
            {
                var prevRow = dgvCreateTickets.Rows[e.Row.Index - 1];
                e.Row.Cells["colCategory"].Value = prevRow.Cells["colCategory"].Value ?? defaultCat;
                e.Row.Cells["colSubCategory"].Value = prevRow.Cells["colSubCategory"].Value ?? defaultSubCat;
            }
            else
            {
                e.Row.Cells["colCategory"].Value = defaultCat;
                e.Row.Cells["colSubCategory"].Value = defaultSubCat;
            }
            if (!string.IsNullOrEmpty(defaultAssigneeId)) e.Row.Cells["colAssignee"].Value = defaultAssigneeId;
        }

        private void DgvCreateTickets_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C) { CopyToClipboard(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.V) { PasteFromClipboard(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.X) { CutToClipboard(); e.Handled = true; }
            else if (e.KeyCode == Keys.Delete)
            {
                SmartDelete(); // Giờ đây phím Delete thường sẽ là Xóa dòng
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
                //DeleteSelectedCells();
            }
            catch { }
        }

        private void DeleteEntireRow()
        {
            if (dgvCreateTickets.SelectedCells.Count > 0)
            {
                var rowIndices = dgvCreateTickets.SelectedCells.Cast<DataGridViewCell>()
                                    .Select(c => c.RowIndex).Distinct().OrderByDescending(r => r).ToList();

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                foreach (int rIndex in rowIndices)
                {
                    if (!dgvCreateTickets.Rows[rIndex].IsNewRow) dgvCreateTickets.Rows.RemoveAt(rIndex);
                }
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                EnsureSufficientRows(100); // Xóa xong tự bù lại cho đủ 30 dòng
                UpdateStatusCount();
            }
        }

        private void ResetEntireTable()
        {
            if (MessageBox.Show("Bạn có chắc muốn xóa SẠCH SÀNH SANH toàn bộ dữ liệu trên bảng để làm lại từ đầu không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                dgvCreateTickets.Rows.Clear();
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                EnsureSufficientRows(100);
                SaveDraft();
                UpdateStatusCount();
            }
        }

        private void SmartDelete()
        {
            try
            {
                if (dgvCreateTickets.SelectedCells.Count > 0)
                {
                    // Lấy danh sách các dòng đang được bôi đen
                    var rowIndices = dgvCreateTickets.SelectedCells.Cast<DataGridViewCell>()
                                        .Select(c => c.RowIndex).Distinct().OrderByDescending(r => r).ToList();

                    dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                    // Thực hiện xóa bốc hơi các dòng đó
                    foreach (int rIndex in rowIndices)
                    {
                        if (!dgvCreateTickets.Rows[rIndex].IsNewRow)
                            dgvCreateTickets.Rows.RemoveAt(rIndex);
                    }

                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;

                    // Tự động bù lại dòng trống ở cuối bảng để luôn giữ mức 30 dòng
                    EnsureSufficientRows(100);
                    UpdateStatusCount();
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
                                // -------------------------------------------------------------
                                // --- TẠM ẨN: KHÔNG CHO PASTE CHỮ LẠ VÀO ---
                                //if (!defaultTitles.Any(t => t.Title.Equals(cellText, StringComparison.OrdinalIgnoreCase)))
                                //{
                                //    defaultTitles.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = cellText });
                                //}
                                //targetCell.Value = cellText;
                                //dgvCreateTickets.Rows[startRow].Cells["colGroup"].Value = "Khác (Nhập tay)";
                                continue;
                            }
                        }
                        else if (colName == "colDesc") targetCell.Value = cellText;
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

        //private void btnAddRow_Click(object sender, EventArgs e)
        //{
        //    if (dgvCreateTickets != null)
        //    {
        //        int oldRowCount = dgvCreateTickets.Rows.Count;
        //        dgvCreateTickets.Rows.Add(5);
        //        dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

        //        try
        //        {
        //            for (int i = oldRowCount - 1; i < dgvCreateTickets.Rows.Count; i++)
        //            {
        //                var targetRow = dgvCreateTickets.Rows[i];
        //                if (targetRow.IsNewRow) continue;

        //                if (!string.IsNullOrEmpty(defaultAssigneeId)) targetRow.Cells["colAssignee"].Value = defaultAssigneeId;

        //                // ÉP CỨNG DỮ LIỆU
        //                //targetRow.Cells["colCategory"].Value = "FO_1_POS";
        //                //targetRow.Cells["colSubCategory"].Value = "POS_Lỗi kết nối";
        //                // ÉP DỮ LIỆU TỪ BỘ NHỚ (Đã bỏ gán cứng "FO_1_POS")
        //                targetRow.Cells["colCategory"].Value = defaultCat;
        //                targetRow.Cells["colSubCategory"].Value = defaultSubCat;
        //            }
        //        }
        //        finally
        //        {
        //            dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
        //        }
        //    }
        //}

        private async void btnCreateTicket_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            int successCount = 0;
            // 🌟 ĐƯA KHAI BÁO LÊN ĐẦU HÀM ĐỂ DÙNG CHUNG CHO TẤT CẢ
            var vtiTag = tagList.FirstOrDefault(t => t.name.Equals("VTI", StringComparison.OrdinalIgnoreCase)) ?? tagList.FirstOrDefault(t => t.name.ToUpper().Contains("VTI"));

            // =========================================================================
            // 🌟 TRẠM KIỂM SOÁT: BẮT BUỘC TYPE ISSUE CHO TAG CON CỦA VTI
            // =========================================================================
            foreach (DataGridViewRow row in dgvCreateTickets.Rows)
            {
                if (row.IsNewRow || (row.Cells["colResult"].Value != null && row.Cells["colResult"].Value.ToString().Contains("☁️"))) continue;

                string tagId = row.Cells["colTag"].Value?.ToString() ?? "";
                string typeIssue = row.Cells["colTypeIssue"].Value?.ToString() ?? "";
                string title = row.Cells["colTitle"].Value?.ToString() ?? "Phiếu trống";

                if (vtiTag != null && !string.IsNullOrEmpty(tagId))
                {
                    var currentTag = tagList.FirstOrDefault(t => t.id == tagId);
                    // CHỈ RÀNG BUỘC NẾU TAG ĐANG CHỌN LÀ TAG CON CỦA VTI (hoặc chính nó)
                    if (currentTag != null && (currentTag.parent_id == vtiTag.id || currentTag.id == vtiTag.id))
                    {
                        if (!string.IsNullOrEmpty(tagId))
                        {
                            var curTag = tagList.FirstOrDefault(t => t.id == tagId);
                            if (curTag != null && (curTag.name.ToUpper().Contains("VTI") || curTag.name.ToUpper().Contains("HLC") || (vtiTag != null && (curTag.parent_id == vtiTag.id || curTag.id == vtiTag.id))))
                            {
                                if (string.IsNullOrEmpty(typeIssue))
                                {
                                    Cursor.Current = Cursors.Default;
                                    MessageBox.Show($"Phiếu '{title}' đang dùng Tag VTI / Highlands.\nBẠN BẮT BUỘC PHẢI CHỌN [Type Issue] TRƯỚC KHI TẠO PHIẾU!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    dgvCreateTickets.CurrentCell = row.Cells["colTypeIssue"];
                                    return; // 🛑 CHẶN ĐỨNG QUÁ TRÌNH TẠO PHIẾU
                                }
                            }
                        }
                    }
                }
            }
            // =========================================================================

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
                        // 🌟 BỔ SUNG: LẤY DỮ LIỆU TYPE ISSUE TRÊN LƯỚI
                        string typeIssue = row.Cells["colTypeIssue"].Value?.ToString() ?? "";
                        string startTime = row.Cells["colStartTime"].Value?.ToString() ?? "";
                        string endTime = row.Cells["colEndTime"].Value?.ToString() ?? "";

                        // 🌟 ĐÓNG GÓI TYPE ISSUE VÀ THỜI GIAN VÀO MÃ HTML GIỐNG HỆT POPUP SỬA PHIẾU
                        string finalDesc = string.IsNullOrEmpty(desc) ? "" : $"<div style=\"font-size: 15px;\">{desc}</div>";
                        if (!string.IsNullOrEmpty(startTime) || !string.IsNullOrEmpty(endTime) || !string.IsNullOrEmpty(typeIssue))
                        {
                            finalDesc += $"<br><br>[TG: {startTime} - {endTime}]<br>[Type: {typeIssue}]";
                        }
                        // 🌟 TRẢ LẠI CÁCH GHI CHỮ ĐỂ OMICRM KHÔNG THỂ XÓA
                        if (!string.IsNullOrEmpty(startTime) || !string.IsNullOrEmpty(endTime))
                        {
                            finalDesc += $"<br><br>[TG: {startTime} - {endTime}]";
                        }

                        // Lấy text đang hiển thị
                        string catName = row.Cells["colCategory"].Value?.ToString() ?? "";
                        string subCatName = row.Cells["colSubCategory"].Value?.ToString() ?? "";
                        string tagId = row.Cells["colTag"].Value?.ToString() ?? "";
                        string empId = row.Cells["colAssignee"].Value?.ToString() ?? "";

                        if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(catName)) continue;

                        row.Cells["colResult"].Value = "⏳ Đang gửi...";

                        // ÂM THẦM DÒ TÌM ID CỦA "FO_1_POS" VÀ "POS_Lỗi kết nối"
                        var catInfo = categoryList.FirstOrDefault(c => c.name == catName);
                        string realCatId = catInfo?._id ?? "";

                        var subCatInfo = catInfo?.types?.FirstOrDefault(t => t.name == subCatName);
                        int typeIndex = subCatInfo != null ? subCatInfo.index : 0;

                        string finalTag = string.IsNullOrEmpty(tagId) ? null : tagId;
                        string finalEmp = string.IsNullOrEmpty(empId) ? null : empId;

                        var createBody = new
                        {
                            name = title,
                            // Gửi finalDesc thay vì desc
                            //description = string.IsNullOrEmpty(finalDesc) ? "" : $"<div style=\"font-size: 15px;\">{finalDesc}</div>",
                            //category_id = realCatId,
                            //description = string.IsNullOrEmpty(desc) ? "" : $"<div style=\"font-size: 15px;\">{desc}</div>",
                            description = finalDesc,
                            category_id = realCatId, // ĐẨY ID LÊN API
                            current_type = typeIndex, // ĐẨY INDEX LÊN API
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
                            // Cần check ruột JSON vì đôi khi OmiCRM lỗi nội bộ vẫn trả HTTP 200
                            if (resStr.Contains("\"id\"") || resStr.Contains("\"_id\"") || resStr.Contains("\"success\":true"))
                            {
                                row.Cells["colResult"].Value = "✅ Thành công";
                                successCount++;

                                // Lấy tên hiển thị để ghi log cho dễ đọc
                                string tagName = tagList.FirstOrDefault(t => t.id == tagId)?.name ?? tagId;
                                string empName = employees.FirstOrDefault(e => e.Id == empId)?.Name ?? empId;

                                // === BẮN DỮ LIỆU LÊN GOOGLE SHEET TẠI ĐÂY ===
                                var currentTag = tagList.FirstOrDefault(t => t.id == tagId);
                                // Nếu bị đổi tên lỡ dính khoảng trắng, fallback sang tìm Tag có chứa chữ "VTI"
                                //var vtiTag = tagList.FirstOrDefault(t => t.name.Equals("VTI", StringComparison.OrdinalIgnoreCase))
                                          //?? tagList.FirstOrDefault(t => t.name.ToUpper().Contains("VTI"));

                                // 🛑 TRẠM KIỂM SOÁT ĐỘC QUYỀN VTI
                                if (currentTag != null && vtiTag != null && (currentTag.parent_id == vtiTag.id || currentTag.id == vtiTag.id))
                                {
                                    string sheetGroup = row.Cells["colGroup"].Value?.ToString() ?? "Khác";
                                    _ = SendToGoogleSheetAsync(title, sheetGroup, desc, tagName, empName);
                                }
                                // ==============================================

                                // GHI LOG THÀNH CÔNG VÀO FILE
                                WriteLog("SUCCESS", $"Tạo phiếu mới thành công: '{title}'", $"Tag: {tagName} | Người xử lý: {empName}");
                            }
                            else
                            {
                                row.Cells["colResult"].Value = "❌ Lỗi ẩn: " + (resStr.Length > 80 ? resStr.Substring(0, 80) : resStr);
                                LogError("Lỗi tạo phiếu ẩn từ API", new Exception(resStr));
                            }
                        }
                        else
                        {
                            if (resStr.StartsWith("<")) row.Cells["colResult"].Value = "❌ Tường lửa chặn / Lỗi Token";
                            else row.Cells["colResult"].Value = "❌ " + (resStr.Length > 80 ? resStr.Substring(0, 80) : resStr);
                            LogError($"Lỗi API ({response.StatusCode})", new Exception(resStr));
                        }
                    }
                    catch (Exception ex) { row.Cells["colResult"].Value = "❌ Lỗi: " + ex.Message; }
                }
            }

            Cursor.Current = Cursors.Default;
            MessageBox.Show($"Xong! Đã tạo thành công {successCount} phiếu.", "Kết quả");

            if (successCount > 0)
            {
                // ÉP BỘ NHỚ QUAY VỀ MẶC ĐỊNH GỐC SAU KHI GỬI THÀNH CÔNG
                defaultCat = "FO_1_POS";
                defaultSubCat = "POS_Lỗi kết nối";
                File.WriteAllText(catCachePath, defaultCat + "|" + defaultSubCat);

                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                for (int i = dgvCreateTickets.Rows.Count - 1; i >= 0; i--)
                {
                    DataGridViewRow currentRow = dgvCreateTickets.Rows[i];
                    if (currentRow.IsNewRow) continue;

                    if (currentRow.Cells["colResult"].Value != null && currentRow.Cells["colResult"].Value.ToString().Contains("✅"))
                    {
                        // Tẩy trắng ô thành công để bạn gõ tiếp
                        currentRow.Cells["colTitle"].Value = null;
                        currentRow.Cells["colGroup"].Value = null;
                        currentRow.Cells["colDesc"].Value = null;
                        currentRow.Cells["colTag"].Value = null;
                        currentRow.Cells["colResult"].Value = null;

                        // 🌟 THÊM 3 DÒNG NÀY ĐỂ TẨY SẠCH CÁC CỘT MỚI
                        currentRow.Cells["colTypeIssue"].Value = null;
                        currentRow.Cells["colStartTime"].Value = null;
                        currentRow.Cells["colEndTime"].Value = null;

                        // Ép cứng về Mặc định (Đoạn này của bạn có sẵn, giữ nguyên)
                        currentRow.Cells["colCategory"].Value = defaultCat;
                        currentRow.Cells["colSubCategory"].Value = defaultSubCat;
                        if (!string.IsNullOrEmpty(defaultAssigneeId)) currentRow.Cells["colAssignee"].Value = defaultAssigneeId;
                    }
                }
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;

                EnsureSufficientRows(100);
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

                                    // GHI LOG THÀNH CÔNG VÀO FILE (Thêm Tên người xử lý)
                                    WriteLog("SUCCESS", $"Đóng phiếu thành công: '{ticket.name}'", $"Người xử lý: {selectedEmp.Name} | ID Phiếu: {realId}");
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
                    if (dgvTickets.Columns.Contains("TenTag")) { dgvTickets.Columns["TenTag"].HeaderText = "Tags"; dgvTickets.Columns["TenTag"].DisplayIndex = 4; }
                    if (dgvTickets.Columns.Contains("tags")) dgvTickets.Columns["tags"].Visible = false;
                    if (dgvTickets.Columns.Contains("assignee_contact_ids")) dgvTickets.Columns["assignee_contact_ids"].Visible = false;

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
        // =========================================================================
        // TÍNH NĂNG MỚI: NHÁY ĐÚP ĐỂ CHỈNH SỬA PHIẾU (NHIỀU TAG, NHIỀU NGƯỜI, CÓ LÝ DO)
        // =========================================================================
        // =========================================================================
        // POP-UP CHỈNH SỬA PHIẾU TÍCH HỢP TÌM KIẾM CHO TAG VÀ NGƯỜI XỬ LÝ
        // =========================================================================
        // =========================================================================
        // POP-UP CHỈNH SỬA PHIẾU (ĐÃ THÊM Ô MÔ TẢ CHI TIẾT)
        // =========================================================================
        private void DgvTickets_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var ticket = dgvTickets.Rows[e.RowIndex].DataBoundItem as TicketItem;
            if (ticket == null) return;

            Form popup = new Form()
            {
                Width = 550,
                Height = 750, // 🌟 Kéo dài Form ra thêm 100px để chứa ô Mô tả
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Chỉnh sửa phiếu: " + ticket.unique_id,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false
            };

            int currentY = 15;

            // 1. Tên phiếu
            Label lblName = new Label() { Left = 20, Top = currentY, Text = "Tên phiếu:", AutoSize = true };
            currentY += 20;
            TextBox txtName = new TextBox() { Left = 20, Top = currentY, Width = 490, Text = ticket.name };
            currentY += 35;

            // 🌟 1.5 TYPE ISSUE (ĐÃ VỨT BỎ DATASOURCE, CHO PHÉP XÓA TRẮNG KHÔNG TỰ PHỤC HỒI)
            Label lblType = new Label() { Left = 20, Top = currentY, Text = "Type Issue:", AutoSize = true };
            currentY += 20;
            ComboBox cbType = new ComboBox() { Left = 20, Top = currentY, Width = 490, DropDownStyle = ComboBoxStyle.DropDown };

            // Chuyển dữ liệu sang mảng chuỗi thường và nạp thẳng vào Items
            var allTypeIssues = defaultTypeIssues.Select(t => t.Text).ToArray();
            cbType.Items.AddRange(allTypeIssues);

            cbType.AutoCompleteMode = AutoCompleteMode.None;

            cbType.TextUpdate += (s_cb, ev_cb) =>
            {
                string keyword = cbType.Text;
                int cursorPos = cbType.SelectionStart;

                cbType.Items.Clear(); // Xóa danh sách cũ đi

                if (string.IsNullOrEmpty(keyword.Trim()))
                {
                    // Nếu xóa sạch chữ -> Nạp lại list gốc nhưng KHÔNG xòe menu, giữ ô trắng tinh
                    cbType.Items.AddRange(allTypeIssues);
                    cbType.DroppedDown = false;
                }
                else
                {
                    // Nếu có gõ chữ -> Lọc tìm kiếm và xòe menu
                    string search = ConvertToUnSignStatic(keyword);
                    var ds = allTypeIssues.Where(x => ConvertToUnSignStatic(x).Contains(search)).ToArray();

                    if (ds.Length > 0) cbType.Items.AddRange(ds);
                    else cbType.Items.Add("Không tìm thấy Type Issue...");

                    cbType.DroppedDown = true;
                }

                cbType.Text = keyword;
                cbType.SelectionStart = Math.Min(cursorPos, cbType.Text.Length);
                Cursor.Current = Cursors.Default;
            };

            // 🌟 NẠP DỮ LIỆU TỪ OMiCRM LÊN (Nếu rỗng thì nó hiện trắng tinh)
            cbType.Text = ticket.TypeIssue ?? "";
            currentY += 35;

            // 🌟 2. MÔ TẢ CHI TIẾT (Ô nhập liệu nhiều dòng)
            Label lblDesc = new Label() { Left = 20, Top = currentY, Text = "Mô tả chi tiết:", AutoSize = true };
            currentY += 20;
            TextBox txtDesc = new TextBox() { Left = 20, Top = currentY, Width = 490, Height = 70, Multiline = true, ScrollBars = ScrollBars.Vertical, Text = ticket.MoTa };
            currentY += 85;

            // 🌟 2.5 THỜI GIAN XỬ LÝ (Áp dụng "phép thuật" gõ thông minh tự động ép số)
            Label lblTime = new Label() { Left = 20, Top = currentY, Text = "Thời gian (Nhận - Xong):", AutoSize = true };
            currentY += 20;
            TextBox txtStart = new TextBox() { Left = 20, Top = currentY, Width = 80, Text = ticket.ThoiGianNhan, MaxLength = 5 };
            Label lblDash = new Label() { Left = 110, Top = currentY + 3, Text = "-", AutoSize = true };
            TextBox txtEnd = new TextBox() { Left = 130, Top = currentY, Width = 80, Text = ticket.ThoiGianXong, MaxLength = 5 };

            txtStart.TextChanged += TimeColumn_TextChanged;
            txtEnd.TextChanged += TimeColumn_TextChanged;
            currentY += 35;

            // 3. Trạng thái & Lý do
            Label lblStatus = new Label() { Left = 20, Top = currentY, Text = "Trạng thái:", AutoSize = true };
            Label lblReason = new Label() { Left = 200, Top = currentY, Text = "Lý do (Nếu chuyển trạng thái):", AutoSize = true };
            currentY += 20;

            ComboBox cbStatus = new ComboBox() { Left = 20, Top = currentY, Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };

            // 🌟 FIX CRASH: Bơm thẳng dữ liệu vào Items để WinForms nhận ngay lập tức
            cbStatus.Items.AddRange(new string[] {
                "Mới tiếp nhận (0)", // Tương đương Index = 0
                "Đang xử lý (1)",    // Tương đương Index = 1
                "Đợi phản hồi (2)",  // Tương đương Index = 2
                "Leo thang (3)",     // Tương đương Index = 3
                "Đã Đóng (4)"        // Tương đương Index = 4
            });

            // Do mã trạng thái OmiCRM (0-4) khớp y chang số thứ tự Index (0-4), ta gán trực tiếp luôn!
            cbStatus.SelectedIndex = (ticket.current_status >= 0 && ticket.current_status <= 4) ? ticket.current_status : 0;

            TextBox txtReason = new TextBox() { Left = 200, Top = currentY, Width = 310, Text = "" };
            currentY += 35;

            // 4. Đổi Người xử lý (CÓ THANH TÌM KIẾM)
            Label lblAssignee = new Label() { Left = 20, Top = currentY, Text = "Người xử lý:", AutoSize = true };
            TextBox txtSearchAssignee = new TextBox() { Left = 250, Top = currentY - 3, Width = 260, Text = "🔍 Tìm người..." };
            txtSearchAssignee.GotFocus += (s, ev) => { if (txtSearchAssignee.Text == "🔍 Tìm người...") txtSearchAssignee.Text = ""; };
            currentY += 25;

            CheckedListBox clbAssignees = new CheckedListBox() { Left = 20, Top = currentY, Width = 490, Height = 100, CheckOnClick = true };
            HashSet<string> checkedAssignees = new HashSet<string>();
            if (ticket.assignee_contact_ids != null)
            {
                foreach (var a in ticket.assignee_contact_ids)
                {
                    string aId = a.ValueKind == System.Text.Json.JsonValueKind.Object && a.TryGetProperty("id", out var idProp) ? idProp.GetString() : a.ToString();
                    checkedAssignees.Add(aId);
                }
            }

            Action loadAssignees = () =>
            {
                clbAssignees.Items.Clear();
                string kw = ConvertToUnSignStatic(txtSearchAssignee.Text == "🔍 Tìm người..." ? "" : txtSearchAssignee.Text);
                var filtered = string.IsNullOrEmpty(kw) ? employees : employees.Where(x => x.UnsignedName.Contains(kw)).ToList();
                foreach (var emp in filtered) clbAssignees.Items.Add(emp, checkedAssignees.Contains(emp.Id));
            };
            loadAssignees();

            clbAssignees.ItemCheck += (s, ev) =>
            {
                var emp = clbAssignees.Items[ev.Index] as Employee;
                if (ev.NewValue == CheckState.Checked) checkedAssignees.Add(emp.Id); else checkedAssignees.Remove(emp.Id);
            };
            txtSearchAssignee.TextChanged += (s, ev) => loadAssignees();
            currentY += 110;

            // 5. Đổi Tag (CÓ THANH TÌM KIẾM CHỐNG LAG)
            Label lblTag = new Label() { Left = 20, Top = currentY, Text = "Tags:", AutoSize = true };
            TextBox txtSearchTag = new TextBox() { Left = 250, Top = currentY - 3, Width = 260, Text = "🔍 Tìm Tag..." };
            txtSearchTag.GotFocus += (s, ev) => { if (txtSearchTag.Text == "🔍 Tìm Tag...") txtSearchTag.Text = ""; };
            currentY += 25;

            CheckedListBox clbTags = new CheckedListBox() { Left = 20, Top = currentY, Width = 490, Height = 150, CheckOnClick = true };
            HashSet<string> checkedTags = new HashSet<string>();
            if (ticket.tags != null)
            {
                foreach (var t in ticket.tags)
                {
                    string tId = t.ValueKind == System.Text.Json.JsonValueKind.Object && t.TryGetProperty("id", out var idProp) ? idProp.GetString() : t.ToString();
                    checkedTags.Add(tId);
                }
            }

            Action loadTags = () =>
            {
                clbTags.Items.Clear();
                string kw = ConvertToUnSignStatic(txtSearchTag.Text == "🔍 Tìm Tag..." ? "" : txtSearchTag.Text);
                var filtered = string.IsNullOrEmpty(kw) ? tagList.Take(150).ToList() : tagList.Where(x => x.UnsignedName.Contains(kw)).Take(150).ToList();
                var checkedItemsInTagList = tagList.Where(t => checkedTags.Contains(t.id)).ToList();
                var finalDisplayList = checkedItemsInTagList.Union(filtered).Distinct().ToList();

                foreach (var tag in finalDisplayList) clbTags.Items.Add(tag, checkedTags.Contains(tag.id));
            };
            loadTags();

            clbTags.ItemCheck += (s, ev) =>
            {
                var tag = clbTags.Items[ev.Index] as TagItem;
                if (ev.NewValue == CheckState.Checked) checkedTags.Add(tag.id); else checkedTags.Remove(tag.id);
            };
            txtSearchTag.TextChanged += (s, ev) => loadTags();
            currentY += 160;

            Button btnSave = new Button() { Text = "Lưu Cập Nhật", Left = 215, Top = currentY, Width = 120, Height = 35, BackColor = Color.LightGreen, FlatStyle = FlatStyle.Flat };

            popup.Controls.Add(lblType);
            popup.Controls.Add(cbType);
            popup.Controls.Add(lblTime); popup.Controls.Add(txtStart); popup.Controls.Add(lblDash); popup.Controls.Add(txtEnd);
            popup.Controls.Add(lblName); popup.Controls.Add(txtName);
            popup.Controls.Add(lblDesc); popup.Controls.Add(txtDesc); // 🌟 Gắn ô Mô tả vào giao diện
            popup.Controls.Add(lblStatus); popup.Controls.Add(cbStatus);
            popup.Controls.Add(lblReason); popup.Controls.Add(txtReason);
            popup.Controls.Add(lblAssignee); popup.Controls.Add(txtSearchAssignee); popup.Controls.Add(clbAssignees);
            popup.Controls.Add(lblTag); popup.Controls.Add(txtSearchTag); popup.Controls.Add(clbTags);
            popup.Controls.Add(btnSave);
            popup.AcceptButton = btnSave;

            btnSave.Click += async (s, ev) =>
            {
                string newName = txtName.Text.Trim();
                string newDesc = txtDesc.Text.Trim();
                string newType = cbType.Text.Trim();
                string newStart = txtStart.Text.Trim(); // Lấy Thời gian nhận
                string newEnd = txtEnd.Text.Trim();     // Lấy Thời gian xong
                int newStatus = cbStatus.SelectedIndex;
                if (string.IsNullOrEmpty(newName)) return;

                // 🌟 FIX 2.1: KIỂM TRA GÕ BẬY (Chữ bắt buộc phải có trong template)
                if (!string.IsNullOrEmpty(newType) && !defaultTypeIssues.Any(t => t.Text.Equals(newType, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show("Type Issue không hợp lệ!\nVui lòng chọn đúng giá trị có trong danh sách.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cbType.Focus();
                    return; // 🛑 Chặn đứng, không cho lưu
                }

                // 🌟 FIX 2.2: RÀNG BUỘC VTI VÀ HIGHLANDS KHÔNG ĐƯỢC TRỐNG
                bool isVtiTicket = false;
                var vtiTag = tagList.FirstOrDefault(t => t.name.Equals("VTI", StringComparison.OrdinalIgnoreCase)) ?? tagList.FirstOrDefault(t => t.name.ToUpper().Contains("VTI"));

                foreach (var tid in checkedTags)
                {
                    var curTag = tagList.FirstOrDefault(t => t.id == tid);
                    if (curTag != null)
                    {
                        // Kiểm tra trọn ổ: Tên chứa VTI, chứa HLC, hoặc Parent là VTI
                        if (curTag.name.ToUpper().Contains("VTI") || curTag.name.ToUpper().Contains("HLC") ||
                            (vtiTag != null && (curTag.parent_id == vtiTag.id || curTag.id == vtiTag.id)))
                        {
                            isVtiTicket = true; break;
                        }
                    }
                }

                if (isVtiTicket && string.IsNullOrEmpty(newType))
                {
                    MessageBox.Show("Phiếu này đang dùng Tag VTI / Highlands.\nBẠN BẮT BUỘC PHẢI CHỌN [Type Issue]!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cbType.Focus();
                    return; // 🛑 Chặn đứng, không cho lưu
                }

                btnSave.Text = "Đang lưu..."; btnSave.Enabled = false;
                // ... Các dòng code gửi API bên dưới giữ nguyên ...
                Cursor.Current = Cursors.WaitCursor;

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);
                    string realId = !string.IsNullOrEmpty(ticket._id) ? ticket._id : ticket.id;
                    bool successName = true, successStatus = true;

                    var selectedAssignees = checkedAssignees.ToArray();
                    var selectedTags = checkedTags.ToArray();
                    string realCatId = !string.IsNullOrEmpty(ticket.category_id) ? ticket.category_id : "65a4b914c08b463df237dcec";

                    // 🌟 ĐÓNG GÓI THÊM TYPE ISSUE VÀ THỜI GIAN VÀO MÔ TẢ ĐỂ OMICRM GIỮ GIÙM
                    string finalDescHTML = string.IsNullOrEmpty(newDesc) ? "" : $"<div style=\"font-size: 15px;\">{newDesc}</div>";
                    if (!string.IsNullOrEmpty(newStart) || !string.IsNullOrEmpty(newEnd) || !string.IsNullOrEmpty(newType))
                    {
                        finalDescHTML += $"<br><br>[TG: {newStart} - {newEnd}]<br>[Type: {newType}]";
                    }

                    var updateContent = new StringContent(System.Text.Json.JsonSerializer.Serialize(new
                    {
                        id = realId,
                        name = newName,
                        description = finalDescHTML, // Gửi HTML có bọc Type/Thời gian
                        category_id = realCatId,
                        current_type = 0,
                        assignee_contact_ids = selectedAssignees,
                        tags = selectedTags,
                        source = "crud"
                    }), Encoding.UTF8, "application/json");

                    var resUpdate = await client.PutAsync("https://ticket-v2-stg.omicrm.com/ticket/update?lng=vi&utm_source=web", updateContent);
                    successName = resUpdate.IsSuccessStatusCode;

                    if (newStatus != ticket.current_status || !string.IsNullOrEmpty(txtReason.Text.Trim()))
                    {
                        string reasonMsg = string.IsNullOrEmpty(txtReason.Text.Trim()) ? "Cập nhật qua Tool" : txtReason.Text.Trim();
                        var statusContent = new StringContent(System.Text.Json.JsonSerializer.Serialize(new { id = realId, status = newStatus, message = reasonMsg }), Encoding.UTF8, "application/json");
                        var resStatus = await client.PutAsync("https://ticket-v2-stg.omicrm.com/ticket/status/update?lng=vi&utm_source=web", statusContent);
                        successStatus = resStatus.IsSuccessStatusCode;
                    }

                    Cursor.Current = Cursors.Default;

                    if (successName && successStatus)
                    {
                        // Cập nhật ngay vào màn hình để Xuất Excel có số liệu
                        ticket.name = newName;
                        ticket.description = finalDescHTML; // Gán lại để MoTa tự cập nhật
                        ticket.TypeIssue = newType;
                        ticket.ThoiGianNhan = newStart;
                        ticket.ThoiGianXong = newEnd;
                        ticket.current_status = newStatus;

                        MessageBox.Show("Cập nhật thành công!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        popup.DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        MessageBox.Show("Có lỗi xảy ra khi cập nhật lên OmiCRM.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        btnSave.Text = "Lưu Cập Nhật"; btnSave.Enabled = true;
                    }
                }
            };
            if (popup.ShowDialog() == DialogResult.OK) dgvTickets.Refresh();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            SearchGlobalTicket();
        }

        private async Task SyncRemoteConfigAsync()
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string data = await client.GetStringAsync(configUrl);
                    var lines = data.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var line in lines)
                    {
                        var cols = line.Split('\t');
                        if (cols.Length < 2) continue;
                        string key = cols[0].Trim().ToUpper();
                        string val = cols[1].Trim().ToUpper();

                        if (key == "STRICT_TITLE") isStrictTitle = (val == "TRUE");
                        if (key == "STRICT_TYPE") isStrictType = (val == "TRUE");
                        if (key == "ALLOW_DELETE") allowDelete = (val == "TRUE");
                        if (key == "MAINTENANCE") isMaintenance = (val == "ON" || val == "TRUE");
                    }
                }

                // 🌟 KIỂM TRA BẢO TRÌ NGAY KHI TẢI XONG
                bool isAdmin = adminEmails.Any(e => loginEmailCached.ToLower().Contains(e.ToLower()));
                if (isMaintenance && !isAdmin)
                {
                    this.Invoke(new Action(() => {
                        MessageBox.Show("Hệ thống đang được Admin bảo trì / cập nhật tính năng mới.\nVui lòng quay lại sau!", "Bảo Trì", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        Application.Exit();
                    }));
                }
            }
            catch { isStrictTitle = true; isStrictType = true; }
        }
        private async void btnAdmin_Click(object sender, EventArgs e)
        {
            Form adminPopup = new Form()
            {
                Width = 420,
                Height = 320,
                Text = "BẢNG ĐIỀU KHIỂN ADMIN",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false
            };

            Label lbl1 = new Label() { Left = 20, Top = 20, Text = "Bật / Tắt các tính năng hệ thống:", Font = new Font("Segoe UI", 9, FontStyle.Bold), AutoSize = true };

            CheckBox chkTitle = new CheckBox() { Left = 30, Top = 50, Text = "Bắt buộc chọn Tên Phiếu từ danh sách", Checked = isStrictTitle, Width = 350 };
            CheckBox chkType = new CheckBox() { Left = 30, Top = 80, Text = "Bắt buộc Type Issue & Ràng buộc Tag VTI", Checked = isStrictType, Width = 350 };
            CheckBox chkDel = new CheckBox() { Left = 30, Top = 110, Text = "Cho phép nhân viên XÓA PHIẾU trên OmiCRM", Checked = allowDelete, Width = 350, ForeColor = Color.DarkRed };
            CheckBox chkMaint = new CheckBox() { Left = 30, Top = 140, Text = "Chế độ Bảo trì (Khóa Tool đối với nhân viên)", Checked = isMaintenance, Width = 350 };

            Button btnSave = new Button()
            {
                Text = "LƯU LẠI & ĐỒNG BỘ",
                Left = 110,
                Top = 200,
                Width = 180,
                Height = 40,
                BackColor = Color.Orange,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            adminPopup.Controls.AddRange(new Control[] { lbl1, chkTitle, chkType, chkDel, chkMaint, btnSave });

            btnSave.Click += async (s_save, e_save) => {
                btnSave.Enabled = false; btnSave.Text = "Đang truyền lệnh...";
                string scriptUrl = "https://script.google.com/macros/s/AKfycbyLj9sFGUe-NWWdDP3pfaNqsn6OnRo-2ZOtSnv4AUFV7Q944yUn1mdmzhaSS3Ls4Kkz/exec";

                try
                {
                    using (HttpClient client = new HttpClient())
                    {
                        await client.PostAsync(scriptUrl, new StringContent(JsonSerializer.Serialize(new { key = "STRICT_TITLE", value = chkTitle.Checked ? "TRUE" : "FALSE" }), Encoding.UTF8, "application/json"));
                        await client.PostAsync(scriptUrl, new StringContent(JsonSerializer.Serialize(new { key = "STRICT_TYPE", value = chkType.Checked ? "TRUE" : "FALSE" }), Encoding.UTF8, "application/json"));
                        await client.PostAsync(scriptUrl, new StringContent(JsonSerializer.Serialize(new { key = "ALLOW_DELETE", value = chkDel.Checked ? "TRUE" : "FALSE" }), Encoding.UTF8, "application/json"));
                        await client.PostAsync(scriptUrl, new StringContent(JsonSerializer.Serialize(new { key = "MAINTENANCE", value = chkMaint.Checked ? "ON" : "OFF" }), Encoding.UTF8, "application/json"));
                    }

                    isStrictTitle = chkTitle.Checked; isStrictType = chkType.Checked; allowDelete = chkDel.Checked; isMaintenance = chkMaint.Checked;
                    MessageBox.Show("Đã phát lệnh thành công! 15 máy anh em sẽ tự động áp dụng luật mới.", "Admin Master", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    adminPopup.Close();
                }
                catch (Exception ex) { MessageBox.Show("Lỗi kết nối Server: " + ex.Message); }
                finally { btnSave.Enabled = true; btnSave.Text = "LƯU LẠI & ĐỒNG BỘ"; }
            };

            adminPopup.ShowDialog();
        }
        // =========================================================================
        // 1. TỰ ĐỘNG VIỆT HÓA CỘT MỌI NƠI (KHÔNG BAO GIỜ BỊ LÒI TIẾNG ANH NỮA)
        // =========================================================================
        private void DgvTickets_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            UpdateStatusCount();

            // =================================================================================
            // 🌟 SẮP XẾP LẠI VỊ TRÍ CỘT: ID -> NHÓM -> TÊN PHIẾU -> NGÀY TẠO -> ...
            // =================================================================================
            if (dgvTickets.Columns.Contains("unique_id")) { dgvTickets.Columns["unique_id"].HeaderText = "ID"; dgvTickets.Columns["unique_id"].DisplayIndex = 0; dgvTickets.Columns["unique_id"].Width = 60; }
            if (dgvTickets.Columns.Contains("group_name")) { dgvTickets.Columns["group_name"].HeaderText = "Nhóm"; dgvTickets.Columns["group_name"].DisplayIndex = 1; dgvTickets.Columns["group_name"].Width = 100; }
            if (dgvTickets.Columns.Contains("name")) { dgvTickets.Columns["name"].HeaderText = "Tên Phiếu"; dgvTickets.Columns["name"].DisplayIndex = 2; }
            if (dgvTickets.Columns.Contains("TypeIssue")) { dgvTickets.Columns["TypeIssue"].HeaderText = "Type Issue"; dgvTickets.Columns["TypeIssue"].DisplayIndex = 3; }
            if (dgvTickets.Columns.Contains("MoTa")) { dgvTickets.Columns["MoTa"].HeaderText = "Mô tả chi tiết"; dgvTickets.Columns["MoTa"].DisplayIndex = 4; }
            if (dgvTickets.Columns.Contains("ThoiGianNhan")) { dgvTickets.Columns["ThoiGianNhan"].HeaderText = "TG Nhận"; dgvTickets.Columns["ThoiGianNhan"].DisplayIndex = 5; dgvTickets.Columns["ThoiGianNhan"].Width = 80; }
            if (dgvTickets.Columns.Contains("ThoiGianXong")) { dgvTickets.Columns["ThoiGianXong"].HeaderText = "TG Xong"; dgvTickets.Columns["ThoiGianXong"].DisplayIndex = 6; dgvTickets.Columns["ThoiGianXong"].Width = 80; }
            if (dgvTickets.Columns.Contains("NgayTao")) { dgvTickets.Columns["NgayTao"].HeaderText = "Ngày Tạo"; dgvTickets.Columns["NgayTao"].DisplayIndex = 7; dgvTickets.Columns["NgayTao"].Width = 120; }
            if (dgvTickets.Columns.Contains("TrangThaiHienThi")) { dgvTickets.Columns["TrangThaiHienThi"].HeaderText = "Trạng Thái"; dgvTickets.Columns["TrangThaiHienThi"].DisplayIndex = 8; dgvTickets.Columns["TrangThaiHienThi"].Width = 100; }
            if (dgvTickets.Columns.Contains("NguoiNhan")) { dgvTickets.Columns["NguoiNhan"].HeaderText = "Người Xử Lý"; dgvTickets.Columns["NguoiNhan"].DisplayIndex = 9; }
            if (dgvTickets.Columns.Contains("TenTag")) { dgvTickets.Columns["TenTag"].HeaderText = "Danh sách Tag"; dgvTickets.Columns["TenTag"].DisplayIndex = 10; }

            // Ẩn toàn bộ các cột dữ liệu rác của hệ thống
            if (dgvTickets.Columns.Contains("tags")) dgvTickets.Columns["tags"].Visible = false;
            if (dgvTickets.Columns.Contains("assignee_contact_ids")) dgvTickets.Columns["assignee_contact_ids"].Visible = false;
            if (dgvTickets.Columns.Contains("description")) dgvTickets.Columns["description"].Visible = false; // Ẩn cột chứa mã HTML rác
            if (dgvTickets.Columns.Contains("_id")) dgvTickets.Columns["_id"].Visible = false;
            if (dgvTickets.Columns.Contains("id")) dgvTickets.Columns["id"].Visible = false;
            if (dgvTickets.Columns.Contains("current_status")) dgvTickets.Columns["current_status"].Visible = false;
            if (dgvTickets.Columns.Contains("category_id")) dgvTickets.Columns["category_id"].Visible = false;
            if (dgvTickets.Columns.Contains("created_date")) dgvTickets.Columns["created_date"].Visible = false;

            dgvTickets.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        // =========================================================================
        // 2. BẤM CHUỘT VÀO TIÊU ĐỀ CỘT ĐỂ SẮP XẾP (A-Z / Z-A) THÔNG MINH
        // =========================================================================
        private bool sortAscending = true;
        private void DgvTickets_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgvTickets.DataSource is List<TicketItem> list && list.Count > 0)
            {
                string colName = dgvTickets.Columns[e.ColumnIndex].Name;

                // Xếp ID theo dạng số (1 -> 9), xếp Tên theo bảng chữ cái (A -> Z), xếp Trạng thái...
                if (colName == "unique_id")
                    list = sortAscending ? list.OrderBy(x => int.TryParse(x.unique_id, out int id) ? id : 0).ToList() : list.OrderByDescending(x => int.TryParse(x.unique_id, out int id) ? id : 0).ToList();
                else if (colName == "name")
                    list = sortAscending ? list.OrderBy(x => x.name).ToList() : list.OrderByDescending(x => x.name).ToList();
                else if (colName == "TrangThaiHienThi")
                    list = sortAscending ? list.OrderBy(x => x.current_status).ToList() : list.OrderByDescending(x => x.current_status).ToList();
                else if (colName == "NguoiNhan")
                    list = sortAscending ? list.OrderBy(x => x.NguoiNhan).ToList() : list.OrderByDescending(x => x.NguoiNhan).ToList();
                else if (colName == "TenTag")
                    list = sortAscending ? list.OrderBy(x => x.TenTag).ToList() : list.OrderByDescending(x => x.TenTag).ToList();

                sortAscending = !sortAscending; // Đảo chiều cho lần click sau
                dgvTickets.DataSource = list;   // Nạp lại bảng
            }
        }
        // =========================================================================
        // TÍNH NĂNG NÚT: ĐỔI MÀU, ẨN/HIỆN PHÂN LOẠI
        // =========================================================================
        private void btnTheme_Click(object sender, EventArgs e)
        {
            Form popup = new Form() { Width = 350, Height = 300, Text = "Phối màu Giao diện", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };

            Label lbl1 = new Label() { Text = "Màu Tiêu đề cột:", Left = 20, Top = 20, AutoSize = true };
            Button btnHeader = new Button() { Left = 150, Top = 15, Width = 150, BackColor = dgvCreateTickets.ColumnHeadersDefaultCellStyle.BackColor.IsEmpty ? SystemColors.Control : dgvCreateTickets.ColumnHeadersDefaultCellStyle.BackColor, FlatStyle = FlatStyle.Flat };

            Label lbl2 = new Label() { Text = "Màu dòng lẻ (Zebra):", Left = 20, Top = 60, AutoSize = true };
            Button btnOdd = new Button() { Left = 150, Top = 55, Width = 150, BackColor = dgvCreateTickets.AlternatingRowsDefaultCellStyle.BackColor.IsEmpty ? Color.WhiteSmoke : dgvCreateTickets.AlternatingRowsDefaultCellStyle.BackColor, FlatStyle = FlatStyle.Flat };

            Label lbl3 = new Label() { Text = "Màu dòng chẵn:", Left = 20, Top = 100, AutoSize = true };
            Button btnEven = new Button() { Left = 150, Top = 95, Width = 150, BackColor = dgvCreateTickets.RowsDefaultCellStyle.BackColor.IsEmpty ? Color.White : dgvCreateTickets.RowsDefaultCellStyle.BackColor, FlatStyle = FlatStyle.Flat };

            Label lbl4 = new Label() { Text = "Màu khi bôi đen:", Left = 20, Top = 140, AutoSize = true };
            Button btnSelect = new Button() { Left = 150, Top = 135, Width = 150, BackColor = dgvCreateTickets.DefaultCellStyle.SelectionBackColor.IsEmpty ? SystemColors.Highlight : dgvCreateTickets.DefaultCellStyle.SelectionBackColor, FlatStyle = FlatStyle.Flat };

            EventHandler pickColor = (s, ev) =>
            {
                Button btn = (Button)s;
                ColorDialog cd = new ColorDialog() { Color = btn.BackColor };
                if (cd.ShowDialog() == DialogResult.OK) btn.BackColor = cd.Color;
            };

            btnHeader.Click += pickColor; btnOdd.Click += pickColor; btnEven.Click += pickColor; btnSelect.Click += pickColor;

            // Nút Mặc định (Mới)
            Button btnDefault = new Button() { Text = "Mặc định", Left = 40, Top = 200, Width = 110, Height = 35, BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            btnDefault.Click += (s, ev) =>
            {
                string themePath = Path.Combine(Application.StartupPath, "theme.txt");
                if (File.Exists(themePath)) File.Delete(themePath); // Xóa bộ nhớ màu

                ResetThemeToDefault(); // Đưa UI về gốc
                MessageBox.Show("Đã khôi phục giao diện về mặc định!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                popup.DialogResult = DialogResult.Abort; // Đóng popup không kích hoạt lưu
            };

            // Nút Lưu cấu hình (Đã đổi sang màu xanh lam đẹp mắt hơn)
            Button btnSave = new Button() { Text = "Lưu cấu hình", Left = 180, Top = 200, Width = 110, Height = 35, BackColor = Color.DodgerBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.OK };

            popup.Controls.Add(lbl1); popup.Controls.Add(btnHeader);
            popup.Controls.Add(lbl2); popup.Controls.Add(btnOdd);
            popup.Controls.Add(lbl3); popup.Controls.Add(btnEven);
            popup.Controls.Add(lbl4); popup.Controls.Add(btnSelect);
            popup.Controls.Add(btnDefault); popup.Controls.Add(btnSave);
            popup.AcceptButton = btnSave;

            if (popup.ShowDialog() == DialogResult.OK)
            {
                ApplyTheme(btnHeader.BackColor, btnOdd.BackColor, btnEven.BackColor, btnSelect.BackColor);
                // Lưu 4 màu vào file để lần sau mở máy vẫn nhớ
                string config = $"{ColorTranslator.ToHtml(btnHeader.BackColor)}|{ColorTranslator.ToHtml(btnOdd.BackColor)}|{ColorTranslator.ToHtml(btnEven.BackColor)}|{ColorTranslator.ToHtml(btnSelect.BackColor)}";
                File.WriteAllText(Path.Combine(Application.StartupPath, "theme.txt"), config);
            }
        }

        private void ApplyTheme(Color header, Color oddRow, Color evenRow, Color selected)
        {
            // Áp dụng cho Bảng Tạo Phiếu
            dgvCreateTickets.EnableHeadersVisualStyles = false;
            dgvCreateTickets.ColumnHeadersDefaultCellStyle.BackColor = header;
            dgvCreateTickets.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvCreateTickets.AlternatingRowsDefaultCellStyle.BackColor = oddRow;
            dgvCreateTickets.RowsDefaultCellStyle.BackColor = evenRow;
            dgvCreateTickets.DefaultCellStyle.SelectionBackColor = selected;

            // Áp dụng cho Bảng Đóng Issue
            if (dgvTickets != null)
            {
                dgvTickets.EnableHeadersVisualStyles = false;
                dgvTickets.ColumnHeadersDefaultCellStyle.BackColor = header;
                dgvTickets.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgvTickets.AlternatingRowsDefaultCellStyle.BackColor = oddRow;
                dgvTickets.RowsDefaultCellStyle.BackColor = evenRow;
                dgvTickets.DefaultCellStyle.SelectionBackColor = selected;
            }
        }

        private void ResetThemeToDefault()
        {
            Action<DataGridView> resetGrid = (grid) =>
            {
                if (grid == null) return;

                // Trả lại toàn bộ giao diện nguyên thủy của Windows Forms
                grid.EnableHeadersVisualStyles = true;
                grid.ColumnHeadersDefaultCellStyle.BackColor = Color.Empty;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.Empty;
                grid.AlternatingRowsDefaultCellStyle.BackColor = Color.Empty;
                grid.RowsDefaultCellStyle.BackColor = Color.Empty;
                grid.DefaultCellStyle.SelectionBackColor = Color.Empty;
            };

            resetGrid(dgvCreateTickets);
            resetGrid(dgvTickets);
        }

        private void btnToggleCategory_Click(object sender, EventArgs e)
        {
            bool isCurrentlyHidden = !dgvCreateTickets.Columns["colCategory"].Visible;
            dgvCreateTickets.Columns["colCategory"].Visible = isCurrentlyHidden;
            dgvCreateTickets.Columns["colSubCategory"].Visible = isCurrentlyHidden;

            Button btn = sender as Button;
            if (btn != null) btn.Text = isCurrentlyHidden ? "👁️ Ẩn Phân Loại" : "👁️ Hiện Phân Loại";
        }
        // =========================================================================
        // ĐỘNG CƠ TỰ ĐỘNG BƠM DÒNG: LUÔN GIỮ BẢNG ĐẦY ẮP MÀ KHÔNG CẦN BẤM NÚT
        // =========================================================================
        private void EnsureSufficientRows(int minRows)
        {
            if (dgvCreateTickets == null) return;

            int emptyCount = 0;
            foreach (DataGridViewRow row in dgvCreateTickets.Rows)
            {
                if (row.IsNewRow) continue;
                if (string.IsNullOrEmpty(row.Cells["colTitle"].Value?.ToString())) emptyCount++;
            }

            if (emptyCount < minRows)
            {
                int rowsToAdd = minRows - emptyCount;
                dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;
                try
                {
                    for (int i = 0; i < rowsToAdd; i++)
                    {
                        int idx = dgvCreateTickets.Rows.Add();
                        var targetRow = dgvCreateTickets.Rows[idx];

                        if (!string.IsNullOrEmpty(defaultAssigneeId)) targetRow.Cells["colAssignee"].Value = defaultAssigneeId;

                        // Tự động điền dữ liệu từ bộ nhớ thay vì điền cứng
                        targetRow.Cells["colCategory"].Value = defaultCat;
                        targetRow.Cells["colSubCategory"].Value = defaultSubCat;
                    }
                }
                finally
                {
                    dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                }
            }
            UpdateStatusCount();
        }
        // =========================================================================
        // TÍNH NĂNG MỚI: XÓA PHIẾU TRỰC TIẾP TRÊN OMICRM (TỪ TAB ĐÓNG ISSUE)
        // =========================================================================
        private async void DeleteSelectedTicketsAsync()
        {
            // 🌟 KIỂM TRA QUYỀN XÓA (Admin thì luôn được xóa)
            bool isAdmin = adminEmails.Any(e => loginEmailCached.ToLower().Contains(e.ToLower()));
            if (!allowDelete && !isAdmin)
            {
                MessageBox.Show("Tính năng XÓA PHIẾU đang bị Admin tạm khóa!", "Từ chối quyền", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            // 1. Lấy danh sách các dòng đang được chọn
            var selectedRows = dgvTickets.SelectedCells.Cast<DataGridViewCell>()
                                .Select(c => c.OwningRow).Distinct().ToList();

            if (selectedRows.Count == 0) return;

            DialogResult confirm = MessageBox.Show(
                $"Bạn có chắc chắn muốn XÓA VĨNH VIỄN {selectedRows.Count} phiếu này khỏi hệ thống OmiCRM không?\n(Hành động này KHÔNG THỂ HOÀN TÁC!)",
                "Cảnh báo Xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Error);

            if (confirm == DialogResult.No) return;

            Cursor.Current = Cursors.WaitCursor;
            List<string> idsToDelete = new List<string>();

            // Lấy ID thật của phiếu
            foreach (var row in selectedRows)
            {
                var ticket = row.DataBoundItem as TicketItem;
                if (ticket != null)
                {
                    string realId = !string.IsNullOrEmpty(ticket._id) ? ticket._id : ticket.id;
                    if (!string.IsNullOrEmpty(realId)) idsToDelete.Add(realId);
                }
            }

            if (idsToDelete.Count > 0)
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

                    var payload = new { ids = idsToDelete };
                    string jsonPayload = JsonSerializer.Serialize(payload);

                    var request = new HttpRequestMessage
                    {
                        Method = HttpMethod.Delete,
                        RequestUri = new Uri("https://ticket-v2-stg.omicrm.com/ticket/delete-many?lng=vi&utm_source=web"),
                        Content = new StringContent(jsonPayload, Encoding.UTF8, "application/json")
                    };

                    try
                    {
                        HttpResponseMessage response = await client.SendAsync(request);
                        string resStr = await response.Content.ReadAsStringAsync();

                        if (response.IsSuccessStatusCode && (resStr.Contains("\"payload\":true") || resStr.Contains("\"payload\": true") || resStr.Contains("\"status_code\":9999") || resStr.Contains("\"status_code\": 9999")))
                        {
                            // Cập nhật lại giao diện: Xóa các dòng bốc hơi khỏi bảng
                            var dataSource = dgvTickets.DataSource as List<TicketItem>;
                            if (dataSource != null)
                            {
                                dataSource.RemoveAll(t => idsToDelete.Contains(!string.IsNullOrEmpty(t._id) ? t._id : t.id));
                                dgvTickets.DataSource = null;
                                dgvTickets.DataSource = dataSource; // Gán lại để bảng tự Refresh
                            }

                            WriteLog("SUCCESS", "Xóa phiếu trên hệ thống", $"Đã xóa {idsToDelete.Count} phiếu. Danh sách ID: {string.Join(", ", idsToDelete)}");
                            MessageBox.Show($"Tuyệt vời! Đã xóa thành công {idsToDelete.Count} phiếu khỏi hệ thống OmiCRM.", "Xóa Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            UpdateStatusCount();
                        }
                        else
                        {
                            MessageBox.Show($"Xóa thất bại!\nOmiCRM trả về: {resStr}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi kết nối mạng: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            Cursor.Current = Cursors.Default;
        }
        // =========================================================================
        // TÍNH NĂNG MỚI: GỬI BACKUP DỮ LIỆU LÊN GOOGLE SHEET CÁ NHÂN
        // =========================================================================
        private async Task SendToGoogleSheetAsync(string title, string group, string desc, string tag, string assignee)
        {
            // Link thật của bạn:
            string scriptUrl = "https://script.google.com/macros/s/AKfycbwlgLm22OkjLgA3lbaRfrE9ug8oX2xYcbvtygsnznD3f6mMivYcA37AKGmGlAb2LbE/exec";

            using (HttpClient client = new HttpClient())
            {
                // Đóng gói 5 biến gửi lên (Sheet tự động thêm Thời gian là 6)
                var body = new
                {
                    title = title,
                    group = group,
                    desc = desc,
                    tag = tag,
                    assignee = assignee
                };

                var content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
                try { await client.PostAsync(scriptUrl, content); } catch { /* Im lặng bỏ qua lỗi mạng */ }
            }
        }
        // =========================================================================
        // TÍNH NĂNG MỚI: XUẤT DỮ LIỆU RA FILE EXCEL (CSV)
        // =========================================================================
        private void ExportTicketsToExcel()
        {
            if (dgvTickets.Rows.Count == 0 || dgvTickets.DataSource == null)
            {
                MessageBox.Show("Không có dữ liệu nào trên bảng để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ==============================================================================
            // 🌟 DANH SÁCH CỘT XUẤT EXCEL (Viết dạng Tuple siêu gọn gàng, dễ nhìn)
            // ==============================================================================
            var availableColumns = new[]
            {
                // -- CÁC CỘT CHÍNH (Sẽ được tick mặc định) --
                (Prop: "unique_id",        Name: "ID Phiếu",                 IsChecked: true),
                (Prop: "group_name",       Name: "Nhóm",                     IsChecked: true),
                (Prop: "name",             Name: "Tên Phiếu",                IsChecked: true),
                (Prop: "MoTa",             Name: "Mô tả chi tiết",           IsChecked: true),
                (Prop: "NgayTao",          Name: "Ngày Tạo",                 IsChecked: true),
                (Prop: "NguoiNhan",        Name: "Người Xử Lý",              IsChecked: true),
                (Prop: "TrangThaiHienThi", Name: "Trạng Thái",               IsChecked: true),
                (Prop: "TenTag",           Name: "Danh sách Tag",            IsChecked: true),
                (Prop: "TypeIssue",        Name: "Type Issue",               IsChecked: true),
                (Prop: "ThoiGianNhan",     Name: "Thời Gian Nhận",           IsChecked: true),
                (Prop: "ThoiGianXong",     Name: "Thời Gian Hoàn Thành",     IsChecked: true),

                // -- CÁC CỘT PHỤ HỆ THỐNG (Không tick sẵn) --
                (Prop: "id",               Name: "Mã ID hệ thống (Ẩn)",      IsChecked: false),
                (Prop: "current_status",   Name: "Mã Trạng thái gốc (Số)",   IsChecked: false),
                (Prop: "category_id",      Name: "Mã Chủ đề (Category ID)",  IsChecked: false),
                (Prop: "created_date",     Name: "Mã thời gian Unix gốc",    IsChecked: false)
            };

            Form popup = new Form() { Width = 400, Height = 580, Text = "Tuỳ chọn Xuất Báo Cáo", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };

            // 🌟 1. THÊM COMBOBOX CHỌN LOẠI BÁO CÁO
            Label lblType = new Label() { Text = "1. Chọn loại báo cáo:", Left = 20, Top = 20, AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            ComboBox cbReportType = new ComboBox() { Left = 20, Top = 45, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList };
            cbReportType.Items.Add("Báo cáo Dữ liệu Chi tiết (Raw Data)");
            cbReportType.Items.Add("Báo cáo Thống kê Tổng hợp (Dashboard)");
            cbReportType.SelectedIndex = 0;

            // 🌟 2. CHỌN CỘT DỮ LIỆU
            Label lbl = new Label() { Text = "2. Tick chọn các trường (cột) muốn xuất:", Left = 20, Top = 90, AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            CheckedListBox clbColumns = new CheckedListBox() { Left = 20, Top = 115, Width = 340, Height = 320, CheckOnClick = true };

            // Nạp danh sách cột vào giao diện
            foreach (var col in availableColumns)
            {
                var item = new ColumnExportItem { PropertyName = col.Prop, HeaderText = col.Name };
                clbColumns.Items.Add(item, col.IsChecked);
            }

            Button btnExport = new Button() { Text = "Xuất file Excel", Left = 120, Top = 460, Width = 140, Height = 40, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.OK, Font = new Font("Segoe UI", 10, FontStyle.Bold) };

            popup.Controls.Add(lblType); popup.Controls.Add(cbReportType);
            popup.Controls.Add(lbl); popup.Controls.Add(clbColumns);
            popup.Controls.Add(btnExport);
            popup.AcceptButton = btnExport;

            if (popup.ShowDialog() == DialogResult.OK)
            {
                var selectedColumns = clbColumns.CheckedItems.Cast<ColumnExportItem>().ToList();
                if (selectedColumns.Count == 0) { MessageBox.Show("Bạn chưa chọn trường nào để xuất!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                bool isDashboard = cbReportType.SelectedIndex == 1;
                string fileNamePrefix = isDashboard ? "Dashboard_ThongKe" : "BaoCao_ChiTiet";

                SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", Title = "Lưu Báo cáo OmiCRM", FileName = $"{fileNamePrefix}_{DateTime.Now:ddMMyyyy_HHmm}.xlsx" };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    try
                    {
                        using (var workbook = new ClosedXML.Excel.XLWorkbook())
                        {
                            var dataSource = dgvTickets.DataSource as List<TicketItem>;

                            // ========================================================
                            // SHEET 1: DỮ LIỆU CHI TIẾT (Luôn luôn tạo dựa trên checkbox)
                            // ========================================================
                            var worksheet = workbook.Worksheets.Add("DuLieuChiTiet");

                            for (int i = 0; i < selectedColumns.Count; i++)
                            {
                                var cell = worksheet.Cell(1, i + 1);
                                cell.Value = selectedColumns[i].HeaderText;
                                cell.Style.Font.Bold = true;
                                cell.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.CornflowerBlue;
                                cell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                                cell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            }

                            int rowIdx = 2;
                            foreach (var ticket in dataSource)
                            {
                                for (int i = 0; i < selectedColumns.Count; i++)
                                {
                                    string propName = selectedColumns[i].PropertyName;
                                    var prop = typeof(TicketItem).GetProperty(propName);
                                    if (prop != null)
                                    {
                                        worksheet.Cell(rowIdx, i + 1).Value = prop.GetValue(ticket, null)?.ToString() ?? "";
                                        worksheet.Cell(rowIdx, i + 1).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                    }
                                }
                                rowIdx++;
                            }
                            worksheet.Columns().AdjustToContents();

                            // ========================================================
                            // SHEET 2: DASHBOARD THỐNG KÊ (ĐÃ NÂNG CẤP CHUẨN DOANH NGHIỆP)
                            // ========================================================
                            if (isDashboard)
                            {
                                var wsDash = workbook.Worksheets.Add("Dashboard Thong Ke");

                                // 🌟 ĐÃ FIX LỖI CRASH: XÓA LỆNH wsDash.Outline.SummaryVBelow = false;
                                // Thư viện sẽ tự động dùng mặc định, nút +/- vẫn hoạt động bình thường!

                                // Trích xuất dữ liệu
                                var reportData = dataSource.Select(t => new
                                {
                                    Date = t.NgayTao.Length >= 10 ? t.NgayTao.Substring(0, 10) : t.NgayTao,
                                    Store = string.IsNullOrEmpty(t.TenTag) ? "Khác" : t.TenTag.Split(',')[0].Trim(),
                                    TypeEn = dictTypeEng.ContainsKey(t.TypeIssue ?? "") ? dictTypeEng[t.TypeIssue] : t.TypeIssue
                                }).ToList();

                                var validData = reportData.Where(x => !string.IsNullOrEmpty(x.TypeEn)).ToList();

                                // =========================================================================
                                // HÀM VẼ BẢNG I, II, III (CÓ KÈM RÚT GỌN +/-)
                                // =========================================================================
                                int DrawTable(int startRow, int startCol, string title, IEnumerable<dynamic> groupedData, string header1, string header2)
                                {
                                    wsDash.Cell(startRow, startCol).Value = title; wsDash.Cell(startRow, startCol).Style.Font.Bold = true; wsDash.Cell(startRow, startCol).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;
                                    wsDash.Cell(startRow + 1, startCol).Value = header1; wsDash.Cell(startRow + 1, startCol + 1).Value = header2; wsDash.Cell(startRow + 1, startCol + 2).Value = "QUANTITY";
                                    var range = wsDash.Range(startRow + 1, startCol, startRow + 1, startCol + 2);
                                    range.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.CornflowerBlue; range.Style.Font.FontColor = ClosedXML.Excel.XLColor.White; range.Style.Font.Bold = true;

                                    int cr = startRow + 2; int total = 0;
                                    foreach (var group in groupedData)
                                    {
                                        // Dòng Nhóm Cha (Màu xám)
                                        wsDash.Cell(cr, startCol).Value = group.Key; wsDash.Cell(cr, startCol).Style.Font.Bold = true;
                                        wsDash.Cell(cr, startCol + 2).Value = group.Total; wsDash.Cell(cr, startCol + 2).Style.Font.Bold = true;
                                        wsDash.Range(cr, startCol, cr, startCol + 2).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                                        total += group.Total; cr++;

                                        int subStart = cr; // Bắt đầu đếm dòng con
                                        foreach (var sub in group.SubItems)
                                        {
                                            wsDash.Cell(cr, startCol + 1).Value = sub.SubKey;
                                            wsDash.Cell(cr, startCol + 2).Value = sub.Count;
                                            cr++;
                                        }
                                        int subEnd = cr - 1; // Kết thúc đếm dòng con

                                        // 🌟 TẠO NÚT +/- ĐÓNG MỞ CHO CÁC DÒNG CON VÀ THU GỌN MẶC ĐỊNH
                                        if (subEnd >= subStart)
                                        {
                                            wsDash.Rows(subStart, subEnd).Group();
                                            wsDash.Rows(subStart, subEnd).Collapse();
                                        }
                                    }

                                    wsDash.Cell(cr, startCol).Value = "Grand Total"; wsDash.Cell(cr, startCol).Style.Font.Bold = true;
                                    wsDash.Cell(cr, startCol + 2).Value = total; wsDash.Cell(cr, startCol + 2).Style.Font.Bold = true;

                                    // Kẻ khung
                                    wsDash.Range(startRow + 1, startCol, cr, startCol + 2).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                    wsDash.Range(startRow + 1, startCol, cr, startCol + 2).Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                    return cr + 2;
                                }

                                var dataI = validData.GroupBy(x => x.Store).Select(g => new { Key = g.Key, Total = g.Count(), SubItems = g.GroupBy(x => x.TypeEn).Select(sg => new { SubKey = sg.Key, Count = sg.Count() }) }).OrderByDescending(x => x.Total).ToList();
                                DrawTable(1, 1, "I. ISSUE BY STORE BY TYPE", dataI, "LOCATION", "TYPE ISSUE");

                                var dataII = validData.GroupBy(x => x.TypeEn).Select(g => new { Key = g.Key, Total = g.Count(), SubItems = g.GroupBy(x => x.Store).Select(sg => new { SubKey = sg.Key, Count = sg.Count() }) }).OrderByDescending(x => x.Total).ToList();
                                DrawTable(1, 5, "II. ISSUE BY TYPE BY STORE", dataII, "TYPE ISSUE", "LOCATION");

                                var dataIII = validData.GroupBy(x => x.Date).Select(g => new { Key = g.Key, Total = g.Count(), SubItems = g.GroupBy(x => x.TypeEn).Select(sg => new { SubKey = sg.Key, Count = sg.Count() }) }).OrderBy(x => x.Key).ToList();
                                DrawTable(1, 9, "III. ISSUE BY DAY BY TYPE", dataIII, "DATE", "TYPE ISSUE");

                                // =========================================================================
                                // BẢNG IV. TOP ISSUE (Kèm cột điền tay)
                                // =========================================================================
                                int colIV = 13; // Cột M
                                wsDash.Cell(1, colIV).Value = "IV. TOP ISSUE"; wsDash.Cell(1, colIV).Style.Font.Bold = true; wsDash.Cell(1, colIV).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;

                                string[] headersIV = { "STT", "TYPE ISSUE", "QUANTITY", "CAUSES", "SOLUTIONS", "PLAN" };
                                for (int i = 0; i < headersIV.Length; i++)
                                {
                                    wsDash.Cell(2, colIV + i).Value = headersIV[i];
                                    wsDash.Cell(2, colIV + i).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.CornflowerBlue;
                                    wsDash.Cell(2, colIV + i).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                                    wsDash.Cell(2, colIV + i).Style.Font.Bold = true;
                                }

                                var topIssues = validData.GroupBy(x => x.TypeEn).Select(g => new { Type = g.Key, Count = g.Count() }).OrderByDescending(x => x.Count).ToList();
                                int rIV = 3; int sttIV = 1;
                                foreach (var issue in topIssues)
                                {
                                    wsDash.Cell(rIV, colIV).Value = sttIV++;
                                    wsDash.Cell(rIV, colIV + 1).Value = issue.Type;
                                    wsDash.Cell(rIV, colIV + 2).Value = issue.Count;
                                    // Chừa trống Cột Causes, Solutions, Plan
                                    rIV++;
                                }
                                wsDash.Range(2, colIV, rIV - 1, colIV + 5).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                wsDash.Range(2, colIV, rIV - 1, colIV + 5).Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                wsDash.Range(2, colIV, rIV - 1, colIV + 5).Style.Alignment.WrapText = true; // Bật WordWrap cho ghi chú

                                // =========================================================================
                                // BẢNG V. TOP 20 STORE ISSUE
                                // =========================================================================
                                int rV = rIV + 2;
                                wsDash.Cell(rV, colIV).Value = "V. TOP 20 STORE ISSUE"; wsDash.Cell(rV, colIV).Style.Font.Bold = true; wsDash.Cell(rV, colIV).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;

                                string[] headersV = { "STT", "LOCATION", "QUANTITY" };
                                for (int i = 0; i < headersV.Length; i++)
                                {
                                    wsDash.Cell(rV + 1, colIV + i).Value = headersV[i];
                                    wsDash.Cell(rV + 1, colIV + i).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.CornflowerBlue;
                                    wsDash.Cell(rV + 1, colIV + i).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                                    wsDash.Cell(rV + 1, colIV + i).Style.Font.Bold = true;
                                }

                                var topStores = validData.GroupBy(x => x.Store).Select(g => new { Store = g.Key, Count = g.Count() }).OrderByDescending(x => x.Count).Take(20).ToList();
                                int crV = rV + 2; int sttV = 1;
                                foreach (var store in topStores)
                                {
                                    wsDash.Cell(crV, colIV).Value = sttV++;
                                    wsDash.Cell(crV, colIV + 1).Value = store.Store;
                                    wsDash.Cell(crV, colIV + 2).Value = store.Count;
                                    crV++;
                                }
                                wsDash.Range(rV + 1, colIV, crV - 1, colIV + 2).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                wsDash.Range(rV + 1, colIV, crV - 1, colIV + 2).Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                                // Chỉnh Auto-size cột cho đẹp
                                wsDash.Columns().AdjustToContents();
                                wsDash.Column(colIV + 3).Width = 35; // Làm rộng cột Causes
                                wsDash.Column(colIV + 4).Width = 35; // Làm rộng cột Solutions
                                wsDash.Column(colIV + 5).Width = 15; // Làm rộng cột Plan
                            }

                            workbook.SaveAs(sfd.FileName);
                        }
                        Cursor.Current = Cursors.Default;
                        string msg = isDashboard ? "Đã xuất thành công Báo cáo Thống kê Dashboard (kèm sheet Dữ liệu chi tiết)!" : "Đã xuất thành công Báo cáo Dữ liệu Chi tiết!";
                        MessageBox.Show(msg, "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Process.Start("explorer.exe", $"/select,\"{sfd.FileName}\"");
                    }
                    catch (Exception ex) { Cursor.Current = Cursors.Default; MessageBox.Show("Lỗi lưu file: \n" + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        //private void btnRefreshOmi_Click(object sender, EventArgs e)
        //{
        //    if (webViewOmicall != null && webViewOmicall.CoreWebView2 != null)
        //    {
        //        // Lệnh "thần thánh" để ép Omicall tải lại trang, hốt tin nhắn bị miss
        //        webViewOmicall.CoreWebView2.Reload();
        //    }
        //}
    }



    public class PredefinedTitle
    {
        public string Group { get; set; }
        public string Title { get; set; }
        public string TypeIssue { get; set; }  // 🌟 MỚI THÊM
        public string TechAction { get; set; } // 🌟 MỚI THÊM

        public override string ToString() { return Title; }

        [System.Text.Json.Serialization.JsonIgnore]
        public string UnsignedTitle => _us ?? (_us = Form1.ConvertToUnSignStatic(Title));
        private string _us;
    }

    public class ComboItem
    {
        public string Text { get; set; }
        public override string ToString() { return Text; }

        [System.Text.Json.Serialization.JsonIgnore]
        public string UnsignedText => _ut ?? (_ut = Form1.ConvertToUnSignStatic(Text));
        private string _ut;
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

        [System.Text.Json.Serialization.JsonIgnore]
        public string UnsignedName => _un ?? (_un = Form1.ConvertToUnSignStatic(Name));
        private string _un;
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
        [System.ComponentModel.Browsable(false)]
        [System.Text.Json.Serialization.JsonPropertyName("_id")]
        public string _id { get; set; }

        [System.ComponentModel.Browsable(false)]
        public string id { get; set; }

        public string unique_id { get; set; }
        public string name { get; set; }

        // ==================================
        [System.ComponentModel.Browsable(false)]
        public int current_status { get; set; }
        [System.ComponentModel.Browsable(false)]
        public long created_date { get; set; } // OmiCRM trả về mã Unix Milliseconds

        public string NgayTao
        {
            get
            {
                if (created_date <= 0) return "";
                // Dịch mã thời gian Unix sang Ngày giờ Việt Nam (dd/MM/yyyy HH:mm)
                return DateTimeOffset.FromUnixTimeMilliseconds(created_date).LocalDateTime.ToString("dd/MM/yyyy HH:mm");
            }
        }

        [System.ComponentModel.Browsable(false)]
        public string description { get; set; } // Trường gốc từ API

        // 🌟 THÊM ĐOẠN NÀY: Tự động tẩy sạch thẻ HTML để hiện chữ thuần túy
        // 🌟 TẨY SẠCH 100% HTML MÃ HÓA TỪ OMICRM
        // 🌟 TẨY SẠCH 100% HTML VÀ THẺ TÀNG HÌNH TỪ OMICRM
        public string MoTa
        {
            get
            {
                if (string.IsNullOrEmpty(description)) return "";

                // 1. Giải mã HTML
                string decodedHTML = System.Net.WebUtility.HtmlDecode(description);

                // 2. LỌC SẠCH RÁC THỜI GIAN VÀ TYPE CŨ (Chống lặp)
                decodedHTML = System.Text.RegularExpressions.Regex.Replace(decodedHTML, @"\[TG:\s*.*?\s*-\s*.*?\]", "");
                decodedHTML = System.Text.RegularExpressions.Regex.Replace(decodedHTML, @"\[Type:\s*.*?\]", "");

                // 3. Xóa sạch mọi thẻ HTML (<...>)
                string cleanText = System.Text.RegularExpressions.Regex.Replace(decodedHTML, "<.*?>", " ");

                return System.Text.RegularExpressions.Regex.Replace(cleanText, @"\s+", " ").Trim();
            }
        }

        // ==========================================================

        [System.ComponentModel.Browsable(false)] // TỰ ĐỘNG ẨN CỘT NÀY TRÊN BẢNG
        public string category_id { get; set; }
        public string group_name { get; set; }

        [System.ComponentModel.Browsable(false)]
        public List<JsonElement> tags { get; set; }

        [System.ComponentModel.Browsable(false)]
        public List<JsonElement> assignee_contact_ids { get; set; }

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
        public string TenTag { get; set; }  // Chứa Tên Tiếng Việt của Tag
        public string TypeIssue { get; set; }
        public string ThoiGianNhan { get; set; }
        public string ThoiGianXong { get; set; }
        public string NguoiNhan { get; set; }
        
    }

    public class ColumnExportItem
    {
        public string PropertyName { get; set; }
        public string HeaderText { get; set; }
        public override string ToString() { return HeaderText; } // Để CheckedListBox hiển thị đúng Tiếng Việt
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
        public string parent_id { get; set; } // TAG CHA
        public override string ToString() { return name; }

        [System.Text.Json.Serialization.JsonIgnore]
        public string UnsignedName => _un ?? (_un = Form1.ConvertToUnSignStatic(name));
        private string _un;
    }
}