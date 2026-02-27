using System;
using System.Collections.Generic;
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
        private string draftFilePath = Path.Combine(Application.StartupPath, "draft.json"); // File nháp

        private List<Employee> employees;
        private List<CategoryItem> categoryList = new List<CategoryItem>();
        private List<TagItem> tagList = new List<TagItem>();

        // Bộ nhớ lưu tên người xử lý để dùng trọn đời cho các dòng thêm mới
        private string defaultAssigneeId = "";

        // Bộ nhớ tạm chống OmiCRM phản hồi chậm (Danh sách đen - khử Ghost Tickets)
        private HashSet<string> recentlyClosedTicketIds = new HashSet<string>();

        // Đồng hồ đếm nhịp Auto-Save
        private System.Windows.Forms.Timer autoSaveTimer;

        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;

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

            if (dgvCreateTickets != null)
            {
                dgvCreateTickets.DataError += DgvCreateTickets_DataError;
                dgvCreateTickets.CellValueChanged += DgvCreateTickets_CellValueChanged;
                dgvCreateTickets.CurrentCellDirtyStateChanged += DgvCreateTickets_CurrentCellDirtyStateChanged;
                dgvCreateTickets.EditingControlShowing += DgvCreateTickets_EditingControlShowing;

                // Bẫy sự kiện: Cứ gõ xong 1 ô hoặc xóa 1 dòng là tự động lưu nháp
                dgvCreateTickets.CellEndEdit += (s, e) => SaveDraft();
                dgvCreateTickets.RowsRemoved += (s, e) => SaveDraft();
            }

            // 1. GẮN CẢNH BÁO KHI BẤM NÚT X (TẮT PHẦN MỀM)
            this.FormClosing += Form1_FormClosing;

            // 2. BẬT HỆ THỐNG AUTO-SAVE NGẦM (5 GIÂY LƯU 1 LẦN CHỐNG CÚP ĐIỆN)
            autoSaveTimer = new System.Windows.Forms.Timer();
            autoSaveTimer.Interval = 5000; // 5000 ms = 5 giây
            autoSaveTimer.Tick += (s, e) => SaveDraft();
            autoSaveTimer.Start(); // Bắt đầu chạy ngầm
        }

        // ==========================================================
        // HỎI ĐÁP TRƯỚC KHI THOÁT & CHỐT DỮ LIỆU CUỐI CÙNG
        // ==========================================================
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Nếu người dùng bấm X hoặc Alt+F4
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show(
                    "Bạn có chắc chắn muốn thoát Tool Ghi Issue không?\n(Dữ liệu đang nhập dở sẽ được tự động lưu nháp an toàn)",
                    "Xác nhận thoát",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                    e.Cancel = true; // Hủy lệnh tắt, ở lại xài tiếp
                    return;
                }
            }

            // Cú chốt: Ép bảng cập nhật nốt cái chữ đang gõ dở (chưa kịp bấm Enter)
            if (dgvCreateTickets != null && dgvCreateTickets.IsCurrentCellInEditMode)
            {
                dgvCreateTickets.EndEdit();
            }

            SaveDraft(); // Lưu lần cuối cùng trước khi nhắm mắt
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
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

                // Đồng bộ tính năng tìm kiếm thả xuống cho ô Người xử lý bên tab Đóng Issue
                cboEmployees.AutoCompleteMode = AutoCompleteMode.None;
                cboEmployees.DropDownStyle = ComboBoxStyle.DropDown;
                cboEmployees.TextUpdate -= cboEmployees_TextUpdate;
                cboEmployees.TextUpdate += cboEmployees_TextUpdate;
            }

            SetupCreateTicketGrid();

            // 1. TỰ ĐỘNG KIỂM TRA UPDATE KHI MỞ APP
            CheckAndDownloadUpdateAsync();

            // 2. KÍCH HOẠT HỆ THỐNG KIỂM TRA TOKEN & ĐĂNG NHẬP
            await PerformLoginSequenceAsync();

            // 3. KHÔI PHỤC BẢN NHÁP (NẾU CÓ BỊ CÚP ĐIỆN)
            LoadDraft();
        }

        // ==========================================================
        // TÌM KIẾM TÊN NGƯỜI XỬ LÝ KHÔNG DẤU TẠI TAB ĐÓNG ISSUE
        // ==========================================================
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

        // ==========================================================
        // CƠ CHẾ AUTO-SAVE (CHỐNG CÚP ĐIỆN)
        // ==========================================================
        private void SaveDraft()
        {
            try
            {
                var drafts = new List<DraftTicket>();
                foreach (DataGridViewRow row in dgvCreateTickets.Rows)
                {
                    if (row.IsNewRow) continue;

                    // Chỉ lưu những dòng có gõ Tiêu đề hoặc chọn Chủ đề
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
                    File.Delete(draftFilePath); // Bảng trống thì xóa file rác đi
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
                        dgvCreateTickets.Rows.Clear(); // Xóa các dòng mặc định

                        // Tạm tắt sự kiện Thác nước để load dữ liệu không bị loạn
                        dgvCreateTickets.CellValueChanged -= DgvCreateTickets_CellValueChanged;

                        foreach (var draft in drafts)
                        {
                            int idx = dgvCreateTickets.Rows.Add();
                            var row = dgvCreateTickets.Rows[idx];

                            row.Cells["colTitle"].Value = draft.Title;
                            row.Cells["colDesc"].Value = draft.Desc;
                            row.Cells["colCategory"].Value = draft.CatId;

                            // Ép lại danh sách Phân loại phụ (SubCat) cho dòng này
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

        // =========================================================================
        // HỆ THỐNG AUTO-UPDATE (CÓ THANH TIẾN TRÌNH PROGRESS BAR)
        // =========================================================================
        private async void CheckAndDownloadUpdateAsync()
        {
            string currentVersion = "1.2"; // Nhớ sửa số này khi lên bản mới
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
                            // === 1. TẠO BẢNG THÔNG BÁO TIẾN TRÌNH ===
                            Form progressForm = new Form()
                            {
                                Text = "Đang tải bản cập nhật...",
                                Size = new Size(400, 120),
                                FormBorderStyle = FormBorderStyle.FixedDialog,
                                StartPosition = FormStartPosition.CenterScreen,
                                ControlBox = false // Khóa nút X để không cho tắt ngang
                            };
                            ProgressBar pb = new ProgressBar() { Left = 20, Top = 20, Width = 340, Height = 25 };
                            Label lbl = new Label() { Left = 20, Top = 50, Width = 340, Text = "Đang kết nối đến GitHub..." };
                            progressForm.Controls.Add(pb);
                            progressForm.Controls.Add(lbl);
                            progressForm.Show();

                            string tempExeName = "GhiIssue_Update.exe";

                            // === 2. TẢI FILE THEO TỪNG CHUNK ĐỂ LẤY PHẦN TRĂM (%) ===
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
                                                    // Tính % và số MB đã tải
                                                    int percent = (int)((double)totalRead / totalBytes.Value * 100);
                                                    pb.Value = percent;
                                                    lbl.Text = $"Đang tải: {totalRead / 1048576} MB / {totalBytes.Value / 1048576} MB ({percent}%)";
                                                    Application.DoEvents(); // Ép giao diện mượt, không bị đơ
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception downloadEx)
                            {
                                // Nếu rớt mạng giữa chừng
                                progressForm.Close();
                                MessageBox.Show("Lỗi tải file: " + downloadEx.Message, "Lỗi mạng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return; // Dừng luôn, không chạy cài đặt nữa
                            }

                            progressForm.Close();

                            // === 3. CHẠY KỊCH BẢN XÓA CŨ - THẾ MỚI ===
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
            catch (Exception ex)
            {
                // MessageBox.Show("Lỗi check Update: " + ex.Message);
            }
        }

        // =========================================================================
        // HỆ THỐNG ĐĂNG NHẬP & LẤY TOKEN (CHUẨN DOANH NGHIỆP)
        // =========================================================================
        private async Task PerformLoginSequenceAsync()
        {
            bool isDataLoaded = false;

            if (File.Exists(tokenFilePath))
            {
                OMICRM_TOKEN = File.ReadAllText(tokenFilePath).Trim();
                isDataLoaded = await LoadDataWithTokenAsync();
            }

            while (!isDataLoaded)
            {
                string email = "", password = "";
                if (ShowLoginDialog(out email, out password))
                {
                    Cursor.Current = Cursors.WaitCursor;
                    bool loginSuccess = await LoginAndGetTokenAsync(email, password);
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

        // =========================================================================
        // HỆ THỐNG GÕ TẮT - TÌM KIẾM KHÔNG DẤU - BẤM ENTER NHẢY Ô
        // =========================================================================
        private void DgvCreateTickets_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox cb)
            {
                cb.DropDownStyle = ComboBoxStyle.DropDown;
                cb.AutoCompleteMode = AutoCompleteMode.None;

                string colName = dgvCreateTickets.CurrentCell.OwningColumn.Name;
                if (colName == "colCategory") cb.DataSource = categoryList;
                else if (colName == "colTag") cb.DataSource = tagList;
                else if (colName == "colAssignee") cb.DataSource = employees;
                else if (colName == "colSubCategory")
                {
                    string catId = dgvCreateTickets.CurrentRow.Cells["colCategory"].Value?.ToString();
                    var cat = categoryList.FirstOrDefault(c => c._id == catId);
                    cb.DataSource = (cat != null && cat.types != null) ? cat.types : new List<SubCategoryItem>();
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
                e.Handled = true; e.SuppressKeyPress = true; SendKeys.Send("{TAB}");
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

            if (colName == "colCategory") cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? categoryList : categoryList.Where(x => ConvertToUnSign(x.name).Contains(searchKeyword)).ToList();
            else if (colName == "colTag") cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? tagList : tagList.Where(x => ConvertToUnSign(x.name).Contains(searchKeyword)).ToList();
            else if (colName == "colAssignee") cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? employees : employees.Where(x => ConvertToUnSign(x.Name).Contains(searchKeyword)).ToList();
            else if (colName == "colSubCategory")
            {
                string catId = dgvCreateTickets.CurrentRow.Cells["colCategory"].Value?.ToString();
                var cat = categoryList.FirstOrDefault(c => c._id == catId);
                if (cat != null && cat.types != null) cb.DataSource = string.IsNullOrEmpty(searchKeyword) ? cat.types : cat.types.Where(x => ConvertToUnSign(x.name).Contains(searchKeyword)).ToList();
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

        // =========================================================================
        // HỆ THỐNG SAO CHÉP "THÁC NƯỚC" TỪ TRÊN XUỐNG
        // =========================================================================
        private void DgvCreateTickets_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string colName = dgvCreateTickets.Columns[e.ColumnIndex].Name;
            var currentRow = dgvCreateTickets.Rows[e.RowIndex];

            if (colName == "colCategory")
            {
                string selectedCatId = currentRow.Cells["colCategory"].Value?.ToString();
                var cellSubCat = (DataGridViewComboBoxCell)currentRow.Cells["colSubCategory"];
                var cat = categoryList.FirstOrDefault(c => c._id == selectedCatId);
                if (cat != null && cat.types != null)
                {
                    cellSubCat.DataSource = cat.types;
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
                                targetCellSubCat.DataSource = cat.types;
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

        // NÂNG CẤP: Vét sạch toàn bộ Tag đang có trên OmiCRM
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

        // ==========================================================
        // KHỞI TẠO BẢNG & TÍNH NĂNG COPY NHƯ EXCEL
        // ==========================================================
        private void SetupCreateTicketGrid()
        {
            if (dgvCreateTickets == null) return;
            dgvCreateTickets.Columns.Clear();
            dgvCreateTickets.AutoGenerateColumns = false;
            dgvCreateTickets.AllowUserToAddRows = true;

            // Bật tính năng copy và bôi đen như Excel
            dgvCreateTickets.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            dgvCreateTickets.MultiSelect = true;
            dgvCreateTickets.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            // Bắt phím Ctrl C, V, X
            dgvCreateTickets.KeyDown -= DgvCreateTickets_KeyDown;
            dgvCreateTickets.KeyDown += DgvCreateTickets_KeyDown;

            // Sinh menu chuột phải
            ContextMenuStrip gridMenu = new ContextMenuStrip();
            gridMenu.Items.Add("Copy (Ctrl+C)", null, (s, e) => CopyToClipboard());
            gridMenu.Items.Add("Paste (Ctrl+V)", null, (s, e) => PasteFromClipboard());
            gridMenu.Items.Add("Cut (Ctrl+X)", null, (s, e) => CutToClipboard());
            dgvCreateTickets.ContextMenuStrip = gridMenu;

            dgvCreateTickets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tiêu đề phiếu", Name = "colTitle", Width = 150 });
            dgvCreateTickets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Mô tả chi tiết", Name = "colDesc", Width = 200 });

            DataGridViewComboBoxColumn colCategory = new DataGridViewComboBoxColumn();
            colCategory.HeaderText = "Chủ đề";
            colCategory.Name = "colCategory";
            colCategory.DisplayMember = "name";
            colCategory.ValueMember = "_id";
            colCategory.Width = 150;
            dgvCreateTickets.Columns.Add(colCategory);

            DataGridViewComboBoxColumn colSubCategory = new DataGridViewComboBoxColumn();
            colSubCategory.HeaderText = "Phân loại";
            colSubCategory.Name = "colSubCategory";
            colSubCategory.DisplayMember = "name";
            colSubCategory.ValueMember = "uuid";
            colSubCategory.Width = 180;
            dgvCreateTickets.Columns.Add(colSubCategory);

            DataGridViewComboBoxColumn colTag = new DataGridViewComboBoxColumn();
            colTag.HeaderText = "Tag";
            colTag.Name = "colTag";
            colTag.DisplayMember = "name";
            colTag.ValueMember = "id";
            colTag.Width = 150;
            dgvCreateTickets.Columns.Add(colTag);

            DataGridViewComboBoxColumn colAssignee = new DataGridViewComboBoxColumn();
            colAssignee.HeaderText = "Người xử lý";
            colAssignee.Name = "colAssignee";
            colAssignee.DataSource = employees;
            colAssignee.DisplayMember = "Name";
            colAssignee.ValueMember = "Id";
            colAssignee.Width = 150;
            dgvCreateTickets.Columns.Add(colAssignee);

            dgvCreateTickets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Kết quả", Name = "colResult", Width = 180, ReadOnly = true });
        }

        // ==========================================================
        // CÁC HÀM XỬ LÝ CHUỘT PHẢI / PHÍM TẮT
        // ==========================================================
        private void DgvCreateTickets_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C) { CopyToClipboard(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.V) { PasteFromClipboard(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.X) { CutToClipboard(); e.Handled = true; }
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
                foreach (DataGridViewCell cell in dgvCreateTickets.SelectedCells)
                {
                    // Cho phép xóa (Cut) toàn bộ các cột, trừ cột Kết quả
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

                // Tạm tắt sự kiện thác nước để tránh bị đè dữ liệu khi đang dán hàng loạt
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

                        // Nếu là cột chữ bình thường -> Dán thẳng
                        if (colName == "colTitle" || colName == "colDesc")
                        {
                            targetCell.Value = cellText;
                        }
                        // Nếu là cột Chủ đề -> Dịch tên thành ID
                        else if (colName == "colCategory")
                        {
                            var match = categoryList.FirstOrDefault(c => c.name.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                            if (match != null)
                            {
                                targetCell.Value = match._id;
                                // Mồi sẵn danh sách Phân loại phụ cho ô bên cạnh để nó kịp nhận dữ liệu
                                var cellSubCat = (DataGridViewComboBoxCell)dgvCreateTickets.Rows[startRow].Cells["colSubCategory"];
                                if (match.types != null)
                                {
                                    cellSubCat.DataSource = match.types;
                                    cellSubCat.DisplayMember = "name";
                                    cellSubCat.ValueMember = "uuid";
                                }
                            }
                        }
                        // Nếu là cột Phân loại phụ -> Dịch tên thành UUID
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
                        // Nếu là cột Tag -> Dịch tên thành ID Tag
                        else if (colName == "colTag")
                        {
                            var match = tagList.FirstOrDefault(t => t.name.Equals(cellText, StringComparison.OrdinalIgnoreCase));
                            if (match != null) targetCell.Value = match.id;
                        }
                        // Nếu là cột Người xử lý -> Dịch Tên thành ID Nhân viên
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
                // Bật lại sự kiện thác nước
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
                                targetCellSubCat.DataSource = cat.types;
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

        // ==========================================================
        // TẠO PHIẾU HÀNG LOẠT VÀ TỰ DỌN DẸP
        // ==========================================================
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

                        if (response.IsSuccessStatusCode)
                        {
                            row.Cells["colResult"].Value = "✅ Thành công";
                            successCount++;
                        }
                        else
                        {
                            string errorDetail = await response.Content.ReadAsStringAsync();
                            if (errorDetail.StartsWith("<")) row.Cells["colResult"].Value = "❌ Tường lửa chặn / Lỗi Token";
                            else row.Cells["colResult"].Value = "❌ " + (errorDetail.Length > 80 ? errorDetail.Substring(0, 80) : errorDetail);
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

                SaveDraft(); // Dọn dẹp xong thì cập nhật lại file nháp luôn cho sạch
            }
        }

        // ==========================================================
        // ĐÓNG ISSUE HÀNG LOẠT (BẢO VỆ AN TOÀN)
        // ==========================================================
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

                    if (!searchResponse.IsSuccessStatusCode)
                    {
                        MessageBox.Show($"Lỗi API Tìm kiếm: {searchResponse.StatusCode}. Có thể Token đã hết hạn!", "Lỗi");
                        Cursor.Current = Cursors.Default;
                        return;
                    }

                    string responseString = await searchResponse.Content.ReadAsStringAsync();
                    OmiResponse omiData = JsonSerializer.Deserialize<OmiResponse>(responseString);

                    List<TicketItem> allTickets = new List<TicketItem>();
                    if (omiData?.payload?.items != null) allTickets = omiData.payload.items;

                    // KHỬ GHOST TICKETS BẰNG DANH SÁCH ĐEN
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

                                    // GHI VÀO DANH SÁCH ĐEN KHI ĐÓNG THÀNH CÔNG
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

    // =========================================================================
    // CÁC LỚP MÔ HÌNH DỮ LIỆU (MODELS)
    // =========================================================================
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