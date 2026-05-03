// =========================================================================
// FILE: Form1_CardUI.cs
// THÊM VÀO PROJECT: GhiIssue → Add → New Item → Class
// Đây là partial class của Form1 — chứa TẤT CẢ code liên quan đến Card UI
// Toàn bộ logic gốc (đăng nhập, API, close issue, omicall...) KHÔNG thay đổi
// =========================================================================
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using System.IO;

namespace GhiIssue
{
    public partial class Form1
    {
        // ==================== FIELD MỚI ====================
        private CardGridPanel _cardPanel;   // Thay thế dgvCreateTickets

        // ==================== SETUP ====================
        /// <summary>
        /// Gọi hàm này thay vì SetupCreateTicketGrid() trong Form1_Load
        ///
        /// Trong Form1_Load, tìm dòng:
        ///     SetupCreateTicketGrid();
        /// Thay bằng:
        ///     SetupCardUI();
        ///
        /// Và tìm dòng:
        ///     LoadDraft();
        /// Thay bằng:
        ///     LoadDraftCards();
        /// </summary>
        private void SetupCardUI()
        {
            // 1. Tạo CardGridPanel
            _cardPanel = new CardGridPanel
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Location = dgvCreateTickets.Location,
                Size = dgvCreateTickets.Size,
            };

            // 2. Ẩn DataGridView cũ
            dgvCreateTickets.Visible = false;
            tabPage1.Controls.Add(_cardPanel);

            // 3. Bơm datasource
            _cardPanel.SetDataSources(
                tagList, categoryList, defaultTitles, defaultTypeIssues, defaultTechActions,
                employees, defaultCat, defaultSubCat, defaultAssigneeId);

            // 4. Auto-save khi data thay đổi
            // TỐI ƯU HÓA: Đẩy việc đếm thẻ và bù thẻ vào luồng chạy ngầm
            _cardPanel.CardsChanged += (s, e) =>
            {
                // Sử dụng BeginInvoke để không làm kẹt giao diện lúc người dùng đang gõ
                this.BeginInvoke(new Action(() =>
                {
                    UpdateStatusCount();
                    // 🌟 SỬA CHÍNH TẠI ĐÂY: Đổi EnsureMinCards(5) thành EnsureMinCards(0)
                    _cardPanel.EnsureMinCards(0);
                }));
            };
            // BẮT SỰ KIỆN KHI BẤM NÚT "GỬI" TRÊN 1 THẺ BẤT KỲ
            _cardPanel.SingleSendRequested += async (s, targetCard) =>
            {
                // Gọi hàm kiểm tra và gửi đúng 1 thẻ này đi (Tương tự logic của btnCreateTicket_CardClick)
                await SendSingleCardAsync(targetCard);
            };

            LoadDraftCards();
            // 5. Bỏ qua nút Thêm phiếu thủ công
            // SetupAddCardButton(); 

            // 6. Mồi sẵn 5 phiếu trống lúc mới bật phần mềm
            _cardPanel.EnsureMinCards(5);
        }

        /// <summary>Nút thêm card mới — đặt cạnh btnCreateTicket</summary>
        private void SetupAddCardButton()
        {
            var btnAdd = new Button
            {
                Text      = "+ Thêm phiếu",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 248, 255),
                ForeColor = Color.SteelBlue,
                Cursor    = Cursors.Hand,
                Size      = new Size(100, 30),
                Location  = new Point(btnBulkEdit.Right + 8, btnBulkEdit.Top),
                Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
                Anchor    = AnchorStyles.Bottom | AnchorStyles.Left,
            };
            btnAdd.FlatAppearance.BorderColor = Color.SteelBlue;
            btnAdd.Click += (s, e) => { _cardPanel.AddCard(); };
            tabPage1.Controls.Add(btnAdd);
            tabPage1.Controls.SetChildIndex(btnAdd, 0);
        }

        // ==================== DATASOURCE REFRESH ====================
        // Gọi các hàm này trong SyncTitlesBackgroundAsync, SyncCategoriesBackgroundAsync, SyncTagsBackgroundAsync
        // thay vì set DataSource trực tiếp cho cột của dgvCreateTickets


        private void CardUI_RefreshTitles()
        {
            if (_cardPanel == null) return;
            _cardPanel.UpdateTitleSource(defaultTitles);
        }

        private void CardUI_RefreshTypeIssues()
        {
            if (_cardPanel == null) return;
            _cardPanel.UpdateTypeIssueSource(defaultTypeIssues);
        }

        private void CardUI_RefreshDescriptions()
        {
            if (_cardPanel == null) return;
            _cardPanel.UpdateDescSource(defaultTechActions);
        }

        private void CardUI_RefreshTags()
        {
            if (_cardPanel == null) return;
            _cardPanel.UpdateTagSource(tagList);
        }

        private void CardUI_RefreshCategories()
        {
            if (_cardPanel == null) return;
            _cardPanel.UpdateCategorySource(categoryList);
        }

        private void CardUI_SetDefaultAssignee(string empId)
        {
            defaultAssigneeId = empId;
            _cardPanel?.SetDefaultAssignee(empId);
        }

        private void CardUI_ToggleCategory(bool show)
        {
            _cardPanel?.SetShowCategory(show);
        }

        // ==================== STATUS COUNT MỚI ====================
        /// <summary>
        /// Gọi hàm này từ UpdateStatusCount() — thêm vào case tab1
        /// if (tabControl1.SelectedTab == tabPage1 && _cardPanel != null)
        ///     myStatusLabel.Text = $"Phiếu chờ: {_cardPanel.PendingCount} / Tổng: {_cardPanel.TotalCount}";
        /// </summary>

        // ==================== DRAFT ====================
        private void SaveDraftCards()
        {
            _cardPanel?.SaveDraft(draftFilePath);
        }

        private void LoadDraftCards()
        {
            if (_cardPanel == null) return;
            _cardPanel.LoadDraft(draftFilePath);
            _cardPanel.EnsureMinCards(5); // Luôn có ít nhất 5 card trống để gõ
        }

        // ==================== TẠO PHIẾU (CARD VERSION) ====================
        /// <summary>
        /// Đây là phiên bản mới của btnCreateTicket_Click.
        /// THAY THẾ toàn bộ handler cũ bằng hàm này.
        ///
        /// Trong Form1_Load (hoặc nơi wire event), đổi:
        ///     btnCreateTicket.Click += btnCreateTicket_Click;
        /// thành:
        ///     btnCreateTicket.Click += btnCreateTicket_CardClick;
        /// </summary>
        private async void btnCreateTicket_CardClick(object sender, EventArgs e)
        {
            if (_cardPanel == null) { MessageBox.Show("Card panel chưa khởi tạo."); return; }

            Cursor.Current  = Cursors.WaitCursor;
            int successCount = 0;

            var vtiTag = tagList.FirstOrDefault(t =>
                t.name.Equals("VTI", StringComparison.OrdinalIgnoreCase)) ??
                tagList.FirstOrDefault(t => t.name.ToUpper().Contains("VTI"));

            // ==================== VALIDATE TRƯỚC ====================
            foreach (var card in _cardPanel.GetPendingCards())
            {
                string start = card.StartTime;
                string end   = card.EndTime;
                string type  = card.TypeIssueName;
                string title = card.TitleText;

                // Auto +2 phút nếu có time nhận nhưng quên time xong
                if (!string.IsNullOrEmpty(start) && string.IsNullOrEmpty(end))
                {
                    if (TimeSpan.TryParseExact(start, @"hh\:mm", null, out TimeSpan st))
                        card.ResultText = ""; // sẽ set qua property EndTime bên card... 
                    // LƯU Ý: TicketCardControl chưa expose setter cho EndTime
                    // → Cần thêm public setter "EndTime { set => txtEndTime.Text = value; }"
                }

                // Validate VTI
                bool isVTI = IsCardTagVTI(card, vtiTag);
                if (isVTI && (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(start) || string.IsNullOrEmpty(end)))
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show(
                        $"Phiếu '{title}' dùng Tag VTI/Highlands.\nBẮT BUỘC điền [Type Issue] và [Thời Gian]!",
                        "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // ==================== GỬI PHIẾU ====================
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
            client.DefaultRequestHeaders.Add("Accept", "application/json, text/plain, */*");
            client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

            foreach (var card in _cardPanel.GetPendingCards().ToList())
            {
                string title   = card.TitleText;
                string desc    = card.DescText;
                string type    = card.TypeIssueName;
                string start   = card.StartTime;
                string end     = card.EndTime;
                string tagId   = card.TagId;
                string empId   = card.AssigneeId;
                string catName = card.CategoryName;

                if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(catName)) continue;

                card.ResultText = "⏳ Đang gửi...";

                // Build description
                string finalDesc = string.IsNullOrEmpty(desc) ? "" : $"<div style=\"font-size: 15px;\">{desc}</div>";
                if (!string.IsNullOrEmpty(start) || !string.IsNullOrEmpty(end) || !string.IsNullOrEmpty(type))
                    finalDesc += $"<br><br>[TG: {start} - {end}]<br>[Type: {type}]";

                // Lookup category
                var catInfo    = categoryList.FirstOrDefault(c => c.name == catName);
                string realCatId = catInfo?._id ?? "";
                var subInfo    = card.GetSelectedSubCategory();
                int typeIndex  = subInfo?.index ?? 0;

                var body = new
                {
                    name        = title,
                    description = finalDesc,
                    category_id = realCatId,
                    current_type = typeIndex,
                    assignee_contact_ids = string.IsNullOrEmpty(empId) ? Array.Empty<string>() : new[] { empId },
                    tags         = string.IsNullOrEmpty(tagId) ? Array.Empty<string>() : new[] { tagId },
                    source       = "crud",
                    priority     = "medium",
                    embed_files  = Array.Empty<string>(),
                    reporter_contact_ids = Array.Empty<string>(),
                    work_list    = Array.Empty<string>(),
                    attribute_structures = Array.Empty<object>(),
                };

                try
                {
                    var content  = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/add?lng=vi&utm_source=web", content);
                    string res   = await response.Content.ReadAsStringAsync();

                    if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized || res.Contains("Unauthorized"))
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Phiên đăng nhập hết hạn! Vui lòng đăng nhập lại.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (File.Exists(tokenFilePath)) File.Delete(tokenFilePath);
                        OMICRM_TOKEN = "";
                        await PerformLoginSequenceAsync();
                        return;
                    }

                    if (response.IsSuccessStatusCode && (res.Contains("\"id\"") || res.Contains("\"_id\"") || res.Contains("\"success\":true")))
                    {
                        card.ResultText = "✅ Thành công";
                        successCount++;
                        LearnSmartTemplate(title, desc);

                        // Lấy ticket ID để log & push sheet
                        string ticketId = "N/A";
                        try { using var doc = JsonDocument.Parse(res); ticketId = doc.RootElement.GetProperty("payload").GetProperty("unique_id").GetString(); } catch { }

                        string tagName = tagList.FirstOrDefault(t => t.id == tagId)?.name ?? tagId;
                        string empName = employees.FirstOrDefault(emp => emp.Id == empId)?.Name ?? empId;
                        WriteLog("SUCCESS", $"Tạo phiếu '{title}'", $"Tag: {tagName} | NXL: {empName} | ID: {ticketId}");

                        // Push Google Sheet nếu VTI
                        bool isVtiPush = IsCardTagVTI(card, vtiTag);
                        if (isVtiPush)
                        {
                            string sheetGroup = card.GroupName;   // ✅ lấy từ AutoFillFromTitle
                            string mainType = card.MainType;    // ✅ lấy từ AutoFillFromTitle
                            string tStatus    = (!string.IsNullOrEmpty(start) && !string.IsNullOrEmpty(end)) ? "Đã Đóng (4)" : "Đang xử lý (1)";
                            _ = SendToGoogleSheetAsync(ticketId, sheetGroup, title, mainType, type, tagName, desc, tStatus, start, end);
                        }
                    }
                    else if (!response.IsSuccessStatusCode && res.StartsWith("<"))
                    {
                        card.ResultText = "❌ Tường lửa / Token hết hạn";
                        LogError($"Lỗi API ({response.StatusCode})");
                    }
                    else
                    {
                        card.ResultText = "❌ " + (res.Length > 80 ? res[..80] : res);
                        LogError($"Lỗi API ({response.StatusCode})", new Exception(res));
                    }
                }
                catch (Exception ex)
                {
                    card.ResultText = "❌ " + ex.Message;
                }
            }

            Cursor.Current = Cursors.Default;
            MessageBox.Show($"Xong! Đã tạo thành công {successCount} phiếu.", "Kết quả");

            // Reset card đã gửi thành công
            if (successCount > 0)
            {
                foreach (var card in _cardPanel.AllCards.Where(c => c.IsDone).ToList())
                    card.ClearAfterSuccess();

                _cardPanel.EnsureMinCards(5);
                SaveDraftCards();
            }
        }

        // ==================== KIỂM TRA TAG VTI (CARD VERSION) ====================
        private bool IsCardTagVTI(TicketCardControl card, TagItem vtiTag)
        {
            if (vtiTag == null || string.IsNullOrEmpty(card.TagId)) return false;
            var curTag = card.GetTagById(card.TagId);
            if (curTag == null) return false;
            return curTag.name.ToUpper().Contains("VTI")
                || curTag.name.ToUpper().Contains("HLC")
                || curTag.parent_id == vtiTag.id
                || curTag.id == vtiTag.id;
        }
        // HÀM CHUYÊN DỤNG ĐỂ GỬI ĐÚNG 1 THẺ LÊN OMICRM
        private async System.Threading.Tasks.Task SendSingleCardAsync(TicketCardControl card)
        {
            if (card == null || !card.HasData || card.IsDone) return;

            Cursor.Current = Cursors.WaitCursor;
            var vtiTag = tagList.FirstOrDefault(t => t.name.Equals("VTI", StringComparison.OrdinalIgnoreCase)) ??
                         tagList.FirstOrDefault(t => t.name.ToUpper().Contains("VTI"));

            string start = card.StartTime;
            string end = card.EndTime;
            string type = card.TypeIssueName;
            string title = card.TitleText;

            // 1. Kiểm tra ràng buộc VTI
            bool isVTI = IsCardTagVTI(card, vtiTag);
            if (isVTI && (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(start) || string.IsNullOrEmpty(end)))
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show($"Phiếu '{title}' dùng Tag VTI/Highlands.\nBẮT BUỘC điền [Type Issue] và [Thời Gian]!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Gói dữ liệu
            card.ResultText = "⏳ Đang gửi...";
            using var client = new System.Net.Http.HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", OMICRM_TOKEN);

            string finalDesc = string.IsNullOrEmpty(card.DescText) ? "" : $"<div style=\"font-size: 15px;\">{card.DescText}</div>";
            if (!string.IsNullOrEmpty(start) || !string.IsNullOrEmpty(end) || !string.IsNullOrEmpty(type))
                finalDesc += $"<br><br>[TG: {start} - {end}]<br>[Type: {type}]";

            var catInfo = categoryList.FirstOrDefault(c => c.name == card.CategoryName);
            string realCatId = catInfo?._id ?? "";
            var subInfo = card.GetSelectedSubCategory();
            int typeIndex = subInfo?.index ?? 0;

            var body = new
            {
                name = title,
                description = finalDesc,
                category_id = realCatId,
                current_type = typeIndex,
                assignee_contact_ids = string.IsNullOrEmpty(card.AssigneeId) ? Array.Empty<string>() : new[] { card.AssigneeId },
                tags = string.IsNullOrEmpty(card.TagId) ? Array.Empty<string>() : new[] { card.TagId },
                source = "crud",
                priority = "medium",
                embed_files = Array.Empty<string>(),
                reporter_contact_ids = Array.Empty<string>(),
                work_list = Array.Empty<string>(),
                attribute_structures = Array.Empty<object>(),
            };

            // 3. Bóp cò gửi đi
            try
            {
                var content = new System.Net.Http.StringContent(System.Text.Json.JsonSerializer.Serialize(body), System.Text.Encoding.UTF8, "application/json");
                var response = await client.PostAsync("https://ticket-v2-stg.omicrm.com/ticket/add?lng=vi&utm_source=web", content);
                string res = await response.Content.ReadAsStringAsync();

                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized || res.Contains("Unauthorized"))
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("Phiên đăng nhập hết hạn! Vui lòng đăng nhập lại.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (response.IsSuccessStatusCode && (res.Contains("\"id\"") || res.Contains("\"_id\"") || res.Contains("\"success\":true")))
                {
                    card.ResultText = "✅ Thành công";
                    LearnSmartTemplate(title, card.DescText);
                    // 🌟 Tuyệt chiêu: Tự động tẩy trắng thẻ sau 1.5 giây để gõ phiếu mới
                    System.Threading.Tasks.Task.Delay(1500).ContinueWith(t => {
                        this.Invoke(new Action(() => { card.ClearAfterSuccess(); _cardPanel.EnsureMinCards(5); }));
                    });

                    SaveDraftCards();
                }
                else card.ResultText = "❌ Lỗi: " + (res.Length > 50 ? res.Substring(0, 50) : res);
            }
            catch (Exception ex) { card.ResultText = "❌ Lỗi mạng: " + ex.Message; }

            Cursor.Current = Cursors.Default;
        }
    }
}
