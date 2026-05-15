using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

namespace GhiIssue
{
    public class CardGridPanel : Panel
    {
        // ==================== FIELDS ====================
        private FlowLayoutPanel _innerPanel;
        private List<TicketCardControl> _cards = new();

        private List<TagItem> _tagList;
        private List<CategoryItem> _categoryList;
        private BindingList<PredefinedTitle> _titleList;
        private BindingList<ComboItem> _typeIssueList;
        private BindingList<ComboItem> _descList;
        private List<Employee> _employeeList;
        private string _defaultCat = "";
        private string _defaultSubCat = "";
        private string _defaultAssigneeId = "";
        private bool _showCategory = false;

        // Biến này để "Đóng băng" các sự kiện thừa khi đang thêm/xóa hàng loạt
        private bool _isUpdatingCards = false;

        public event EventHandler CardsChanged;
        public event EventHandler<TicketCardControl> SingleSendRequested;

        // ==================== CONSTRUCTOR ====================
        public CardGridPanel()
        {
            this.AutoScroll = true;
            this.BackColor = Color.FromArgb(240, 242, 245);
            this.Padding = new Padding(8);

            _innerPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(0),
                Margin = new Padding(0),
            };
            this.Controls.Add(_innerPanel);
            this.SizeChanged += CardGridPanel_SizeChanged;
        }
        private void CardGridPanel_SizeChanged(object sender, EventArgs e)
        {
            UpdateCardWidths();
        }
        private void UpdateCardWidths()
        {
            if (_innerPanel == null || _innerPanel.Controls.Count == 0) return;

            _innerPanel.SuspendLayout(); // Tạm dừng vẽ UI để tránh giật lag

            // Lấy chiều rộng hiện tại của Panel cha, trừ đi padding (khoảng 30px để có lề 2 bên)
            int padding = 30;
            int targetWidth = this.ClientSize.Width - padding;

            // Trừ hao thêm độ rộng của thanh cuộn dọc (nếu có xuất hiện) để tránh bị tràn gây ra thanh cuộn ngang
            if (this.VerticalScroll.Visible || _innerPanel.VerticalScroll.Visible)
            {
                targetWidth -= SystemInformation.VerticalScrollBarWidth;
            }

            // Không cho phép Card bị bóp quá nhỏ gây vỡ layout (ví dụ: tối thiểu 500px)
            if (targetWidth < 500) targetWidth = 500;

            // Gán lại chiều rộng mới cho toàn bộ Card đang có
            foreach (Control ctrl in _innerPanel.Controls)
            {
                if (ctrl is TicketCardControl card)
                {
                    card.Width = targetWidth;
                }
            }

            _innerPanel.ResumeLayout(); // Tiếp tục vẽ lại UI
        }

        // ==================== BƠM DATASOURCE ====================
        public void SetDataSources(
            List<TagItem> tags, List<CategoryItem> categories, BindingList<PredefinedTitle> titles,
            BindingList<ComboItem> typeIssues, BindingList<ComboItem> techActions, List<Employee> employees,
            string defaultCat = "", string defaultSubCat = "", string defaultAssigneeId = "")
        {
            _tagList = tags; _categoryList = categories; _titleList = titles;
            _typeIssueList = typeIssues; _descList = techActions; _employeeList = employees;
            _defaultCat = defaultCat; _defaultSubCat = defaultSubCat; _defaultAssigneeId = defaultAssigneeId;

            foreach (var card in _cards.ToList())
                card.SetDataSources(tags, categories, titles, typeIssues, techActions, employees);
        }

        public void UpdateTitleSource(BindingList<PredefinedTitle> titles) { _titleList = titles; foreach (var c in _cards.ToList()) c.RefreshTitleSource(titles); }
        public void UpdateTypeIssueSource(BindingList<ComboItem> items) { _typeIssueList = items; foreach (var c in _cards.ToList()) c.RefreshTypeIssueSource(items); }
        public void UpdateDescSource(BindingList<ComboItem> items)
        {
            _descList = items;
            foreach (var c in _cards.ToList())
                c.RefreshDescSource(items);
        }
        public void UpdateTagSource(List<TagItem> tags) { _tagList = tags; foreach (var c in _cards.ToList()) c.RefreshTagSource(tags); }
        public void UpdateCategorySource(List<CategoryItem> cats) { _categoryList = cats; foreach (var c in _cards.ToList()) c.RefreshCategorySource(cats); }
        public void SetDefaultAssignee(string empId) { _defaultAssigneeId = empId; }
        public void SetShowCategory(bool show) { _showCategory = show; foreach (var c in _cards.ToList()) c.ShowCategory = show; }

        // ==================== LOGIC THÊM/XÓA CỐT LÕI (CHỐNG LAG) ====================

        // Hàm thêm thẻ rỗng không làm kích hoạt sự kiện thừa
        private TicketCardControl AddCardCore(DraftTicket draft = null)
        {
            var card = new TicketCardControl();
            card.CardIndex = _cards.Count;
            card.ShowCategory = _showCategory;
            card.Margin = new Padding(0, 0, 0, 10);

            if (_tagList != null)
                card.SetDataSources(_tagList, _categoryList, _titleList, _typeIssueList, _descList, _employeeList);

            card.SetDefaultCategory(_defaultCat, _defaultSubCat);
            card.SetDefaultAssignee(_defaultAssigneeId);

            if (draft != null)
                card.FromDraft(draft, _defaultCat, _defaultSubCat, _defaultAssigneeId);

            card.DeleteRequested += OnCardDeleteRequested;
            card.SendRequested += (s, e) => SingleSendRequested?.Invoke(this, card);

            card.DataChanged += (s, e) =>
            {
                if (!_isUpdatingCards) CardsChanged?.Invoke(this, EventArgs.Empty);
            };

            card.AssigneeChanged += (s, newEmpId) =>
            {
                _defaultAssigneeId = newEmpId;
                bool found = false;
                foreach (var c in _cards)
                {
                    if (c == card) { found = true; continue; }
                    if (found) c.SetDefaultAssignee(newEmpId);
                }
            };

            _cards.Add(card);
            _innerPanel.Controls.Add(card);
            card.SendToBack();

            return card;
        }

        public TicketCardControl AddCard(DraftTicket draft = null)
        {
            if (!_isUpdatingCards) _innerPanel.SuspendLayout();

            var card = AddCardCore(draft);

            if (!_isUpdatingCards)
            {
                ReindexCards();
                _innerPanel.ResumeLayout();
                CardsChanged?.Invoke(this, EventArgs.Empty);
            }
            return card;
        }

        private void OnCardDeleteRequested(object sender, EventArgs e)
        {
            if (sender is TicketCardControl card) RemoveCard(card);
        }

        public void RemoveCard(TicketCardControl card)
        {
            if (!_cards.Contains(card)) return;

            bool wasUpdating = _isUpdatingCards;
            _isUpdatingCards = true;
            if (!wasUpdating) _innerPanel.SuspendLayout();

            // 1. Gỡ sự kiện và xóa thẻ khỏi bộ nhớ / giao diện
            card.DeleteRequested -= OnCardDeleteRequested;
            _cards.Remove(card);
            _innerPanel.Controls.Remove(card);
            card.Dispose();

            if (!wasUpdating)
            {
                // 2. Đánh lại số thứ tự cho các thẻ còn lại
                ReindexCards();

                _innerPanel.ResumeLayout();
                _isUpdatingCards = false;

                // 🌟 SỬA LOGIC Ở ĐÂY: CHỈ bù lại 5 thẻ NẾU đã xóa sạch trơn không còn thẻ nào
                if (_cards.Count == 0)
                {
                    EnsureMinCards(5);
                }

                // Báo hiệu dữ liệu thay đổi để Form1 update tổng số dòng
                CardsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        // ==================== KIỂM SOÁT SỐ LƯỢNG THẺ ====================
        public void EnsureMinCards(int minTotal)
        {
            if (_isUpdatingCards) return;
            _isUpdatingCards = true;
            _innerPanel.SuspendLayout();

            bool added = false;

            // ✅ Tính 1 lần duy nhất, không scan lại trong vòng lặp
            int activeCnt = _cards.Count(c => !c.IsDone);

            while (activeCnt < minTotal)
            {
                AddCardCore();
                activeCnt++;
                added = true;
            }

            // Nếu thẻ cuối đã có dữ liệu → thêm 1 thẻ trống để gõ tiếp
            var activeCards = _cards.Where(c => !c.IsDone).ToList();
            if (activeCards.Count > 0 && activeCards.Last().HasData)
            {
                AddCardCore();
                added = true;
            }

            if (added) ReindexCards();

            _innerPanel.ResumeLayout();
            _isUpdatingCards = false;

            // ✅ Cập nhật chiều rộng ngay sau khi thêm card — fix lỗi scale
            if (added) UpdateCardWidths();

            if (added) CardsChanged?.Invoke(this, EventArgs.Empty);
        }

        public void ClearEmptyCards()
        {
            if (_isUpdatingCards) return;
            _isUpdatingCards = true;
            _innerPanel.SuspendLayout();

            // Xóa toàn bộ thẻ rỗng và chưa gửi
            var emptyCards = _cards.Where(c => !c.HasData && !c.IsDone).ToList();
            foreach (var c in emptyCards)
            {
                c.DeleteRequested -= OnCardDeleteRequested;
                _cards.Remove(c);
                _innerPanel.Controls.Remove(c);
                c.Dispose();
            }

            ReindexCards();
            _innerPanel.ResumeLayout();
            _isUpdatingCards = false;

            // Gọi lại hàm để bù cho đủ 5 thẻ mặc định
            EnsureMinCards(5);
            CardsChanged?.Invoke(this, EventArgs.Empty);
        }

        public void ClearAllCardsCore()
        {
            foreach (var c in _cards.ToList())
            {
                c.DeleteRequested -= OnCardDeleteRequested;
                _innerPanel.Controls.Remove(c);
                c.Dispose();
            }
            _cards.Clear();
        }

        private void ReindexCards()
        {
            for (int i = 0; i < _cards.Count; i++) _cards[i].CardIndex = i;
        }

        // ==================== QUERY VÀ DRAFT ====================
        public IEnumerable<TicketCardControl> GetPendingCards() => _cards.Where(c => c.HasData && !c.IsDone);
        public IReadOnlyList<TicketCardControl> AllCards => _cards.AsReadOnly();
        public int TotalCount => _cards.Count;
        public int PendingCount => _cards.Count(c => !c.IsDone);

        public void SaveDraft(string filePath)
        {
            try
            {
                var drafts = _cards.Where(c => !c.IsDone && c.HasData).Select(c => c.ToDraft()).Where(d => d != null).ToList();
                if (drafts.Count > 0)
                {
                    string tmp = filePath + ".tmp";
                    File.WriteAllText(tmp, JsonSerializer.Serialize(drafts));
                    if (File.Exists(filePath)) File.Replace(tmp, filePath, filePath + ".bak"); else File.Move(tmp, filePath);
                }
                else if (File.Exists(filePath)) File.Delete(filePath);
            }
            catch { }
        }

        //public void LoadDraft(string filePath)
        //{
        //    if (!File.Exists(filePath)) return;
        //    try
        //    {
        //        string json = File.ReadAllText(filePath);
        //        var drafts = JsonSerializer.Deserialize<List<DraftTicket>>(json);
        //        if (drafts == null || drafts.Count == 0) return;

        //        _isUpdatingCards = true;
        //        _innerPanel.SuspendLayout();

        //        ClearAllCardsCore(); // Dọn dẹp siêu tốc không báo event
        //        foreach (var d in drafts) AddCardCore(d);

        //        ReindexCards();
        //        _innerPanel.ResumeLayout();
        //        _isUpdatingCards = false;
        //    }
        //    catch { }
        //}
        public void LoadDraft(string filePath)
        {
            if (!File.Exists(filePath)) return;
            try
            {
                string json = File.ReadAllText(filePath);
                var drafts = JsonSerializer.Deserialize<List<DraftTicket>>(json);
                if (drafts == null || drafts.Count == 0) return;

                _isUpdatingCards = true;
                _innerPanel.SuspendLayout();

                ClearAllCardsCore();

                // 🌟 BỘ LỌC CỰC MẠNH: Chỉ tải những bản nháp THỰC SỰ CÓ CHỮ (Chặn đứng 100 dòng trống từ file cũ)
                var validDrafts = drafts.Where(d => !string.IsNullOrWhiteSpace(d.Title) || !string.IsNullOrWhiteSpace(d.TagId)).ToList();

                foreach (var d in validDrafts) AddCardCore(d);

                ReindexCards();
                _innerPanel.ResumeLayout();
                _isUpdatingCards = false;
            }
            catch { }
        }
        // Hàm này nhận lệnh từ Form1 và truyền màu xuống tất cả các thẻ
        public void ApplyThemeToCards(Color cardBackColor)
        {
            foreach (var card in _cards)
            {
                card.ApplyTheme(cardBackColor);
            }
        }
        // Hàm truyền màu
        public void ApplyColorToCards(Color c)
        {
            foreach (Control ctrl in _innerPanel.Controls)
            {
                if (ctrl is TicketCardControl card) card.SetCardColor(c);
            }
        }

        // Hàm truyền ảnh
        public void ApplyImageToCards(string path)
        {
            foreach (Control ctrl in _innerPanel.Controls)
            {
                if (ctrl is TicketCardControl card) card.SetCardBackground(path);
            }
        }
        // Hàm đổi màu cho cái Khung nền lớn
        public void SetPanelColor(Color c)
        {
            this.BackgroundImage = null; // Xoá ảnh nền nếu đang có
            this.BackColor = c;

            // Ép cái panel chứa card bên trong phải trong suốt thì mới thấy màu của panel cha
            if (_innerPanel != null) _innerPanel.BackColor = Color.Transparent;
        }

        // Hàm chèn ảnh cho cái Khung nền lớn
        public void SetPanelImage(string path)
        {
            if (File.Exists(path))
            {
                this.BackgroundImage = Image.FromFile(path);
                this.BackgroundImageLayout = ImageLayout.Stretch;

                // Làm trong suốt panel con để lộ ảnh nền ra
                if (_innerPanel != null) _innerPanel.BackColor = Color.Transparent;
            }
        }
    }
}