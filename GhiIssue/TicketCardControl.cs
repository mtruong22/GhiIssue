using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace GhiIssue
{

    public partial class TicketCardControl : UserControl
    {
        // NHÚNG API WINDOWS ĐỂ TẠO CHỮ MỜ CHO COMBOBOX
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lParam);
        private const int CB_SETCUEBANNER = 0x1703;
        // ==================== NGUỒN DỮ LIỆU ====================
        private List<TagItem> _tagList;
        private List<CategoryItem> _categoryList;
        private BindingList<PredefinedTitle> _titleList;
        private BindingList<ComboItem> _typeIssueList;
        private BindingList<ComboItem> _descList;
        private List<Employee> _employeeList;
        private bool _suspendDataChanged;

        // ==================== EVENTS ====================
        public event EventHandler DeleteRequested;
        public event EventHandler DataChanged;
        public event EventHandler<string> AssigneeChanged;
        public event EventHandler SendRequested;

        // ==================== PROPERTIES VÀ CONTROL ====================
        private int _cardIndex;
        private Panel pnlLeftStrip;
        private Label lblIndex;
        private ComboBox cboTag;
        private ComboBox cboTypeIssue;
        private ComboBox cboTitle;
        private TextBox txtStartTime;
        private TextBox txtEndTime;
        private Label label1;
        private ComboBox cboAssignee;
        private Button btnDelete;
        private ComboBox cboDesc;
        private ComboBox cboCategory;
        private ComboBox cboSubCategory;
        private Label lblResult;
        private Button btnSend;
        private Button btnSuggest;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int CardIndex
        {
            get => _cardIndex;
            set { _cardIndex = value; if (lblIndex != null) lblIndex.Text = $"#{value + 1:D2}"; }
        }

        public string TagId => cboTag.SelectedValue?.ToString() ?? "";
        public string TitleText => cboTitle.Text.Trim();
        public string GroupName { get; private set; } = "Khác (Nhập tay)";
        public string MainType { get; private set; } = "";
        public string CategoryName => cboCategory.Text.Trim();
        public string SubCategoryName => cboSubCategory.Text.Trim();
        public string TypeIssueName => cboTypeIssue.Text.Trim();
        public string DescText => cboDesc.Text.Trim();

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string StartTime
        {
            get => txtStartTime.Text.Trim();
            set { if (txtStartTime != null) txtStartTime.Text = value ?? ""; }
        }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string EndTime
        {
            get => txtEndTime.Text.Trim();
            set { if (txtEndTime != null) txtEndTime.Text = value ?? ""; }
        }

        public string AssigneeId => cboAssignee.SelectedValue?.ToString() ?? "";
        public bool IsDone => lblResult.Text.Contains("☁️") || lblResult.Text.Contains("✅");
        public bool HasData => !string.IsNullOrEmpty(TitleText) || !string.IsNullOrEmpty(TagId);

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string ResultText
        {
            get => lblResult?.Text ?? "";
            set => this.Invoke(() => ApplyResult(value));
        }

        private bool _showCategory;
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool ShowCategory
        {
            get => _showCategory;
            set
            {
                _showCategory = value;
                if (cboCategory != null)
                {
                    cboCategory.Visible = value;
                    cboSubCategory.Visible = value;
                }
            }
        }

        public TicketCardControl()
        {
            InitializeComponent();
            SetStyle(ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);

            // Set Placeholder cho ComboBox
            SendMessage(cboTag.Handle, CB_SETCUEBANNER, (IntPtr)0, "🏷️ Chọn Tag...");
            SendMessage(cboTitle.Handle, CB_SETCUEBANNER, (IntPtr)0, "🔎 Tiêu đề phiếu...");
            SendMessage(cboTypeIssue.Handle, CB_SETCUEBANNER, (IntPtr)0, "Type Issue...");
            SendMessage(cboDesc.Handle, CB_SETCUEBANNER, (IntPtr)0, "📝 Mô tả chi tiết...");
            SendMessage(cboAssignee.Handle, CB_SETCUEBANNER, (IntPtr)0, "👤 Người XL...");

            // Set Placeholder cho TextBox Thời gian
            txtStartTime.PlaceholderText = "HH:mm";
            txtEndTime.PlaceholderText = "HH:mm";

            WireEvents();
        }

        // ==================== GẮN RÀNG BUỘC VÀ SỰ KIỆN ====================
        private void WireEvents()
        {
            btnDelete.Click += (s, e) => DeleteRequested?.Invoke(this, EventArgs.Empty);
            this.Paint += TicketCardControl_Paint;

            if (btnSend != null) btnSend.Click += (s, e) => SendRequested?.Invoke(this, EventArgs.Empty);
            cboCategory.SelectedIndexChanged += (s, e) => { RefreshSubCategorySource(); FireDataChanged(); };
            cboAssignee.SelectedIndexChanged += (s, e) =>
            {
                if (!_suspendDataChanged) AssigneeChanged?.Invoke(this, AssigneeId);
                FireDataChanged();
            };

            // Bắt sự kiện Format giờ
            txtStartTime.TextChanged += TimeBox_TextChanged;
            txtEndTime.TextChanged += TimeBox_TextChanged;
            txtStartTime.KeyPress += TimeBox_KeyPress;
            txtEndTime.KeyPress += TimeBox_KeyPress;

            // Bắt sự kiện Tìm kiếm
            cboTitle.TextUpdate += CboTitle_TextUpdate;
            cboTag.TextUpdate += CboTag_TextUpdate;
            cboTypeIssue.TextUpdate += CboTypeIssue_TextUpdate;
            cboDesc.TextUpdate += CboDesc_TextUpdate;
            cboAssignee.TextUpdate += CboAssignee_TextUpdate;

            // Thêm các dòng này vào trong WireEvents()
            cboTypeIssue.SelectedIndexChanged += CboTypeIssue_SelectedIndexChanged;
            cboDesc.DoubleClick += CboDesc_DoubleClick;
            txtStartTime.Leave += TxtStartTime_Leave;

            cboTitle.DropDown += AdjustDropDownWidth;
            cboTag.DropDown += AdjustDropDownWidth;
            cboTypeIssue.DropDown += AdjustDropDownWidth;
            cboDesc.DropDown += AdjustDropDownWidth;

            cboTitle.Validating += CboTitle_Validating;
            cboTypeIssue.Validating += CboTypeIssue_Validating;
            cboTag.Validating += (s, e) => { UpdateStripColor(); FilterTypeIssues(); AutoFillFromTitle(); FireDataChanged(); };

            // Mapping tự động
            cboTitle.SelectedIndexChanged += (s, e) => { AutoFillFromTitle(); FireDataChanged(); };
            cboTag.SelectedIndexChanged += (s, e) => { UpdateStripColor(); AutoFillFromTitle(); FireDataChanged(); };
            txtStartTime.Leave += (s, e) => FireDataChanged();
            txtEndTime.Leave += (s, e) => FireDataChanged();
            foreach (Control c in new Control[] { cboTag, cboTitle, cboTypeIssue, txtStartTime, txtEndTime, cboAssignee, cboDesc, cboCategory, cboSubCategory })
                c.Leave += (s, e) => FireDataChanged();
        }

        // 🌟 KIỂM TRA XEM TAG HIỆN TẠI CÓ PHẢI VTI KHÔNG
        private bool IsCurrentTagVTI()
        {
            var vtiTag = _tagList?.FirstOrDefault(t => t.name.Equals("VTI", StringComparison.OrdinalIgnoreCase)) ?? _tagList?.FirstOrDefault(t => t.name.ToUpper().Contains("VTI"));
            var curTag = GetSelectedTag();
            if (curTag != null)
            {
                if (curTag.name.ToUpper().Contains("VTI") || curTag.name.ToUpper().Contains("HLC") || (vtiTag != null && (curTag.parent_id == vtiTag.id || curTag.id == vtiTag.id)))
                    return true;
            }
            return false;
        }

        // 🌟 LỌC TYPE ISSUE THEO TAG ĐANG CHỌN
        private void FilterTypeIssues()
        {
            if (_typeIssueList == null) return;
            bool isVTI = IsCurrentTagVTI();

            // Lọc: Lấy các Type "Chung", và Type khớp với VTI hoặc Khách lẻ
            var filtered = _typeIssueList.Where(t =>
                string.IsNullOrEmpty(t.Classification) ||
                t.Classification.IndexOf("Chung", StringComparison.OrdinalIgnoreCase) >= 0 ||
                (isVTI && t.Classification.IndexOf("VTI", StringComparison.OrdinalIgnoreCase) >= 0) ||
                (!isVTI && t.Classification.IndexOf("Khách", StringComparison.OrdinalIgnoreCase) >= 0)
            ).ToList();

            string currentText = cboTypeIssue.Text;
            BindCombo(cboTypeIssue, new BindingList<ComboItem>(filtered), "Text", "Text");
            cboTypeIssue.Text = currentText; // Phục hồi lại chữ đang gõ
        }

        // 🌟 BẮT LỖI TÊN PHIẾU NẾU ADMIN BẬT STRICT MODE
        private void CboTitle_Validating(object sender, CancelEventArgs e)
        {
            if (Form1.isStrictTitle && !string.IsNullOrEmpty(cboTitle.Text))
            {
                if (_titleList != null && !_titleList.Any(t => t.Title.Equals(cboTitle.Text, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show("Không tìm thấy nội dung trong template!\nAdmin đang bật chế độ bắt buộc chọn từ danh sách.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cboTitle.Text = "";
                    cboTitle.SelectedIndex = -1;
                }
            }
        }

        // 🌟 BẮT LỖI TYPE ISSUE NẾU ADMIN BẬT STRICT MODE
        private void CboTypeIssue_Validating(object sender, CancelEventArgs e)
        {
            if (Form1.isStrictType && !string.IsNullOrEmpty(cboTypeIssue.Text))
            {
                if (_typeIssueList != null && !_typeIssueList.Any(t => t.Text.Equals(cboTypeIssue.Text, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show("Type Issue bắt buộc phải chọn theo Template có sẵn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cboTypeIssue.Text = "";
                    cboTypeIssue.SelectedIndex = -1;
                }
            }
        }

        private void TicketCardControl_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var pen = new Pen(Color.FromArgb(210, 215, 230), 1f);
            g.DrawRectangle(pen, new Rectangle(1, 1, this.Width - 2, this.Height - 2));
        }

        // ==================== LOGIC TÌM KIẾM COMBOBOX ====================
        // ==================== LOGIC TÌM KIẾM COMBOBOX ====================
        private void CboTitle_TextUpdate(object sender, EventArgs e)
        {
            if (_titleList == null) return;
            string keyword = cboTitle.Text;
            string kw = Form1.ConvertToUnSignStatic(keyword);
            int pos = cboTitle.SelectionStart;
            cboTitle.TextUpdate -= CboTitle_TextUpdate;

            var filtered = string.IsNullOrEmpty(kw) ? _titleList.ToList() : _titleList.Where(t => Form1.ConvertToUnSignStatic(t.Title).Contains(kw)).ToList();

            // 🌟 FIX CRASH: Ép vào 1 dòng để ComboBox không bị lỗi khi rỗng
            if (filtered.Count == 0) filtered.Add(new PredefinedTitle { Title = "Không tìm thấy template...", Group = "Khác" });

            cboTitle.DataSource = filtered; cboTitle.DisplayMember = "Title"; cboTitle.ValueMember = "Title";

            cboTitle.DroppedDown = filtered.Count > 0;
            cboTitle.Text = keyword;
            cboTitle.SelectionStart = Math.Min(pos, cboTitle.Text.Length);
            cboTitle.TextUpdate += CboTitle_TextUpdate;
        }

        private void CboTag_TextUpdate(object sender, EventArgs e)
        {
            if (_tagList == null) return;
            string keyword = cboTag.Text;
            string kw = Form1.ConvertToUnSignStatic(keyword);
            int pos = cboTag.SelectionStart;
            cboTag.TextUpdate -= CboTag_TextUpdate;

            var filtered = string.IsNullOrEmpty(kw) ? _tagList.Take(100).ToList() : _tagList.Where(t => Form1.ConvertToUnSignStatic(t.name).Contains(kw)).Take(100).ToList();

            // 🌟 FIX CRASH: Nhét luôn chữ đang gõ vào danh sách để không bị rỗng
            if (!string.IsNullOrEmpty(keyword) && !filtered.Any(x => x.name.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                filtered.Insert(0, new TagItem { name = keyword, id = keyword });
            else if (filtered.Count == 0) filtered.Add(new TagItem { name = keyword, id = keyword });

            cboTag.DataSource = filtered; cboTag.DisplayMember = "name"; cboTag.ValueMember = "id";

            cboTag.DroppedDown = filtered.Count > 0;
            cboTag.Text = keyword;
            cboTag.SelectionStart = Math.Min(pos, cboTag.Text.Length);
            cboTag.TextUpdate += CboTag_TextUpdate;
        }

        private void CboTypeIssue_TextUpdate(object sender, EventArgs e)
        {
            if (_typeIssueList == null) return;
            string keyword = cboTypeIssue.Text;
            string kw = Form1.ConvertToUnSignStatic(keyword);
            int pos = cboTypeIssue.SelectionStart;
            cboTypeIssue.TextUpdate -= CboTypeIssue_TextUpdate;

            var filtered = string.IsNullOrEmpty(kw) ? _typeIssueList.ToList() : _typeIssueList.Where(t => Form1.ConvertToUnSignStatic(t.Text).Contains(kw)).ToList();

            // 🌟 FIX CRASH: Cảnh báo không tìm thấy nhưng vẫn cho gõ tay
            if (filtered.Count == 0) filtered.Add(new ComboItem { Text = "Không tìm thấy Type Issue..." });
            if (!filtered.Any(x => x.Text.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                filtered.Insert(0, new ComboItem { Text = keyword });

            cboTypeIssue.DataSource = filtered; cboTypeIssue.DisplayMember = "Text"; cboTypeIssue.ValueMember = "Text";

            cboTypeIssue.DroppedDown = filtered.Count > 0;
            cboTypeIssue.Text = keyword;
            cboTypeIssue.SelectionStart = Math.Min(pos, cboTypeIssue.Text.Length);
            cboTypeIssue.TextUpdate += CboTypeIssue_TextUpdate;
        }

        private void CboDesc_TextUpdate(object sender, EventArgs e)
        {
            if (_descList == null) return;
            string keyword = cboDesc.Text;
            string kw = Form1.ConvertToUnSignStatic(keyword);
            int pos = cboDesc.SelectionStart;
            cboDesc.TextUpdate -= CboDesc_TextUpdate;

            var filtered = string.IsNullOrEmpty(kw) ? _descList.ToList() : _descList.Where(t => Form1.ConvertToUnSignStatic(t.Text).Contains(kw)).ToList();

            // 🌟 FIX CRASH: Nhét chữ đang gõ vào
            if (filtered.Count == 0) filtered.Add(new ComboItem { Text = keyword });

            cboDesc.DataSource = filtered; cboDesc.DisplayMember = "Text"; cboDesc.ValueMember = "Text";

            cboDesc.DroppedDown = filtered.Count > 0;
            cboDesc.Text = keyword;
            cboDesc.SelectionStart = Math.Min(pos, cboDesc.Text.Length);
            cboDesc.TextUpdate += CboDesc_TextUpdate;
        }

        private void CboAssignee_TextUpdate(object sender, EventArgs e)
        {
            if (_employeeList == null) return;
            string keyword = cboAssignee.Text;
            string kw = Form1.ConvertToUnSignStatic(keyword);
            int pos = cboAssignee.SelectionStart;
            cboAssignee.TextUpdate -= CboAssignee_TextUpdate;

            var filtered = string.IsNullOrEmpty(kw) ? _employeeList.ToList() : _employeeList.Where(t => Form1.ConvertToUnSignStatic(t.Name).Contains(kw)).ToList();

            // 🌟 FIX CRASH: Nhét tên lạ vào danh sách
            if (!string.IsNullOrEmpty(keyword) && !filtered.Any(x => x.Name.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                filtered.Insert(0, new Employee { Name = keyword, Id = keyword });
            else if (filtered.Count == 0) filtered.Add(new Employee { Name = keyword, Id = keyword });

            cboAssignee.DataSource = filtered; cboAssignee.DisplayMember = "Name"; cboAssignee.ValueMember = "Id";

            cboAssignee.DroppedDown = filtered.Count > 0;
            cboAssignee.Text = keyword;
            cboAssignee.SelectionStart = Math.Min(pos, cboAssignee.Text.Length);
            cboAssignee.TextUpdate += CboAssignee_TextUpdate;
        }

        // ==================== AUTO FILL TỪ TITLE & TAG ====================
        private void AutoFillFromTitle()
        {
            if (_titleList == null || string.IsNullOrEmpty(cboTitle.Text)) return;
            var matched = _titleList.FirstOrDefault(t => t.Title.Equals(cboTitle.Text, StringComparison.OrdinalIgnoreCase));

            if (matched != null)
            {
                // 1. Map Nhóm (Group)
                GroupName = matched.Group;

                // 2. Map Mô tả (Ưu tiên điền nếu trống)
                if (matched.PossibleDescriptions != null && matched.PossibleDescriptions.Count == 1)
                {
                    cboDesc.Text = matched.PossibleDescriptions[0];
                }
                else if (!string.IsNullOrEmpty(matched.DefaultDesc) && string.IsNullOrEmpty(cboDesc.Text))
                {
                    cboDesc.Text = matched.DefaultDesc;
                }

                // 3. Map Type Issue chuẩn xác theo VTI/Khách Lẻ
                bool isVTI = IsCurrentTagVTI();
                string mappedType = isVTI ? matched.TypeIssueVTI : matched.TypeIssueKhachLe;

                if (!string.IsNullOrEmpty(mappedType))
                {
                    cboTypeIssue.Text = mappedType;
                    var typeMatch = _typeIssueList?.FirstOrDefault(t => t.Text.Equals(mappedType, StringComparison.OrdinalIgnoreCase));
                    MainType = typeMatch?.MainType ?? "";
                }
                else
                {
                    // Nếu Sheet không cấu hình Type Issue thì xóa trắng
                    cboTypeIssue.Text = "";
                    MainType = "";
                }
            }
            else
            {
                GroupName = "Khác (Nhập tay)";
                MainType = "";
            }
        }

        // ==================== TIME FORMATTER (Tự động thêm :) ====================
        private void TimeBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ':' && e.KeyChar != (char)Keys.Back)
                e.Handled = true;
        }

        private bool isTimeFormatting = false;
        private void TimeBox_TextChanged(object sender, EventArgs e)
        {
            if (isTimeFormatting) return;
            TextBox txt = sender as TextBox;
            if (txt == null) return;

            isTimeFormatting = true;

            // Chỉ lấy số
            string raw = new string(txt.Text.Where(char.IsDigit).ToArray());
            if (raw.Length > 4) raw = raw.Substring(0, 4);

            string newText = "";
            if (raw.Length > 0)
            {
                // Thêm Giờ
                newText += raw.Substring(0, Math.Min(2, raw.Length));
                // Thêm Phút (có kèm dấu :)
                if (raw.Length >= 3)
                    newText += ":" + raw.Substring(2, Math.Min(2, raw.Length - 2));
            }

            if (txt.Text != newText)
            {
                txt.Text = newText;
                txt.SelectionStart = txt.Text.Length; // Đẩy con trỏ về cuối
            }

            isTimeFormatting = false;
            //FireDataChanged();
        }

        // ==================== BƠM DATASOURCE VÀ CÁC LOGIC KHÁC ====================
        public void SetDataSources(
            List<TagItem> tags, List<CategoryItem> categories, BindingList<PredefinedTitle> titles,
            BindingList<ComboItem> typeIssues, BindingList<ComboItem> techActions, List<Employee> employees)
        {
            _suspendDataChanged = true;
            _tagList = tags; _categoryList = categories; _titleList = titles;
            _typeIssueList = typeIssues; _descList = techActions; _employeeList = employees;

            BindCombo(cboTag, new List<TagItem>(tags), "name", "id");
            BindCombo(cboTitle, titles, "Title", "Title");
            BindCombo(cboTypeIssue, new BindingList<ComboItem>(typeIssues.ToList()), "Text", "Text");
            BindCombo(cboDesc, new BindingList<ComboItem>(techActions.ToList()), "Text", "Text");
            BindCombo(cboAssignee, new List<Employee>(employees), "Name", "Id");
            BindCombo(cboCategory, new List<CategoryItem>(categories), "name", "name");

            cboTag.SelectedIndex = -1; cboTag.Text = "";
            cboTitle.SelectedIndex = -1; cboTitle.Text = "";
            cboTypeIssue.SelectedIndex = -1; cboTypeIssue.Text = "";
            cboDesc.SelectedIndex = -1; cboDesc.Text = "";
            cboAssignee.SelectedIndex = -1; cboAssignee.Text = "";

            RefreshSubCategorySource();
            _suspendDataChanged = false;
        }

        private static void BindCombo(ComboBox cbo, object source, string display, string value)
        {
            cbo.DataSource = null;
            cbo.DataSource = source;
            cbo.DisplayMember = display;
            cbo.ValueMember = value;
        }

        public void RefreshTitleSource(BindingList<PredefinedTitle> titles)
        {
            string old = cboTitle.Text; _titleList = titles;
            BindCombo(cboTitle, titles, "Title", "Title");
            if (!string.IsNullOrEmpty(old)) cboTitle.Text = old; else { cboTitle.SelectedIndex = -1; cboTitle.Text = ""; }
        }

        public void RefreshTypeIssueSource(BindingList<ComboItem> items)
        {
            string old = cboTypeIssue.Text; _typeIssueList = items;
            BindCombo(cboTypeIssue, new BindingList<ComboItem>(items.ToList()), "Text", "Text");
            if (!string.IsNullOrEmpty(old)) cboTypeIssue.Text = old; else { cboTypeIssue.SelectedIndex = -1; cboTypeIssue.Text = ""; }
        }

        public void RefreshTagSource(List<TagItem> tags)
        {
            string oldId = TagId; _tagList = tags;
            BindCombo(cboTag, new List<TagItem>(tags), "name", "id");
            if (!string.IsNullOrEmpty(oldId)) try { cboTag.SelectedValue = oldId; } catch { cboTag.SelectedIndex = -1; cboTag.Text = ""; }
            else { cboTag.SelectedIndex = -1; cboTag.Text = ""; }
        }

        public void RefreshDescSource(BindingList<ComboItem> items)
        {
            string old = cboDesc.Text;
            _descList = items;
            BindCombo(cboDesc, new BindingList<ComboItem>(items.ToList()), "Text", "Text");
            if (!string.IsNullOrEmpty(old)) cboDesc.Text = old; else { cboDesc.SelectedIndex = -1; cboDesc.Text = ""; }
        }

        public void RefreshCategorySource(List<CategoryItem> cats)
        {
            string oldCat = cboCategory.Text; string oldSub = cboSubCategory.Text; _categoryList = cats;
            BindCombo(cboCategory, new List<CategoryItem>(cats), "name", "name");
            if (!string.IsNullOrEmpty(oldCat)) cboCategory.Text = oldCat; else { cboCategory.SelectedIndex = -1; cboCategory.Text = ""; }
            RefreshSubCategorySource();
            if (!string.IsNullOrEmpty(oldSub)) cboSubCategory.Text = oldSub; else { cboSubCategory.SelectedIndex = -1; cboSubCategory.Text = ""; }
        }

        private void RefreshSubCategorySource()
        {
            if (_categoryList == null) return;
            var cat = _categoryList.FirstOrDefault(c => c.name == cboCategory.Text);
            var subs = cat?.types ?? new List<SubCategoryItem>();
            string old = cboSubCategory.Text;
            BindCombo(cboSubCategory, new List<SubCategoryItem>(subs), "name", "name");
            cboSubCategory.Text = old;
        }

        private void UpdateStripColor()
        {
            if (pnlLeftStrip == null || _tagList == null) return;
            var tag = _tagList.FirstOrDefault(t => t.id == TagId);
            if (tag == null) { pnlLeftStrip.BackColor = Color.FromArgb(180, 190, 210); return; }

            string n = tag.name.ToUpper();
            pnlLeftStrip.BackColor = n.Contains("VTI") ? Color.CornflowerBlue
                                   : n.Contains("HLC") ? Color.MediumSeaGreen
                                   : n.Contains("MAIL") ? Color.DarkOrange
                                   : n.Contains("ZALO") ? Color.MediumSlateBlue
                                   : Color.FromArgb(140, 150, 175);
        }

        private void ApplyResult(string text)
        {
            lblResult.Text = text;
            if (text.Contains("✅") || text.Contains("☁️"))
            {
                this.BackColor = Color.FromArgb(235, 255, 240);
                pnlLeftStrip.BackColor = Color.MediumSeaGreen;
                lblResult.ForeColor = Color.DarkGreen;
            }
            else if (text.Contains("❌"))
            {
                this.BackColor = Color.FromArgb(255, 242, 242);
                pnlLeftStrip.BackColor = Color.IndianRed;
                lblResult.ForeColor = Color.DarkRed;
            }
            else if (text.Contains("⏳"))
            {
                this.BackColor = Color.FromArgb(255, 252, 220);
                lblResult.ForeColor = Color.DarkGoldenrod;
            }
            else lblResult.ForeColor = Color.SlateGray;
        }

        public void ClearAfterSuccess()
        {
            _suspendDataChanged = true;
            cboTag.SelectedIndex = -1; cboTag.Text = "";
            cboTitle.SelectedIndex = -1; cboTitle.Text = "";
            cboTypeIssue.SelectedIndex = -1; cboTypeIssue.Text = "";
            cboDesc.SelectedIndex = -1; cboDesc.Text = "";
            txtStartTime.Text = ""; txtEndTime.Text = ""; lblResult.Text = "";
            this.BackColor = Color.White; UpdateStripColor();
            _suspendDataChanged = false;
        }

        public void SetDefaultCategory(string catName, string subCatName)
        {
            if (!string.IsNullOrEmpty(catName)) { cboCategory.Text = catName; RefreshSubCategorySource(); }
            if (!string.IsNullOrEmpty(subCatName)) cboSubCategory.Text = subCatName;
        }

        public void SetDefaultAssignee(string empId)
        {
            if (string.IsNullOrEmpty(empId)) return;
            try { cboAssignee.SelectedValue = empId; } catch { cboAssignee.Text = ""; }
        }

        public DraftTicket ToDraft()
        {
            if (!HasData) return null;
            return new DraftTicket
            {
                Title = TitleText,
                Desc = DescText,
                CatId = CategoryName,
                SubCatId = SubCategoryName,
                TagId = TagId,
                EmpId = AssigneeId,
                TypeIssue = TypeIssueName,   // 🌟 Lưu thêm Type Issue
                StartTime = StartTime,       // 🌟 Lưu thêm Giờ nhận
                EndTime = EndTime            // 🌟 Lưu thêm Giờ xong
            };
        }

        public void FromDraft(DraftTicket d, string defaultCat, string defaultSubCat, string defaultAssigneeId)
        {
            if (d == null) return;
            _suspendDataChanged = true;

            if (!string.IsNullOrEmpty(d.Title) && _titleList != null && !_titleList.Any(t => t.Title.Equals(d.Title, StringComparison.OrdinalIgnoreCase)))
            {
                _titleList.Insert(0, new PredefinedTitle { Group = "Khác (Nhập tay)", Title = d.Title });
                BindCombo(cboTitle, _titleList, "Title", "Title");
            }

            cboTitle.Text = d.Title ?? "";
            cboDesc.Text = d.Desc ?? "";

            // 🌟 Đọc Nháp điền lại Type Issue và Thời Gian
            cboTypeIssue.Text = d.TypeIssue ?? "";
            txtStartTime.Text = d.StartTime ?? "";
            txtEndTime.Text = d.EndTime ?? "";

            cboCategory.Text = !string.IsNullOrEmpty(d.CatId) ? d.CatId : defaultCat; RefreshSubCategorySource();
            cboSubCategory.Text = !string.IsNullOrEmpty(d.SubCatId) ? d.SubCatId : defaultSubCat;

            if (!string.IsNullOrEmpty(d.TagId)) try { cboTag.SelectedValue = d.TagId; UpdateStripColor(); } catch { }
            string empId = !string.IsNullOrEmpty(d.EmpId) ? d.EmpId : defaultAssigneeId;
            if (!string.IsNullOrEmpty(empId)) try { cboAssignee.SelectedValue = empId; } catch { }

            _suspendDataChanged = false;
        }

        public void ApplyTheme(Color panelBack) { if (!IsDone) this.BackColor = panelBack; }
        private void FireDataChanged() { if (!_suspendDataChanged) DataChanged?.Invoke(this, EventArgs.Empty); }
        public TagItem GetSelectedTag() => _tagList?.FirstOrDefault(t => t.id == TagId);
        public SubCategoryItem GetSelectedSubCategory()
        {
            var cat = _categoryList?.FirstOrDefault(c => c.name == CategoryName);
            return cat?.types?.FirstOrDefault(s => s.name == SubCategoryName);
        }
        public TagItem GetTagById(string id) => _tagList?.FirstOrDefault(t => t.id == id);

        private void InitializeComponent()
        {
            pnlLeftStrip = new Panel();
            lblIndex = new Label();
            cboTag = new ComboBox();
            cboTypeIssue = new ComboBox();
            cboTitle = new ComboBox();
            txtStartTime = new TextBox();
            txtEndTime = new TextBox();
            label1 = new Label();
            cboAssignee = new ComboBox();
            btnDelete = new Button();
            cboDesc = new ComboBox();
            cboCategory = new ComboBox();
            cboSubCategory = new ComboBox();
            lblResult = new Label();
            btnSend = new Button();
            btnSuggest = new Button();
            SuspendLayout();
            // 
            // pnlLeftStrip
            // 
            pnlLeftStrip.BackColor = Color.LightGray;
            pnlLeftStrip.Dock = DockStyle.Left;
            pnlLeftStrip.Location = new Point(0, 0);
            pnlLeftStrip.Name = "pnlLeftStrip";
            pnlLeftStrip.Size = new Size(6, 76);
            pnlLeftStrip.TabIndex = 0;
            // 
            // lblIndex
            // 
            lblIndex.AutoSize = true;
            lblIndex.Location = new Point(24, 9);
            lblIndex.Name = "lblIndex";
            lblIndex.Size = new Size(26, 15);
            lblIndex.TabIndex = 1;
            lblIndex.Text = "#01";
            // 
            // cboTag
            // 
            cboTag.FlatStyle = FlatStyle.Flat;
            cboTag.FormattingEnabled = true;
            cboTag.Location = new Point(56, 3);
            cboTag.Name = "cboTag";
            cboTag.Size = new Size(242, 23);
            cboTag.TabIndex = 2;
            // 
            // cboTypeIssue
            // 
            cboTypeIssue.FlatStyle = FlatStyle.Flat;
            cboTypeIssue.FormattingEnabled = true;
            cboTypeIssue.Location = new Point(304, 32);
            cboTypeIssue.Name = "cboTypeIssue";
            cboTypeIssue.Size = new Size(242, 23);
            cboTypeIssue.TabIndex = 3;
            // 
            // cboTitle
            // 
            cboTitle.FlatStyle = FlatStyle.Flat;
            cboTitle.FormattingEnabled = true;
            cboTitle.Location = new Point(304, 3);
            cboTitle.Name = "cboTitle";
            cboTitle.Size = new Size(242, 23);
            cboTitle.TabIndex = 4;
            // 
            // txtStartTime
            // 
            txtStartTime.Location = new Point(552, 3);
            txtStartTime.Name = "txtStartTime";
            txtStartTime.Size = new Size(60, 23);
            txtStartTime.TabIndex = 5;
            // 
            // txtEndTime
            // 
            txtEndTime.Location = new Point(677, 3);
            txtEndTime.Name = "txtEndTime";
            txtEndTime.Size = new Size(60, 23);
            txtEndTime.TabIndex = 6;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(632, 9);
            label1.Name = "label1";
            label1.Size = new Size(17, 15);
            label1.TabIndex = 7;
            label1.Text = "→";
            // 
            // cboAssignee
            // 
            cboAssignee.FlatStyle = FlatStyle.Flat;
            cboAssignee.FormattingEnabled = true;
            cboAssignee.Location = new Point(743, 3);
            cboAssignee.Name = "cboAssignee";
            cboAssignee.Size = new Size(242, 23);
            cboAssignee.TabIndex = 8;
            // 
            // btnDelete
            // 
            btnDelete.FlatAppearance.BorderSize = 0;
            btnDelete.FlatStyle = FlatStyle.Flat;
            btnDelete.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnDelete.ForeColor = Color.Red;
            btnDelete.Location = new Point(1098, 28);
            btnDelete.Name = "btnDelete";
            btnDelete.Size = new Size(20, 20);
            btnDelete.TabIndex = 9;
            btnDelete.Text = "✖";
            btnDelete.UseVisualStyleBackColor = true;
            // 
            // cboDesc
            // 
            cboDesc.FlatStyle = FlatStyle.Flat;
            cboDesc.FormattingEnabled = true;
            cboDesc.Location = new Point(56, 32);
            cboDesc.Name = "cboDesc";
            cboDesc.Size = new Size(242, 23);
            cboDesc.TabIndex = 10;
            // 
            // cboCategory
            // 
            cboCategory.FlatStyle = FlatStyle.Flat;
            cboCategory.FormattingEnabled = true;
            cboCategory.Location = new Point(679, 32);
            cboCategory.Name = "cboCategory";
            cboCategory.Size = new Size(150, 23);
            cboCategory.TabIndex = 11;
            // 
            // cboSubCategory
            // 
            cboSubCategory.FlatStyle = FlatStyle.Flat;
            cboSubCategory.FormattingEnabled = true;
            cboSubCategory.Location = new Point(835, 32);
            cboSubCategory.Name = "cboSubCategory";
            cboSubCategory.Size = new Size(150, 23);
            cboSubCategory.TabIndex = 12;
            // 
            // lblResult
            // 
            lblResult.AutoSize = true;
            lblResult.Location = new Point(1032, 31);
            lblResult.Name = "lblResult";
            lblResult.Size = new Size(60, 15);
            lblResult.TabIndex = 14;
            lblResult.Text = "Trạng thái";
            // 
            // btnSend
            // 
            btnSend.BackColor = Color.ForestGreen;
            btnSend.FlatStyle = FlatStyle.Flat;
            btnSend.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnSend.Location = new Point(1124, 0);
            btnSend.Name = "btnSend";
            btnSend.Size = new Size(76, 76);
            btnSend.TabIndex = 15;
            btnSend.Text = "Tạo phiếu";
            btnSend.UseVisualStyleBackColor = false;
            // 
            // btnSuggest
            // 
            btnSuggest.Cursor = Cursors.Hand;
            btnSuggest.FlatStyle = FlatStyle.Flat;
            btnSuggest.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSuggest.Location = new Point(552, 32);
            btnSuggest.Name = "btnSuggest";
            btnSuggest.Size = new Size(60, 24);
            btnSuggest.TabIndex = 16;
            btnSuggest.Text = "✨Gợi ý";
            btnSuggest.UseVisualStyleBackColor = true;
            btnSuggest.Click += btnSuggest_Click;
            // 
            // TicketCardControl
            // 
            BackColor = Color.White;
            Controls.Add(btnSuggest);
            Controls.Add(btnSend);
            Controls.Add(lblResult);
            Controls.Add(cboSubCategory);
            Controls.Add(cboCategory);
            Controls.Add(cboDesc);
            Controls.Add(btnDelete);
            Controls.Add(cboAssignee);
            Controls.Add(label1);
            Controls.Add(txtEndTime);
            Controls.Add(txtStartTime);
            Controls.Add(cboTitle);
            Controls.Add(cboTypeIssue);
            Controls.Add(cboTag);
            Controls.Add(lblIndex);
            Controls.Add(pnlLeftStrip);
            Name = "TicketCardControl";
            Size = new Size(1200, 76);
            ResumeLayout(false);
            PerformLayout();

        }
        // TÍNH NĂNG 4: DOUBLE CLICK MỞ POPUP NHẬP MÔ TẢ
        private void CboDesc_DoubleClick(object sender, EventArgs e)
        {
            Form frm = new Form { Text = "Nhập Mô tả chi tiết", Width = 500, Height = 300, StartPosition = FormStartPosition.CenterParent, ShowIcon = false };
            TextBox txt = new TextBox { Multiline = true, Dock = DockStyle.Fill, Text = cboDesc.Text, Font = new Font("Segoe UI", 10), ScrollBars = ScrollBars.Vertical };
            Button btnOk = new Button { Text = "LƯU MÔ TẢ", Dock = DockStyle.Bottom, Height = 45, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnOk.Click += (s, ev) => { cboDesc.Text = txt.Text; frm.Close(); FireDataChanged(); };
            frm.Controls.Add(txt); frm.Controls.Add(btnOk);
            frm.ShowDialog();
        }

        // TÍNH NĂNG 5: TỰ ĐỘNG KÉO DÀI DANH SÁCH NẾU CHỮ QUÁ DÀI (CHỐNG TRÀN)
        private void AdjustDropDownWidth(object sender, EventArgs e)
        {
            if (sender is ComboBox cb && cb.Items.Count > 0)
            {
                int maxWidth = cb.Width;
                using (Graphics g = cb.CreateGraphics())
                {
                    foreach (var item in cb.Items)
                    {
                        int width = (int)g.MeasureString(item.ToString(), cb.Font).Width + 30; // +30px cho thanh cuộn
                        if (width > maxWidth) maxWidth = width;
                    }
                }
                cb.DropDownWidth = maxWidth;
            }
        }

        // TÍNH NĂNG 6: RỜI Ô START TIME TỰ ĐỘNG +2 PHÚT VÀO END TIME
        private void TxtStartTime_Leave(object sender, EventArgs e)
        {
            if (txtStartTime.Text.Length == 5 && string.IsNullOrEmpty(txtEndTime.Text))
            {
                if (TimeSpan.TryParseExact(txtStartTime.Text, @"hh\:mm", null, out TimeSpan st))
                {
                    txtEndTime.Text = st.Add(TimeSpan.FromMinutes(2)).ToString(@"hh\:mm");
                    FireDataChanged();
                }
            }
        }
        private void CboTypeIssue_SelectedIndexChanged(object sender, EventArgs e)
        {
            var typeMatch = _typeIssueList?.FirstOrDefault(t => t.Text == cboTypeIssue.Text);
            MainType = typeMatch?.MainType ?? "";
            FireDataChanged();
        }

        private void btnSuggest_Click(object sender, EventArgs e)
        {
            string smartTemplatePath = System.IO.Path.Combine(Application.StartupPath, "smart_templates.json");
            List<SmartTemplate> templates = new List<SmartTemplate>();

            if (System.IO.File.Exists(smartTemplatePath))
            {
                try { templates = System.Text.Json.JsonSerializer.Deserialize<List<SmartTemplate>>(System.IO.File.ReadAllText(smartTemplatePath)); } catch { }
            }

            ContextMenuStrip menu = new ContextMenuStrip();

            if (templates.Count == 0)
            {
                menu.Items.Add(new ToolStripMenuItem("Chưa có dữ liệu...") { Enabled = false });
            }
            else
            {
                var topTemplates = templates.OrderByDescending(t => t.UseCount).Take(7).ToList();
                foreach (var tpl in topTemplates)
                {
                    string menuText = $"🔥 [{tpl.UseCount} lần] {tpl.Title}";
                    string tooltipDesc = tpl.Desc.Length > 50 ? tpl.Desc.Substring(0, 50) + "..." : tpl.Desc;

                    var menuItem = new ToolStripMenuItem(menuText) { ToolTipText = tooltipDesc };
                    menuItem.Click += (s, ev) =>
                    {
                        cboTitle.Text = tpl.Title;
                        cboDesc.Text = tpl.Desc;
                        AutoFillFromTitle();
                        FireDataChanged();
                        cboTag.Focus();
                    };
                    menu.Items.Add(menuItem);
                }
            }

            // 🌟 THÊM NÚT RESET LỊCH SỬ HỌC MÁY
            menu.Items.Add(new ToolStripSeparator()); // Đường kẻ ngang
            var resetItem = new ToolStripMenuItem("🗑️ Xoá lịch sử gợi ý (Reset)") { ForeColor = Color.IndianRed };
            resetItem.Click += (s, ev) =>
            {
                if (MessageBox.Show("Bạn có chắc muốn xoá toàn bộ lịch sử gợi ý không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    if (System.IO.File.Exists(smartTemplatePath)) System.IO.File.Delete(smartTemplatePath);
                    MessageBox.Show("Đã xoá sạch lịch sử!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            menu.Items.Add(resetItem);

            Button btn = sender as Button;
            menu.Show(btn, new Point(0, btn.Height));
        }
        // Hàm này giúp thẻ Card tự đổi màu nền của chính nó
        //public void ApplyTheme(Color backColor)
        //{
        //    if (!IsDone) // Nếu chưa làm xong thì mới đổi màu
        //    {
        //        this.BackColor = backColor;
        //    }
        //}
    }
}