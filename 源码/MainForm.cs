using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CaoXing
{
    internal class UndoAction
    {
        public string Kind;          // add / delete / batchadd
        public Student Student;
        public int ModuleIndex;
        public int Index;
        public RecordItem Record;
        public List<BatchUndoItem> Batch;
    }

    internal class BatchUndoItem
    {
        public Student Student;
        public int ModuleIndex;
        public int Index;
        public RecordItem Record;
    }

    public class MainForm : Form
    {
        private ClassData _data;
        private string _filePath = "";
        private string _startupFile = "";
        private bool _dirty;
        private bool _loadingUi;
        private Student _selected;
        private int _selectedModule;
        private readonly List<UndoAction> _undo = new List<UndoAction>();
        private AppConfig _config;
        private bool _closeConfirmed;

        private SplitContainer _split;
        private Panel _welcome;
        private CardPanel _welcomeCard;
        private TableLayoutPanel _content;
        private CardPanel _cardStudent;
        private CardPanel _cardAdd;
        private CardPanel _cardRecords;

        private DataGridView _gridStudents;
        private DataGridView _gridRecords;
        private TextBox _txtSearch;
        private Label _lblStudentCount;
        private Label _lblStudentName;
        private Label _lblStudentMeta;
        private Label _lblAverage;
        private Label _lblRecordsTitle;
        private Label _lblRecordsScore;
        private Label _lblEmptyRecords;
        private Label _lblNoStudents;
        private Label _lblPreview;
        private Label _lblHeaderSub;
        private AvatarPanel _avatar;
        private FlowLayoutPanel _chipPanel;
        private ComboBox _cmbStudent;
        private ComboBox _cmbModule;
        private ComboBox _cmbAspect;
        private NumericUpDown _numDelta;
        private TextBox _txtItem;
        private Button _btnAdd;
        private Button _btnDelete;
        private Button _btnUndo;
        private Button _btnSave;
        private Button _btnSaveAs;
        private Button _btnImport;
        private Button _btnBatch;
        private Button _btnOpen;
        private Button _btnNew;
        private Button _btnLog;
        private Button _btnHelp;
        private Button _btnTotal;

        private ToolStripStatusLabel _statusFile;
        private ToolStripStatusLabel _statusCount;
        private ToolStripStatusLabel _statusLog;

        public MainForm(string startupFile)
        {
            _startupFile = startupFile ?? "";
            _config = AppConfig.Load();
            InitializeForm();
            BuildHeader();
            BuildStatus();
            BuildBody();
            WireEvents();
            ApplyData(null, "", false);
            if (_config.Width > 0 && _config.Height > 0)
            {
                Rectangle workArea = Screen.PrimaryScreen.WorkingArea;
                Width = Math.Min(_config.Width, workArea.Width);
                Height = Math.Min(_config.Height, workArea.Height);
            }
            if (_config.Maximized) WindowState = FormWindowState.Maximized;
            Shown += delegate
            {
                Theme.SetCue(_txtSearch, "搜索学号 / 姓名");
                Theme.SetCue(_txtItem, "例如：升旗仪式、全勤、迟到……");
                if (!string.IsNullOrEmpty(_startupFile) && File.Exists(_startupFile)) OpenPath(_startupFile);

            };
        }

        private void InitializeForm()
        {
            Text = "班级操行分管理系统";
            StartPosition = FormStartPosition.CenterScreen;
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            int w = Math.Min(1300, wa.Width - 40);
            int h = Math.Min(830, wa.Height - 40);
            if (w < 1080) w = 1080;
            if (h < 680) h = 680;
            MinimumSize = new Size(1080, 680);
            Size = new Size(w, h);
            BackColor = Theme.Page;
            Font = Theme.Font(9.5f);
            KeyPreview = true;
            DoubleBuffered = true;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
        }

        // ---------------- UI 构建 ----------------

        private void BuildHeader()
        {
            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 74;
            header.BackColor = Theme.Navy;
            Controls.Add(header);
            header.BringToFront();

            Label title = new Label();
            title.Text = "班级操行分管理";
            title.AutoSize = true;
            title.ForeColor = Color.White;
            title.Font = Theme.Font(17f, FontStyle.Bold);
            title.Location = new Point(24, 12);
            header.Controls.Add(title);

            _lblHeaderSub = new Label();
            _lblHeaderSub.Text = "每月汇总 · 逐条留痕 · 自动计算";
            _lblHeaderSub.AutoSize = true;
            _lblHeaderSub.ForeColor = Theme.HeaderSub;
            _lblHeaderSub.Font = Theme.Font(9f);
            _lblHeaderSub.Location = new Point(26, 45);
            header.Controls.Add(_lblHeaderSub);

            FlowLayoutPanel right = new FlowLayoutPanel();
            right.Dock = DockStyle.Right;
            right.Width = 704;
            right.FlowDirection = FlowDirection.LeftToRight;
            right.WrapContents = false;
            right.BackColor = Theme.Navy;
            right.Padding = new Padding(0, 20, 16, 0);
            header.Controls.Add(right);

            _btnOpen = HeaderButton("打开表格", 84, false);
            _btnSave = HeaderButton("保存", 62, true);
            _btnSaveAs = HeaderButton("另存为", 76, false);
            _btnImport = HeaderButton("导入学生", 80, false);
            _btnBatch = HeaderButton("批量添加", 80, false);
            _btnNew = HeaderButton("新建班级", 80, false);
            _btnUndo = HeaderButton("撤销", 62, false);
            _btnLog = HeaderButton("日志", 62, false);
            _btnHelp = HeaderButton("帮助", 62, false);
            right.Controls.Add(_btnOpen);
            right.Controls.Add(_btnSave);
            right.Controls.Add(_btnSaveAs);
            right.Controls.Add(_btnImport);
            right.Controls.Add(_btnBatch);
            right.Controls.Add(_btnNew);
            right.Controls.Add(_btnUndo);
            right.Controls.Add(_btnLog);
            right.Controls.Add(_btnHelp);
        }

        private Button HeaderButton(string text, int width, bool primary)
        {
            Button b = new Button();
            b.Text = text;
            b.Width = width;
            b.Height = 34;
            b.Name = "btn_" + text;
            Theme.HeaderButton(b, primary);
            return b;
        }

        private void BuildStatus()
        {
            StatusStrip status = new StatusStrip();
            status.Dock = DockStyle.Bottom;
            status.BackColor = Theme.StatusBg;
            status.SizingGrip = false;
            status.Font = Theme.Font(9f);
            status.Padding = new Padding(12, 0, 12, 0);

            _statusFile = new ToolStripStatusLabel();
            _statusFile.Spring = true;
            _statusFile.TextAlign = ContentAlignment.MiddleLeft;
            _statusFile.ForeColor = Theme.Muted;

            _statusCount = new ToolStripStatusLabel();
            _statusCount.AutoSize = false;
            _statusCount.Width = 150;
            _statusCount.TextAlign = ContentAlignment.MiddleRight;
            _statusCount.ForeColor = Theme.Muted;

            _statusLog = new ToolStripStatusLabel();
            _statusLog.AutoSize = false;
            _statusLog.Width = 330;
            _statusLog.TextAlign = ContentAlignment.MiddleRight;
            _statusLog.ForeColor = Theme.Muted;
            _statusLog.IsLink = true;
            _statusLog.ToolTipText = "点击打开日志文件夹";

            status.Items.Add(_statusFile);
            status.Items.Add(_statusCount);
            status.Items.Add(_statusLog);
            Controls.Add(status);
            _statusLog.Click += delegate { OpenLogFolder(); };
        }

        private void BuildBody()
        {
            _welcome = new Panel();
            _welcome.Dock = DockStyle.Fill;
            _welcome.BackColor = Theme.Page;
            _welcomeCard = new CardPanel();
            _welcomeCard.Size = new Size(680, 340);
            _welcome.Controls.Add(_welcomeCard);
            _welcome.Resize += delegate { CenterWelcome(); };
            Controls.Add(_welcome);

            Label wTitle = new Label();
            wTitle.Text = "欢迎使用班级操行分管理系统";
            wTitle.Font = Theme.Font(20f, FontStyle.Bold);
            wTitle.ForeColor = Theme.Text;
            wTitle.AutoSize = true;
            wTitle.Location = new Point(20, 18);
            _welcomeCard.Controls.Add(wTitle);

            Label wSub = new Label();
            wSub.Text = "打开上个月的表格继续记录，或者新建一个班级。";
            wSub.Font = Theme.Font(10f);
            wSub.ForeColor = Theme.Muted;
            wSub.AutoSize = true;
            wSub.Location = new Point(22, 58);
            _welcomeCard.Controls.Add(wSub);

            Label bullets = new Label();
            bullets.Text = "•  打开 Excel 模板后，原有学生、模块分和明细会自动识别\n" +
                           "•  选择 同学 → 模块 → 方面，填写事项和加减分，一键添加\n" +
                           "•  模块分与学期平均分自动同步，所有更改写入每月独立日志";
            bullets.Font = Theme.Font(10f);
            bullets.ForeColor = Theme.BodyText;
            bullets.AutoSize = true;
            bullets.Location = new Point(24, 96);
            _welcomeCard.Controls.Add(bullets);

            Label hint = new Label();
            hint.Text = "已适配 8 模块 / 5 方面 / 初始 80 分的模板，也可在“新建班级”中自定义。";
            hint.Font = Theme.Font(9f);
            hint.ForeColor = Theme.Muted;
            hint.AutoSize = true;
            hint.Location = new Point(24, 186);
            _welcomeCard.Controls.Add(hint);

            Button wOpen = new Button();
            wOpen.Text = "打开上个月的表格";
            wOpen.Size = new Size(190, 40);
            wOpen.Location = new Point(24, 232);
            Theme.PrimaryButton(wOpen);
            wOpen.Click += delegate { OpenFile(); };
            _welcomeCard.Controls.Add(wOpen);

            Button wNew = new Button();
            wNew.Text = "新建班级";
            wNew.Size = new Size(120, 40);
            wNew.Location = new Point(226, 232);
            Theme.SecondaryButton(wNew);
            wNew.Click += delegate { NewClass(); };
            _welcomeCard.Controls.Add(wNew);

            Button wImport = new Button();
            wImport.Text = "导入学生";
            wImport.Size = new Size(110, 40);
            wImport.Location = new Point(358, 232);
            Theme.SecondaryButton(wImport);
            wImport.Enabled = false;
            wImport.Click += delegate { ImportStudents(); };
            _welcomeCard.Controls.Add(wImport);

            Button wHelp = new Button();
            wHelp.Text = "使用说明";
            wHelp.Size = new Size(110, 40);
            wHelp.Location = new Point(480, 232);
            Theme.SecondaryButton(wHelp);
            wHelp.Click += delegate { OpenHelp(); };
            _welcomeCard.Controls.Add(wHelp);

            _split = new SplitContainer();
            _split.Dock = DockStyle.Fill;
            _split.BackColor = Theme.Border;
            _split.Size = new Size(1200, 700);
            _split.SplitterWidth = 1;
            _split.FixedPanel = FixedPanel.Panel1;
            _split.Panel1MinSize = 300;
            _split.Panel2MinSize = 320;
            _split.SplitterDistance = 372;
            _split.Panel1.BackColor = Color.White;
            _split.Panel1.Padding = new Padding(14, 12, 14, 12);
            _split.Panel2.BackColor = Theme.Page;
            _split.Panel2.Padding = new Padding(16, 14, 16, 14);
            Controls.Add(_split);
            _split.BringToFront();

            BuildStudentPanel();
            BuildContentPanel();
        }

        private void BuildStudentPanel()
        {
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 2;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 82f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.BackColor = Color.White;
            _split.Panel1.Controls.Add(layout);

            Panel top = new Panel();
            top.Dock = DockStyle.Fill;
            top.BackColor = Color.White;
            layout.Controls.Add(top, 0, 0);

            Label title = new Label();
            title.Text = "学生名单";
            title.Font = Theme.Font(12f, FontStyle.Bold);
            title.ForeColor = Theme.Text;
            title.AutoSize = true;
            title.Location = new Point(0, 8);
            top.Controls.Add(title);

            _lblStudentCount = new Label();
            _lblStudentCount.Text = "共 0 人";
            _lblStudentCount.Font = Theme.Font(9f);
            _lblStudentCount.ForeColor = Theme.Muted;
            _lblStudentCount.AutoSize = true;
            _lblStudentCount.Location = new Point(80, 12);
            top.Controls.Add(_lblStudentCount);

            _txtSearch = new TextBox();
            _txtSearch.Location = new Point(0, 44);
            _txtSearch.Height = 30;
            _txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Theme.StyleTextBox(_txtSearch);
            _txtSearch.Width = _split.Panel1.ClientSize.Width - 28;
            top.Controls.Add(_txtSearch);

            Panel gridHost = new Panel();
            gridHost.Dock = DockStyle.Fill;
            gridHost.BackColor = Color.White;
            layout.Controls.Add(gridHost, 0, 1);

            _gridStudents = new DataGridView();
            _gridStudents.Dock = DockStyle.Fill;
            _gridStudents.Name = "gridStudents";
            Theme.StyleGrid(_gridStudents);
            _gridStudents.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _gridStudents.Columns.Add(MakeColumn("序号", 10));
            _gridStudents.Columns.Add(MakeColumn("学号", 36));
            _gridStudents.Columns.Add(MakeColumn("姓名", 26));
            _gridStudents.Columns.Add(MakeColumn("平均分", 28));
            _gridStudents.ColumnHeadersHeight = 40;
            _gridStudents.DefaultCellStyle.Padding = new Padding(4, 0, 0, 0);
            _gridStudents.ColumnHeadersDefaultCellStyle.Padding = new Padding(4, 0, 0, 0);
            gridHost.Controls.Add(_gridStudents);

            _lblNoStudents = new Label();
            _lblNoStudents.Dock = DockStyle.Fill;
            _lblNoStudents.TextAlign = ContentAlignment.MiddleCenter;
            _lblNoStudents.ForeColor = Theme.Muted;
            _lblNoStudents.Font = Theme.Font(10f);
            _lblNoStudents.Text = "还没有学生\r\n请点击顶部“导入学生”";
            _lblNoStudents.BackColor = Color.White;
            _lblNoStudents.Visible = false;
            gridHost.Controls.Add(_lblNoStudents);

            _split.Panel1.Resize += delegate
            {
                _txtSearch.Width = _split.Panel1.ClientSize.Width - 28;
            };
        }

        private static DataGridViewTextBoxColumn MakeColumn(string name, int weight)
        {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = name;
            c.Name = name;
            c.FillWeight = weight;
            c.SortMode = DataGridViewColumnSortMode.NotSortable;
            c.ReadOnly = true;
            return c;
        }

        private void BuildContentPanel()
        {
            _content = new TableLayoutPanel();
            _content.Dock = DockStyle.Fill;
            _content.ColumnCount = 1;
            _content.RowCount = 3;
            _content.BackColor = Theme.Page;
            _content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _content.RowStyles.Add(new RowStyle(SizeType.Absolute, 132f));
            _content.RowStyles.Add(new RowStyle(SizeType.Absolute, 196f));
            _content.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            _content.Visible = false;
            _split.Panel2.Controls.Add(_content);

            BuildStudentCard();
            BuildAddCard();
            BuildRecordsCard();
        }

        private void BuildStudentCard()
        {
            _cardStudent = new CardPanel();
            _cardStudent.Dock = DockStyle.Fill;
            _cardStudent.Margin = new Padding(0, 0, 0, 10);
            _cardStudent.Padding = new Padding(16);
            _content.Controls.Add(_cardStudent, 0, 0);

            _avatar = new AvatarPanel();
            _avatar.Size = new Size(58, 58);
            _avatar.Location = new Point(18, 16);
            _cardStudent.Controls.Add(_avatar);

            _lblStudentName = new Label();
            _lblStudentName.Text = "未选择学生";
            _lblStudentName.Font = Theme.Font(16f, FontStyle.Bold);
            _lblStudentName.ForeColor = Theme.Text;
            _lblStudentName.AutoSize = true;
            _lblStudentName.Location = new Point(92, 14);
            _cardStudent.Controls.Add(_lblStudentName);

            _lblStudentMeta = new Label();
            _lblStudentMeta.Text = "请在左侧学生名单中选择";
            _lblStudentMeta.Font = Theme.Font(9f);
            _lblStudentMeta.ForeColor = Theme.Muted;
            _lblStudentMeta.AutoSize = true;
            _lblStudentMeta.Location = new Point(94, 48);
            _cardStudent.Controls.Add(_lblStudentMeta);

            Panel right = new Panel();
            right.Name = "avgPanel";
            right.Size = new Size(196, 72);
            right.BackColor = Color.White;
            right.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            right.Location = new Point(_cardStudent.ClientSize.Width - 212, 10);
            _cardStudent.Controls.Add(right);

            Label avgTitle = new Label();
            avgTitle.Text = "学期平均分";
            avgTitle.Font = Theme.Font(9f);
            avgTitle.ForeColor = Theme.Muted;
            avgTitle.AutoSize = true;
            avgTitle.Location = new Point(0, 0);
            right.Controls.Add(avgTitle);

            _lblAverage = new Label();
            _lblAverage.Text = "--";
            _lblAverage.Font = Theme.Font(22f, FontStyle.Bold);
            _lblAverage.ForeColor = Theme.Accent;
            _lblAverage.AutoSize = true;
            _lblAverage.Location = new Point(0, 16);
            right.Controls.Add(_lblAverage);

            _btnTotal = new Button();
            _btnTotal.Text = "查看总表";
            _btnTotal.Size = new Size(84, 28);
            _btnTotal.Location = new Point(108, 42);
            Theme.SecondaryButton(_btnTotal);
            _btnTotal.Font = Theme.Font(9f);
            _btnTotal.Click += delegate { ShowTotal(); };
            right.Controls.Add(_btnTotal);

            _chipPanel = new FlowLayoutPanel();
            _chipPanel.Location = new Point(18, 82);
            _chipPanel.Height = 40;
            _chipPanel.Width = _cardStudent.ClientSize.Width - 36;
            _chipPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _chipPanel.WrapContents = false;
            _chipPanel.AutoScroll = true;
            _chipPanel.BackColor = Color.White;
            _cardStudent.Controls.Add(_chipPanel);

            _cardStudent.Resize += delegate
            {
                right.Left = _cardStudent.ClientSize.Width - right.Width - 20;
                _chipPanel.Width = _cardStudent.ClientSize.Width - 36;
            };
        }

        private void BuildAddCard()
        {
            _cardAdd = new CardPanel();
            _cardAdd.Dock = DockStyle.Fill;
            _cardAdd.Margin = new Padding(0, 0, 0, 10);
            _cardAdd.Padding = new Padding(14, 10, 14, 10);
            _content.Controls.Add(_cardAdd, 0, 1);

            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.ColumnCount = 6;
            t.RowCount = 6;
            t.BackColor = Color.White;
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190f));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112f));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 94f));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92f));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 26f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 18f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 32f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 18f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            _cardAdd.Controls.Add(t);

            Label title = new Label();
            title.Text = "添加操行记录";
            title.Font = Theme.Font(12f, FontStyle.Bold);
            title.ForeColor = Theme.Text;
            title.Dock = DockStyle.Fill;
            t.Controls.Add(title, 0, 0);
            t.SetColumnSpan(title, 6);

            string[] labels = { "同学", "模块", "方面", "加减分" };
            for (int i = 0; i < labels.Length; i++)
            {
                Label l = new Label();
                l.Text = labels[i];
                l.Font = Theme.Font(9f);
                l.ForeColor = Theme.Muted;
                l.Dock = DockStyle.Fill;
                t.Controls.Add(l, i, 1);
            }

            _cmbStudent = new ComboBox();
            _cmbStudent.Name = "cmbStudent";
            _cmbStudent.Dock = DockStyle.Fill;
            _cmbStudent.Margin = new Padding(0, 0, 8, 0);
            Theme.StyleCombo(_cmbStudent);
            _cmbStudent.DisplayMember = "Display";
            t.Controls.Add(_cmbStudent, 0, 2);

            _cmbModule = new ComboBox();
            _cmbModule.Name = "cmbModule";
            _cmbModule.Dock = DockStyle.Fill;
            _cmbModule.Margin = new Padding(0, 0, 8, 0);
            Theme.StyleCombo(_cmbModule);
            t.Controls.Add(_cmbModule, 1, 2);

            _cmbAspect = new ComboBox();
            _cmbAspect.Name = "cmbAspect";
            _cmbAspect.Dock = DockStyle.Fill;
            _cmbAspect.Margin = new Padding(0, 0, 8, 0);
            Theme.StyleCombo(_cmbAspect);
            t.Controls.Add(_cmbAspect, 2, 2);

            _numDelta = new NumericUpDown();
            _numDelta.Name = "numDelta";
            _numDelta.Dock = DockStyle.Fill;
            _numDelta.Margin = new Padding(0, 0, 8, 0);
            _numDelta.Minimum = -100m;
            _numDelta.Maximum = 100m;
            _numDelta.DecimalPlaces = 2;
            _numDelta.Increment = 0.5m;
            _numDelta.Value = 1m;
            _numDelta.Font = Theme.Font(10f);
            t.Controls.Add(_numDelta, 3, 2);

            _btnAdd = new Button();
            _btnAdd.Name = "btnAdd";
            _btnAdd.Text = "添加记录";
            _btnAdd.Dock = DockStyle.Fill;
            _btnAdd.Margin = new Padding(4, 0, 0, 0);
            Theme.PrimaryButton(_btnAdd);
            t.Controls.Add(_btnAdd, 4, 2);
            t.SetColumnSpan(_btnAdd, 2);

            Label itemLabel = new Label();
            itemLabel.Text = "事项";
            itemLabel.Font = Theme.Font(9f);
            itemLabel.ForeColor = Theme.Muted;
            itemLabel.Dock = DockStyle.Fill;
            t.Controls.Add(itemLabel, 0, 3);
            t.SetColumnSpan(itemLabel, 6);

            _txtItem = new TextBox();
            _txtItem.Name = "txtItem";
            _txtItem.Dock = DockStyle.Fill;
            _txtItem.Margin = new Padding(0, 0, 0, 0);
            Theme.StyleTextBox(_txtItem);
            t.Controls.Add(_txtItem, 0, 4);
            t.SetColumnSpan(_txtItem, 6);

            _lblPreview = new Label();
            _lblPreview.Text = "填写事项并输入加减分后，这里会显示添加后的分数变化。";
            _lblPreview.Font = Theme.Font(9f);
            _lblPreview.ForeColor = Theme.Muted;
            _lblPreview.Dock = DockStyle.Fill;
            _lblPreview.TextAlign = ContentAlignment.MiddleLeft;
            t.Controls.Add(_lblPreview, 0, 5);
            t.SetColumnSpan(_lblPreview, 6);
        }

        private void BuildRecordsCard()
        {
            _cardRecords = new CardPanel();
            _cardRecords.Dock = DockStyle.Fill;
            _cardRecords.Margin = new Padding(0);
            _cardRecords.Padding = new Padding(14, 10, 14, 12);
            _content.Controls.Add(_cardRecords, 0, 2);

            TableLayoutPanel header = new TableLayoutPanel();
            header.Dock = DockStyle.Top;
            header.Height = 36;
            header.BackColor = Color.White;
            header.ColumnCount = 3;
            header.RowCount = 1;
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170f));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 98f));
            _cardRecords.Controls.Add(header);

            _lblRecordsTitle = new Label();
            _lblRecordsTitle.Text = "模块记录";
            _lblRecordsTitle.Font = Theme.Font(12f, FontStyle.Bold);
            _lblRecordsTitle.ForeColor = Theme.Text;
            _lblRecordsTitle.Dock = DockStyle.Fill;
            _lblRecordsTitle.TextAlign = ContentAlignment.MiddleLeft;
            header.Controls.Add(_lblRecordsTitle, 0, 0);

            _lblRecordsScore = new Label();
            _lblRecordsScore.Text = "";
            _lblRecordsScore.Font = Theme.Font(9f);
            _lblRecordsScore.ForeColor = Theme.Muted;
            _lblRecordsScore.Dock = DockStyle.Fill;
            _lblRecordsScore.TextAlign = ContentAlignment.MiddleLeft;
            header.Controls.Add(_lblRecordsScore, 1, 0);

            _btnDelete = new Button();
            _btnDelete.Text = "删除选中";
            _btnDelete.Dock = DockStyle.Fill;
            _btnDelete.Margin = new Padding(0, 3, 0, 3);
            _btnDelete.Enabled = false;
            Theme.SecondaryButton(_btnDelete);
            _btnDelete.Font = Theme.Font(9f);
            header.Controls.Add(_btnDelete, 2, 0);

            _gridRecords = new DataGridView();
            _gridRecords.Name = "gridRecords";
            _gridRecords.Dock = DockStyle.Fill;
            Theme.StyleGrid(_gridRecords);
            _gridRecords.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _gridRecords.Columns.Add(MakeColumn("方面", 20));
            _gridRecords.Columns.Add(MakeColumn("事项", 52));
            _gridRecords.Columns.Add(MakeColumn("加减分", 14));
            _gridRecords.Columns.Add(MakeColumn("时间", 24));
            _gridRecords.AllowUserToResizeRows = false;
            _cardRecords.Controls.Add(_gridRecords);

            _lblEmptyRecords = new Label();
            _lblEmptyRecords.Dock = DockStyle.Fill;
            _lblEmptyRecords.TextAlign = ContentAlignment.MiddleCenter;
            _lblEmptyRecords.Font = Theme.Font(10f);
            _lblEmptyRecords.ForeColor = Theme.Muted;
            _lblEmptyRecords.BackColor = Color.White;
            _lblEmptyRecords.Text = "该模块还没有记录\r\n添加后会显示在这里";
            _lblEmptyRecords.Visible = false;
            _cardRecords.Controls.Add(_lblEmptyRecords);
        }

        private void CenterWelcome()
        {
            if (_welcomeCard == null) return;
            _welcomeCard.Left = Math.Max(0, (_welcome.ClientSize.Width - _welcomeCard.Width) / 2);
            _welcomeCard.Top = Math.Max(0, (_welcome.ClientSize.Height - _welcomeCard.Height) / 2);
        }

        private void WireEvents()
        {
            _btnOpen.Click += delegate { OpenFile(); };
            _btnSave.Click += delegate { SaveCurrent(false); };
            _btnSaveAs.Click += delegate { SaveCurrent(true); };
            _btnImport.Click += delegate { ImportStudents(); };
            _btnBatch.Click += delegate { BatchAdd(); };
            _btnNew.Click += delegate { NewClass(); };
            _btnUndo.Click += delegate { Undo(); };
            _btnLog.Click += delegate { OpenLogFolder(); };
            _btnHelp.Click += delegate { OpenHelp(); };
            _btnAdd.Click += delegate { AddRecord(); };
            _btnDelete.Click += delegate { DeleteRecord(); };
            _btnTotal.Click += delegate { ShowTotal(); };

            _gridStudents.SelectionChanged += delegate
            {
                if (_loadingUi) return;
                if (_gridStudents.SelectedRows.Count == 0) return;
                Student s = _gridStudents.SelectedRows[0].Tag as Student;
                if (s == null) return;
                if (_selected == s) return;
                SelectStudent(s, true, false);
            };

            _cmbStudent.SelectedIndexChanged += delegate
            {
                if (_loadingUi) return;
                Student s = _cmbStudent.SelectedItem as Student;
                if (s == null) return;
                if (_selected == s) return;
                SelectStudent(s, false, true);
            };

            _cmbModule.SelectedIndexChanged += delegate
            {
                if (_loadingUi) return;
                _selectedModule = _cmbModule.SelectedIndex;
                RebuildAspectCombo();
                UpdateRecords();
                UpdateChips();
                UpdatePreview();
            };

            _cmbAspect.SelectedIndexChanged += delegate
            {
                if (_loadingUi) return;
                UpdatePreview();
            };

            _numDelta.ValueChanged += delegate { UpdatePreview(); };
            _txtItem.TextChanged += delegate { UpdatePreview(); };
            _txtItem.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    AddRecord();
                }
            };

            _gridRecords.CellFormatting += delegate(object sender, DataGridViewCellFormattingEventArgs e)
            {
                if (e.RowIndex < 0) return;
                if (e.ColumnIndex == 2)
                {
                    string text = Convert.ToString(e.Value);
                    if (text.StartsWith("-")) e.CellStyle.ForeColor = Theme.Red;
                    else if (text.StartsWith("+")) e.CellStyle.ForeColor = Theme.Green;
                    e.CellStyle.Font = Theme.Font(9.5f, FontStyle.Bold);
                }
            };

            _gridRecords.SelectionChanged += delegate
            {
                _btnDelete.Enabled = _gridRecords.SelectedRows.Count > 0;
            };

            _txtSearch.TextChanged += delegate { RefreshStudentGrid(); };

            KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.S)
                {
                    e.SuppressKeyPress = true;
                    SaveCurrent(false);
                }
                else if (e.Control && e.KeyCode == Keys.O)
                {
                    e.SuppressKeyPress = true;
                    OpenFile();
                }
                else if (e.Control && e.KeyCode == Keys.N)
                {
                    e.SuppressKeyPress = true;
                    NewClass();
                }
                else if (e.Control && e.KeyCode == Keys.Z)
                {
                    e.SuppressKeyPress = true;
                    Undo();
                }
            };

            FormClosing += MainForm_FormClosing;
        }

        // ---------------- 数据与界面状态 ----------------

        private void ApplyData(ClassData data, string path, bool dirty)
        {
            _data = data;
            _filePath = path ?? "";
            _dirty = dirty;
            _selected = null;
            _selectedModule = 0;
            _undo.Clear();

            _loadingUi = true;
            try
            {
                _cmbStudent.Items.Clear();
                _cmbModule.Items.Clear();
                _cmbAspect.Items.Clear();
                _gridStudents.Rows.Clear();
                _gridRecords.Rows.Clear();

                if (_data != null)
                {
                    for (int i = 0; i < _data.Students.Count; i++) _cmbStudent.Items.Add(_data.Students[i]);
                    for (int i = 0; i < _data.Modules.Count; i++) _cmbModule.Items.Add(_data.Modules[i].Name);
                    if (_cmbModule.Items.Count > 0) _cmbModule.SelectedIndex = 0;
                    if (_cmbStudent.Items.Count > 0) _cmbStudent.SelectedIndex = 0;
                }
                _loadingUi = false;
                RefreshStudentGrid();
                if (_data != null && _data.Students.Count > 0)
                {
                    SelectStudent(_data.Students[0], true, true);
                }
                else
                {
                    _selected = null;
                    UpdateStudentCard();
                    UpdateRecords();
                }
            }
            finally
            {
                _loadingUi = false;
            }

            bool hasData = _data != null;
            _split.Visible = hasData;
            _content.Visible = hasData;
            _welcome.Visible = !hasData;
            if (!hasData) CenterWelcome();

            if (_data != null && !string.IsNullOrEmpty(_data.ClassName))
            {
                // keep class name
            }
            UpdateChrome();
            UpdateButtons();
            if (!hasData) _lblHeaderSub.Text = "每月汇总 · 逐条留痕 · 自动计算";
        }

        private void UpdateChrome()
        {
            string title = "班级操行分管理系统";
            if (_data != null)
            {
                if (!string.IsNullOrEmpty(_data.ClassName)) title += " - " + _data.ClassName;
                if (_dirty) title += " *";
            }
            Text = title;

            if (_data == null)
            {
                _lblHeaderSub.Text = "每月汇总 · 逐条留痕 · 自动计算";
                _statusFile.Text = "尚未打开表格";
                _statusCount.Text = "";
                _statusLog.Text = "日志：" + CxLog.MonthFileName(DateTime.Now);
                return;
            }
            _lblHeaderSub.Text = string.Format("{0} · {1} 人 · {2} 个模块 · 初始 {3} 分",
                _data.ClassName, _data.Students.Count, _data.ModuleCount, Format.Decimal(_data.InitialScore));
            _statusFile.Text = string.IsNullOrEmpty(_filePath) ? "尚未保存到表格（Ctrl+S 保存）" : _filePath;
            _statusCount.Text = string.Format("{0} 人 / {1} 模块", _data.Students.Count, _data.ModuleCount);
            _statusLog.Text = "日志：" + CxLog.MonthFileName(DateTime.Now);
            _statusLog.ToolTipText = CurrentLogDir();
        }

        private static string ShortPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return "";
            if (path.Length <= 46) return path;
            return "…" + path.Substring(path.Length - 45);
        }

        private void UpdateButtons()
        {
            bool hasData = _data != null;
            _btnSave.Enabled = hasData;
            _btnSaveAs.Enabled = hasData;
            _btnImport.Enabled = hasData;
            _btnBatch.Enabled = hasData;
            _btnUndo.Enabled = hasData && _undo.Count > 0;
            _btnAdd.Enabled = hasData && _selected != null;
            _btnTotal.Enabled = hasData && _selected != null;
            _btnDelete.Enabled = hasData && _gridRecords.SelectedRows.Count > 0;
        }

        private void RefreshStudentGrid()
        {
            if (_gridStudents == null) return;
            string filter = _txtSearch == null ? "" : _txtSearch.Text.Trim();
            Student keep = _selected;
            _loadingUi = true;
            try
            {
                _gridStudents.Rows.Clear();
                if (_data != null)
                {
                    for (int i = 0; i < _data.Students.Count; i++)
                    {
                        Student s = _data.Students[i];
                        if (filter.Length > 0 &&
                            s.No.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0 &&
                            s.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0) continue;
                        int idx = _gridStudents.Rows.Add(s.Seq, s.No, s.Name, Format.Decimal(_data.GetAverage(s)));
                        _gridStudents.Rows[idx].Tag = s;
                    }
                }
                _gridStudents.ClearSelection();
                if (keep != null)
                {
                    for (int i = 0; i < _gridStudents.Rows.Count; i++)
                    {
                        if (object.ReferenceEquals(_gridStudents.Rows[i].Tag, keep))
                        {
                            _gridStudents.Rows[i].Selected = true;
                            break;
                        }
                    }
                }
                if (_gridStudents.SelectedRows.Count == 0 && _gridStudents.Rows.Count > 0)
                    _gridStudents.Rows[0].Selected = true;
            }
            finally
            {
                _loadingUi = false;
            }
            _lblStudentCount.Text = string.Format("共 {0} 人", _data == null ? 0 : _data.Students.Count);
            bool empty = _gridStudents.Rows.Count == 0;
            _lblNoStudents.Visible = empty;
            _gridStudents.Visible = !empty;
            if (empty)
            {
                _lblNoStudents.Text = (_data == null || _data.Students.Count == 0)
                    ? "还没有学生\r\n请点击顶部“导入学生”"
                    : "没有匹配的学生\r\n换个关键词试试";
                _lblNoStudents.BringToFront();
            }
        }

        private void SelectStudent(Student s, bool syncCombo, bool syncGrid)
        {
            if (_loadingUi) return;
            _loadingUi = true;
            try
            {
                _selected = s;
                if (syncCombo && s != null && _cmbStudent.Items.Contains(s))
                    _cmbStudent.SelectedItem = s;
                if (syncGrid)
                {
                    _gridStudents.ClearSelection();
                    for (int i = 0; i < _gridStudents.Rows.Count; i++)
                    {
                        if (object.ReferenceEquals(_gridStudents.Rows[i].Tag, s))
                        {
                            _gridStudents.CurrentCell = _gridStudents.Rows[i].Cells[0];
                            _gridStudents.Rows[i].Selected = true;
                            _gridStudents.FirstDisplayedScrollingRowIndex = i;
                            break;
                        }
                    }
                }
                if (_cmbModule.Items.Count > 0 && _cmbModule.SelectedIndex < 0) _cmbModule.SelectedIndex = 0;
                _selectedModule = _cmbModule.SelectedIndex;
                if (_selectedModule < 0) _selectedModule = 0;
                UpdateStudentCard();
                RebuildAspectCombo();
                UpdateRecords();
                UpdateChips();
                UpdatePreview();
                UpdateButtons();

            }
            finally
            {
                _loadingUi = false;
            }
        }

        private void UpdateStudentCard()
        {
            if (_data == null)
            {
                _lblStudentName.Text = "未选择学生";
                _lblStudentMeta.Text = "请先打开表格或新建班级";
                _lblAverage.Text = "--";
                _avatar.Initial = "?";
                _chipPanel.Controls.Clear();
                return;
            }
            if (_selected == null)
            {
                _lblStudentName.Text = "未选择学生";
                _lblStudentMeta.Text = "请在左侧学生名单中选择";
                _lblAverage.Text = "--";
                _avatar.Initial = "?";
                _chipPanel.Controls.Clear();
                return;
            }
            _lblStudentName.Text = _selected.Name;
            _lblStudentMeta.Text = string.Format("学号 {0}　｜　模块总分 {1}　｜　{2} 个模块",
                string.IsNullOrEmpty(_selected.No) ? "—" : _selected.No,
                Format.Decimal(_data.GetModuleSum(_selected)),
                _data.ModuleCount);
            _lblAverage.Text = Format.Decimal(_data.GetAverage(_selected));
            _avatar.Initial = _selected.Name;
        }

        private void UpdateChips()
        {
            if (_chipPanel == null) return;
            _chipPanel.SuspendLayout();
            _chipPanel.Controls.Clear();
            if (_data != null && _selected != null)
            {
                for (int i = 0; i < _data.ModuleCount; i++)
                {
                    int moduleIndex = i;
                    Button b = new Button();
                    b.Text = string.Format("{0} {1}分", _data.Modules[i].Name, Format.Decimal2(_data.GetModuleScore(_selected, i)));
                    b.Size = new Size(98, 40);
                    b.Margin = new Padding(0, 0, 6, 0);
                    b.FlatStyle = FlatStyle.Flat;
                    b.FlatAppearance.BorderSize = 1;
                    b.FlatAppearance.BorderColor = i == _selectedModule ? Theme.Accent : Theme.Border;
                    b.BackColor = i == _selectedModule ? Theme.Accent : Theme.ChipBg;
                    b.ForeColor = i == _selectedModule ? Color.White : Theme.Text;
                    b.Font = Theme.Font(8.5f);
                    b.UseVisualStyleBackColor = false;
                    b.Cursor = Cursors.Hand;
                    b.Click += delegate
                    {
                        if (moduleIndex < _cmbModule.Items.Count) _cmbModule.SelectedIndex = moduleIndex;
                    };
                    _chipPanel.Controls.Add(b);
                }
            }
            _chipPanel.ResumeLayout();
        }

        private void RebuildAspectCombo()
        {
            if (_loadingUi) return;
            _loadingUi = true;
            try
            {
                int old = _cmbAspect.SelectedIndex;
                _cmbAspect.Items.Clear();
                if (_data != null && _selectedModule >= 0 && _selectedModule < _data.ModuleCount)
                {
                    int count = _data.Modules[_selectedModule].AspectCount;
                    for (int i = 1; i <= count; i++) _cmbAspect.Items.Add(ClassData.AspectLabel(i));
                    if (_cmbAspect.Items.Count > 0)
                        _cmbAspect.SelectedIndex = Math.Min(Math.Max(old, 0), _cmbAspect.Items.Count - 1);
                }
            }
            finally
            {
                _loadingUi = false;
            }
        }

        private void UpdateRecords()
        {
            if (_gridRecords == null) return;
            bool hasModule = _data != null && _selected != null && _selectedModule >= 0 && _selectedModule < _data.ModuleCount;
            _gridRecords.Rows.Clear();
            int count = 0;
            if (hasModule)
            {
                ModuleData m = _selected.Modules[_selectedModule];
                count = m.Records.Count;
                for (int i = 0; i < m.Records.Count; i++)
                {
                    RecordItem r = m.Records[i];
                    string aspect = string.IsNullOrEmpty(r.AspectLabel) ? "未标注" : r.AspectLabel;
                    string item = string.IsNullOrEmpty(r.Item) ? "（未填写事项）" : r.Item;
                    string time = r.Time.HasValue ? r.Time.Value.ToString("MM-dd HH:mm") : "—";
                    int idx = _gridRecords.Rows.Add(aspect, item, r.DeltaText, time);
                    _gridRecords.Rows[idx].Tag = i;
                }
                _lblRecordsTitle.Text = _data.Modules[_selectedModule].Name + " 记录";
                _lblRecordsScore.Text = string.Format("当前 {0} 分　·　初始 {1} 分　·　{2} 条",
                    Format.Decimal2(_data.GetModuleScore(_selected, _selectedModule)),
                    Format.Decimal(_data.InitialScore), count);
            }
            else
            {
                _lblRecordsTitle.Text = "模块记录";
                _lblRecordsScore.Text = "";
            }
            _lblEmptyRecords.Text = hasModule
                ? "该模块还没有记录\r\n填写上方表单并点击“添加记录”"
                : "请先选择学生";
            _lblEmptyRecords.Visible = count == 0;
            if (_lblEmptyRecords.Visible) _lblEmptyRecords.BringToFront();
            else _gridRecords.BringToFront();
            _btnDelete.Enabled = count > 0 && _gridRecords.SelectedRows.Count > 0;
            if (count > 0) _gridRecords.ClearSelection();
        }

        private void UpdatePreview()
        {
            if (_lblPreview == null) return;
            if (_data == null || _selected == null || _selectedModule < 0 || _selectedModule >= _data.ModuleCount)
            {
                _lblPreview.Text = "填写事项并输入加减分后，这里会显示添加后的分数变化。";
                _lblPreview.ForeColor = Theme.Muted;
                return;
            }
            decimal before = _data.GetModuleScore(_selected, _selectedModule);
            decimal delta = _numDelta.Value;
            decimal after = before + delta;
            decimal avgBefore = _data.GetAverage(_selected);
            decimal sum = _data.GetModuleSum(_selected);
            decimal avgAfter = _data.ModuleCount == 0 ? 0m : (sum - before + after) / _data.ModuleCount;
            _lblPreview.Text = string.Format("添加后：{0} {1} → {2} 分　｜　平均分 {3} → {4}",
                _data.Modules[_selectedModule].Name,
                Format.Decimal2(before), Format.Decimal2(after),
                Format.Decimal(avgBefore), Format.Decimal(avgAfter));
            _lblPreview.ForeColor = delta < 0 ? Theme.Red : (delta > 0 ? Theme.Green : Theme.Muted);
        }

        // ---------------- 增删记录 ----------------

        private void AddRecord()
        {
            if (_data == null)
            {
                MessageBox.Show(this, "请先打开表格或新建班级。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_selected == null)
            {
                MessageBox.Show(this, "请先在左侧选择一个学生。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_selectedModule < 0 || _selectedModule >= _data.ModuleCount) return;
            string item = _txtItem.Text.Trim();
            if (item.Length == 0)
            {
                _txtItem.Focus();
                MessageBox.Show(this, "请填写事项内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            decimal delta = _numDelta.Value;
            if (delta == 0m)
            {
                MessageBox.Show(this, "加减分不能为 0。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int aspect = _cmbAspect.SelectedIndex + 1;
            if (aspect <= 0) aspect = 1;

            decimal before = _data.GetModuleScore(_selected, _selectedModule);
            decimal avgBefore = _data.GetAverage(_selected);
            RecordItem r = _data.AddRecord(_selected, _selectedModule, aspect, item, delta, DateTime.Now);
            decimal after = _data.GetModuleScore(_selected, _selectedModule);
            decimal avgAfter = _data.GetAverage(_selected);

            UndoAction u = new UndoAction();
            u.Kind = "add";
            u.Student = _selected;
            u.ModuleIndex = _selectedModule;
            u.Index = _selected.Modules[_selectedModule].Records.Count - 1;
            u.Record = r;
            _undo.Add(u);
            _dirty = true;

            Log("添加记录",
                "学号=" + _selected.No,
                "姓名=" + _selected.Name,
                "模块=" + _data.Modules[_selectedModule].Name,
                "方面=" + ClassData.AspectLabel(aspect),
                "事项=" + item,
                "加减分=" + Format.ScoreWithSign(delta),
                "模块分=" + Format.Decimal2(before) + "→" + Format.Decimal2(after),
                "平均分=" + Format.Decimal(avgBefore) + "→" + Format.Decimal(avgAfter));

            _txtItem.Clear();
            _numDelta.Value = 1m;
            _txtItem.Focus();
            UpdateAfterChange(true);
            _statusFile.Text = string.Format("已添加：{0} · {1} · {2}（{3} 分）",
                _selected.Name, _data.Modules[_selectedModule].Name, item, Format.ScoreWithSign(delta));
        }

        private void DeleteRecord()
        {
            if (_data == null || _selected == null) return;
            if (_gridRecords.SelectedRows.Count == 0) return;
            int index;
            if (!int.TryParse(Convert.ToString(_gridRecords.SelectedRows[0].Tag), out index)) return;
            ModuleData m = _selected.Modules[_selectedModule];
            if (index < 0 || index >= m.Records.Count) return;
            RecordItem r = m.Records[index];
            string aspect = string.IsNullOrEmpty(r.AspectLabel) ? "未标注" : r.AspectLabel;
            string item = string.IsNullOrEmpty(r.Item) ? "（未填写事项）" : r.Item;
            DialogResult confirm = MessageBox.Show(this,
                string.Format("确定删除这条记录吗？\r\n\r\n{0}　{1}　（{2} 分）", aspect, item, Format.Decimal(r.Delta)),
                "删除记录", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            decimal before = _data.GetModuleScore(_selected, _selectedModule);
            decimal avgBefore = _data.GetAverage(_selected);
            m.Records.RemoveAt(index);
            m.Dirty = true;
            m.Score = _data.GetModuleScore(_selected, _selectedModule);
            decimal after = m.Score;
            decimal avgAfter = _data.GetAverage(_selected);
            _dirty = true;

            UndoAction u = new UndoAction();
            u.Kind = "delete";
            u.Student = _selected;
            u.ModuleIndex = _selectedModule;
            u.Index = index;
            u.Record = r;
            _undo.Add(u);

            Log("删除记录",
                "学号=" + _selected.No,
                "姓名=" + _selected.Name,
                "模块=" + _data.Modules[_selectedModule].Name,
                "方面=" + aspect,
                "事项=" + item,
                "加减分=" + Format.ScoreWithSign(r.Delta),
                "模块分=" + Format.Decimal2(before) + "→" + Format.Decimal2(after),
                "平均分=" + Format.Decimal(avgBefore) + "→" + Format.Decimal(avgAfter));

            UpdateAfterChange(false);
        }

        private void Undo()
        {
            if (_data == null || _undo.Count == 0) return;
            UndoAction u = _undo[_undo.Count - 1];
            _undo.RemoveAt(_undo.Count - 1);

            if (u.Kind == "batchadd")
            {
                if (u.Batch == null || u.Batch.Count == 0)
                {
                    UpdateButtons();
                    return;
                }
                for (int i = u.Batch.Count - 1; i >= 0; i--)
                {
                    BatchUndoItem bi = u.Batch[i];
                    if (bi.Student == null) continue;
                    ModuleData bm = bi.Student.Modules[bi.ModuleIndex];
                    if (bi.Index >= 0 && bi.Index < bm.Records.Count && object.ReferenceEquals(bm.Records[bi.Index], bi.Record))
                        bm.Records.RemoveAt(bi.Index);
                    else
                        bm.Records.Remove(bi.Record);
                    bm.Dirty = true;
                    bm.Score = _data.GetModuleScore(bi.Student, bi.ModuleIndex);
                    Log("撤销批量添加记录",
                        "学号=" + bi.Student.No,
                        "姓名=" + bi.Student.Name,
                        "模块=" + _data.Modules[bi.ModuleIndex].Name,
                        "事项=" + (bi.Record == null ? "" : bi.Record.Item),
                        "加减分=" + (bi.Record == null ? "" : Format.ScoreWithSign(bi.Record.Delta)));
                }
                Log("撤销批量添加",
                    "人数=" + u.Batch.Count,
                    "模块=" + _data.Modules[u.Batch[0].ModuleIndex].Name,
                    "学生=" + DescribeStudentsFromBatch(u.Batch));
                _dirty = true;
                Student focus = u.Batch[0].Student;
                if (focus == null)
                {
                    UpdateChrome();
                    UpdateButtons();
                    return;
                }
                _loadingUi = true;
                try
                {
                    _selectedModule = u.Batch[0].ModuleIndex;
                    if (_cmbModule.Items.Count > _selectedModule) _cmbModule.SelectedIndex = _selectedModule;
                }
                finally
                {
                    _loadingUi = false;
                }
                RefreshStudentGrid();
                SelectStudent(focus, true, true);
                UpdateChrome();
                UpdateButtons();
                _statusFile.Text = "已撤销上一步批量添加";
                return;
            }

            ModuleData m = u.Student.Modules[u.ModuleIndex];
            if (u.Kind == "add")
            {
                if (u.Index >= 0 && u.Index < m.Records.Count && object.ReferenceEquals(m.Records[u.Index], u.Record))
                    m.Records.RemoveAt(u.Index);
                else
                    m.Records.Remove(u.Record);
                Log("撤销添加", "学号=" + u.Student.No, "姓名=" + u.Student.Name,
                    "模块=" + _data.Modules[u.ModuleIndex].Name, "事项=" + u.Record.Item,
                    "加减分=" + Format.ScoreWithSign(u.Record.Delta));
            }
            else
            {
                int index = Math.Min(Math.Max(u.Index, 0), m.Records.Count);
                m.Records.Insert(index, u.Record);
                Log("撤销删除", "学号=" + u.Student.No, "姓名=" + u.Student.Name,
                    "模块=" + _data.Modules[u.ModuleIndex].Name, "事项=" + u.Record.Item,
                    "加减分=" + Format.ScoreWithSign(u.Record.Delta));
            }
            m.Dirty = true;
            m.Score = _data.GetModuleScore(u.Student, u.ModuleIndex);
            _dirty = true;
            if (!object.ReferenceEquals(u.Student, _selected))
            {
                SelectStudent(u.Student, true, true);
            }
            _selectedModule = u.ModuleIndex;
            if (_cmbModule.Items.Count > u.ModuleIndex) _cmbModule.SelectedIndex = u.ModuleIndex;
            UpdateAfterChange(false);
            _statusFile.Text = "已撤销上一步操作";
        }

        private static string DescribeStudentsFromBatch(List<BatchUndoItem> list)
        {
            List<Student> students = new List<Student>();
            for (int i = 0; i < list.Count; i++)
                if (list[i].Student != null) students.Add(list[i].Student);
            return DescribeStudents(students);
        }

        private void UpdateAfterChange(bool selectLast)
        {
            RefreshStudentRow(_selected);
            UpdateStudentCard();
            UpdateChips();
            UpdateRecords();
            UpdatePreview();
            UpdateChrome();
            UpdateButtons();
            if (selectLast && _gridRecords.Rows.Count > 0)
            {
                _gridRecords.ClearSelection();
                _gridRecords.Rows[_gridRecords.Rows.Count - 1].Selected = true;
            }
        }

        private void RefreshStudentRow(Student s)
        {
            if (s == null || _data == null) return;
            for (int i = 0; i < _gridStudents.Rows.Count; i++)
            {
                if (object.ReferenceEquals(_gridStudents.Rows[i].Tag, s))
                {
                    _gridStudents.Rows[i].Cells[3].Value = Format.Decimal(_data.GetAverage(s));
                    break;
                }
            }
        }

        // ---------------- 文件操作 ----------------

        private bool SaveCurrent(bool forceSaveAs)
        {
            if (_data == null) return false;
            string target = _filePath;
            if (forceSaveAs || string.IsNullOrEmpty(target))
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.Title = "保存操行分表格";
                dlg.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                dlg.DefaultExt = "xlsx";
                dlg.AddExtension = true;
                dlg.FileName = SaveService.SuggestFileName(_data);
                string dir = "";
                if (!string.IsNullOrEmpty(_filePath)) dir = Path.GetDirectoryName(_filePath);
                if (string.IsNullOrEmpty(dir) && !string.IsNullOrEmpty(_config.LastFolder) && Directory.Exists(_config.LastFolder))
                    dir = _config.LastFolder;
                if (string.IsNullOrEmpty(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                dlg.InitialDirectory = dir;
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                target = dlg.FileName;
                _config.LastFolder = Path.GetDirectoryName(target);
            }
            try
            {
                Cursor = Cursors.WaitCursor;
                bool useTemplate = _data.FromTemplate && !string.IsNullOrEmpty(_data.SourceFile) && File.Exists(_data.SourceFile);
                string templateSource = _data.SourceFile;
                string backup = SaveService.Save(_data, target, templateSource, _data.RosterDirty, useTemplate);
                _filePath = target;
                _data.SourceFile = target;
                _data.FromTemplate = true;
                _data.RosterDirty = false;
                _dirty = false;
                _config.LastFolder = Path.GetDirectoryName(target);
                Log("保存表格",
                    "文件=" + target,
                    "备份=" + (string.IsNullOrEmpty(backup) ? "无" : backup),
                    "学生数=" + _data.Students.Count);
                UpdateChrome();
                UpdateButtons();
                _statusFile.Text = "已保存：" + target;
                return true;
            }
            catch (XlsxException ex)
            {
                MessageBox.Show(this, ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "保存失败：\r\n" + ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void OpenFile()
        {
            if (!ConfirmDiscard()) return;
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "打开操行分表格";
            dlg.Filter = "Excel 工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*";
            string dir = _config.LastFolder;
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            dlg.InitialDirectory = dir;
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            OpenPath(dlg.FileName);
        }

        private void OpenPath(string path)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                ClassData d = XlsxReader.LoadClass(path);
                ApplyData(d, path, false);
                _config.LastFolder = Path.GetDirectoryName(path);
                Log("打开表格",
                    "文件=" + path,
                    "学生数=" + d.Students.Count,
                    "模块数=" + d.ModuleCount,
                    "初始分=" + Format.Decimal(d.InitialScore));
                _statusFile.Text = "已打开：" + path;
                if (d.Warnings.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("表格已打开，但有以下需要注意的地方：");
                    sb.AppendLine();
                    int max = Math.Min(d.Warnings.Count, 8);
                    for (int i = 0; i < max; i++) sb.AppendLine("· " + d.Warnings[i]);
                    if (d.Warnings.Count > max) sb.AppendLine("· 还有 " + (d.Warnings.Count - max) + " 条提示……");
                    MessageBox.Show(this, sb.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (XlsxException ex)
            {
                MessageBox.Show(this, ex.Message, "无法打开", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开该表格：\r\n" + ex.Message, "无法打开", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void NewClass()
        {
            if (!ConfirmDiscard()) return;
            using (NewClassForm dlg = new NewClassForm())
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                ClassData d = dlg.Result;
                ApplyData(d, "", true);
                Log("新建班级",
                    "班级=" + d.ClassName,
                    "模块数=" + d.ModuleCount,
                    "初始分=" + Format.Decimal(d.InitialScore),
                    "方面数=" + JoinInts(d));
                _statusFile.Text = "未保存：请点击“保存”生成 Excel 表格";
                DialogResult imp = MessageBox.Show(this,
                    "班级已创建。现在导入学生名单吗？\r\n\r\n也可以稍后点击顶部“导入学生”。",
                    "导入学生", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (imp == DialogResult.Yes) ImportStudents();
            }
        }

        private static string JoinInts(ClassData d)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < d.ModuleCount; i++)
            {
                if (i > 0) sb.Append('/');
                sb.Append(d.Modules[i].AspectCount);
            }
            return sb.ToString();
        }

        private void BatchAdd()
        {
            if (_data == null || _data.Students.Count == 0)
            {
                MessageBox.Show(this, "请先打开表格或导入学生。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (BatchAddForm dlg = new BatchAddForm(_data))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Result == null) return;
                BatchAddRequest req = dlg.Result;
                if (req.Students == null || req.Students.Count == 0) return;

                DateTime now = DateTime.Now;
                string batchId = now.ToString("yyyyMMddHHmmss");
                List<BatchUndoItem> items = new List<BatchUndoItem>();
                for (int i = 0; i < req.Students.Count; i++)
                {
                    Student s = req.Students[i];
                    decimal before = _data.GetModuleScore(s, req.ModuleIndex);
                    decimal avgBefore = _data.GetAverage(s);
                    RecordItem r = _data.AddRecord(s, req.ModuleIndex, req.AspectIndex, req.Item, req.Delta, now);
                    decimal after = _data.GetModuleScore(s, req.ModuleIndex);
                    decimal avgAfter = _data.GetAverage(s);

                    BatchUndoItem bi = new BatchUndoItem();
                    bi.Student = s;
                    bi.ModuleIndex = req.ModuleIndex;
                    bi.Index = s.Modules[req.ModuleIndex].Records.Count - 1;
                    bi.Record = r;
                    items.Add(bi);

                    Log("批量添加记录",
                        "批次=" + batchId,
                        "学号=" + s.No,
                        "姓名=" + s.Name,
                        "模块=" + _data.Modules[req.ModuleIndex].Name,
                        "方面=" + ClassData.AspectLabel(req.AspectIndex),
                        "事项=" + req.Item,
                        "加减分=" + Format.ScoreWithSign(req.Delta),
                        "模块分=" + Format.Decimal2(before) + "→" + Format.Decimal2(after),
                        "平均分=" + Format.Decimal(avgBefore) + "→" + Format.Decimal(avgAfter));
                }

                UndoAction u = new UndoAction();
                u.Kind = "batchadd";
                u.Batch = items;
                _undo.Add(u);
                _dirty = true;

                Log("批量添加",
                    "批次=" + batchId,
                    "人数=" + items.Count,
                    "模块=" + _data.Modules[req.ModuleIndex].Name,
                    "方面=" + ClassData.AspectLabel(req.AspectIndex),
                    "事项=" + req.Item,
                    "加减分=" + Format.ScoreWithSign(req.Delta),
                    "学生=" + DescribeStudents(req.Students));

                Student focus = _selected;
                if (focus == null || !ContainsStudent(req.Students, focus)) focus = req.Students[0];
                _loadingUi = true;
                try
                {
                    _selectedModule = req.ModuleIndex;
                    if (_cmbModule.Items.Count > req.ModuleIndex) _cmbModule.SelectedIndex = req.ModuleIndex;
                }
                finally
                {
                    _loadingUi = false;
                }
                RefreshStudentGrid();
                SelectStudent(focus, true, true);
                UpdateChrome();
                UpdateButtons();
                _statusFile.Text = string.Format("已为 {0} 名同学添加：{1} · {2}（{3} 分）",
                    items.Count, _data.Modules[req.ModuleIndex].Name, req.Item, Format.ScoreWithSign(req.Delta));
            }
        }

        private static bool ContainsStudent(List<Student> list, Student s)
        {
            for (int i = 0; i < list.Count; i++)
                if (object.ReferenceEquals(list[i], s)) return true;
            return false;
        }

        private static string DescribeStudents(List<Student> list)
        {
            StringBuilder sb = new StringBuilder();
            int max = Math.Min(list.Count, 8);
            for (int i = 0; i < max; i++)
            {
                if (i > 0) sb.Append('、');
                sb.Append(list[i].Name);
            }
            if (list.Count > max) sb.Append("等").Append(list.Count).Append("人");
            return sb.ToString();
        }

        private void ImportStudents()
        {
            if (_data == null)
            {
                MessageBox.Show(this, "请先打开表格或新建班级。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (ImportStudentsForm dlg = new ImportStudentsForm(_data.Students.Count, _data.ClassName))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                RosterResult result = dlg.Result;
                if (result == null || result.Students.Count == 0) return;
                bool replace = dlg.Mode == "replace";
                if (replace && _data.Students.Count > 0)
                {
                    DialogResult c = MessageBox.Show(this,
                        "替换名单会移除当前所有学生及其操行记录（仅影响当前程序，未保存前原文件不变）。\r\n\r\n确定继续吗？",
                        "替换名单", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (c != DialogResult.Yes) return;
                }

                int added = 0;
                int skipped = 0;
                if (replace)
                {
                    _data.Students.Clear();
                    for (int i = 0; i < result.Students.Count; i++)
                    {
                        StudentSeed seed = result.Students[i];
                        _data.NewStudent(seed.No, seed.Name);
                        added++;
                    }
                }
                else
                {
                    HashSet<string> exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < _data.Students.Count; i++)
                        if (!string.IsNullOrEmpty(_data.Students[i].No)) exists.Add(_data.Students[i].No);
                    for (int i = 0; i < result.Students.Count; i++)
                    {
                        StudentSeed seed = result.Students[i];
                        if (exists.Contains(seed.No)) { skipped++; continue; }
                        exists.Add(seed.No);
                        _data.NewStudent(seed.No, seed.Name);
                        added++;
                    }
                }
                _data.RosterDirty = true;
                _dirty = true;
                _undo.Clear();
                _selected = null;
                _loadingUi = true;
                try
                {
                    _cmbStudent.Items.Clear();
                    for (int i = 0; i < _data.Students.Count; i++) _cmbStudent.Items.Add(_data.Students[i]);
                    _loadingUi = false;
                    RefreshStudentGrid();
                    if (_data.Students.Count > 0) SelectStudent(_data.Students[0], true, true);
                    else
                    {
                        UpdateStudentCard();
                        UpdateRecords();
                    }
                }
                finally
                {
                    _loadingUi = false;
                }
                UpdateChrome();
                UpdateButtons();
                Log("导入学生",
                    "模式=" + (replace ? "替换" : "追加"),
                    "本次导入=" + added,
                    "跳过重复=" + skipped,
                    "当前人数=" + _data.Students.Count);
                _statusFile.Text = string.Format("已导入 {0} 人{1}", added, skipped > 0 ? "（跳过重复 " + skipped + " 人）" : "");
                if (_data.FromTemplate)
                {
                    MessageBox.Show(this,
                        "名单已更新。\r\n保存时会按当前名单重新排列数据行，表头、列宽和合并样式会保留。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // ---------------- 辅助 ----------------

        private string CurrentLogDir()
        {
            return CxLog.ResolveDirectory(_filePath, Application.StartupPath);
        }

        private void Log(string action, params string[] fields)
        {
            string path = CxLog.Write(CurrentLogDir(), action, fields);
            if (!string.IsNullOrEmpty(path))
            {
                _statusLog.ToolTipText = path;
            }
        }

        private void OpenLogFolder()
        {
            try
            {
                string dir = CurrentLogDir();
                Directory.CreateDirectory(dir);
                Process.Start("explorer.exe", "\"" + dir + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开日志文件夹：\r\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OpenHelp()
        {
            string dir = Application.StartupPath;
            string[] names = { "使用说明.md", "README.md" };
            for (int i = 0; i < names.Length; i++)
            {
                string p = Path.Combine(dir, names[i]);
                if (File.Exists(p))
                {
                    try { Process.Start(p); return; }
                    catch { }
                }
            }
            MessageBox.Show(this,
                "没有找到“使用说明.md”。\r\n\r\n主要操作：\r\n1. 点击“打开表格”打开上个月的模板，或“新建班级”；\r\n2. 在左侧选择学生；\r\n3. 选择模块、方面，填写事项和加减分；\r\n4. 点击“添加记录”，再点击“保存”。\r\n\r\n日志位于表格同目录的“操行日志”文件夹。",
                "使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowTotal()
        {
            if (_data == null) return;
            using (TotalForm dlg = new TotalForm(_data)) dlg.ShowDialog(this);
        }

        private bool ConfirmDiscard()
        {
            if (!_dirty) return true;
            DialogResult r = MessageBox.Show(this,
                "当前有未保存的修改，是否先保存？",
                "未保存", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return false;
            if (r == DialogResult.Yes) return SaveCurrent(false);
            return true;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_closeConfirmed) return;
            if (!ConfirmDiscard())
            {
                e.Cancel = true;
                return;
            }
            _closeConfirmed = true;
            try
            {
                _config.Width = Width;
                _config.Height = Height;
                _config.Maximized = WindowState == FormWindowState.Maximized;
                _config.Save();
            }
            catch { }
        }
    }
}














