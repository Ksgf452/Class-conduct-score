using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CaoXing
{
    internal class NewClassForm : Form
    {
        public ClassData Result;

        private TextBox _txtName;
        private NumericUpDown _numInitial;
        private NumericUpDown _numModules;
        private DataGridView _grid;
        private bool _syncing;

        public NewClassForm()
        {
            DialogTheme.Apply(this);
            Text = "新建班级";
            Size = new Size(620, 650);
            Build();
        }

        private void Build()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 2;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 66f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            Controls.Add(root);

            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 66;
            header.BackColor = Theme.Navy;
            root.Controls.Add(header, 0, 0);

            Label title = new Label();
            title.Text = "新建班级";
            title.Font = Theme.Font(16f, FontStyle.Bold);
            title.ForeColor = Color.White;
            title.AutoSize = true;
            title.Location = new Point(22, 12);
            header.Controls.Add(title);

            Label sub = new Label();
            sub.Text = "设置模块数量、每个模块的方面数量和初始分";
            sub.Font = Theme.Font(9f);
            sub.ForeColor = Theme.HeaderSub;
            sub.AutoSize = true;
            sub.Location = new Point(24, 41);
            header.Controls.Add(sub);

            Panel body = new Panel();
            body.Dock = DockStyle.Fill;
            body.Padding = new Padding(24, 18, 24, 12);
            body.BackColor = Theme.Page;
            root.Controls.Add(body, 0, 1);

            Label nameLabel = new Label();
            nameLabel.Text = "班级名称";
            nameLabel.Font = Theme.Font(9f);
            nameLabel.ForeColor = Theme.Muted;
            nameLabel.AutoSize = true;
            nameLabel.Location = new Point(0, 4);
            body.Controls.Add(nameLabel);

            _txtName = new TextBox();
            _txtName.Location = new Point(0, 26);
            _txtName.Width = 240;
            _txtName.Text = "班级";
            Theme.StyleTextBox(_txtName);
            body.Controls.Add(_txtName);

            Label initLabel = new Label();
            initLabel.Text = "模块初始分";
            initLabel.Font = Theme.Font(9f);
            initLabel.ForeColor = Theme.Muted;
            initLabel.AutoSize = true;
            initLabel.Location = new Point(270, 4);
            body.Controls.Add(initLabel);

            _numInitial = new NumericUpDown();
            _numInitial.Location = new Point(270, 26);
            _numInitial.Width = 100;
            _numInitial.Minimum = 0m;
            _numInitial.Maximum = 1000m;
            _numInitial.DecimalPlaces = 1;
            _numInitial.Value = 80m;
            _numInitial.Font = Theme.Font(10f);
            body.Controls.Add(_numInitial);

            Label moduleLabel = new Label();
            moduleLabel.Text = "模块数量";
            moduleLabel.Font = Theme.Font(9f);
            moduleLabel.ForeColor = Theme.Muted;
            moduleLabel.AutoSize = true;
            moduleLabel.Location = new Point(400, 4);
            body.Controls.Add(moduleLabel);

            _numModules = new NumericUpDown();
            _numModules.Location = new Point(400, 26);
            _numModules.Width = 80;
            _numModules.Minimum = 1m;
            _numModules.Maximum = 20m;
            _numModules.Value = 8m;
            _numModules.Font = Theme.Font(10f);
            body.Controls.Add(_numModules);

            Label gridLabel = new Label();
            gridLabel.Text = "模块名称与方面数量";
            gridLabel.Font = Theme.Font(9f);
            gridLabel.ForeColor = Theme.Muted;
            gridLabel.AutoSize = true;
            gridLabel.Location = new Point(0, 74);
            body.Controls.Add(gridLabel);

            _grid = new DataGridView();
            _grid.Location = new Point(0, 96);
            _grid.Size = new Size(540, 330);
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Theme.StyleGrid(_grid);
            _grid.ReadOnly = false;
            _grid.EditMode = DataGridViewEditMode.EditOnEnter;
            _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            _grid.MultiSelect = false;
            _grid.ColumnHeadersHeight = 34;
            _grid.RowTemplate.Height = 30;
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "模块名称", Name = "name", FillWeight = 70 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "方面数量", Name = "aspects", FillWeight = 20 });
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.AllowUserToAddRows = false;
            body.Controls.Add(_grid);

            Label tip = new Label();
            tip.Text = "提示：模块名称可自定义；方面数量表示该模块下“方面一……方面N”的数量，通常为 5。";
            tip.Font = Theme.Font(8.5f);
            tip.ForeColor = Theme.Muted;
            tip.AutoSize = false;
            tip.Location = new Point(0, 436);
            tip.Size = new Size(540, 38);
            tip.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            body.Controls.Add(tip);

            Button ok = new Button();
            ok.Text = "创建班级";
            ok.Size = new Size(120, 38);
            ok.Location = new Point(300, 494);
            ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.PrimaryButton(ok);
            ok.Click += delegate { Create(); };
            body.Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.Size = new Size(100, 38);
            cancel.Location = new Point(432, 494);
            cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.SecondaryButton(cancel);
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            body.Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;

            _numModules.ValueChanged += delegate { SyncRows(); };
            SyncRows();
        }

        private void SyncRows()
        {
            if (_syncing) return;
            _syncing = true;
            try
            {
                int count = (int)_numModules.Value;
                while (_grid.Rows.Count < count)
                {
                    int i = _grid.Rows.Count;
                    _grid.Rows.Add("模块" + Cn.Number(i + 1), "5");
                }
                while (_grid.Rows.Count > count)
                    _grid.Rows.RemoveAt(_grid.Rows.Count - 1);
            }
            finally
            {
                _syncing = false;
            }
        }

        private void Create()
        {
            string className = _txtName.Text.Trim();
            if (className.Length == 0) className = "班级";
            List<string> names = new List<string>();
            List<int> aspects = new List<int>();
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                string name = Convert.ToString(_grid.Rows[i].Cells[0].Value).Trim();
                if (name.Length == 0)
                {
                    MessageBox.Show(this, "第 " + (i + 1) + " 个模块名称不能为空。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (names.Contains(name))
                {
                    MessageBox.Show(this, "模块名称重复：" + name, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int a;
                if (!int.TryParse(Convert.ToString(_grid.Rows[i].Cells[1].Value).Trim(), out a) || a < 1 || a > 30)
                {
                    MessageBox.Show(this, "第 " + (i + 1) + " 个模块的方面数量必须是 1—30 的整数。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                names.Add(name);
                aspects.Add(a);
            }

            ClassData d = new ClassData();
            d.ClassName = className;
            d.InitialScore = _numInitial.Value;
            for (int i = 0; i < names.Count; i++) d.Modules.Add(new ModuleDef(names[i], aspects[i]));
            Result = d;
            DialogResult = DialogResult.OK;
        }
    }

    internal class ImportStudentsForm : Form
    {
        public RosterResult Result;
        public string Mode = "replace";

        private RadioButton _rdoFile;
        private RadioButton _rdoPaste;
        private Panel _filePanel;
        private Panel _pastePanel;
        private TextBox _txtPath;
        private TextBox _txtPaste;
        private DataGridView _preview;
        private Label _lblInfo;
        private Button _btnImport;
        private RadioButton _rdoReplace;
        private RadioButton _rdoAppend;
        private int _currentCount;

        public ImportStudentsForm(int currentCount, string className)
        {
            _currentCount = currentCount;
            DialogTheme.Apply(this);
            Text = "导入学生名单";
            Size = new Size(680, 650);
            Build(className);
        }

        private void Build(string className)
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 2;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 66f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            Controls.Add(root);
            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 66;
            header.BackColor = Theme.Navy;
            root.Controls.Add(header, 0, 0);

            Label title = new Label();
            title.Text = "导入学生名单";
            title.Font = Theme.Font(16f, FontStyle.Bold);
            title.ForeColor = Color.White;
            title.AutoSize = true;
            title.Location = new Point(22, 12);
            header.Controls.Add(title);

            Label sub = new Label();
            sub.Text = "读取学号与姓名，支持 Excel / CSV / TXT 或直接粘贴";
            sub.Font = Theme.Font(9f);
            sub.ForeColor = Theme.HeaderSub;
            sub.AutoSize = true;
            sub.Location = new Point(24, 41);
            header.Controls.Add(sub);

            Panel body = new Panel();
            body.Dock = DockStyle.Fill;
            body.Padding = new Padding(24, 16, 24, 12);
            body.BackColor = Theme.Page;
            root.Controls.Add(body, 0, 1);

            _rdoFile = new RadioButton();
            _rdoFile.Text = "从文件导入";
            _rdoFile.Checked = true;
            _rdoFile.AutoSize = true;
            _rdoFile.Location = new Point(0, 2);
            _rdoFile.Font = Theme.Font(10f);
            body.Controls.Add(_rdoFile);

            _rdoPaste = new RadioButton();
            _rdoPaste.Text = "粘贴名单";
            _rdoPaste.AutoSize = true;
            _rdoPaste.Location = new Point(130, 2);
            _rdoPaste.Font = Theme.Font(10f);
            body.Controls.Add(_rdoPaste);

            _filePanel = new Panel();
            _filePanel.Location = new Point(0, 30);
            _filePanel.Size = new Size(620, 72);
            _filePanel.BackColor = Theme.Page;
            body.Controls.Add(_filePanel);

            _txtPath = new TextBox();
            _txtPath.Location = new Point(0, 4);
            _txtPath.Width = 500;
            _txtPath.ReadOnly = true;
            Theme.StyleTextBox(_txtPath);
            _filePanel.Controls.Add(_txtPath);

            Button browse = new Button();
            browse.Text = "浏览…";
            browse.Size = new Size(100, 30);
            browse.Location = new Point(512, 2);
            Theme.SecondaryButton(browse);
            browse.Click += delegate { Browse(); };
            _filePanel.Controls.Add(browse);

            Label fileHint = new Label();
            fileHint.Text = "表头包含“学号”“姓名”时会自动识别列；没有表头则按前两列读取。";
            fileHint.Font = Theme.Font(8.5f);
            fileHint.ForeColor = Theme.Muted;
            fileHint.AutoSize = true;
            fileHint.Location = new Point(0, 42);
            _filePanel.Controls.Add(fileHint);

            _pastePanel = new Panel();
            _pastePanel.Location = new Point(0, 30);
            _pastePanel.Size = new Size(620, 130);
            _pastePanel.BackColor = Theme.Page;
            _pastePanel.Visible = false;
            body.Controls.Add(_pastePanel);

            _txtPaste = new TextBox();
            _txtPaste.Multiline = true;
            _txtPaste.ScrollBars = ScrollBars.Vertical;
            _txtPaste.Location = new Point(0, 0);
            _txtPaste.Size = new Size(620, 126);
            _txtPaste.Font = Theme.Font(9.5f);
            _txtPaste.Text = "学号\t姓名\r\n2024001\t张三\r\n2024002\t李四";
            Theme.StyleTextBox(_txtPaste);
            _pastePanel.Controls.Add(_txtPaste);

            Label modeLabel = new Label();
            modeLabel.Text = "导入方式";
            modeLabel.Font = Theme.Font(9f);
            modeLabel.ForeColor = Theme.Muted;
            modeLabel.AutoSize = true;
            modeLabel.Location = new Point(0, 170);
            body.Controls.Add(modeLabel);

            _rdoAppend = new RadioButton();
            _rdoAppend.Text = "追加到现有名单（跳过重复学号）";
            _rdoAppend.AutoSize = true;
            _rdoAppend.Location = new Point(0, 192);
            _rdoAppend.Font = Theme.Font(9.5f);
            body.Controls.Add(_rdoAppend);

            _rdoReplace = new RadioButton();
            _rdoReplace.Text = "替换现有名单";
            _rdoReplace.AutoSize = true;
            _rdoReplace.Location = new Point(280, 192);
            _rdoReplace.Font = Theme.Font(9.5f);
            body.Controls.Add(_rdoReplace);
            if (_currentCount > 0) _rdoAppend.Checked = true;
            else _rdoReplace.Checked = true;
            if (_currentCount == 0) _rdoReplace.Text = "替换现有名单（当前为空）";

            Label previewLabel = new Label();
            previewLabel.Text = "预览";
            previewLabel.Font = Theme.Font(9f);
            previewLabel.ForeColor = Theme.Muted;
            previewLabel.AutoSize = true;
            previewLabel.Location = new Point(0, 226);
            body.Controls.Add(previewLabel);

            _preview = new DataGridView();
            _preview.Location = new Point(0, 248);
            _preview.Size = new Size(620, 230);
            _preview.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Theme.StyleGrid(_preview);
            _preview.RowTemplate.Height = 28;
            _preview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "学号", FillWeight = 40 });
            _preview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "姓名", FillWeight = 30 });
            _preview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            body.Controls.Add(_preview);

            _lblInfo = new Label();
            _lblInfo.Text = "点击“读取预览”检查名单。";
            _lblInfo.Font = Theme.Font(9f);
            _lblInfo.ForeColor = Theme.Muted;
            _lblInfo.AutoSize = false;
            _lblInfo.Location = new Point(0, 484);
            _lblInfo.Size = new Size(620, 22);
            _lblInfo.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            body.Controls.Add(_lblInfo);

            Button previewBtn = new Button();
            previewBtn.Text = "读取预览";
            previewBtn.Size = new Size(110, 36);
            previewBtn.Location = new Point(274, 516);
            previewBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.SecondaryButton(previewBtn);
            previewBtn.Click += delegate { Parse(); };
            body.Controls.Add(previewBtn);

            _btnImport = new Button();
            _btnImport.Text = "导入";
            _btnImport.Size = new Size(100, 36);
            _btnImport.Location = new Point(396, 516);
            _btnImport.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnImport.Enabled = false;
            Theme.PrimaryButton(_btnImport);
            _btnImport.Click += delegate { DoImport(); };
            body.Controls.Add(_btnImport);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.Size = new Size(100, 36);
            cancel.Location = new Point(508, 516);
            cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.SecondaryButton(cancel);
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            body.Controls.Add(cancel);

            CancelButton = cancel;
            _rdoFile.CheckedChanged += delegate { UpdateSource(); };
            _rdoPaste.CheckedChanged += delegate { UpdateSource(); };
        }

        private void UpdateSource()
        {
            _filePanel.Visible = _rdoFile.Checked;
            _pastePanel.Visible = _rdoPaste.Checked;
            Result = null;
            _preview.Rows.Clear();
            _lblInfo.Text = _rdoFile.Checked ? "点击“读取预览”检查名单。" : "粘贴后点击“读取预览”。";
            _btnImport.Enabled = false;
        }

        private void Browse()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "选择名单文件";
            dlg.Filter = "名单文件 (*.xlsx;*.xlsm;*.csv;*.txt)|*.xlsx;*.xlsm;*.csv;*.txt|所有文件 (*.*)|*.*";
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                _txtPath.Text = dlg.FileName;
                Parse();
            }
        }

        private void Parse()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                RosterResult r;
                if (_rdoFile.Checked)
                {
                    if (string.IsNullOrEmpty(_txtPath.Text)) { MessageBox.Show(this, "请先选择名单文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                    r = RosterReader.ReadFile(_txtPath.Text);
                }
                else
                {
                    r = RosterReader.ReadPaste(_txtPaste.Text);
                }
                Result = r;
                _preview.Rows.Clear();
                int max = Math.Min(r.Students.Count, 200);
                for (int i = 0; i < max; i++)
                    _preview.Rows.Add(r.Students[i].No, r.Students[i].Name);
                _lblInfo.Text = r.Detail + (r.Students.Count > max ? "，仅预览前 200 条" : "");
                _btnImport.Enabled = true;
            }
            catch (XlsxException ex)
            {
                Result = null;
                _btnImport.Enabled = false;
                MessageBox.Show(this, ex.Message, "读取失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Result = null;
                _btnImport.Enabled = false;
                MessageBox.Show(this, "读取失败：\r\n" + ex.Message, "读取失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void DoImport()
        {
            if (Result == null)
            {
                Parse();
                if (Result == null) return;
            }
            bool replace = _rdoReplace.Checked;
            if (replace && _currentCount > 0)
            {
                DialogResult r = MessageBox.Show(this,
                    "替换名单会移除现有学生及其记录（未保存前原文件不变）。确定继续吗？",
                    "替换名单", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (r != DialogResult.Yes) return;
            }
            Mode = replace ? "replace" : "append";
            DialogResult = DialogResult.OK;
        }
    }

    internal class TotalForm : Form
    {
        private readonly ClassData _data;
        private DataGridView _grid;
        private Label _lblStats;
        private CheckBox _chkSort;

        public TotalForm(ClassData data)
        {
            _data = data;
            Text = "班级操行分总表";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1100, 700);
            MinimumSize = new Size(900, 560);
            BackColor = Theme.Page;
            Font = Theme.Font(9.5f);
            Build();
            Fill();
        }

        private void Build()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 2;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            Controls.Add(root);

            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 64;
            header.BackColor = Theme.Navy;
            root.Controls.Add(header, 0, 0);

            Label title = new Label();
            title.Text = "班级操行分总表";
            title.Font = Theme.Font(16f, FontStyle.Bold);
            title.ForeColor = Color.White;
            title.AutoSize = true;
            title.Location = new Point(22, 11);
            header.Controls.Add(title);

            Label sub = new Label();
            sub.Text = string.Format("{0} · {1} 人 · {2} 个模块 · 初始 {3} 分",
                _data.ClassName, _data.Students.Count, _data.ModuleCount, Format.Decimal(_data.InitialScore));
            sub.Font = Theme.Font(9f);
            sub.ForeColor = Theme.HeaderSub;
            sub.AutoSize = true;
            sub.Location = new Point(24, 40);
            header.Controls.Add(sub);

            Panel body = new Panel();
            body.Dock = DockStyle.Fill;
            body.Padding = new Padding(18, 12, 18, 12);
            body.BackColor = Theme.Page;
            root.Controls.Add(body, 0, 1);

            TableLayoutPanel content = new TableLayoutPanel();
            content.Dock = DockStyle.Fill;
            content.ColumnCount = 1;
            content.RowCount = 4;
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 44f));
            content.BackColor = Theme.Page;
            body.Controls.Add(content);

            _chkSort = new CheckBox();
            _chkSort.Text = "按平均分从高到低排序";
            _chkSort.Dock = DockStyle.Fill;
            _chkSort.Font = Theme.Font(9.5f);
            _chkSort.CheckedChanged += delegate { Fill(); };
            content.Controls.Add(_chkSort, 0, 0);

            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            Theme.StyleGrid(_grid);
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.Columns.Add(Col("序号", 8));
            _grid.Columns.Add(Col("学号", 20));
            _grid.Columns.Add(Col("姓名", 14));
            for (int i = 0; i < _data.ModuleCount; i++)
            {
                DataGridViewTextBoxColumn mc = Col(_data.Modules[i].Name, 9);
                mc.MinimumWidth = 58;
                _grid.Columns.Add(mc);
            }
            _grid.Columns.Add(Col("平均分", 10));
            content.Controls.Add(_grid, 0, 1);

            _lblStats = new Label();
            _lblStats.Dock = DockStyle.Fill;
            _lblStats.TextAlign = ContentAlignment.MiddleLeft;
            _lblStats.Font = Theme.Font(9f);
            _lblStats.ForeColor = Theme.Muted;
            content.Controls.Add(_lblStats, 0, 2);

            Panel footer = new Panel();
            footer.Dock = DockStyle.Fill;
            footer.BackColor = Theme.Page;
            content.Controls.Add(footer, 0, 3);

            Button export = new Button();
            export.Text = "导出 CSV";
            export.Size = new Size(100, 34);
            export.Location = new Point(0, 5);
            Theme.SecondaryButton(export);
            export.Click += delegate { ExportCsv(); };
            footer.Controls.Add(export);

            Button close = new Button();
            close.Text = "关闭";
            close.Size = new Size(90, 34);
            close.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            close.Location = new Point(footer.Width - 90, 5);
            Theme.PrimaryButton(close);
            close.Click += delegate { Close(); };
            footer.Controls.Add(close);
            footer.Resize += delegate { close.Left = footer.ClientSize.Width - close.Width; };
            _grid.CellFormatting += delegate(object sender, DataGridViewCellFormattingEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 3) return;
                decimal v;
                if (e.Value == null || !decimal.TryParse(Convert.ToString(e.Value), NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return;
                e.CellStyle.ForeColor = v < _data.InitialScore ? Theme.Red : (v > _data.InitialScore ? Theme.Green : Theme.Text);
                if (e.ColumnIndex == _grid.Columns.Count - 1) e.CellStyle.Font = Theme.Font(9.5f, FontStyle.Bold);
            };
        }

        private static DataGridViewTextBoxColumn Col(string name, int weight)
        {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = name;
            c.Name = name;
            c.FillWeight = weight;
            c.SortMode = DataGridViewColumnSortMode.NotSortable;
            c.ReadOnly = true;
            return c;
        }

        private void Fill()
        {
            if (_data == null) return;
            _grid.Rows.Clear();
            List<Student> list = new List<Student>(_data.Students);
            if (_chkSort.Checked)
            {
                list.Sort(delegate(Student a, Student b)
                {
                    int c = _data.GetAverage(b).CompareTo(_data.GetAverage(a));
                    if (c != 0) return c;
                    return a.Seq.CompareTo(b.Seq);
                });
            }
            for (int i = 0; i < list.Count; i++)
            {
                Student s = list[i];
                List<object> values = new List<object>();
                values.Add(s.Seq);
                values.Add(s.No);
                values.Add(s.Name);
                for (int m = 0; m < _data.ModuleCount; m++) values.Add(Format.Decimal(_data.GetModuleScore(s, m)));
                values.Add(Format.Decimal(_data.GetAverage(s)));
                int idx = _grid.Rows.Add(values.ToArray());
                _grid.Rows[idx].Tag = s;
            }
            decimal total = 0m;
            decimal max = decimal.MinValue;
            decimal min = decimal.MaxValue;
            for (int i = 0; i < _data.Students.Count; i++)
            {
                decimal a = _data.GetAverage(_data.Students[i]);
                total += a;
                if (a > max) max = a;
                if (a < min) min = a;
            }
            decimal avg = _data.Students.Count == 0 ? 0m : total / _data.Students.Count;
            _lblStats.Text = string.Format("班级平均分 {0}　｜　最高 {1}　｜　最低 {2}　｜　{3} 名学生",
                Format.Decimal(avg), _data.Students.Count == 0 ? "—" : Format.Decimal(max),
                _data.Students.Count == 0 ? "—" : Format.Decimal(min), _data.Students.Count);
        }

        private void ExportCsv()
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "导出班级总表";
            dlg.Filter = "CSV 文件 (*.csv)|*.csv";
            dlg.FileName = (string.IsNullOrEmpty(_data.ClassName) ? "班级" : _data.ClassName) + "操行分总表.csv";
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                StringBuilder sb = new StringBuilder();
                List<string> headers = new List<string>();
                headers.Add("序号"); headers.Add("学号"); headers.Add("姓名");
                for (int m = 0; m < _data.ModuleCount; m++) headers.Add(_data.Modules[m].Name);
                headers.Add("平均分");
                sb.AppendLine(JoinCsv(headers));
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    List<string> vals = new List<string>();
                    for (int c = 0; c < row.Cells.Count; c++) vals.Add(Convert.ToString(row.Cells[c].Value));
                    sb.AppendLine(JoinCsv(vals));
                }
                File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true));
                MessageBox.Show(this, "已导出：\r\n" + dlg.FileName, "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "导出失败：\r\n" + ex.Message, "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string JoinCsv(List<string> values)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < values.Count; i++)
            {
                if (i > 0) sb.Append(',');
                string v = values[i] ?? "";
                if (v.IndexOf(',') >= 0 || v.IndexOf('"') >= 0 || v.IndexOf('\n') >= 0 || v.IndexOf('\r') >= 0)
                    sb.Append('"').Append(v.Replace("\"", "\"\"")).Append('"');
                else sb.Append(v);
            }
            return sb.ToString();
        }
    }
}





