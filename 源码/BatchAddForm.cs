using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace CaoXing
{
    public class BatchAddRequest
    {
        public List<Student> Students = new List<Student>();
        public int ModuleIndex;
        public int AspectIndex;
        public string Item;
        public decimal Delta;
    }

    /// <summary>选择多名学生，为每人添加同一条操行记录。</summary>
    internal class BatchAddForm : Form
    {
        private readonly ClassData _data;
        public BatchAddRequest Result;

        private CheckedListBox _list;
        private ComboBox _cmbModule;
        private ComboBox _cmbAspect;
        private TextBox _txtItem;
        private NumericUpDown _numDelta;
        private Label _lblCount;
        private Label _lblPreview;
        private Button _btnAdd;
        private bool _loading;

        public BatchAddForm(ClassData data)
        {
            _data = data;
            DialogTheme.Apply(this);
            Text = "批量添加操行记录";
            Size = new Size(840, 660);
            Build();
        }

        private void Build()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 3;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 66f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56f));
            Controls.Add(root);

            Panel header = new Panel();
            header.Dock = DockStyle.Fill;
            header.BackColor = Theme.Navy;
            root.Controls.Add(header, 0, 0);

            Label title = new Label();
            title.Text = "批量添加操行记录";
            title.Font = Theme.Font(16f, FontStyle.Bold);
            title.ForeColor = Color.White;
            title.AutoSize = true;
            title.Location = new Point(22, 12);
            header.Controls.Add(title);

            Label sub = new Label();
            sub.Text = "为多名同学或全班添加同一条记录，每位同学分别累计";
            sub.Font = Theme.Font(9f);
            sub.ForeColor = Theme.HeaderSub;
            sub.AutoSize = true;
            sub.Location = new Point(24, 41);
            header.Controls.Add(sub);

            TableLayoutPanel body = new TableLayoutPanel();
            body.Dock = DockStyle.Fill;
            body.ColumnCount = 2;
            body.RowCount = 1;
            body.Padding = new Padding(20, 14, 20, 10);
            body.BackColor = Theme.Page;
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 350f));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.Controls.Add(body, 0, 1);

            BuildStudentCard(body);
            BuildEditorCard(body);

            FlowLayoutPanel footer = new FlowLayoutPanel();
            footer.Dock = DockStyle.Fill;
            footer.FlowDirection = FlowDirection.RightToLeft;
            footer.WrapContents = false;
            footer.Padding = new Padding(0, 10, 18, 0);
            footer.BackColor = Theme.Page;
            root.Controls.Add(footer, 0, 2);

            _btnAdd = new Button();
            _btnAdd.Text = "为 0 人添加";
            _btnAdd.Size = new Size(160, 36);
            _btnAdd.Enabled = false;
            Theme.PrimaryButton(_btnAdd);
            _btnAdd.Click += delegate { DoAdd(); };
            footer.Controls.Add(_btnAdd);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.Size = new Size(90, 36);
            cancel.Margin = new Padding(0, 0, 10, 0);
            Theme.SecondaryButton(cancel);
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            footer.Controls.Add(cancel);

            AcceptButton = _btnAdd;
            CancelButton = cancel;

            RebuildAspects();
            UpdateSummary(0);
        }

        private void BuildStudentCard(TableLayoutPanel body)
        {
            CardPanel card = new CardPanel();
            card.Dock = DockStyle.Fill;
            card.Margin = new Padding(0, 0, 12, 0);
            card.Padding = new Padding(14);
            body.Controls.Add(card, 0, 0);

            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.ColumnCount = 1;
            t.RowCount = 3;
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 34f));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            card.Controls.Add(t);

            Label title = new Label();
            title.Text = "选择学生";
            title.Font = Theme.Font(12f, FontStyle.Bold);
            title.ForeColor = Theme.Text;
            title.Dock = DockStyle.Fill;
            title.TextAlign = ContentAlignment.MiddleLeft;
            t.Controls.Add(title, 0, 0);

            Panel tools = new Panel();
            tools.Dock = DockStyle.Fill;
            tools.BackColor = Color.White;
            t.Controls.Add(tools, 0, 1);

            Button all = new Button();
            all.Text = "全选（全班）";
            all.Size = new Size(100, 28);
            all.Location = new Point(0, 2);
            all.Font = Theme.Font(8.5f);
            Theme.SecondaryButton(all);
            all.Click += delegate { SetAll(true); };
            tools.Controls.Add(all);

            Button none = new Button();
            none.Text = "全不选";
            none.Size = new Size(70, 28);
            none.Location = new Point(108, 2);
            none.Font = Theme.Font(8.5f);
            Theme.SecondaryButton(none);
            none.Click += delegate { SetAll(false); };
            tools.Controls.Add(none);

            _lblCount = new Label();
            _lblCount.Text = "已选 0 人";
            _lblCount.Font = Theme.Font(9f);
            _lblCount.ForeColor = Theme.Muted;
            _lblCount.AutoSize = true;
            _lblCount.Location = new Point(186, 8);
            tools.Controls.Add(_lblCount);

            _list = new CheckedListBox();
            _list.Dock = DockStyle.Fill;
            _list.CheckOnClick = true;
            _list.DisplayMember = "Display";
            _list.BorderStyle = BorderStyle.None;
            _list.Font = Theme.Font(10f);
            _list.BackColor = Color.White;
            _list.ForeColor = Theme.Text;
            _list.IntegralHeight = false;
            _list.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
            {
                if (_loading) return;
                BeginInvoke((MethodInvoker)delegate { UpdateSummary(_list.CheckedItems.Count); });
            };
            for (int i = 0; i < _data.Students.Count; i++) _list.Items.Add(_data.Students[i]);
            t.Controls.Add(_list, 0, 2);
        }

        private void BuildEditorCard(TableLayoutPanel body)
        {
            CardPanel card = new CardPanel();
            card.Dock = DockStyle.Fill;
            card.Margin = new Padding(0);
            card.Padding = new Padding(16);
            body.Controls.Add(card, 1, 0);

            Label title = new Label();
            title.Text = "记录内容";
            title.Font = Theme.Font(12f, FontStyle.Bold);
            title.ForeColor = Theme.Text;
            title.AutoSize = true;
            title.Location = new Point(0, 0);
            card.Controls.Add(title);

            Label moduleLabel = new Label();
            moduleLabel.Text = "模块";
            moduleLabel.Font = Theme.Font(9f);
            moduleLabel.ForeColor = Theme.Muted;
            moduleLabel.AutoSize = true;
            moduleLabel.Location = new Point(0, 42);
            card.Controls.Add(moduleLabel);

            _cmbModule = new ComboBox();
            _cmbModule.Location = new Point(0, 62);
            _cmbModule.Width = 150;
            Theme.StyleCombo(_cmbModule);
            for (int i = 0; i < _data.Modules.Count; i++) _cmbModule.Items.Add(_data.Modules[i].Name);
            if (_cmbModule.Items.Count > 0) _cmbModule.SelectedIndex = 0;
            _cmbModule.SelectedIndexChanged += delegate
            {
                if (_loading) return;
                RebuildAspects();
            };
            card.Controls.Add(_cmbModule);

            Label aspectLabel = new Label();
            aspectLabel.Text = "方面";
            aspectLabel.Font = Theme.Font(9f);
            aspectLabel.ForeColor = Theme.Muted;
            aspectLabel.AutoSize = true;
            aspectLabel.Location = new Point(162, 42);
            card.Controls.Add(aspectLabel);

            _cmbAspect = new ComboBox();
            _cmbAspect.Location = new Point(162, 62);
            _cmbAspect.Width = 110;
            Theme.StyleCombo(_cmbAspect);
            card.Controls.Add(_cmbAspect);

            Label deltaLabel = new Label();
            deltaLabel.Text = "加减分";
            deltaLabel.Font = Theme.Font(9f);
            deltaLabel.ForeColor = Theme.Muted;
            deltaLabel.AutoSize = true;
            deltaLabel.Location = new Point(284, 42);
            card.Controls.Add(deltaLabel);

            _numDelta = new NumericUpDown();
            _numDelta.Location = new Point(284, 62);
            _numDelta.Width = 90;
            _numDelta.Minimum = -100m;
            _numDelta.Maximum = 100m;
            _numDelta.DecimalPlaces = 2;
            _numDelta.Increment = 0.5m;
            _numDelta.Value = 1m;
            _numDelta.Font = Theme.Font(10f);
            _numDelta.ValueChanged += delegate { UpdateSummary(_list.CheckedItems.Count); };
            card.Controls.Add(_numDelta);

            Label itemLabel = new Label();
            itemLabel.Text = "事项";
            itemLabel.Font = Theme.Font(9f);
            itemLabel.ForeColor = Theme.Muted;
            itemLabel.AutoSize = true;
            itemLabel.Location = new Point(0, 108);
            card.Controls.Add(itemLabel);

            _txtItem = new TextBox();
            _txtItem.Location = new Point(0, 128);
            _txtItem.Width = 350;
            _txtItem.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Theme.StyleTextBox(_txtItem);
            _txtItem.TextChanged += delegate { UpdateSummary(_list.CheckedItems.Count); };
            card.Controls.Add(_txtItem);
            Theme.SetCue(_txtItem, "例如：运动会集体项目、主题班会、志愿服务……");

            _lblPreview = new Label();
            _lblPreview.Text = "勾选同学并填写记录后，这里会显示本次批量添加的内容。";
            _lblPreview.Font = Theme.Font(9.5f);
            _lblPreview.ForeColor = Theme.Muted;
            _lblPreview.AutoSize = false;
            _lblPreview.Location = new Point(0, 176);
            _lblPreview.Size = new Size(350, 96);
            _lblPreview.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            card.Controls.Add(_lblPreview);

            Label note = new Label();
            note.Text = "每位同学的该模块分、平均分都会分别更新；\r\n添加后可用一次“撤销”整体撤回本次批量操作。";
            note.Font = Theme.Font(8.5f);
            note.ForeColor = Theme.Muted;
            note.AutoSize = false;
            note.Location = new Point(0, 286);
            note.Size = new Size(350, 46);
            note.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            card.Controls.Add(note);

            card.Resize += delegate
            {
                _txtItem.Width = Math.Max(100, card.ClientSize.Width - 32);
                _lblPreview.Width = _txtItem.Width;
                note.Width = _txtItem.Width;
            };
        }

        private void RebuildAspects()
        {
            int moduleIndex = _cmbModule.SelectedIndex;
            if (moduleIndex < 0 || moduleIndex >= _data.Modules.Count) return;
            int old = _cmbAspect.SelectedIndex;
            _loading = true;
            try
            {
                _cmbAspect.Items.Clear();
                int count = _data.Modules[moduleIndex].AspectCount;
                for (int i = 1; i <= count; i++) _cmbAspect.Items.Add(ClassData.AspectLabel(i));
                if (_cmbAspect.Items.Count > 0)
                    _cmbAspect.SelectedIndex = Math.Min(Math.Max(old, 0), _cmbAspect.Items.Count - 1);
            }
            finally
            {
                _loading = false;
            }
            UpdateSummary(_list.CheckedItems.Count);
        }

        private void SetAll(bool check)
        {
            _loading = true;
            try
            {
                for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, check);
            }
            finally
            {
                _loading = false;
            }
            UpdateSummary(check ? _list.Items.Count : 0);
        }

        private void UpdateSummary(int count)
        {
            _lblCount.Text = "已选 " + count + " 人";
            _btnAdd.Text = "为 " + count + " 人添加";
            string item = _txtItem == null ? "" : _txtItem.Text.Trim();
            bool ok = count > 0 && item.Length > 0 && _numDelta.Value != 0m;
            _btnAdd.Enabled = ok;
            if (!ok)
            {
                _lblPreview.Text = count == 0
                    ? "请先在左侧勾选需要添加的同学。"
                    : (item.Length == 0 ? "请填写事项内容。" : "加减分不能为 0。");
                _lblPreview.ForeColor = Theme.Muted;
                return;
            }
            string moduleName = _cmbModule.SelectedIndex >= 0 ? _data.Modules[_cmbModule.SelectedIndex].Name : "";
            string aspect = _cmbAspect.SelectedIndex >= 0 ? ClassData.AspectLabel(_cmbAspect.SelectedIndex + 1) : "";
            _lblPreview.Text = string.Format(
                "将为 {0} 名同学添加：\r\n{1} · {2} · {3}\r\n\r\n每人该模块分 {4}，平均分分别自动更新。",
                count, moduleName, aspect, item, Format.ScoreWithSign(_numDelta.Value));
            _lblPreview.ForeColor = _numDelta.Value < 0 ? Theme.Red : Theme.Accent;
        }

        private void DoAdd()
        {
            int count = _list.CheckedItems.Count;
            string item = _txtItem.Text.Trim();
            if (count == 0) { MessageBox.Show(this, "请至少勾选一名同学。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (item.Length == 0) { _txtItem.Focus(); MessageBox.Show(this, "请填写事项内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (_numDelta.Value == 0m) { MessageBox.Show(this, "加减分不能为 0。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (_cmbModule.SelectedIndex < 0 || _cmbAspect.SelectedIndex < 0) return;

            string moduleName = _data.Modules[_cmbModule.SelectedIndex].Name;
            string aspect = ClassData.AspectLabel(_cmbAspect.SelectedIndex + 1);
            string message = string.Format(
                "确定为 {0} 名同学添加吗？\r\n\r\n模块：{1}\r\n方面：{2}\r\n事项：{3}\r\n加减分：{4}\r\n\r\n每位同学对应模块分都会 {5}。",
                count, moduleName, aspect, item, Format.ScoreWithSign(_numDelta.Value), Format.ScoreWithSign(_numDelta.Value));
            if (MessageBox.Show(this, message, "批量添加", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            BatchAddRequest req = new BatchAddRequest();
            for (int i = 0; i < _list.CheckedItems.Count; i++)
            {
                Student s = _list.CheckedItems[i] as Student;
                if (s != null) req.Students.Add(s);
            }
            if (req.Students.Count == 0) return;
            req.ModuleIndex = _cmbModule.SelectedIndex;
            req.AspectIndex = _cmbAspect.SelectedIndex + 1;
            req.Item = item;
            req.Delta = _numDelta.Value;
            Result = req;
            DialogResult = DialogResult.OK;
        }
    }
}



