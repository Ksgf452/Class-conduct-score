using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CaoXing
{
    internal static class Theme
    {
        // Palette derived from the lavender / cream / gold logo.
        public static readonly Color Page = Color.FromArgb(247, 242, 250);
        public static readonly Color Card = Color.White;
        public static readonly Color Navy = Color.FromArgb(46, 35, 64);
        public static readonly Color NavyLight = Color.FromArgb(63, 48, 88);
        public static readonly Color Accent = Color.FromArgb(124, 92, 196);
        public static readonly Color AccentHover = Color.FromArgb(106, 74, 174);
        public static readonly Color Text = Color.FromArgb(43, 35, 51);
        public static readonly Color BodyText = Color.FromArgb(76, 64, 88);
        public static readonly Color Muted = Color.FromArgb(126, 114, 140);
        public static readonly Color Border = Color.FromArgb(230, 217, 239);
        public static readonly Color SoftBlue = Color.FromArgb(242, 235, 250);
        public static readonly Color Green = Color.FromArgb(46, 139, 87);
        public static readonly Color Red = Color.FromArgb(192, 57, 43);
        public static readonly Color Gold = Color.FromArgb(217, 164, 59);

        public static readonly Color HeaderSub = Color.FromArgb(217, 200, 235);
        public static readonly Color HeaderBorder = Color.FromArgb(93, 72, 118);
        public static readonly Color StatusBg = Color.FromArgb(239, 231, 247);
        public static readonly Color GridHeaderBg = Color.FromArgb(244, 238, 250);
        public static readonly Color GridHeaderFg = Color.FromArgb(92, 74, 108);
        public static readonly Color GridSelectionBg = Color.FromArgb(231, 221, 247);
        public static readonly Color GridSelectionFg = Color.FromArgb(74, 47, 110);
        public static readonly Color RowAlt = Color.FromArgb(251, 249, 253);
        public static readonly Color ChipBg = Color.FromArgb(245, 239, 250);
        public static readonly Color AvatarFrom = Color.FromArgb(183, 155, 238);
        public static readonly Color AvatarTo = Color.FromArgb(124, 92, 196);

        public static Font Font(float size)
        {
            return new Font("Microsoft YaHei UI", size, FontStyle.Regular, GraphicsUnit.Point);
        }

        public static Font Font(float size, FontStyle style)
        {
            return new Font("Microsoft YaHei UI", size, style, GraphicsUnit.Point);
        }

        public static void PrimaryButton(Button b)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = Accent;
            b.ForeColor = Color.White;
            b.Font = Font(10f, FontStyle.Bold);
            b.Cursor = Cursors.Hand;
            b.UseVisualStyleBackColor = false;
        }

        public static void SecondaryButton(Button b)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = Border;
            b.BackColor = Color.White;
            b.ForeColor = Text;
            b.Font = Font(9.5f);
            b.Cursor = Cursors.Hand;
            b.UseVisualStyleBackColor = false;
        }

        public static void HeaderButton(Button b, bool primary)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = primary ? 0 : 1;
            b.FlatAppearance.BorderColor = HeaderBorder;
            b.BackColor = primary ? Accent : Navy;
            b.ForeColor = Color.White;
            b.Font = Font(9.5f);
            b.Cursor = Cursors.Hand;
            b.UseVisualStyleBackColor = false;
            b.Height = 34;
            b.Margin = new Padding(4, 0, 0, 0);
        }

        public static void StyleGrid(DataGridView g)
        {
            g.EnableHeadersVisualStyles = false;
            g.BorderStyle = BorderStyle.None;
            g.BackgroundColor = Color.White;
            g.GridColor = Color.FromArgb(238, 229, 245);
            g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            g.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            g.RowHeadersVisible = false;
            g.AllowUserToAddRows = false;
            g.AllowUserToDeleteRows = false;
            g.AllowUserToResizeRows = false;
            g.AllowUserToOrderColumns = false;
            g.MultiSelect = false;
            g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            g.ReadOnly = true;
            g.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            g.RowTemplate.Height = 36;
            g.ColumnHeadersHeight = 36;
            g.ColumnHeadersDefaultCellStyle.BackColor = GridHeaderBg;
            g.ColumnHeadersDefaultCellStyle.ForeColor = GridHeaderFg;
            g.ColumnHeadersDefaultCellStyle.Font = Font(9f, FontStyle.Bold);
            g.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            g.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 0, 0, 0);
            g.DefaultCellStyle.BackColor = Color.White;
            g.DefaultCellStyle.ForeColor = Text;
            g.DefaultCellStyle.Font = Font(9.5f);
            g.DefaultCellStyle.SelectionBackColor = GridSelectionBg;
            g.DefaultCellStyle.SelectionForeColor = GridSelectionFg;
            g.DefaultCellStyle.Padding = new Padding(8, 0, 0, 0);
            g.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            g.AlternatingRowsDefaultCellStyle.BackColor = RowAlt;
            g.ShowCellToolTips = true;
        }

        public static void StyleCombo(ComboBox c)
        {
            c.FlatStyle = FlatStyle.Flat;
            c.DropDownStyle = ComboBoxStyle.DropDownList;
            c.Font = Font(9.5f);
            c.BackColor = Color.White;
            c.ForeColor = Text;
        }

        public static void StyleTextBox(TextBox t)
        {
            t.BorderStyle = BorderStyle.FixedSingle;
            t.Font = Font(10f);
            t.BackColor = Color.White;
            t.ForeColor = Text;
        }

        public static void SetCue(TextBox t, string cue)
        {
            try
            {
                t.Handle.ToString();
                SendMessage(t.Handle, 0x1501, (IntPtr)1, cue);
            }
            catch { }
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
    }

    internal class CardPanel : Panel
    {
        public int Radius = 12;
        public Color BorderColor = Theme.Border;

        public CardPanel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Padding = new Padding(16);
        }

        protected override void OnResize(EventArgs eventargs)
        {
            base.OnResize(eventargs);
            try
            {
                using (GraphicsPath path = Rounded(ClientRectangle, Radius))
                {
                    Region old = Region;
                    Region = new Region(path);
                    if (old != null) old.Dispose();
                }
            }
            catch { }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle r = ClientRectangle;
            r.Width -= 1;
            r.Height -= 1;
            using (GraphicsPath path = Rounded(r, Radius))
            using (Pen pen = new Pen(BorderColor, 1f))
            {
                e.Graphics.DrawPath(pen, path);
            }
        }

        internal static GraphicsPath Rounded(Rectangle r, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = radius * 2;
            if (d <= 0) { path.AddRectangle(r); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    internal class AvatarPanel : Panel
    {
        private string _text = "?";

        public AvatarPanel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            BackColor = Color.Transparent;
        }

        public string Initial
        {
            get { return _text; }
            set { _text = string.IsNullOrEmpty(value) ? "?" : value.Substring(0, 1); Invalidate(); }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle r = new Rectangle(2, 2, Width - 5, Height - 5);
            using (LinearGradientBrush brush = new LinearGradientBrush(r, Theme.AvatarFrom, Theme.AvatarTo, 45f))
                e.Graphics.FillEllipse(brush, r);
            using (Font f = Theme.Font(15f, FontStyle.Bold))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(_text, f, Brushes.White, r, sf);
            }
        }
    }

    internal static class DialogTheme
    {
        public static void Apply(Form f)
        {
            f.Font = Theme.Font(9.5f);
            f.BackColor = Theme.Page;
            f.StartPosition = FormStartPosition.CenterParent;
            f.FormBorderStyle = FormBorderStyle.FixedDialog;
            f.MaximizeBox = false;
            f.MinimizeBox = false;
            f.ShowInTaskbar = false;
        }
    }
}
