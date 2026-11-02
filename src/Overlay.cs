using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OscpointHelper
{
    /// <summary>
    /// Always-on-top VT remaining-time box. Drag to move (unless locked); right-click for size/lock.
    /// Never takes focus, so it doesn't interfere with the slide show.
    /// </summary>
    public class Overlay : Form
    {
        readonly Settings _settings;
        string _time = "00:00";
        bool _warn;
        bool _setupMode;

        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, int w, int l);
        const int WM_NCLBUTTONDOWN = 0xA1, HTCAPTION = 0x2;
        const int WS_EX_NOACTIVATE = 0x08000000, WS_EX_TOOLWINDOW = 0x80, WS_EX_TOPMOST = 0x8;

        public Overlay(Settings settings)
        {
            _settings = settings;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;
            Location = new Point(settings.Overlay.X, settings.Overlay.Y);
            Opacity = Math.Clamp(settings.Overlay.Opacity, 0.2, 1.0);

            var menu = new ContextMenuStrip();
            menu.Items.Add("Bigger", null, (s, e) => ScaleBox(1.15f));
            menu.Items.Add("Smaller", null, (s, e) => ScaleBox(1 / 1.15f));
            var lockItem = new ToolStripMenuItem("Lock position") { CheckOnClick = true, Checked = settings.Overlay.Locked };
            lockItem.CheckedChanged += (s, e) => { _settings.Overlay.Locked = lockItem.Checked; _settings.Save(); };
            menu.Items.Add(lockItem);
            menu.Items.Add("Finish positioning", null, (s, e) => SetSetupMode(false));
            ContextMenuStrip = menu;

            MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Left || _settings.Overlay.Locked) return;
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
                _settings.Overlay.X = Left;
                _settings.Overlay.Y = Top;
                _settings.Save();
            };
            Relayout();
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST;
                return cp;
            }
        }

        public void ScaleBox(float factor)
        {
            _settings.Overlay.FontSize = Math.Clamp(_settings.Overlay.FontSize * factor, 12f, 400f);
            _settings.Save();
            Relayout();
        }

        public bool SetupMode => _setupMode;

        /// <summary>Setup mode: shows a placeholder so the box can be positioned and sized.</summary>
        public void SetSetupMode(bool on)
        {
            _setupMode = on;
            if (on) { _time = "00:00"; _warn = false; ShowBox(); Relayout(); }
        }

        public void ShowBox()
        {
            if (!Visible) Show();
            TopMost = true; // re-assert over the presenter view
        }

        public void ShowRemaining(int remainingMs)
        {
            if (_setupMode) return;
            int secs = (int)Math.Ceiling(remainingMs / 1000.0);
            var ts = TimeSpan.FromSeconds(secs);
            var text = ts.TotalHours >= 1 ? ts.ToString(@"h\:mm\:ss") : ts.ToString(@"mm\:ss");
            bool warn = secs <= _settings.Overlay.WarnSeconds;
            if (text != _time || warn != _warn)
            {
                bool relayout = text.Length != _time.Length;
                _time = text;
                _warn = warn;
                if (relayout) Relayout(); else Invalidate();
            }
            ShowBox();
        }

        public void HideBox()
        {
            if (_setupMode) return;
            if (Visible) Hide();
        }

        Font TimeFont => new Font(_settings.Overlay.FontName, _settings.Overlay.FontSize, FontStyle.Bold, GraphicsUnit.Pixel);
        Font LabelFont => new Font(_settings.Overlay.FontName, Math.Max(9f, _settings.Overlay.FontSize * 0.22f), FontStyle.Regular, GraphicsUnit.Pixel);

        void Relayout()
        {
            using var g = CreateGraphics();
            using var tf = TimeFont;
            using var lf = LabelFont;
            var ts = g.MeasureString(_time.Length > 5 ? "0:00:00" : "00:00", tf);
            var ls = _settings.Overlay.ShowLabel ? g.MeasureString(_settings.Overlay.Label, lf) : SizeF.Empty;
            int pad = (int)(_settings.Overlay.FontSize * 0.2f);
            Size = new Size((int)Math.Max(ts.Width, ls.Width) + pad * 2, (int)(ts.Height + ls.Height) + pad);
            Invalidate();
        }

        static Color Parse(string hex, Color fallback)
        {
            try { return ColorTranslator.FromHtml(hex); } catch { return fallback; }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var o = _settings.Overlay;
            var bg = _warn ? Parse(o.WarnBackColour, Color.DarkRed) : Parse(o.BackColour, Color.FromArgb(30, 30, 30));
            var fg = Parse(o.TextColour, Color.White);
            e.Graphics.Clear(bg);
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            using var tf = TimeFont;
            using var lf = LabelFont;
            using var brush = new SolidBrush(fg);
            float y = 0;
            int pad = (int)(o.FontSize * 0.2f);
            if (o.ShowLabel)
            {
                var label = _setupMode ? "DRAG ME · right-click to size" : o.Label;
                var ls = e.Graphics.MeasureString(label, lf);
                e.Graphics.DrawString(label, lf, brush, (Width - ls.Width) / 2, pad / 2f);
                y = ls.Height;
            }
            var ts = e.Graphics.MeasureString(_time, tf);
            e.Graphics.DrawString(_time, tf, brush, (Width - ts.Width) / 2, y);
        }
    }
}
