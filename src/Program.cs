using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;

namespace OscpointHelper
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            using var mutex = new Mutex(true, "OSCPointHelper.SingleInstance", out bool first);
            if (!first)
            {
                MessageBox.Show("OSCPoint Helper is already running (see the system tray).", "OSCPoint Helper");
                return;
            }
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new HelperContext());
        }
    }

    /// <summary>
    /// Tray app. OSC arrives on a background thread and is marshalled to the UI (STA) thread,
    /// where all PowerPoint COM and UI Automation calls happen.
    /// </summary>
    class HelperContext : ApplicationContext
    {
        public const string Version = "0.1.1";

        readonly Settings _settings = Settings.Load();
        readonly PowerPoint _ppt = new PowerPoint();
        readonly PresenterView _pv;
        readonly Overlay _overlay;
        readonly NotifyIcon _tray;
        readonly Control _ui = new Control();
        readonly System.Windows.Forms.Timer _mediaTimer = new System.Windows.Forms.Timer { Interval = 200 };
        readonly System.Windows.Forms.Timer _statusTimer = new System.Windows.Forms.Timer { Interval = 2000 };
        OscServer _osc;
        IPEndPoint _replyTo;
        string _lastShowsJson = "";
        bool _overlayForcedHidden;

        public HelperContext()
        {
            var _ = _ui.Handle; // force handle creation so BeginInvoke works
            _pv = new PresenterView(_ppt);
            _overlay = new Overlay(_settings);

            _tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "OSCPoint Helper",
                Visible = true,
                ContextMenuStrip = BuildMenu(),
            };

            try
            {
                _osc = new OscServer(_settings.ListenPort);
                _osc.MessageReceived += (m, ep) => _ui.BeginInvoke(new Action(() => Handle(m, ep)));
                Log.Write($"OSCPoint Helper {Version} listening on UDP {_settings.ListenPort}");
            }
            catch (Exception ex)
            {
                Log.Write($"Could not open UDP port {_settings.ListenPort}: {ex.Message}");
                MessageBox.Show($"Could not listen on UDP port {_settings.ListenPort}.\n{ex.Message}\n\nChange ListenPort in {Settings.FilePath}",
                    "OSCPoint Helper", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            _mediaTimer.Tick += (s, e) => MediaTick();
            _mediaTimer.Start();
            _statusTimer.Tick += (s, e) => StatusTick();
            _statusTimer.Start();
        }

        ContextMenuStrip BuildMenu()
        {
            var m = new ContextMenuStrip();
            var header = new ToolStripMenuItem($"OSCPoint Helper {Version}  ·  UDP {_settings.ListenPort}") { Enabled = false };
            m.Items.Add(header);
            m.Items.Add(new ToolStripSeparator());

            var overlayOn = new ToolStripMenuItem("VT timer overlay") { CheckOnClick = true, Checked = _settings.Overlay.Enabled };
            overlayOn.CheckedChanged += (s, e) => { _settings.Overlay.Enabled = overlayOn.Checked; _settings.Save(); if (!overlayOn.Checked) _overlay.HideBox(); };
            m.Items.Add(overlayOn);
            m.Items.Add("Position / size the overlay…", null, (s, e) => _overlay.SetSetupMode(true));
            m.Items.Add(new ToolStripSeparator());
            m.Items.Add("Dump Presenter View controls (for troubleshooting)", null, (s, e) =>
            {
                var path = _pv.DumpTree();
                Process.Start(new ProcessStartInfo("notepad.exe", "\"" + path + "\"") { UseShellExecute = true });
            });
            m.Items.Add("Open settings file", null, (s, e) =>
                Process.Start(new ProcessStartInfo("notepad.exe", "\"" + Settings.FilePath + "\"") { UseShellExecute = true }));
            m.Items.Add("Open log", null, (s, e) =>
                Process.Start(new ProcessStartInfo("notepad.exe", "\"" + Log.FilePath + "\"") { UseShellExecute = true }));
            m.Items.Add(new ToolStripSeparator());
            m.Items.Add("Exit", null, (s, e) => ExitThread());
            return m;
        }

        protected override void ExitThreadCore()
        {
            _tray.Visible = false;
            _osc?.Dispose();
            _overlay.Close();
            base.ExitThreadCore();
        }

        // ---------------- OSC ----------------

        void Reply(string address, params object[] args)
        {
            if (_osc == null) return;
            IPEndPoint to = null;
            if (!string.IsNullOrWhiteSpace(_settings.ReplyHost) && IPAddress.TryParse(_settings.ReplyHost, out var ip))
                to = new IPEndPoint(ip, _settings.ReplyPort);
            else if (_replyTo != null)
                to = _replyTo;
            if (to != null) _osc.Send(to, new OscMessage(address, args));
        }

        static object Utf8(string s)
        {
            foreach (var c in s) if (c > 127) return Encoding.UTF8.GetBytes(s);
            return s;
        }

        void Handle(OscMessage m, IPEndPoint from)
        {
            const string P = "/oscpoint/helper/";
            if (!m.Address.StartsWith(P, StringComparison.OrdinalIgnoreCase)) return;
            _replyTo = new IPEndPoint(from.Address, _settings.ReplyPort);
            var cmd = m.Address.Substring(P.Length).ToLowerInvariant();
            Log.Write("OSC in: " + m);

            try
            {
                switch (cmd)
                {
                    case "ping":
                        Reply(P + "pong", Version);
                        SendStatus();
                        break;

                    case "customshow/list":
                        _lastShowsJson = "";
                        SendCustomShows();
                        break;

                    case "customshow/start":
                        // args: name [, endRunning(1/0) default 1]
                        _ppt.StartCustomShow(m.StringArg(0), m.IntArg(1, 1) != 0);
                        break;

                    case "goto/build":
                        _ppt.GotoBuild(m.IntArg(0));
                        break;

                    case "goto/slidebuild":
                        _ppt.GotoSlideBuild(m.IntArg(0, 1), m.IntArg(1, 0));
                        break;

                    case "notes/scroll":
                        // args: delta lines (+down / -up) [, page 1/0] - works with a rotary's +/- steps
                        _pv.ScrollNotes(m.IntArg(0, 1) * Math.Max(1, _settings.ScrollStepsPerClick), m.IntArg(1, 0) != 0);
                        break;
                    case "notes/scroll/up": _pv.ScrollNotes(-Math.Max(1, m.IntArg(0, 1)), false); break;
                    case "notes/scroll/down": _pv.ScrollNotes(Math.Max(1, m.IntArg(0, 1)), false); break;
                    case "notes/page/up": _pv.ScrollNotes(-1, true); break;
                    case "notes/page/down": _pv.ScrollNotes(1, true); break;
                    case "notes/top": _pv.ScrollNotesTo(true); break;
                    case "notes/bottom": _pv.ScrollNotesTo(false); break;

                    case "notes/font/larger": _pv.NotesFont(true, m.IntArg(0, 1)); break;
                    case "notes/font/smaller": _pv.NotesFont(false, m.IntArg(0, 1)); break;
                    case "notes/font":
                        // signed step count, for rotaries
                        int n = m.IntArg(0, 1);
                        if (n != 0) _pv.NotesFont(n > 0, Math.Abs(n));
                        break;

                    case "overlay/show": _overlayForcedHidden = false; _settings.Overlay.Enabled = true; _settings.Save(); break;
                    case "overlay/hide": _overlayForcedHidden = true; _overlay.HideBox(); break;
                    case "overlay/toggle":
                        _overlayForcedHidden = !_overlayForcedHidden;
                        if (_overlayForcedHidden) _overlay.HideBox();
                        break;
                    case "overlay/larger": _overlay.ScaleBox(1.15f); break;
                    case "overlay/smaller": _overlay.ScaleBox(1 / 1.15f); break;
                    case "overlay/setup": _overlay.SetSetupMode(m.IntArg(0, 1) != 0); break;

                    default:
                        Log.Write("Unknown helper command: " + m.Address);
                        break;
                }
            }
            catch (Exception ex) { Log.Write($"Error handling {m.Address}: {ex.Message}"); }
        }

        // ---------------- periodic work ----------------

        void MediaTick()
        {
            try
            {
                if (!_settings.Overlay.Enabled || _overlayForcedHidden) { _overlay.HideBox(); return; }
                if (_overlay.SetupMode) return;
                var st = _ppt.Media();
                bool show = st.HasMedia && (st.Playing || (st.Paused && _settings.Overlay.ShowWhenPaused));
                if (show) _overlay.ShowRemaining(st.RemainingMs);
                else _overlay.HideBox();
            }
            catch (Exception ex) { Log.Write("MediaTick: " + ex.Message); }
        }

        void StatusTick()
        {
            bool running = _ppt.IsRunning;
            _tray.Text = running ? "OSCPoint Helper – PowerPoint connected" : "OSCPoint Helper – PowerPoint not running";
            if (_replyTo == null) return;
            SendStatus();
            SendCustomShows();
        }

        void SendStatus()
        {
            bool running = _ppt.IsRunning;
            Reply("/oscpoint/helper/status", running ? (_ppt.IsShowRunning ? "slideshow" : "powerpoint") : "no-powerpoint");
        }

        void SendCustomShows()
        {
            var json = JsonSerializer.Serialize(_ppt.IsRunning ? _ppt.CustomShowNames() : new List<string>());
            if (json == _lastShowsJson) return; // only send on change
            _lastShowsJson = json;
            Reply("/oscpoint/helper/customshows", Utf8(json));
        }
    }
}
