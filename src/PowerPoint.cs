using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OscpointHelper
{
    public class MediaStatus
    {
        public bool HasMedia;
        public bool Playing;
        public bool Paused;
        public int RemainingMs;
    }

    /// <summary>
    /// Late-bound COM access to the running PowerPoint (no interop assemblies needed).
    /// All calls must be made on the UI (STA) thread.
    /// </summary>
    public class PowerPoint
    {
        // PowerPoint constants
        const int ppShowAll = 1;
        const int ppShowNamedSlideShow = 3;
        const int msoMedia = 16;
        const int msoPlaceholder = 14;
        const int ppPlaying = 2;
        const int ppPaused = 1;

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        static extern int CLSIDFromProgID(string progId, out Guid clsid);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        dynamic _app;

        dynamic App
        {
            get
            {
                if (_app != null)
                {
                    try { var _ = _app.Version; return _app; }
                    catch { _app = null; }
                }
                try
                {
                    if (CLSIDFromProgID("PowerPoint.Application", out var clsid) != 0) return null;
                    GetActiveObject(ref clsid, IntPtr.Zero, out object o);
                    _app = o;
                    Log.Write("Attached to PowerPoint " + (string)_app.Version);
                }
                catch { _app = null; }
                return _app;
            }
        }

        public bool IsRunning => App != null;

        dynamic ActivePresentation
        {
            get
            {
                var app = App;
                if (app == null) return null;
                try { return app.Presentations.Count > 0 ? app.ActivePresentation : null; }
                catch { return null; }
            }
        }

        /// <summary>The slide show window for the active presentation, else the first one running.</summary>
        dynamic ShowWindow
        {
            get
            {
                var app = App;
                if (app == null) return null;
                try
                {
                    if (app.SlideShowWindows.Count == 0) return null;
                    try
                    {
                        var pres = ActivePresentation;
                        if (pres != null) { var sw = pres.SlideShowWindow; if (sw != null) return sw; }
                    }
                    catch { /* active deck isn't the one in slide show */ }
                    return app.SlideShowWindows.Item(1);
                }
                catch { return null; }
            }
        }

        public bool IsShowRunning => ShowWindow != null;

        // ---------- custom shows ----------

        public List<string> CustomShowNames()
        {
            var names = new List<string>();
            var pres = ActivePresentation;
            if (pres == null) return names;
            try
            {
                var shows = pres.SlideShowSettings.NamedSlideShows;
                int n = shows.Count;
                for (int i = 1; i <= n; i++) names.Add((string)shows.Item(i).Name);
            }
            catch (Exception ex) { Log.Write("CustomShowNames: " + ex.Message); }
            return names;
        }

        public bool StartCustomShow(string name, bool endRunningShow)
        {
            var pres = ActivePresentation;
            if (pres == null) { Log.Write("StartCustomShow: no presentation open"); return false; }

            // case-insensitive match against the real names
            string match = null;
            foreach (var n in CustomShowNames())
                if (string.Equals(n, name.Trim(), StringComparison.OrdinalIgnoreCase)) { match = n; break; }
            if (match == null) { Log.Write($"StartCustomShow: no custom show called \"{name}\""); return false; }

            try
            {
                if (endRunningShow)
                {
                    var w = ShowWindow;
                    if (w != null) { w.View.Exit(); }
                }

                var s = pres.SlideShowSettings;
                // remember the deck's own settings so we don't permanently change them
                int oldRange = s.RangeType;
                string oldName = "";
                try { oldName = s.SlideShowName; } catch { }
                bool wasSaved = false;
                try { wasSaved = pres.Saved != 0; } catch { }

                s.RangeType = ppShowNamedSlideShow;
                s.SlideShowName = match;
                s.Run();

                try
                {
                    s.RangeType = oldRange;
                    if (oldRange == ppShowNamedSlideShow && oldName != "") s.SlideShowName = oldName;
                    if (wasSaved) pres.Saved = -1; // don't leave the deck flagged as modified
                }
                catch (Exception ex) { Log.Write("StartCustomShow: restore settings: " + ex.Message); }

                Log.Write($"Started custom show \"{match}\"");
                return true;
            }
            catch (Exception ex)
            {
                Log.Write("StartCustomShow failed: " + ex.Message);
                return false;
            }
        }

        // ---------- builds ----------

        /// <summary>Jump to click/build n on the current slide (0 = before first build).</summary>
        public bool GotoBuild(int click)
        {
            var w = ShowWindow;
            if (w == null) { Log.Write("GotoBuild: no slide show running"); return false; }
            try
            {
                w.View.GotoClick(Math.Max(0, click));
                return true;
            }
            catch (Exception ex) { Log.Write("GotoBuild failed: " + ex.Message); return false; }
        }

        /// <summary>Go to slide (presentation slide index) then to a build on it.</summary>
        public bool GotoSlideBuild(int slide, int click)
        {
            var w = ShowWindow;
            if (w == null) { Log.Write("GotoSlideBuild: no slide show running"); return false; }
            try
            {
                var view = w.View;
                int current = 0;
                try { current = view.Slide.SlideIndex; } catch { }
                if (current != slide) view.GotoSlide(slide, 0 /* msoFalse: don't reset */);
                if (click > 0 || current == slide) view.GotoClick(Math.Max(0, click));
                return true;
            }
            catch (Exception ex) { Log.Write("GotoSlideBuild failed: " + ex.Message); return false; }
        }

        public (int slide, int click, int clicks) Position()
        {
            var w = ShowWindow;
            if (w == null) return (0, 0, 0);
            try
            {
                var v = w.View;
                return ((int)v.Slide.SlideIndex, (int)v.GetClickIndex(), (int)v.GetClickCount());
            }
            catch { return (0, 0, 0); }
        }

        // ---------- media (for the remaining-time overlay) ----------

        public MediaStatus Media()
        {
            var st = new MediaStatus();
            var w = ShowWindow;
            if (w == null) return st;
            try
            {
                var view = w.View;
                var slide = view.Slide;
                var shapes = slide.Shapes;
                int count = shapes.Count;
                MediaStatus best = null;
                for (int i = 1; i <= count; i++)
                {
                    var sh = shapes.Item(i);
                    if (!IsMedia(sh)) continue;
                    var ms = ReadPlayer(view, sh);
                    if (ms == null) continue;
                    // prefer the one that's playing, then paused
                    if (best == null || (ms.Playing && !best.Playing) || (ms.Paused && !best.Playing && !best.Paused))
                        best = ms;
                }
                return best ?? st;
            }
            catch { return st; }
        }

        static bool IsMedia(dynamic sh)
        {
            try
            {
                int t = sh.Type;
                if (t == msoMedia) return true;
                if (t == msoPlaceholder)
                {
                    int ct = sh.PlaceholderFormat.ContainedType;
                    return ct == msoMedia;
                }
            }
            catch { }
            return false;
        }

        static MediaStatus ReadPlayer(dynamic view, dynamic sh)
        {
            try
            {
                var player = view.Player(sh.Id);
                int state = player.State;
                int pos = player.CurrentPosition; // ms
                var mf = sh.MediaFormat;
                int length = mf.Length;
                int end = 0, start = 0;
                try { end = mf.EndPoint; start = mf.StartPoint; } catch { }
                if (end <= 0) end = length;
                return new MediaStatus
                {
                    HasMedia = true,
                    Playing = state == ppPlaying,
                    Paused = state == ppPaused && pos > start,
                    RemainingMs = Math.Max(0, end - pos),
                };
            }
            catch { return null; }
        }

        public int ProcessId()
        {
            try
            {
                var ps = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
                return ps.Length > 0 ? ps[0].Id : 0;
            }
            catch { return 0; }
        }
    }
}
