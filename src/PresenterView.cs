using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace OscpointHelper
{
    /// <summary>
    /// Drives PowerPoint's Presenter View through Windows UI Automation (the object model has no
    /// API for the notes pane). This depends on PowerPoint's UI, so it is the fragile part:
    /// use DumpTree() to see what this machine's PowerPoint actually exposes.
    /// </summary>
    public class PresenterView
    {
        readonly PowerPoint _ppt;
        AutomationElement _notesCache;

        public PresenterView(PowerPoint ppt) { _ppt = ppt; }

        [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(POINT p);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        const uint WM_MOUSEWHEEL = 0x020A;

        AutomationElement FindWindow()
        {
            int pid = _ppt.ProcessId();
            if (pid == 0) return null;
            var cond = new PropertyCondition(AutomationElement.ProcessIdProperty, pid);
            var wins = AutomationElement.RootElement.FindAll(TreeScope.Children, cond);
            AutomationElement byName = null;
            foreach (AutomationElement w in wins)
            {
                string cls = "", name = "";
                try { cls = w.Current.ClassName ?? ""; name = w.Current.Name ?? ""; } catch { continue; }
                if (cls == "PodiumParent") return w; // Presenter View window class (as far as known)
                if (byName == null && name.IndexOf("Presenter View", StringComparison.OrdinalIgnoreCase) >= 0) byName = w;
            }
            return byName;
        }

        static string Describe(AutomationElement e)
        {
            try
            {
                var c = e.Current;
                return $"{c.ControlType.ProgrammaticName.Replace("ControlType.", "")} name=\"{c.Name}\" class=\"{c.ClassName}\" id=\"{c.AutomationId}\"";
            }
            catch { return "(unavailable)"; }
        }

        AutomationElement FindNotes()
        {
            if (_notesCache != null)
            {
                try { var rect = _notesCache.Current.BoundingRectangle; if (!rect.IsEmpty) return _notesCache; }
                catch { }
                _notesCache = null;
            }
            var win = FindWindow();
            if (win == null) return null;

            var scrollable = win.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.IsScrollPatternAvailableProperty, true));
            AutomationElement best = null;
            double bestArea = 0;
            foreach (AutomationElement e in scrollable)
            {
                string label = "";
                try { label = (e.Current.Name + " " + e.Current.AutomationId).ToLowerInvariant(); } catch { continue; }
                if (label.Contains("note")) { best = e; break; }
                var r = e.Current.BoundingRectangle;
                // notes pane is on the right-hand side; ignore the slide thumbnails strip etc.
                if (r.Width * r.Height > bestArea) { bestArea = r.Width * r.Height; best = e; }
            }
            if (best == null)
            {
                // no scroll pattern exposed: fall back to anything called notes, for wheel scrolling
                var all = win.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                foreach (AutomationElement e in all)
                {
                    string label = "";
                    try { label = (e.Current.Name + " " + e.Current.AutomationId).ToLowerInvariant(); } catch { continue; }
                    if (label.Contains("note") && !label.Contains("click to add")) { best = e; break; }
                }
            }
            if (best != null) Log.Write("Notes pane: " + Describe(best));
            _notesCache = best;
            return best;
        }

        /// <summary>lines &gt; 0 scrolls down, &lt; 0 scrolls up. page = true uses page-sized steps.</summary>
        public bool ScrollNotes(int lines, bool page)
        {
            if (lines == 0) return true;
            var notes = FindNotes();
            if (notes == null) { Log.Write("ScrollNotes: Presenter View notes pane not found"); return false; }

            try
            {
                if (notes.TryGetCurrentPattern(ScrollPattern.Pattern, out object p))
                {
                    var sp = (ScrollPattern)p;
                    var amount = lines > 0
                        ? (page ? ScrollAmount.LargeIncrement : ScrollAmount.SmallIncrement)
                        : (page ? ScrollAmount.LargeDecrement : ScrollAmount.SmallDecrement);
                    for (int i = 0; i < Math.Abs(lines); i++)
                    {
                        if (sp.Current.VerticallyScrollable) sp.ScrollVertical(amount);
                    }
                    return true;
                }
            }
            catch (Exception ex) { Log.Write("ScrollPattern failed, trying mouse wheel: " + ex.Message); }

            // Fallback: post wheel messages to the notes pane without moving the real mouse
            try
            {
                var r = notes.Current.BoundingRectangle;
                var pt = new POINT { X = (int)(r.Left + r.Width / 2), Y = (int)(r.Top + r.Height / 2) };
                var hwnd = WindowFromPoint(pt);
                int notches = Math.Abs(lines) * (page ? 5 : 1);
                int delta = lines > 0 ? -120 : 120;
                for (int i = 0; i < notches; i++)
                {
                    var wParam = (IntPtr)(delta << 16);
                    var lParam = (IntPtr)((pt.Y << 16) | (pt.X & 0xFFFF));
                    PostMessage(hwnd, WM_MOUSEWHEEL, wParam, lParam);
                }
                return true;
            }
            catch (Exception ex) { Log.Write("ScrollNotes failed: " + ex.Message); return false; }
        }

        public bool ScrollNotesTo(bool top)
        {
            var notes = FindNotes();
            if (notes == null) return false;
            try
            {
                if (notes.TryGetCurrentPattern(ScrollPattern.Pattern, out object p))
                {
                    ((ScrollPattern)p).SetScrollPercent(ScrollPattern.NoScroll, top ? 0 : 100);
                    return true;
                }
            }
            catch (Exception ex) { Log.Write("ScrollNotesTo: " + ex.Message); }
            return ScrollNotes(top ? -50 : 50, true);
        }

        /// <summary>Presses Presenter View's own "make text larger/smaller" buttons.</summary>
        public bool NotesFont(bool larger, int times)
        {
            var win = FindWindow();
            if (win == null) { Log.Write("NotesFont: Presenter View not found"); return false; }
            var buttons = win.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            string[] keys = larger
                ? new[] { "larger", "increase", "bigger", "grow" }
                : new[] { "smaller", "decrease", "shrink" };
            foreach (AutomationElement b in buttons)
            {
                string label = "";
                try { label = (b.Current.Name + " " + b.Current.HelpText + " " + b.Current.AutomationId).ToLowerInvariant(); }
                catch { continue; }
                foreach (var k in keys)
                {
                    if (!label.Contains(k)) continue;
                    if (b.TryGetCurrentPattern(InvokePattern.Pattern, out object ip))
                    {
                        for (int i = 0; i < Math.Max(1, times); i++) ((InvokePattern)ip).Invoke();
                        return true;
                    }
                }
            }
            Log.Write("NotesFont: text size button not found - run 'Dump Presenter View controls' and send the file");
            return false;
        }

        /// <summary>Writes Presenter View's UI Automation tree to a text file (for fixing the above).</summary>
        public string DumpTree()
        {
            var path = Path.Combine(Settings.Folder, "presenter-view-controls.txt");
            var sb = new StringBuilder();
            var win = FindWindow();
            if (win == null)
            {
                sb.AppendLine("Presenter View window not found. Top-level PowerPoint windows:");
                int pid = _ppt.ProcessId();
                if (pid != 0)
                    foreach (AutomationElement w in AutomationElement.RootElement.FindAll(TreeScope.Children,
                                 new PropertyCondition(AutomationElement.ProcessIdProperty, pid)))
                        sb.AppendLine("  " + Describe(w));
            }
            else Walk(win, 0, sb);
            File.WriteAllText(path, sb.ToString());
            Log.Write("Wrote " + path);
            return path;
        }

        static void Walk(AutomationElement e, int depth, StringBuilder sb)
        {
            if (depth > 12) return;
            var patterns = "";
            try
            {
                foreach (var p in e.GetSupportedPatterns()) patterns += p.ProgrammaticName.Replace("PatternIdentifiers.Pattern", "") + " ";
            }
            catch { }
            string rect = "";
            try { var r = e.Current.BoundingRectangle; rect = $"[{(int)r.Left},{(int)r.Top} {(int)r.Width}x{(int)r.Height}]"; } catch { }
            sb.AppendLine(new string(' ', depth * 2) + Describe(e) + " " + rect + (patterns != "" ? "  {" + patterns.Trim() + "}" : ""));
            var walker = TreeWalker.RawViewWalker;
            AutomationElement child = null;
            try { child = walker.GetFirstChild(e); } catch { }
            while (child != null)
            {
                Walk(child, depth + 1, sb);
                try { child = walker.GetNextSibling(child); } catch { break; }
            }
        }
    }
}
