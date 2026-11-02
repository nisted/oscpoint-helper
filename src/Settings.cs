using System;
using System.IO;
using System.Text.Json;

namespace OscpointHelper
{
    public class OverlaySettings
    {
        public bool Enabled { get; set; } = true;
        /// <summary>Show while media is paused as well as playing.</summary>
        public bool ShowWhenPaused { get; set; } = true;
        public int X { get; set; } = 40;
        public int Y { get; set; } = 40;
        public float FontSize { get; set; } = 48f;
        public string FontName { get; set; } = "Segoe UI";
        public string TextColour { get; set; } = "#FFFFFF";
        public string BackColour { get; set; } = "#1E1E1E";
        public string WarnBackColour { get; set; } = "#C00000";
        public int WarnSeconds { get; set; } = 10;
        public double Opacity { get; set; } = 0.9;
        public bool ShowLabel { get; set; } = true;
        public string Label { get; set; } = "VT REMAINING";
        public bool Locked { get; set; } = false;
    }

    public class Settings
    {
        /// <summary>UDP port the helper listens on for OSC commands.</summary>
        public int ListenPort { get; set; } = 35553;
        /// <summary>Port feedback is sent back to, on the IP that last sent a command (the Companion module's listening port).</summary>
        public int ReplyPort { get; set; } = 35550;
        /// <summary>Optional fixed feedback IP; blank = reply to whoever last sent a command.</summary>
        public string ReplyHost { get; set; } = "";
        public int ScrollStepsPerClick { get; set; } = 1;
        public OverlaySettings Overlay { get; set; } = new OverlaySettings();

        public static string Folder =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OSCPointHelper");
        public static string FilePath => Path.Combine(Folder, "settings.json");

        static readonly JsonSerializerOptions Opts = new JsonSerializerOptions { WriteIndented = true };

        public static Settings Load()
        {
            try
            {
                Directory.CreateDirectory(Folder);
                if (File.Exists(FilePath))
                    return JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath)) ?? new Settings();
            }
            catch (Exception ex) { Log.Write("Settings load failed, using defaults: " + ex.Message); }
            var s = new Settings();
            s.Save();
            return s;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Folder);
                File.WriteAllText(FilePath, JsonSerializer.Serialize(this, Opts));
            }
            catch (Exception ex) { Log.Write("Settings save failed: " + ex.Message); }
        }
    }

    public static class Log
    {
        static readonly object Lock = new object();
        public static string FilePath => Path.Combine(Settings.Folder, "helper.log");

        public static void Write(string msg)
        {
            var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + msg;
            System.Diagnostics.Debug.WriteLine(line);
            lock (Lock)
            {
                try
                {
                    Directory.CreateDirectory(Settings.Folder);
                    var fi = new FileInfo(FilePath);
                    if (fi.Exists && fi.Length > 2_000_000) fi.Delete();
                    File.AppendAllText(FilePath, line + Environment.NewLine);
                }
                catch { /* logging must never break the app */ }
            }
        }
    }
}
