using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace OscpointHelper
{
    /// <summary>Minimal OSC 1.0 message encode/decode (i, f, s, b; bundles unpacked).</summary>
    public class OscMessage
    {
        public string Address;
        public List<object> Args = new List<object>();

        public OscMessage(string address, params object[] args)
        {
            Address = address;
            if (args != null) Args.AddRange(args);
        }

        public int IntArg(int i, int fallback = 0)
        {
            if (i >= Args.Count) return fallback;
            var a = Args[i];
            if (a is int n) return n;
            if (a is float f) return (int)Math.Round(f);
            if (a is string s && int.TryParse(s, out var p)) return p;
            return fallback;
        }

        public string StringArg(int i, string fallback = "")
        {
            if (i >= Args.Count) return fallback;
            var a = Args[i];
            if (a is string s) return s;
            if (a is byte[] b) return Encoding.UTF8.GetString(b); // OSCPoint convention: blob = UTF-8 text
            return Convert.ToString(a, System.Globalization.CultureInfo.InvariantCulture);
        }

        public override string ToString()
        {
            var parts = new List<string> { Address };
            foreach (var a in Args) parts.Add(a is byte[] b ? "\"" + Encoding.UTF8.GetString(b) + "\"" : Convert.ToString(a));
            return string.Join(" ", parts);
        }
    }

    public static class Osc
    {
        static void Pad(MemoryStream ms) { while (ms.Length % 4 != 0) ms.WriteByte(0); }

        static void WriteString(MemoryStream ms, string s)
        {
            var b = Encoding.UTF8.GetBytes(s);
            ms.Write(b, 0, b.Length);
            ms.WriteByte(0);
            Pad(ms);
        }

        static void WriteInt(MemoryStream ms, int v)
        {
            var b = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(v));
            ms.Write(b, 0, 4);
        }

        public static byte[] Encode(OscMessage m)
        {
            using var ms = new MemoryStream();
            WriteString(ms, m.Address);
            var tags = new StringBuilder(",");
            foreach (var a in m.Args)
            {
                if (a is int) tags.Append('i');
                else if (a is float) tags.Append('f');
                else if (a is byte[]) tags.Append('b');
                else tags.Append('s');
            }
            WriteString(ms, tags.ToString());
            foreach (var a in m.Args)
            {
                switch (a)
                {
                    case int i: WriteInt(ms, i); break;
                    case float f:
                        WriteInt(ms, BitConverter.ToInt32(BitConverter.GetBytes(f), 0));
                        break;
                    case byte[] b:
                        WriteInt(ms, b.Length);
                        ms.Write(b, 0, b.Length);
                        Pad(ms);
                        break;
                    default: WriteString(ms, Convert.ToString(a)); break;
                }
            }
            return ms.ToArray();
        }

        static string ReadString(byte[] d, ref int p)
        {
            int start = p;
            while (p < d.Length && d[p] != 0) p++;
            var s = Encoding.UTF8.GetString(d, start, p - start);
            p++;
            while (p % 4 != 0) p++;
            return s;
        }

        static int ReadInt(byte[] d, ref int p)
        {
            int v = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(d, p));
            p += 4;
            return v;
        }

        public static List<OscMessage> Decode(byte[] d)
        {
            var list = new List<OscMessage>();
            DecodeInto(d, 0, d.Length, list);
            return list;
        }

        static void DecodeInto(byte[] all, int offset, int length, List<OscMessage> list)
        {
            var d = new byte[length];
            Array.Copy(all, offset, d, 0, length);
            int p = 0;
            if (d.Length >= 8 && Encoding.ASCII.GetString(d, 0, 7) == "#bundle")
            {
                p = 16; // "#bundle\0" + 8-byte timetag
                while (p + 4 <= d.Length)
                {
                    int size = ReadInt(d, ref p);
                    if (size <= 0 || p + size > d.Length) break;
                    DecodeInto(d, p, size, list);
                    p += size;
                }
                return;
            }

            var msg = new OscMessage(ReadString(d, ref p));
            if (p >= d.Length) { list.Add(msg); return; }
            var tags = ReadString(d, ref p);
            for (int t = 1; t < tags.Length; t++)
            {
                switch (tags[t])
                {
                    case 'i': msg.Args.Add(ReadInt(d, ref p)); break;
                    case 'f':
                        msg.Args.Add(BitConverter.ToSingle(BitConverter.GetBytes(ReadInt(d, ref p)), 0));
                        break;
                    case 's': msg.Args.Add(ReadString(d, ref p)); break;
                    case 'b':
                        int len = ReadInt(d, ref p);
                        var b = new byte[len];
                        Array.Copy(d, p, b, 0, len);
                        p += len;
                        while (p % 4 != 0) p++;
                        msg.Args.Add(b);
                        break;
                    case 'T': msg.Args.Add(1); break;
                    case 'F': msg.Args.Add(0); break;
                    default: break; // unsupported tag: ignore
                }
            }
            list.Add(msg);
        }
    }

    /// <summary>UDP OSC server. Callback runs on a background thread.</summary>
    public class OscServer : IDisposable
    {
        readonly UdpClient _udp;
        readonly UdpClient _send = new UdpClient();
        bool _running = true;

        public event Action<OscMessage, IPEndPoint> MessageReceived;

        public OscServer(int port)
        {
            _udp = new UdpClient(new IPEndPoint(IPAddress.Any, port));
            var t = new System.Threading.Thread(Loop) { IsBackground = true, Name = "OSC listener" };
            t.Start();
        }

        void Loop()
        {
            while (_running)
            {
                try
                {
                    var ep = new IPEndPoint(IPAddress.Any, 0);
                    var data = _udp.Receive(ref ep);
                    foreach (var m in Osc.Decode(data)) MessageReceived?.Invoke(m, ep);
                }
                catch (ObjectDisposedException) { return; }
                catch (Exception ex) { Log.Write("OSC receive error: " + ex.Message); }
            }
        }

        public void Send(IPEndPoint to, OscMessage m)
        {
            try
            {
                var b = Osc.Encode(m);
                _send.Send(b, b.Length, to);
            }
            catch (Exception ex) { Log.Write("OSC send error: " + ex.Message); }
        }

        public void Dispose()
        {
            _running = false;
            _udp.Dispose();
            _send.Dispose();
        }
    }
}
