// WordNote —— 一边看视频一边记 Word 笔记的小工具
// 截图后自动插入当前 Word 文档的光标处并换行，不用切窗口、不用 Ctrl+V。
// 编译：build.cmd（.NET Framework 自带 csc，无需安装任何东西）
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace WordNote
{
    // ---------------- 配置 ----------------
    internal static class Config
    {
        public static string SnipKey = "Ctrl+Alt+S";
        public static string ReselectKey = "Ctrl+Alt+Shift+S";
        public static string PlayKey = "Ctrl+Alt+Space";
        public static string FocusKey = "Ctrl+Alt+W";
        public static string AutoKey = "Ctrl+Alt+G";
        // 速记框：随时按一下弹出来打字，回车自动写进 Word
        public static string NoteKey = "Ctrl+Alt+N";
        // 速记框上次的位置 / 大小 / 透明度
        public static int NoteX = -1;
        public static int NoteY = -1;
        public static int NoteW = -1;
        public static int NoteH = -1;
        public static double NoteOpacity = 0.72;
        public static int NoteRadius = 14;
        public static double NoteFont = 12;
        public static string NoteTheme = "dark";
        public static string NoteCue = "输入文字，回车自动写进 Word";
        public static string NoteHint = "↵ 写入 Word   ·   按住这里拖动   ·   边角可缩放";
        // 回车之后：true = 常驻（不关，继续打，Esc / 热键才收起）；false = 老行为，写完就收
        public static bool NoteStay = true;
        public static bool AutoMode = true;
        public static int LinesAfter = 1;
        public static int ImageWidthPercent = 100;
        // 走“剪贴板 + Ctrl+V”备用通道时，把图缩到这个像素宽度（96dpi 下约 375 磅）
        public static int PasteWidthPx = 500;
        public static Rectangle Region = Rectangle.Empty;
        public static string WordProgID = "Word.Application";
        public static string WordClass = "OpusApp";
        // 用户指定的插入目标窗口（0 = 自动挑最合适的可见 Word 窗口）
        public static long WordHwnd = 0;
        public static string WordTitle = "";

        public static string IniPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordNote.ini");
        }

        public static void Load()
        {
            string ini = IniPath();
            if (!File.Exists(ini)) { Save(); return; }
            try
            {
                string[] lines = File.ReadAllLines(ini, Encoding.UTF8);
                foreach (string raw in lines)
                {
                    string line = raw.Trim();
                    if (line.Length == 0) continue;
                    if (line[0] == ';' || line[0] == '#' || line[0] == '[') continue;
                    int eq = line.IndexOf('=');
                    if (eq < 0) continue;
                    string k = line.Substring(0, eq).Trim().ToLowerInvariant();
                    string v = line.Substring(eq + 1).Trim();
                    switch (k)
                    {
                        case "snip": SnipKey = v; break;
                        case "reselect": ReselectKey = v; break;
                        case "play_pause": PlayKey = v; break;
                        case "focus": FocusKey = v; break;
                        case "auto": AutoKey = v; break;
                        case "note": NoteKey = v; break;
                        case "note_x":
                            {
                                int n;
                                if (int.TryParse(v, out n)) NoteX = n;
                                break;
                            }
                        case "note_y":
                            {
                                int n;
                                if (int.TryParse(v, out n)) NoteY = n;
                                break;
                            }
                        case "note_w":
                            {
                                int n;
                                if (int.TryParse(v, out n) && n >= 200) NoteW = n;
                                break;
                            }
                        case "note_h":
                            {
                                int n;
                                if (int.TryParse(v, out n) && n >= 60) NoteH = n;
                                break;
                            }
                        case "note_opacity":
                            {
                                double d;
                                if (double.TryParse(v, System.Globalization.NumberStyles.Float,
                                        System.Globalization.CultureInfo.InvariantCulture, out d) && d >= 0.3 && d <= 1.0)
                                    NoteOpacity = d;
                                break;
                            }
                        case "note_radius":
                            {
                                int n;
                                if (int.TryParse(v, out n) && n >= 0 && n <= 40) NoteRadius = n;
                                break;
                            }
                        case "note_font":
                            {
                                double d;
                                if (double.TryParse(v, System.Globalization.NumberStyles.Float,
                                        System.Globalization.CultureInfo.InvariantCulture, out d) && d >= 8 && d <= 24)
                                    NoteFont = d;
                                break;
                            }
                        case "note_theme":
                            NoteTheme = (v.ToLowerInvariant() == "light") ? "light" : "dark";
                            break;
                        case "note_cue": NoteCue = v; break;
                        case "note_hint": NoteHint = v; break;
                        case "note_stay": NoteStay = (v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase)); break;
                        case "auto_mode": AutoMode = (v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase)); break;
                        case "lines_after": int.TryParse(v, out LinesAfter); break;
                        case "image_width_percent":
                            {
                                int n;
                                if (int.TryParse(v, out n) && n >= 20 && n <= 100) ImageWidthPercent = n;
                                break;
                            }
                        case "paste_width_px":
                            {
                                int n;
                                if (int.TryParse(v, out n) && n >= 200 && n <= 1600) PasteWidthPx = n;
                                break;
                            }
                        case "region":
                            {
                                string[] p = v.Split(',');
                                if (p.Length == 4)
                                {
                                    int x, y, w, h;
                                    if (int.TryParse(p[0], out x) && int.TryParse(p[1], out y) &&
                                        int.TryParse(p[2], out w) && int.TryParse(p[3], out h) && w > 0 && h > 0)
                                        Region = new Rectangle(x, y, w, h);
                                }
                                break;
                            }
                        case "word_progid": WordProgID = v; break;
                        case "word_class": WordClass = v; break;
                case "word_hwnd": try { WordHwnd = long.Parse(v); } catch { WordHwnd = 0; } break;
                case "word_title": WordTitle = v; break;
                    }
                }
            }
            catch { }
        }

        public static void Save()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("; WordNote 配置 —— 改完关掉程序重新打开生效");
                sb.AppendLine("; 热键写法：Ctrl+Alt+Z / Shift+F8 / Alt+Space ...");
                sb.AppendLine();
                sb.AppendLine("snip=" + SnipKey);
                sb.AppendLine("reselect=" + ReselectKey);
                sb.AppendLine("play_pause=" + PlayKey);
                sb.AppendLine("focus=" + FocusKey);
                sb.AppendLine("auto=" + AutoKey);
                sb.AppendLine("note=" + NoteKey);
                sb.AppendLine("; 速记框上次的位置和大小（拖动/缩放后自动记下，-1 = 自动）");
                sb.AppendLine("note_x=" + NoteX);
                sb.AppendLine("note_y=" + NoteY);
                sb.AppendLine("note_w=" + NoteW);
                sb.AppendLine("note_h=" + NoteH);
                sb.AppendLine("; 速记框外观：透明度 0.3-1.0 / 圆角 px / 字号 pt / 主题 dark|light");
                sb.AppendLine("note_opacity=" + NoteOpacity.ToString("0.0#", System.Globalization.CultureInfo.InvariantCulture));
                sb.AppendLine("note_radius=" + NoteRadius);
                sb.AppendLine("note_font=" + NoteFont.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture));
                sb.AppendLine("note_theme=" + NoteTheme);
                sb.AppendLine("; 输入框里的灰色提示字 / 底部那行小字");
                sb.AppendLine("note_cue=" + NoteCue);
                sb.AppendLine("note_hint=" + NoteHint);
                sb.AppendLine("; 1=回车后保持常驻（继续打，Esc 关） 0=写完自动收起");
                sb.AppendLine("note_stay=" + (NoteStay ? "1" : "0"));
                sb.AppendLine();
                sb.AppendLine("; 1=开 0=关：任何截图工具截的图都自动进 Word");
                sb.AppendLine("auto_mode=" + (AutoMode ? "1" : "0"));
                sb.AppendLine("; 图片后面空几行");
                sb.AppendLine("lines_after=" + LinesAfter);
                sb.AppendLine("; 图片宽度占版心宽度的百分比 20-100（COM 通道用）");
                sb.AppendLine("image_width_percent=" + ImageWidthPercent);
                sb.AppendLine("; 备用通道（剪贴板+Ctrl+V）把图缩到的像素宽度");
                sb.AppendLine("paste_width_px=" + PasteWidthPx);
                sb.AppendLine("; 秒截区域 x,y,w,h（留 0 表示还没设置）");
                sb.AppendLine("region=" + Region.X + "," + Region.Y + "," + Region.Width + "," + Region.Height);
                sb.AppendLine();
                sb.AppendLine("; 用 WPS 的话把下面改成 KWPS.Application（不确定就先用 Word）");
                sb.AppendLine("word_progid=" + WordProgID);
                sb.AppendLine("word_class=" + WordClass);
                sb.AppendLine("; 插入目标窗口（托盘里点“选择 Word 窗口…”设置；0=自动挑）");
                sb.AppendLine("word_hwnd=" + WordHwnd);
                sb.AppendLine("word_title=" + WordTitle);
                File.WriteAllText(IniPath(), sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        public static void OpenIni()
        {
            try
            {
                if (!File.Exists(IniPath())) Save();
                Process.Start("notepad.exe", "\"" + IniPath() + "\"");
            }
            catch { }
        }

        public static string TempDir()
        {
            // 优先放程序目录下（自己的文件夹一般可写），不行再退回系统临时目录
            string p = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "_shots");
            try { Directory.CreateDirectory(p); return p; }
            catch
            {
                p = Path.Combine(Path.GetTempPath(), "WordNote");
                try { Directory.CreateDirectory(p); return p; }
                catch { return AppDomain.CurrentDomain.BaseDirectory; }
            }
        }

        public static string TempFile()
        {
            Directory.CreateDirectory(TempDir());
            return Path.Combine(TempDir(), DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".png");
        }

        // 调试日志（只写入程序目录下的 WordNote.log，出问题看这个文件）
        public static void Log(string s)
        {
            try
            {
                string p = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordNote.log");
                if (File.Exists(p) && new FileInfo(p).Length > 300000) File.Delete(p);
                File.AppendAllText(p, DateTime.Now.ToString("MM-dd HH:mm:ss.fff") + "  " + s + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }

        public static void CleanTemp()
        {
            try
            {
                Directory.CreateDirectory(TempDir());
                foreach (string f in Directory.GetFiles(TempDir(), "*.png"))
                {
                    try
                    {
                        if (File.GetLastWriteTime(f) < DateTime.Now.AddDays(-2)) File.Delete(f);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }

    // ---------------- Win32 ----------------
    internal static class Native
    {
        [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mods, uint vk);

        [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr hdc, int x, int y, int w, int h,
            IntPtr hdcSrc, int x1, int y1, int rop);

        // 抓屏：BitBlt 直接拿屏幕 DC，比 Graphics.CopyFromScreen 更可控（也省得用非法枚举值）
        public static bool CaptureRect(Bitmap bmp, Rectangle r)
        {
            const int SRCCOPY = 0x00CC0020;
            bool ok = false;
            using (Graphics g = Graphics.FromImage(bmp))
            {
                IntPtr dst = g.GetHdc();
                IntPtr src = GetDC(IntPtr.Zero);
                try { ok = BitBlt(dst, 0, 0, r.Width, r.Height, src, r.X, r.Y, SRCCOPY); }
                finally { ReleaseDC(IntPtr.Zero, src); g.ReleaseHdc(dst); }
            }
            return ok;
        }

        [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder s, int max);
        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
        [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr p);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
        public delegate bool EnumProc(IntPtr hWnd, IntPtr p);
        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

        // Word 会有好几个 OpusApp 窗口（隐藏的、屏外的），FindWindow 拿到的往往是没用的那个。
        // 这里只在“可见 + 有标题”的窗口里挑：优先前台那个，否则挑面积最大的。
        public static IntPtr BestWordWindow(string cls)
        {
            IntPtr fg = GetForegroundWindow();
            if (fg != IntPtr.Zero && ClassName(fg) == cls && IsWindowVisible(fg) && Title(fg).Length > 0)
                return fg;

            IntPtr best = IntPtr.Zero;
            int bestArea = 0;
            EnumWindows(delegate(IntPtr h, IntPtr p)
            {
                if (!IsWindowVisible(h)) return true;
                if (ClassName(h) != cls) return true;
                if (Title(h).Length == 0) return true;
                RECT r;
                if (!GetWindowRect(h, out r)) return true;
                int area = (r.Right - r.Left) * (r.Bottom - r.Top);
                if (area > bestArea) { bestArea = area; best = h; }
                return true;
            }, IntPtr.Zero);
            return best;
        }

        public static string Title(IntPtr hWnd)
        {
            try
            {
                StringBuilder sb = new StringBuilder(GetWindowTextLength(hWnd) + 2);
                GetWindowText(hWnd, sb, sb.Capacity);
                return sb.ToString();
            }
            catch { return ""; }
        }

        // ---- 通过 MSAA 的“原生对象模型”拿到指定窗口所属的 Word 实例 ----
        // Word 的编辑窗格（_WwG）支持 OBJID_NATIVEOM，拿到的就是那扇窗口自己的 Window 对象，
        // 不用 ROT，也不怕同时开着好几个 Word。
        [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id,
            ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        private static Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0u;

        public static object NativeOM(IntPtr hWnd)
        {
            object o = null;
            Guid g = IID_IDispatch;
            int hr = AccessibleObjectFromWindow(hWnd, OBJID_NATIVEOM, ref g, out o);
            if (hr != 0 || o == null)
                throw new Exception("AccessibleObjectFromWindow 失败 hr=0x" + hr.ToString("X8"));
            return o;
        }

        // 在 Word 主窗口里找到承载文档的编辑窗格（_WwG）
        public static IntPtr FindEditPane(IntPtr frame)
        {
            if (ClassName(frame) == "_WwG") return frame;
            IntPtr pane = IntPtr.Zero;
            EnumChildWindows(frame, delegate(IntPtr h, IntPtr p)
            {
                if (ClassName(h) == "_WwG") { pane = h; return false; }
                return true;
            }, IntPtr.Zero);
            return pane;
        }

        // 用户指定的窗口还在就用它；否则按标题找；再否则自动挑
        public static IntPtr ResolveWordWindow()
        {
            if (Config.WordHwnd != 0)
            {
                IntPtr picked = (IntPtr)Config.WordHwnd;
                if (IsWindowVisible(picked) && ClassName(picked) == Config.WordClass && Title(picked).Length > 0)
                    return picked;
            }
            if (Config.WordTitle != null && Config.WordTitle.Length > 0)
            {
                IntPtr byTitle = IntPtr.Zero;
                EnumWindows(delegate(IntPtr h, IntPtr p)
                {
                    if (!IsWindowVisible(h)) return true;
                    if (ClassName(h) != Config.WordClass) return true;
                    if (Title(h) == Config.WordTitle) { byTitle = h; return false; }
                    return true;
                }, IntPtr.Zero);
                if (byTitle != IntPtr.Zero) return byTitle;
            }
            return BestWordWindow(Config.WordClass);
        }
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr FindWindow(string cls, string title);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hWnd, bool altTab);
        [DllImport("user32.dll")] public static extern bool AddClipboardFormatListener(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool RemoveClipboardFormatListener(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetClipboardOwner();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, IntPtr extra);
        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();

        public static string ClassName(IntPtr h)
        {
            StringBuilder sb = new StringBuilder(256);
            GetClassName(h, sb, sb.Capacity);
            return sb.ToString();
        }
    }

    // ---------------- 隐藏消息窗口（收热键 + 剪贴板通知） ----------------
    internal class MsgWindow : Form
    {
        public Action<int> Hotkey;
        public Action ClipUpdate;

        public MsgWindow()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(1, 1);
            Location = new Point(-32000, -32000);
            CreateHandle();
            Native.AddClipboardFormatListener(Handle);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            e.Cancel = true;
            base.OnFormClosing(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)          // WM_HOTKEY
            {
                Action<int> h = Hotkey;
                if (h != null) h(m.WParam.ToInt32());
            }
            else if (m.Msg == 0x031D)     // WM_CLIPBOARDUPDATE
            {
                Action c = ClipUpdate;
                if (c != null) c();
            }
            base.WndProc(ref m);
        }
    }

    // ---------------- 速记框 ----------------
    // 造型参考 Raycast / Spotlight 的悬浮输入面板：
    // 无边框 + 圆角 + 可调透明度 + 细提示条（按住可拖动）+ 边角可缩放
    internal class NoteBox : Form
    {
        public Action<string> Submitted;
        public Action Cancelled;
        public Action Moved;

        private TextBox box;
        private Label hint;
        private Label cue;      // 输入框里的灰色提示字（自己画的，深浅主题都好控）

        private const int PAD = 10;      // 内容边距
        private const int EDGE = 6;      // 可缩放的边缘宽度

        private const int WM_NCHITTEST = 0x84;
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int WM_GETMINMAXINFO = 0x24;
        private const int WM_EXITSIZEMOVE = 0x232;
        private const int HTCAPTION = 2;
        private const int HTLEFT = 10, HTRIGHT = 11, HTTOP = 12, HTTOPLEFT = 13, HTTOPRIGHT = 14,
                          HTBOTTOM = 15, HTBOTTOMLEFT = 16, HTBOTTOMRIGHT = 17;
        private const int EM_SETCUEBANNER = 0x1501;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)]
        private struct MINMAXINFO { public POINT Reserved, MaxSize, MaxPosition, MinTrackSize, MaxTrackSize; }

        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr h, int msg, int w, int l);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageStr(IntPtr h, int msg, int w, string l);

        [DllImport("imm32.dll")] private static extern IntPtr ImmGetContext(IntPtr hWnd);
        [DllImport("imm32.dll")] private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
        [DllImport("imm32.dll", CharSet = CharSet.Unicode)] private static extern int ImmGetCompositionString(IntPtr hIMC, int index, IntPtr buf, int len);
        private const int GCS_COMPSTR = 8;

        public NoteBox()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;
            Size = new Size(Config.NoteW >= 360 ? Config.NoteW : 680, Config.NoteH >= 78 ? Config.NoteH : 82);

            box = new TextBox();
            box.BorderStyle = BorderStyle.None;
            box.KeyDown += OnBoxKeyDown;
            box.TextChanged += delegate { UpdateCue(); };
            Controls.Add(box);

            cue = new Label();
            cue.AutoSize = false;
            cue.TextAlign = ContentAlignment.MiddleLeft;
            cue.Cursor = Cursors.IBeam;
            cue.Click += delegate { box.Focus(); };
            Controls.Add(cue);
            cue.BringToFront();

            hint = new Label();
            hint.AutoSize = false;
            hint.AutoEllipsis = true;   // 框拉窄了文字就显示省略号，不硬切
            hint.Cursor = Cursors.SizeAll;
            hint.MouseDown += OnDragDown;
            Controls.Add(hint);

            Resize += delegate { LayoutChildren(); ApplyRound(); };

            // 右下角蓝色徽标兼做“缩放角”，按住就能拉大小
            MouseDown += OnGripDown;
            MouseMove += OnGripMove;
            MouseUp += OnGripUp;
            Cursor = Cursors.Default;

            ApplyStyle();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            ApplyCue();
        }

        private void ApplyCue()
        {
            try { SendMessageStr(box.Handle, EM_SETCUEBANNER, 1, Config.NoteCue); } catch { }
        }

        // 主题配色：深色 / 浅色
        private void ThemeColors(out Color panel, out Color text, out Color hintColor, out Color border)
        {
            if (Config.NoteTheme == "light")
            {
                panel = Color.FromArgb(252, 252, 252);
                text = Color.FromArgb(30, 30, 30);
                hintColor = Color.FromArgb(120, 120, 120);
                border = Color.FromArgb(60, 0, 0, 0);
            }
            else
            {
                panel = Color.FromArgb(28, 28, 32);
                text = Color.FromArgb(245, 245, 248);
                hintColor = Color.FromArgb(140, 140, 150);
                border = Color.FromArgb(80, 255, 255, 255);
            }
        }

        // 改样式（设置界面里拖动滑块时实时调用）
        public void ApplyStyle()
        {
            Color panel, text, hintColor, border;
            ThemeColors(out panel, out text, out hintColor, out border);

            BackColor = panel;
            box.BackColor = panel;
            box.ForeColor = text;
            box.Font = new Font("Microsoft YaHei", (float)Config.NoteFont);
            cue.BackColor = panel;
            cue.ForeColor = hintColor;
            cue.Font = box.Font;
            cue.Text = Config.NoteCue;
            hint.ForeColor = hintColor;
            hint.Font = new Font("Microsoft YaHei", (float)Math.Max(7.5, Config.NoteFont - 3.5));
            hint.Text = Config.NoteHint;
            if (Config.NoteStay) hint.Text = Config.NoteHint + " · Esc 关";
            try { Opacity = Config.NoteOpacity; } catch { }

            LayoutChildren();
            ApplyCue();
            ApplyRound();
            Invalidate();
        }

        private void LayoutChildren()
        {
            int textH = box.Font.Height + 4;
            if (textH > ClientSize.Height - 34) textH = Math.Max(18, ClientSize.Height - 34);
            box.Location = new Point(PAD + 2, PAD);
            box.Size = new Size(Math.Max(60, ClientSize.Width - 2 * PAD - 4), textH);

            cue.Location = new Point(PAD + 4, PAD);
            cue.Size = new Size(Math.Max(60, ClientSize.Width - 2 * PAD - 8), textH);
            UpdateCue();

            int hintH = 26;   // 底边整条都能按住拖
            hint.Location = new Point(PAD + 2, ClientSize.Height - hintH);
            int hintW = ClientSize.Width - 2 * PAD - 4 - 26;   // 右下角给徽标留位置
            hint.Size = new Size(Math.Max(60, hintW), hintH);
            hint.TextAlign = ContentAlignment.MiddleLeft;
        }

        private bool gripDrag;
        private int gripX, gripY, gripW, gripH;

        private bool InBadge(Point p)
        {
            int bs = 18;
            int bx = Width - PAD - bs - 1;
            int by = Height - PAD - bs - 1;
            return p.X >= bx - 4 && p.X <= bx + bs + 4 && p.Y >= by - 4 && p.Y <= by + bs + 4;
        }

        private void OnGripDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (!InBadge(e.Location)) return;
            gripDrag = true;
            gripX = Cursor.Position.X;
            gripY = Cursor.Position.Y;
            gripW = Width;
            gripH = Height;
            Capture = true;
            Cursor = Cursors.SizeNWSE;
        }

        private void OnGripMove(object sender, MouseEventArgs e)
        {
            if (!gripDrag) return;
            int w = gripW + (Cursor.Position.X - gripX);
            int h = gripH + (Cursor.Position.Y - gripY);
            if (w < 360) w = 360;
            if (h < 78) h = 78;
            Size = new Size(w, h);
        }

        private void OnGripUp(object sender, MouseEventArgs e)
        {
            if (!gripDrag) return;
            gripDrag = false;
            Capture = false;
            Cursor = Cursors.Default;
            SaveGeom();
        }

        private void UpdateCue()
        {
            if (cue == null || box == null) return;
            cue.Visible = box.TextLength == 0;
            if (cue.Visible) cue.BringToFront();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80;   // WS_EX_TOOLWINDOW，不进 Alt+Tab
                return cp;
            }
        }

        private static System.Drawing.Drawing2D.GraphicsPath RoundPath(Rectangle b, int r)
        {
            System.Drawing.Drawing2D.GraphicsPath p = new System.Drawing.Drawing2D.GraphicsPath();
            p.AddArc(b.Left, b.Top, r * 2, r * 2, 180, 90);
            p.AddArc(b.Right - r * 2 - 1, b.Top, r * 2, r * 2, 270, 90);
            p.AddArc(b.Right - r * 2 - 1, b.Bottom - r * 2 - 1, r * 2, r * 2, 0, 90);
            p.AddArc(b.Left, b.Bottom - r * 2 - 1, r * 2, r * 2, 90, 90);
            p.CloseFigure();
            return p;
        }

        private void ApplyRound()
        {
            if (Width <= 0 || Height <= 0) return;
            Region old = Region;
            int r = Config.NoteRadius;
            if (r < 0) r = 0;
            if (r > Math.Min(Width, Height) / 2) r = Math.Min(Width, Height) / 2;
            if (r <= 0) { Region = null; }
            else
            {
                using (System.Drawing.Drawing2D.GraphicsPath p = RoundPath(new Rectangle(0, 0, Width, Height), r))
                    Region = new Region(p);
            }
            if (old != null) old.Dispose();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            int r = Config.NoteRadius;
            if (r > Math.Min(Width, Height) / 2) r = Math.Min(Width, Height) / 2;
            Color panel, text, hintColor, border;
            ThemeColors(out panel, out text, out hintColor, out border);
            using (Pen pen = new Pen(border))
            {
                if (r > 0)
                {
                    using (System.Drawing.Drawing2D.GraphicsPath p = RoundPath(new Rectangle(0, 0, Width - 1, Height - 1), r))
                        e.Graphics.DrawPath(pen, p);
                }
                else
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                }
            }

            // 右下角小徽标：蓝底 + 三条白线（“笔记”意象）
            int bs = 18;
            int bx = Width - PAD - bs - 1;
            int by = Height - PAD - bs - 1;
            if (bx > 40)
            {
                Rectangle br = new Rectangle(bx, by, bs, bs);
                using (System.Drawing.Drawing2D.GraphicsPath gp = RoundPath(br, 4))
                {
                    using (SolidBrush fill = new SolidBrush(Color.FromArgb(235, 92, 141, 240)))
                        e.Graphics.FillPath(fill, gp);
                    using (Pen white = new Pen(Color.FromArgb(240, 255, 255, 255), 1.6f))
                    {
                        e.Graphics.DrawLine(white, bx + 4f, by + 5.5f, bx + 13.5f, by + 5.5f);
                        e.Graphics.DrawLine(white, bx + 4f, by + 9.5f, bx + 13.5f, by + 9.5f);
                        e.Graphics.DrawLine(white, bx + 4f, by + 13.5f, bx + 9.5f, by + 13.5f);
                    }
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCHITTEST)
            {
                long lp = m.LParam.ToInt64();
                int sx = unchecked((short)(lp & 0xFFFF));
                int sy = unchecked((short)((lp >> 16) & 0xFFFF));
                Point p = PointToClient(new Point(sx, sy));
                bool l = p.X < EDGE, r = p.X >= ClientSize.Width - EDGE;
                bool t = p.Y < EDGE, b = p.Y >= ClientSize.Height - EDGE;
                int ht = 0;
                if (t && l) ht = HTTOPLEFT;
                else if (t && r) ht = HTTOPRIGHT;
                else if (b && l) ht = HTBOTTOMLEFT;
                else if (b && r) ht = HTBOTTOMRIGHT;
                else if (l) ht = HTLEFT;
                else if (r) ht = HTRIGHT;
                else if (t) ht = HTTOP;
                else if (b) ht = HTBOTTOM;
                if (ht != 0) { m.Result = (IntPtr)ht; return; }
                if (InBadge(p)) { m.Result = (IntPtr)1; return; }   // 徽标：留给缩放角自己的鼠标逻辑
                // 其余自绘区域：整块面板都能拖动
                m.Result = (IntPtr)HTCAPTION;
                return;
            }
            else if (m.Msg == WM_GETMINMAXINFO)
            {
                MINMAXINFO mmi = (MINMAXINFO)m.GetLParam(typeof(MINMAXINFO));
                mmi.MinTrackSize.X = 360;
                mmi.MinTrackSize.Y = 78;
                Marshal.StructureToPtr(mmi, m.LParam, false);
            }
            else if (m.Msg == WM_EXITSIZEMOVE)
            {
                ApplyRound();
                SaveGeom();
            }
            base.WndProc(ref m);
        }

        private void SaveGeom()
        {
            // 存之前夹到工作区，避免记忆出一个永远在屏幕外的坐标
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            int x = Left, y = Top;
            if (x + Width > wa.Right) x = wa.Right - Width;
            if (y + Height > wa.Bottom) y = wa.Bottom - Height;
            if (x < wa.Left) x = wa.Left;
            if (y < wa.Top) y = wa.Top;
            Config.NoteX = x;
            Config.NoteY = y;
            Config.NoteW = Width;
            Config.NoteH = Height;
            if (Moved != null) Moved();
        }

        private void OnDragDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            SaveGeom();   // 拖动是模态的，返回时窗口已在最终位置
        }

        private static bool IsComposing(IntPtr hWnd)
        {
            IntPtr hIMC = ImmGetContext(hWnd);
            if (hIMC == IntPtr.Zero) return false;
            try { return ImmGetCompositionString(hIMC, GCS_COMPSTR, IntPtr.Zero, 0) > 0; }
            finally { ImmReleaseContext(hWnd, hIMC); }
        }

        private void OnBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                Hide();
                if (Cancelled != null) Cancelled();
                return;
            }
            if (e.KeyCode != Keys.Enter) return;

            // 输入法还在拼写时，回车是“上屏”，不要当提交
            if (IsComposing(box.Handle)) return;

            e.SuppressKeyPress = true;
            string text = box.Text.Trim();

            if (Config.NoteStay)
            {
                // 常驻模式：交出去之后清空输入框，窗口留着继续打，Esc 才关
                if (text.Length > 0 && Submitted != null) Submitted(text);
                if (box.TextLength > 0) box.Clear();
                UpdateCue();
                box.Focus();
                return;
            }

            Hide();
            if (text.Length == 0) return;
            if (Submitted != null) Submitted(text);
        }

        // 让打字焦点回到输入框（常驻模式下提交完要用）
        public void FocusInput()
        {
            try { PasteEngine.ForceForeground(Handle); } catch { }
            box.Focus();
            box.SelectionStart = box.TextLength;
        }

        public string CurrentText
        {
            get { return box.Text; }
        }

        public void ShowBox(string preset)
        {
            int w = Config.NoteW >= 360 ? Config.NoteW : 680;
            int h = Config.NoteH >= 78 ? Config.NoteH : 82;
            ClientSize = new Size(w, h);

            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            int x = Config.NoteX;
            int y = Config.NoteY;
            // 有记忆就用记忆的位置，只要大致还在屏幕里就“夹”回工作区，而不是丢掉
            if (x == -1 || y == -1 || x < wa.Left - 400 || y < wa.Top - 400 || x > wa.Right || y > wa.Bottom)
            {
                x = wa.Left + (wa.Width - Width) / 2;
                y = wa.Top + (int)(wa.Height * 0.62);
            }
            else
            {
                if (x + Width > wa.Right) x = wa.Right - Width;
                if (y + Height > wa.Bottom) y = wa.Bottom - Height;
                if (x < wa.Left) x = wa.Left;
                if (y < wa.Top) y = wa.Top;
            }
            Location = new Point(x, y);
            if (x != Config.NoteX || y != Config.NoteY) { Config.NoteX = x; Config.NoteY = y; Config.Save(); }

            ApplyStyle();
            box.Text = preset == null ? "" : preset;
            Show();
            PasteEngine.ForceForeground(Handle);
            box.Focus();
            box.SelectionStart = box.TextLength;
        }
    }

    // ---------------- 速记框外观设置 ----------------
    internal class NoteSettingsForm : Form
    {
        private TrackBar op, rad, fnt;
        private ComboBox theme;
        private TextBox cue, hintText;
        private Label opVal, radVal, fntVal;
        private CheckBox stay;

        public Action Applied;

        private double oldOpacity, oldFont;
        private int oldRadius;
        private string oldTheme, oldCue, oldHint;
        private bool oldStay;

        public NoteSettingsForm()
        {
            Text = "速记框外观";
            try { Icon = MainApp.AppIcon(); } catch { }
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = true;
            TopMost = true;
            Font = new Font("Microsoft YaHei", 9f);
            ClientSize = new Size(460, 378);

            oldOpacity = Config.NoteOpacity; oldFont = Config.NoteFont;
            oldRadius = Config.NoteRadius; oldTheme = Config.NoteTheme;
            oldCue = Config.NoteCue; oldHint = Config.NoteHint;
            oldStay = Config.NoteStay;

            int y = 16;
            MakeLabel("透明度", y);
            op = new TrackBar();
            op.Minimum = 30; op.Maximum = 100; op.TickFrequency = 5;
            op.Value = (int)Math.Round(Config.NoteOpacity * 100);
            op.Location = new Point(96, y - 6); op.Size = new Size(270, 32);
            op.ValueChanged += delegate { opVal.Text = op.Value + "%"; Live(); };
            Controls.Add(op);
            opVal = new Label(); opVal.Text = op.Value + "%";
            opVal.Location = new Point(372, y - 2); opVal.AutoSize = true;
            Controls.Add(opVal);
            y += 44;

            MakeLabel("圆角", y);
            rad = new TrackBar();
            rad.Minimum = 0; rad.Maximum = 24; rad.TickFrequency = 2;
            rad.Value = Config.NoteRadius < 0 ? 0 : Config.NoteRadius;
            rad.Location = new Point(96, y - 6); rad.Size = new Size(270, 32);
            rad.ValueChanged += delegate { radVal.Text = rad.Value + " px"; Live(); };
            Controls.Add(rad);
            radVal = new Label(); radVal.Text = rad.Value + " px";
            radVal.Location = new Point(372, y - 2); radVal.AutoSize = true;
            Controls.Add(radVal);
            y += 44;

            MakeLabel("字号", y);
            fnt = new TrackBar();
            fnt.Minimum = 9; fnt.Maximum = 20; fnt.TickFrequency = 1;
            fnt.Value = (int)Math.Round(Config.NoteFont);
            fnt.Location = new Point(96, y - 6); fnt.Size = new Size(270, 32);
            fnt.ValueChanged += delegate { fntVal.Text = fnt.Value + " pt"; Live(); };
            Controls.Add(fnt);
            fntVal = new Label(); fntVal.Text = fnt.Value + " pt";
            fntVal.Location = new Point(372, y - 2); fntVal.AutoSize = true;
            Controls.Add(fntVal);
            y += 44;

            MakeLabel("主题", y);
            theme = new ComboBox();
            theme.DropDownStyle = ComboBoxStyle.DropDownList;
            theme.Items.Add("深色");
            theme.Items.Add("浅色");
            theme.SelectedIndex = (Config.NoteTheme == "light") ? 1 : 0;
            theme.Location = new Point(96, y - 3); theme.Size = new Size(120, 26);
            theme.SelectedIndexChanged += delegate { Live(); };
            Controls.Add(theme);
            y += 40;

            MakeLabel("输入提示", y);
            cue = new TextBox();
            cue.Text = Config.NoteCue;
            cue.Location = new Point(96, y); cue.Size = new Size(340, 26);
            cue.TextChanged += delegate { Live(); };
            Controls.Add(cue);
            y += 36;

            MakeLabel("底部提示", y);
            hintText = new TextBox();
            hintText.Text = Config.NoteHint;
            hintText.Location = new Point(96, y); hintText.Size = new Size(340, 26);
            hintText.TextChanged += delegate { Live(); };
            Controls.Add(hintText);
            y += 44;

            Button def = new Button(); def.Text = "恢复默认";
            def.Location = new Point(16, y + 34); def.Size = new Size(90, 32);
            def.Click += delegate { RestoreDefaults(); };
            Controls.Add(def);

            Button cancel = new Button(); cancel.Text = "取消";
            cancel.Location = new Point(250, y + 34); cancel.Size = new Size(90, 32);
            cancel.Click += delegate { RestoreOld(); Close(); };
            Controls.Add(cancel);

            Button ok = new Button(); ok.Text = "保存并应用";
            ok.Location = new Point(350, y + 34); ok.Size = new Size(94, 32);
            ok.Click += delegate { Config.Save(); if (Applied != null) Applied(); Close(); };
            Controls.Add(ok);

            stay = new CheckBox();
            stay.Text = "回车后保持常驻（继续打，Esc 或热键才收起）";
            stay.Location = new Point(18, y + 4);
            stay.Size = new Size(340, 24);
            stay.Checked = Config.NoteStay;
            stay.CheckedChanged += delegate { Live(); };
            Controls.Add(stay);
        }

        private void MakeLabel(string text, int y)
        {
            Label l = new Label();
            l.Text = text; l.AutoSize = true; l.Location = new Point(18, y + 3);
            Controls.Add(l);
        }

        private void RestoreDefaults()
        {
            op.Value = 72; rad.Value = 14; fnt.Value = 12;
            theme.SelectedIndex = 0;
            cue.Text = "输入文字，回车自动写进 Word";
            hintText.Text = "↵ 写入 Word   ·   按住这里拖动   ·   边角可缩放";
            stay.Checked = true;
            Live();
        }

        private void ApplyToConfig()
        {
            Config.NoteOpacity = op.Value / 100.0;
            Config.NoteRadius = rad.Value;
            Config.NoteFont = fnt.Value;
            Config.NoteTheme = (theme.SelectedIndex == 1) ? "light" : "dark";
            Config.NoteCue = cue.Text;
            Config.NoteHint = hintText.Text;
            Config.NoteStay = stay.Checked;
        }

        private void Live()
        {
            ApplyToConfig();
            if (Applied != null) Applied();
        }

        private void RestoreOld()
        {
            Config.NoteOpacity = oldOpacity; Config.NoteRadius = oldRadius;
            Config.NoteFont = oldFont; Config.NoteTheme = oldTheme;
            Config.NoteCue = oldCue; Config.NoteHint = oldHint;
            Config.NoteStay = oldStay;
            if (Applied != null) Applied();
        }
    }

    // ---------------- 选择插入目标的 Word 窗口 ----------------
    internal class WordPickForm : Form
    {
        private ListView list;
        private long pickedHwnd;
        private string pickedTitle;

        public WordPickForm()
        {
            Text = "选择插入目标的 Word 窗口";
            try { Icon = MainApp.AppIcon(); } catch { }
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false; MinimizeBox = false;
            ShowInTaskbar = true;
            TopMost = true;
            Font = new Font("Microsoft YaHei", 9f);
            ClientSize = new Size(560, 320);

            Label tip = new Label();
            tip.Text = "截图和速记会写进这里选中的那扇 Word 窗口。开着好几个 Word 时，用它指定目标。";
            tip.AutoSize = false;
            tip.Location = new Point(14, 12); tip.Size = new Size(532, 36);
            Controls.Add(tip);

            list = new ListView();
            list.View = View.Details;
            list.FullRowSelect = true;
            list.MultiSelect = false;
            list.HideSelection = false;
            list.Columns.Add("文档窗口", 420);
            list.Columns.Add("进程 ID", 90);
            list.Location = new Point(14, 52);
            list.Size = new Size(532, 200);
            list.DoubleClick += delegate { Accept(); };
            Controls.Add(list);

            Button refresh = new Button(); refresh.Text = "刷新";
            refresh.Location = new Point(14, 268); refresh.Size = new Size(90, 32);
            refresh.Click += delegate { Fill(); };
            Controls.Add(refresh);

            Button cancel = new Button(); cancel.Text = "取消";
            cancel.Location = new Point(352, 268); cancel.Size = new Size(90, 32);
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(cancel);

            Button ok = new Button(); ok.Text = "用这个";
            ok.Location = new Point(452, 268); ok.Size = new Size(94, 32);
            ok.Click += delegate { Accept(); };
            Controls.Add(ok);

            AcceptButton = ok;
            CancelButton = cancel;

            Fill();
        }

        private void Fill()
        {
            list.Items.Clear();
            ListViewItem autoItem = new ListViewItem("（自动：挑最合适的那个）");
            autoItem.SubItems.Add("");
            autoItem.Tag = (long)0;
            list.Items.Add(autoItem);

            Native.EnumWindows(delegate(IntPtr h, IntPtr p)
            {
                if (!Native.IsWindowVisible(h)) return true;
                if (Native.ClassName(h) != Config.WordClass) return true;
                string t = Native.Title(h);
                if (t.Length == 0) return true;
                uint pid;
                Native.GetWindowThreadProcessId(h, out pid);
                ListViewItem it = new ListViewItem(t);
                it.SubItems.Add(pid.ToString());
                it.Tag = (long)h;
                list.Items.Add(it);
                return true;
            }, IntPtr.Zero);

            // 选中当前在用的那个
            int sel = 0;
            for (int i = 0; i < list.Items.Count; i++)
            {
                long tag = (long)list.Items[i].Tag;
                if (tag == Config.WordHwnd && tag != 0) { sel = i; break; }
            }
            if (list.Items.Count > 0) list.Items[sel].Selected = true;
        }

        private void Accept()
        {
            if (list.SelectedItems.Count == 0) { DialogResult = DialogResult.Cancel; Close(); return; }
            ListViewItem it = list.SelectedItems[0];
            pickedHwnd = (long)it.Tag;
            pickedTitle = (pickedHwnd == 0) ? "" : it.Text;
            Config.WordHwnd = pickedHwnd;
            Config.WordTitle = pickedTitle;
            Config.Save();
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ---------------- 框选区域的遮罩窗口 ----------------
    internal class RegionForm : Form
    {
        public Rectangle Selected;
        private Point anchor;
        private bool dragging;
        private Rectangle box;

        public RegionForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Bounds = SystemInformation.VirtualScreen;
            BackColor = Color.Black;
            Opacity = 0.30;
            TopMost = true;
            ShowInTaskbar = false;
            Cursor = Cursors.Cross;
            KeyPreview = true;
            DoubleBuffered = true;
        }

        private static Rectangle MakeRect(Point a, Point b)
        {
            return new Rectangle(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Abs(a.X - b.X), Math.Abs(a.Y - b.Y));
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Activate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            anchor = e.Location;
            dragging = true;
            box = new Rectangle(anchor, Size.Empty);
            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (!dragging) return;
            box = MakeRect(anchor, e.Location);
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            dragging = false;
            box = MakeRect(anchor, e.Location);
            if (box.Width >= 8 && box.Height >= 8)
            {
                Selected = new Rectangle(box.X + Bounds.X, box.Y + Bounds.Y, box.Width, box.Height);
                DialogResult = DialogResult.OK;
            }
            else
            {
                DialogResult = DialogResult.Cancel;
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) DialogResult = DialogResult.Cancel;
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (box.Width > 0 && box.Height > 0)
            {
                using (Pen pen = new Pen(Color.FromArgb(255, 0, 174, 255), 2))
                    e.Graphics.DrawRectangle(pen, box);
            }
            string hint = "拖动框选要截的范围    Esc 取消    " + box.Width + " x " + box.Height;
            using (Font f = new Font("Microsoft YaHei", 12f))
            using (Brush b = new SolidBrush(Color.White))
            {
                e.Graphics.DrawString(hint, f, b, new PointF(24, 24));
            }
        }
    }

    // ---------------- Word 接口 ----------------
    internal class WordBridge
    {
        private dynamic win;         // 目标窗口自己的 Word.Window 对象
        private IntPtr winFor = IntPtr.Zero;

        public bool Running
        {
            get { return Native.ResolveWordWindow() != IntPtr.Zero; }
        }

        // 附着到“用户选中的那扇 Word 窗口”所属的实例：
        // 通过 MSAA 原生对象模型直接拿那扇窗口的 Window 对象，不再依赖运行对象表（ROT）
        private bool Attach()
        {
            IntPtr target = Native.ResolveWordWindow();
            if (target == IntPtr.Zero) return false;
            if (win != null && winFor == target) return true;
            try
            {
                IntPtr pane = Native.FindEditPane(target);
                if (pane == IntPtr.Zero)
                {
                    Config.Log("com: 窗口 " + Native.Title(target) + " 找不到编辑窗格");
                    win = null; winFor = IntPtr.Zero;
                    return false;
                }
                win = Native.NativeOM(pane);
                winFor = target;
                try { Config.Log("com: 已连上 [" + (string)win.Document.Name + "]"); }
                catch { Config.Log("com: 已连上 " + Native.Title(target)); }
                return true;
            }
            catch (Exception ex)
            {
                Config.Log("com: 连接失败 " + ex.Message);
                win = null; winFor = IntPtr.Zero;
                return false;
            }
        }

        private dynamic Selection()
        {
            dynamic s = win.Selection;
            return s;
        }

        // 把图片插到当前文档的光标处，图片后自动换行
        public bool Insert(string pngPath)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
            if (!Attach()) return false;
            try
            {
                dynamic doc = win.Document;
                dynamic sel = Selection();

                // 如果当前段落里已经有字，先另起一行，让图片独占一行
                try
                {
                    string txt = (string)sel.Paragraphs.Item(1).Range.Text;
                    if (txt != null && txt.Trim().Length > 0) sel.TypeParagraph();
                }
                catch { }

                dynamic shape = sel.InlineShapes.AddPicture(pngPath, false, true, Type.Missing);

                // 太宽就缩到版心宽度以内
                try
                {
                    double sw = (double)shape.Width;
                    double sh = (double)shape.Height;
                    if (sw > 0)
                    {
                        double avail = ((double)doc.PageSetup.PageWidth - (double)doc.PageSetup.LeftMargin
                                        - (double)doc.PageSetup.RightMargin - 6.0) * (Config.ImageWidthPercent / 100.0);
                        if (sw > avail)
                        {
                            double ratio = sh / sw;
                            shape.Width = (float)avail;
                            shape.Height = (float)(avail * ratio);
                        }
                    }
                }
                catch { }

                // 光标移到图片后面并换行（不抢焦点，COM 调用本身不会激活 Word 窗口）
                try
                {
                    sel.Collapse(0);
                    for (int i = 0; i < Config.LinesAfter; i++) sel.TypeParagraph();
                }
                catch { }
                return true;
            }
            catch (Exception ex)
            {
                Config.Log("insert COM error (" + (attempt + 1) + "/2): " + ex.Message);
                win = null; winFor = IntPtr.Zero;   // 连接坏了就丢掉，下一轮重新附着
            }
            }
            return false;
        }

        // 把一段文字打到光标处，然后换行
        public bool InsertText(string text)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
                if (!Attach()) return false;
                try
                {
                    dynamic sel = Selection();
                    sel.TypeText(text);
                    sel.TypeParagraph();
                    return true;
                }
                catch (Exception ex)
                {
                    Config.Log("note COM error (" + (attempt + 1) + "/2): " + ex.Message);
                    win = null; winFor = IntPtr.Zero;
                }
            }
            return false;
        }
    }

    // ---------------- 备用通道：剪贴板 + Ctrl+V ----------------
    // 用户自己打开的 Word 不会登记到运行对象表（ROT），COM 就拿不到它。
    // 这时退回到“把图放进剪贴板 → 激活 Word → 发 Ctrl+V、回车”。
    internal static class PasteEngine
    {
        public static DateTime SilentUntil = DateTime.MinValue;

        public static bool Paste(string pngPath, out string err)
        {
            err = null;
            SilentUntil = DateTime.Now.AddSeconds(3);
            try
            {
                Bitmap src = new Bitmap(pngPath);
                Bitmap dst = Resize(src, Config.PasteWidthPx);
                src.Dispose();

                DataObject data = new DataObject();
                data.SetData(DataFormats.Bitmap, dst);
                Clipboard.SetDataObject(data, true);
                Config.Log("paste: clipboard " + dst.Width + "x" + dst.Height);
                dst.Dispose();

                IntPtr word = Native.BestWordWindow(Config.WordClass);
                if (word == IntPtr.Zero) { err = "Word 没打开（或没有可见的 Word 窗口）"; return false; }
                Config.Log("paste: target " + Native.Title(word));
                if (Native.IsIconic(word)) Native.ShowWindow(word, 9);   // SW_RESTORE
                if (!ForceForeground(word)) { err = "Word 没能切到前台"; return false; }
                Thread.Sleep(80);

                SendCtrlV();
                Thread.Sleep(200);
                SendEnter();
                Config.Log("paste: keys sent, fg=" + Native.Title(Native.GetForegroundWindow()));
                return true;
            }
            catch (Exception ex)
            {
                err = ex.GetType().Name + " " + ex.Message;
                return false;
            }
        }

        // 把一段文字粘进 Word 光标处（可用于任意中文，不依赖输入法）
        public static bool PasteText(string text, out string err)
        {
            err = null;
            SilentUntil = DateTime.Now.AddSeconds(3);
            try
            {
                DataObject data = new DataObject();
                data.SetData(DataFormats.UnicodeText, text);
                Clipboard.SetDataObject(data, true);
                Config.Log("paste text: clipboard " + text.Length + " chars");

                IntPtr word = Native.BestWordWindow(Config.WordClass);
                if (word == IntPtr.Zero) { err = "Word 没打开（或没有可见的 Word 窗口）"; return false; }
                Config.Log("paste text: target " + Native.Title(word));
                if (Native.IsIconic(word)) Native.ShowWindow(word, 9);
                if (!ForceForeground(word)) { err = "Word 没能切到前台"; return false; }
                Thread.Sleep(120);

                SendCtrlV();
                Thread.Sleep(220);
                SendEnter();
                Config.Log("paste text: keys sent, fg=" + Native.Title(Native.GetForegroundWindow()));
                return true;
            }
            catch (Exception ex)
            {
                err = ex.GetType().Name + " " + ex.Message;
                return false;
            }
        }

        public static bool ForceForeground(IntPtr h)
        {
            if (Native.GetForegroundWindow() == h) return true;
            Native.SetForegroundWindow(h);
            Thread.Sleep(150);
            if (Native.GetForegroundWindow() == h) return true;

            // 1) 把输入队列挂到前台线程上再试
            IntPtr fg = Native.GetForegroundWindow();
            uint pid;
            uint ft = Native.GetWindowThreadProcessId(fg, out pid);
            uint mt = Native.GetCurrentThreadId();
            bool attached = false;
            try
            {
                attached = Native.AttachThreadInput(mt, ft, true);
                Native.SetForegroundWindow(h);
                Thread.Sleep(200);
            }
            finally
            {
                if (attached) Native.AttachThreadInput(mt, ft, false);
            }
            if (Native.GetForegroundWindow() == h) return true;

            // 2) 先敲一个几乎没人用的键（F15），系统会认作“用户有输入”，从而放行切换
            Native.keybd_event(0x7E, 0, 0, IntPtr.Zero);
            Native.keybd_event(0x7E, 0, 2, IntPtr.Zero);
            Thread.Sleep(60);
            Native.SetForegroundWindow(h);
            Thread.Sleep(200);
            if (Native.GetForegroundWindow() == h) return true;

            // 3) 最后一招
            Native.SwitchToThisWindow(h, true);
            Thread.Sleep(200);
            return Native.GetForegroundWindow() == h;
        }

        private static Bitmap Resize(Bitmap src, int targetW)
        {
            if (src.Width <= targetW) return new Bitmap(src);
            int w = src.Width;
            int h = src.Height;
            int nh = (int)Math.Round(h * (targetW / (double)w));
            if (nh < 8) nh = 8;
            Bitmap dst = new Bitmap(targetW, nh, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(dst))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(src, new Rectangle(0, 0, targetW, nh));
            }
            return dst;
        }

        private static void SendCtrlV()
        {
            Native.keybd_event(0x11, 0, 0, IntPtr.Zero);          // Ctrl down
            Native.keybd_event(0x56, 0, 0, IntPtr.Zero);          // V down
            Native.keybd_event(0x56, 0, 2, IntPtr.Zero);          // V up
            Native.keybd_event(0x11, 0, 2, IntPtr.Zero);          // Ctrl up
        }

        private static void SendEnter()
        {
            Native.keybd_event(0x0D, 0, 0, IntPtr.Zero);          // Enter down
            Native.keybd_event(0x0D, 0, 2, IntPtr.Zero);          // Enter up
        }
    }

    // ---------------- 主程序 ----------------
    internal class MainApp : ApplicationContext
    {
        private const int ID_SNIP = 1, ID_RESELECT = 2, ID_PLAY = 3, ID_FOCUS = 4, ID_AUTO = 5, ID_NOTE = 6;

        private NotifyIcon tray;
        private MsgWindow msg;
        private WordBridge word = new WordBridge();
        private ToolStripMenuItem autoItem;
        private DateTime lastInsert = DateTime.MinValue;
        private string lastHash = "";
        private DateTime lastHashAt = DateTime.MinValue;
        private IntPtr lastOther = IntPtr.Zero;
        private bool cleaned;
        private NoteBox noteBox;
        private IntPtr notePrev = IntPtr.Zero;

        public MainApp()
        {
            Config.Log("start");
            Config.CleanTemp();
            msg = new MsgWindow();
            msg.Hotkey += OnHotkey;
            msg.ClipUpdate += OnClipUpdate;
            Config.Log("msg window handle=" + msg.Handle);
            BuildTray();
            RegisterHotkeys();
            Application.ApplicationExit += delegate { Cleanup(); };
            UpdateTrayText();

            foreach (string a in Environment.GetCommandLineArgs())
            {
                if (a == "--note-settings") { ShowNoteSettings(); break; }
                if (a == "--pick-word") { PickWordWindow(); break; }
            }
        }

        // ---- 托盘 ----
        private static Icon cachedAppIcon;

        // 托盘用 exe 里嵌的图标（蓝色小笔记），没有就退回默认图标
        internal static Icon AppIcon()
        {
            if (cachedAppIcon != null) return cachedAppIcon;
            try { cachedAppIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { cachedAppIcon = null; }
            if (cachedAppIcon == null) cachedAppIcon = SystemIcons.Application;
            return cachedAppIcon;
        }

        private void BuildTray()
        {
            ContextMenuStrip m = new ContextMenuStrip();
            autoItem = new ToolStripMenuItem("自动插入：截图后自动写进 Word");
            autoItem.CheckOnClick = true;
            autoItem.Checked = Config.AutoMode;
            autoItem.Click += delegate { Config.AutoMode = autoItem.Checked; Config.Save(); UpdateTrayText(); };
            m.Items.Add(autoItem);
            m.Items.Add(new ToolStripMenuItem("现在截一张", null, delegate { Snip(false); }));
            m.Items.Add(new ToolStripMenuItem("速记框外观…", null, delegate { ShowNoteSettings(); }));
            m.Items.Add(new ToolStripMenuItem("重新框选截图区域", null, delegate { Snip(true); }));
            m.Items.Add(new ToolStripMenuItem("把剪贴板里的图片插入 Word", null, delegate { InsertClipboardImage(); }));
            m.Items.Add(new ToolStripMenuItem("选择插入目标的 Word 窗口…", null, delegate { PickWordWindow(); }));
            m.Items.Add(new ToolStripSeparator());
            m.Items.Add(new ToolStripMenuItem("打开配置文件（改热键）", null, delegate { Config.OpenIni(); }));
            m.Items.Add(new ToolStripSeparator());
            m.Items.Add(new ToolStripMenuItem("退出", null, delegate { ExitAll(); }));

            tray = new NotifyIcon();
            tray.Icon = AppIcon();
            tray.Text = "WordNote";
            tray.Visible = true;
            tray.ContextMenuStrip = m;
            tray.DoubleClick += delegate { Snip(false); };
        }

        private void UpdateTrayText()
        {
            if (tray == null) return;
            try { tray.Text = "WordNote  自动插入:" + (Config.AutoMode ? "开" : "关"); } catch { }
        }

        private void Notify(string text)
        {
            if (tray == null) return;
            try { tray.ShowBalloonTip(1500, "WordNote", text, ToolTipIcon.Info); } catch { }
        }

        // ---- 热键 ----
        private void RegisterHotkeys()
        {
            StringBuilder bad = new StringBuilder();
            Reg(Config.SnipKey, ID_SNIP, bad);
            Reg(Config.ReselectKey, ID_RESELECT, bad);
            Reg(Config.PlayKey, ID_PLAY, bad);
            Reg(Config.FocusKey, ID_FOCUS, bad);
            Reg(Config.AutoKey, ID_AUTO, bad);
            Reg(Config.NoteKey, ID_NOTE, bad);
            if (bad.Length > 0)
                Notify("这些热键没注册上（可能被别的软件占了）：" + bad + "  —— 改 WordNote.ini 后重启");
            else
                Notify("已启动。" + Config.SnipKey + " 截图，" + Config.NoteKey + " 速记，都会自动进 Word。");
        }

        private void Reg(string combo, int id, StringBuilder bad)
        {
            uint mods, vk;
            bool ok = ParseCombo(combo, out mods, out vk) && Native.RegisterHotKey(msg.Handle, id, mods, vk);
            Config.Log("hotkey " + combo + " id=" + id + " -> " + ok);
            if (!ok)
            {
                if (bad.Length > 0) bad.Append("、");
                bad.Append(combo);
            }
        }

        private static bool ParseCombo(string combo, out uint mods, out uint vk)
        {
            mods = 0;
            vk = 0;
            if (string.IsNullOrEmpty(combo)) return false;
            string[] parts = combo.Split('+');
            int last = parts.Length - 1;
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i].Trim();
                if (p.Length == 0) return false;
                string low = p.ToLowerInvariant();
                if (i < last)
                {
                    if (low == "ctrl" || low == "control") mods |= 0x0002;
                    else if (low == "alt") mods |= 0x0001;
                    else if (low == "shift") mods |= 0x0004;
                    else if (low == "win") mods |= 0x0008;
                    else return false;
                }
                else
                {
                    if (p.Length == 1)
                    {
                        char c = char.ToUpperInvariant(p[0]);
                        if (c >= 'A' && c <= 'Z') vk = (uint)(0x41 + (c - 'A'));
                        else if (c >= '0' && c <= '9') vk = (uint)(0x30 + (c - '0'));
                        else return false;
                    }
                    else
                    {
                        Keys k;
                        if (!Enum.TryParse<Keys>(p, true, out k)) return false;
                        vk = (uint)k;
                    }
                }
            }
            return vk != 0;
        }

        private void OnHotkey(int id)
        {
            Config.Log("hotkey id=" + id);
            switch (id)
            {
                case ID_SNIP: Snip(false); break;
                case ID_RESELECT: Snip(true); break;
                case ID_PLAY: MediaPlayPause(); break;
                case ID_FOCUS: ToggleFocus(); break;
                case ID_NOTE: ShowNoteBox(); break;
                case ID_AUTO:
                    {
                        Config.AutoMode = !Config.AutoMode;
                        if (autoItem != null) autoItem.Checked = Config.AutoMode;
                        Config.Save();
                        UpdateTrayText();
                        Notify("自动插入：" + (Config.AutoMode ? "开" : "关"));
                        break;
                    }
            }
        }

        // ---- 速记框 ----
        private void ShowNoteBox()
        {
            EnsureNoteBox();
            if (!noteBox.Visible) notePrev = Native.GetForegroundWindow();
            noteBox.ShowBox(noteBox.CurrentText);
        }

        private void EnsureNoteBox()
        {
            if (noteBox != null) return;
            noteBox = new NoteBox();
            noteBox.Submitted += SubmitNote;
            noteBox.Cancelled += delegate { Config.Log("note cancelled"); };
            noteBox.Moved += delegate
            {
                Config.Save();
                Config.Log("note box " + Config.NoteX + "," + Config.NoteY + " " + Config.NoteW + "x" + Config.NoteH);
            };
        }

        // ---- 速记框外观 ----
        private void ShowNoteSettings()
        {
            EnsureNoteBox();
            NoteSettingsForm f = new NoteSettingsForm();
            f.Applied += delegate { noteBox.ApplyStyle(); };
            f.ShowDialog();
            Config.Save();
            if (!noteBox.Visible) Notify("外观已保存。按 " + Config.NoteKey + " 看效果。");
        }

        private void PickWordWindow()
        {
            using (WordPickForm f = new WordPickForm())
            {
                if (f.ShowDialog() != DialogResult.OK) return;
            }
            string t = Config.WordHwnd == 0 ? "自动（挑最合适的）" : Config.WordTitle;
            Config.Log("word target -> " + t);
            Notify("插入目标已设为：" + t);
        }

        private void SubmitNote(string text)
        {
            if (text == null) return;
            text = text.Trim();
            if (text.Length == 0) return;

            bool comOk = false;
            try
            {
                comOk = word.InsertText(text);
                Config.Log("note COM -> " + comOk);
            }
            catch (Exception ex)
            {
                Config.Log("note COM error: " + ex.Message);
            }

            if (!comOk)
            {
                string err = null;
                bool ok = PasteEngine.PasteText(text, out err);
                Config.Log("note PASTE -> " + ok + (ok ? "" : " (" + err + ")"));
                if (!ok)
                {
                    Notify("速记没写进 Word：" + (err == null ? "未知原因" : err));
                    if (noteBox != null) noteBox.ShowBox(text);   // 字留着，别白打
                    return;
                }
            }

            // 常驻模式：焦点交回速记框，接着打下一条
            if (Config.NoteStay)
            {
                if (noteBox != null && noteBox.Visible) noteBox.FocusInput();
                return;
            }

            // 写完把焦点还回你刚才看的窗口（视频），不用手动切回来
            if (notePrev != IntPtr.Zero && notePrev != Native.GetForegroundWindow())
                PasteEngine.ForceForeground(notePrev);
        }

        // ---- 截图 ----
        private void Snip(bool reselect)
        {
            if (reselect || Config.Region.Width < 8 || Config.Region.Height < 8)
            {
                Rectangle? r = PickRegion();
                if (r == null) return;
                Config.Region = r.Value;
                Config.Save();
                Thread.Sleep(140);
            }
            CaptureAndInsert(Config.Region);
        }

        private Rectangle? PickRegion()
        {
            Rectangle? r = null;
            using (RegionForm f = new RegionForm())
            {
                if (f.ShowDialog() == DialogResult.OK) r = f.Selected;
            }
            Thread.Sleep(140);
            return r;
        }

        private void CaptureAndInsert(Rectangle r)
        {
            // 自己的速记框如果在屏幕上，先收起来再抓：半透明分层窗口参与 BitBlt 会抓出竖条纹残影
            bool hadBox = noteBox != null && noteBox.Visible;
            string keep = hadBox ? noteBox.CurrentText : null;
            if (hadBox) noteBox.Hide();
            try
            {
                if (hadBox) Thread.Sleep(90);   // 让 DWM 重画一帧
                using (Bitmap bmp = new Bitmap(r.Width, r.Height, PixelFormat.Format32bppArgb))
                {
                    bool ok = Native.CaptureRect(bmp, r);
                    if (!ok)
                    {
                        // BitBlt 失败时退回 GDI+ 的抓屏
                        using (Graphics g = Graphics.FromImage(bmp))
                            g.CopyFromScreen(r.X, r.Y, 0, 0, new Size(r.Width, r.Height), CopyPixelOperation.SourceCopy);
                    }
                    string png = Config.TempFile();
                    bmp.Save(png, ImageFormat.Png);
                    lastInsert = DateTime.Now;
                    InsertFile(png);
                }
            }
            catch (Exception ex)
            {
                Notify("截图失败：" + ex.Message);
            }
            finally
            {
                // 截完把速记框放回来（常驻就该一直待在屏幕上），原来打的字也留着
                if (hadBox && noteBox != null) noteBox.ShowBox(keep);
            }
        }

        private void InsertFile(string png)
        {
            bool comOk = false;
            try
            {
                comOk = word.Insert(png);
                Config.Log("insert COM -> " + comOk);
            }
            catch (Exception ex)
            {
                Config.Log("insert COM error: " + ex.Message);
            }
            if (comOk) return;

            string err = null;
            bool ok = PasteEngine.Paste(png, out err);
            Config.Log("insert PASTE -> " + ok + (ok ? "" : " (" + err + ")"));
            if (!ok) Notify("插入失败：" + (err == null ? "未知原因" : err));
        }

        // ---- 剪贴板自动插入 ----
        private void OnClipUpdate()
        {
            if (DateTime.Now < PasteEngine.SilentUntil) { Config.Log("clip: silent (self paste)"); return; }
            if (!Config.AutoMode) { Config.Log("clip: auto off"); return; }
            if ((DateTime.Now - lastInsert).TotalMilliseconds < 700) { Config.Log("clip: debounce"); return; }
            try
            {
                IntPtr owner = Native.GetClipboardOwner();
                string ownerClass = owner == IntPtr.Zero ? "<null>" : Native.ClassName(owner);
                bool hasImg = Clipboard.ContainsImage();
                Config.Log("clip: owner=" + ownerClass + " image=" + hasImg);
                if (owner != IntPtr.Zero && ownerClass == Config.WordClass) return; // 在 Word 里复制图片，别回插
                if (!hasImg) return;
                Image img = Clipboard.GetImage();
                if (img == null) return;
                string png = Config.TempFile();
                using (Bitmap bmp = new Bitmap(img)) bmp.Save(png, ImageFormat.Png);
                img.Dispose();

                // PrtScn / 截图工具会把同一张图往剪贴板写两次（间隔可能 >700ms），
                // 按内容指纹去重，避免一次插入两张一模一样的画
                string hash = HashFile(png);
                if (hash == lastHash && (DateTime.Now - lastHashAt).TotalSeconds < 8)
                {
                    Config.Log("clip: duplicate image, skip");
                    try { File.Delete(png); } catch { }
                    return;
                }
                lastHash = hash;
                lastHashAt = DateTime.Now;

                lastInsert = DateTime.Now;
                InsertFile(png);
            }
            catch (Exception ex)
            {
                Config.Log("clip error: " + ex.Message);
            }
        }

        private static string HashFile(string path)
        {
            try
            {
                using (FileStream fs = File.OpenRead(path))
                {
                    byte[] h = System.Security.Cryptography.SHA1.Create().ComputeHash(fs);
                    return BitConverter.ToString(h);
                }
            }
            catch { return "<no-hash>"; }
        }

        private void InsertClipboardImage()
        {
            try
            {
                if (!Clipboard.ContainsImage()) { Notify("剪贴板里没有图片"); return; }
                Image img = Clipboard.GetImage();
                if (img == null) return;
                string png = Config.TempFile();
                using (Bitmap bmp = new Bitmap(img)) bmp.Save(png, ImageFormat.Png);
                img.Dispose();
                lastInsert = DateTime.Now;
                InsertFile(png);
            }
            catch (Exception ex)
            {
                Notify("插入失败：" + ex.Message);
            }
        }

        // ---- 焦点 / 播放 ----
        private void ToggleFocus()
        {
            IntPtr wh = Native.BestWordWindow(Config.WordClass);
            if (wh == IntPtr.Zero) { Notify("Word 没打开（或没有可见的 Word 窗口）"); return; }
            IntPtr fg = Native.GetForegroundWindow();
            if (fg == wh)
            {
                if (lastOther != IntPtr.Zero && Native.IsWindowVisible(lastOther)) Native.SetForegroundWindow(lastOther);
                else Notify("还没记住视频窗口，先点一下视频");
            }
            else
            {
                if (fg != IntPtr.Zero) lastOther = fg;
                Native.SetForegroundWindow(wh);
            }
        }

        private static void MediaPlayPause()
        {
            Native.keybd_event(0xB3, 0, 1, IntPtr.Zero);
            Native.keybd_event(0xB3, 0, 3, IntPtr.Zero);
        }

        // ---- 退出 ----
        private void ExitAll()
        {
            Cleanup();
            ExitThread();
        }

        private void Cleanup()
        {
            if (cleaned) return;
            cleaned = true;
            try { Native.RemoveClipboardFormatListener(msg.Handle); } catch { }
            for (int id = ID_SNIP; id <= ID_NOTE; id++)
            {
                try { Native.UnregisterHotKey(msg.Handle, id); } catch { }
            }
            try { if (noteBox != null) { noteBox.Close(); noteBox.Dispose(); } } catch { }
            try { msg.Dispose(); } catch { }
            Config.Log("exit");
            try { if (tray != null) { tray.Visible = false; tray.Dispose(); } } catch { }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Native.SetProcessDPIAware();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Config.Load();
            Application.Run(new MainApp());
        }
    }
}
