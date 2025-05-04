using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Drawing;


namespace MyChatGPTLauncher
{
    static class Program
    {
        // Win32 API declarations
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SetActiveWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);
        [DllImport("user32.dll")] private static extern bool TranslateMessage([In] ref MSG lpMsg);
        [DllImport("user32.dll")] private static extern IntPtr DispatchMessage([In] ref MSG lpMsg);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        // Window state API
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);
        private const int SW_RESTORE = 9;
        // Mouse event API
        [DllImport("user32.dll")] private static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")] private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT lpPoint);
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_INIT_ID = 9000;  // 初期プロンプト送信用
        private const int HOTKEY_EXEC_ID = 9001;  // コマンド実行用
        private const int HOTKEY_SS_ID   = 9002;  // スクリーンショット用
        private const int HOTKEY_DUMP_ID = 9003; // ダンプ用
        private const uint MOD_CTRL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint VK_G = 0x47; // Gキー
        private const uint VK_H = 0x48; // Hキー
        private const uint VK_S = 0x53; // Sキー
        private const uint VK_D = 0x44; // Dキー

        // SetWindowPos flags
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        private const uint MOUSEEVENTF_LEFTUP = 0x04;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hWnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public System.Drawing.Point pt;
        }
        private struct POINT {
            public int x;
            public int y;
        }
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [STAThread]
        static void Main()
        {
            // ホットキー登録
            if (!RegisterHotKey(IntPtr.Zero, HOTKEY_INIT_ID, MOD_CTRL | MOD_SHIFT, VK_G) ||
                !RegisterHotKey(IntPtr.Zero, HOTKEY_EXEC_ID, MOD_CTRL | MOD_SHIFT, VK_H) ||
                !RegisterHotKey(IntPtr.Zero, HOTKEY_SS_ID,   MOD_CTRL | MOD_SHIFT, VK_S) ||
                !RegisterHotKey(IntPtr.Zero, HOTKEY_DUMP_ID, MOD_CTRL | MOD_SHIFT, VK_D))
            {
                Console.WriteLine("ホットキーの登録に失敗しました。");
                return;
            }
            Console.WriteLine("Ctrl+Shift+G: 初期プロンプト送信  |  Ctrl+Shift+H: コマンド実行  |  Ctrl+Shift+S: スクリーンショット | Ctrl+Shift+D: Dump Tree");

            // メッセージループ
            while (GetMessage(out MSG msg, IntPtr.Zero, 0, 0) != 0)
            {
                if (msg.message == WM_HOTKEY)
                {
                    int id = (int)msg.wParam;
                    switch (id) {
                        case HOTKEY_INIT_ID:
                            ActivateChatGPT();
                            SendInitialPrompt();
                            break;
                        
                        case HOTKEY_EXEC_ID:
                            //ActivateChatGPT();
                            var response = FetchChatGPTResponse();
                            Console.WriteLine($"[Debug] FetchChatGPTResponse returned: '{response}'");
                            ExecuteCommand(response);
                            break;
                        case HOTKEY_SS_ID:
                            CaptureAndPasteScreenshot();
                            break;
                        case HOTKEY_DUMP_ID:
                            DumpAutomationTree();
                            break;
                    }
                }
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }

            UnregisterHotKey(IntPtr.Zero, HOTKEY_INIT_ID);
            UnregisterHotKey(IntPtr.Zero, HOTKEY_EXEC_ID);
            UnregisterHotKey(IntPtr.Zero, HOTKEY_SS_ID);
            UnregisterHotKey(IntPtr.Zero, HOTKEY_DUMP_ID);
        }

        private static void ActivateChatGPT()
        {
            IntPtr hWnd = GetChatGPTWindowHandle();
            if (hWnd == IntPtr.Zero)
            {
                Process.Start(@"C:\Program Files\ChatGPT\ChatGPT.exe");
                Thread.Sleep(1000);
                hWnd = GetChatGPTWindowHandle();
                if (hWnd == IntPtr.Zero) return;
            }
            // ウィンドウが最小化されていたら復元
            if (IsIconic(hWnd))
                ShowWindow(hWnd, SW_RESTORE);

            uint targetThread = GetWindowThreadProcessId(hWnd, out _);
            uint currentThread = GetCurrentThreadId();
            AttachThreadInput(currentThread, targetThread, true);

            SetForegroundWindow(hWnd);
            SetActiveWindow(hWnd);
            SetFocus(hWnd);
            SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

            AttachThreadInput(currentThread, targetThread, false);
            Thread.Sleep(200);
        }

        private static IntPtr GetChatGPTWindowHandle()
        {
            foreach (var proc in Process.GetProcessesByName("ChatGPT"))
            {
                if (proc.MainWindowHandle != IntPtr.Zero)
                    return proc.MainWindowHandle;
                IntPtr found = IntPtr.Zero;
                EnumWindows((hWnd, lParam) =>
                {
                    GetWindowThreadProcessId(hWnd, out uint pid);
                    if (pid == proc.Id)
                    {
                        int len = GetWindowTextLength(hWnd);
                        if (len > 0)
                        {
                            var sb = new StringBuilder(len + 1);
                            GetWindowText(hWnd, sb, sb.Capacity);
                            if (sb.ToString().Contains("ChatGPT")) { found = hWnd; return false; }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
                if (found != IntPtr.Zero) return found;
            }
            return IntPtr.Zero;
        }
        private static void SendInitialPrompt()
        {
            // 初期プロンプトを行単位で送信し、Shift+Enterで改行
            var lines = new[]
            {
                "あなたはWindows環境で動作する自動化アシスタントです。ユーザーの入力に従い、キーボード／マウス操作やファイル操作を自動化してください。利用可能なコマンドは以下のとおりです。必ずこの形式で、コマンドのみを返してください（余計な説明は禁止）。",
                "open <path>            → アプリやファイルを開く",
                "type <text>            → 指定テキストを入力",
                "enter                  → Enterキーを押下",
                "click <x>,<y>          → 指定座標をクリック (事前に必ず screenshot コマンドを実行してください)",
                "wait <milliseconds>    → 指定ミリ秒待機",
                "screenshot             → 画面全体をキャプチャして貼り付け (これだけでは送信されません)",
                "prompt <text>          → 指定テキストを入力して送信",
                "準備ができたら、次のコマンドをお待ちします。"
            };

            foreach (var line in lines)
            {
                SendKeys.SendWait(line);
                // Shift+Enter で改行（送信せずに改行だけ）
                SendKeys.SendWait("+{ENTER}");
                Thread.Sleep(5);  // キー入力のタイミング調整
            }
            // 最後に通常の Enter で送信
            SendKeys.SendWait("{ENTER}");
        }

        private static string FetchChatGPTResponse()
        {
            // コンテナ要素 (ControlType.Document、LocalizedControlType="ドキュメント")
            IntPtr hWnd = GetChatGPTWindowHandle();
            if (hWnd == IntPtr.Zero) return string.Empty;

            var root = AutomationElement.FromHandle(hWnd);
            if (root == null) return string.Empty;

            // Container を探す
            var containerCondition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document),
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "ドキュメント")
            );
            var containers = root.FindAll(TreeScope.Descendants, containerCondition);
            if (containers == null || containers.Count == 0) return string.Empty;
            var container = containers[containers.Count - 1];

            // Anchor を探し、その後の Text 要素を収集
            var all = container.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            bool foundAnchor = false;
            var lines = new List<string>();

            foreach (AutomationElement el in all)
            {
                var ctr = el.Current;
                // Anchor: ControlType.Text + Name=='ChatGPT:' + LocalizedControlType=='テキスト'
                if (!foundAnchor &&
                    ctr.ControlType == ControlType.Text &&
                    ctr.Name == "ChatGPT:" &&
                    ctr.LocalizedControlType == "テキスト")
                {
                    foundAnchor = true;
                    continue;
                }
                if (foundAnchor)
                {
                    // 終了条件: ControlType.Button, LocalizedControlType=='ボタン'
                    if (ctr.ControlType == ControlType.Button && ctr.LocalizedControlType == "ボタン")
                        break;

                    // 出力行の収集: ControlType.Text, LocalizedControlType=='テキスト'
                    if (ctr.ControlType == ControlType.Text && ctr.LocalizedControlType == "テキスト")
                    {
                        var text = ctr.Name?.Trim();
                        if (!string.IsNullOrEmpty(text))
                            lines.Add(text);
                    }
                }
            }

            return lines.Count > 0 ? string.Join(Environment.NewLine, lines) : string.Empty;
        }
        private static void FocusChatInput()
        {
            IntPtr hWnd = GetChatGPTWindowHandle();
            if (hWnd == IntPtr.Zero) return;

            var root = AutomationElement.FromHandle(hWnd);
            if (root == null) return;

            POINT originalPos;
            if(!GetCursorPos(out originalPos))
                originalPos = new POINT {x = 0, y = 0};

            // 「プロンプトを送信する」ボタン要素を検索
            var sendBtnCondition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                new PropertyCondition(AutomationElement.NameProperty, "音声入力ボタン"),
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "ボタン")
            );
            var sendBtn = root.FindFirst(TreeScope.Descendants, sendBtnCondition);
            if (sendBtn != null)
            {
                // ボタン上部 10px をクリック
                var rect = sendBtn.Current.BoundingRectangle;
                int x = (int)(rect.Left + rect.Width / 2);
                int y = (int)(rect.Top) - 10;
                SetCursorPos(x, y);
                mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
                Thread.Sleep(3);
                SetCursorPos(originalPos.x, originalPos.y);
                return;
            }

            // フォールバック: ウィンドウ下部付近をクリック
            if (GetWindowRect(hWnd, out Rectangle rectFb))
            {
                int x = rectFb.Left + rectFb.Width / 2;
                int y = rectFb.Bottom - 50;
                SetCursorPos(x, y);
                mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
                Thread.Sleep(3);
            }
            SetCursorPos(originalPos.x, originalPos.y);
        }

        private static Bitmap CaptureScreen()
        {
            try
            {
                var bounds = Screen.PrimaryScreen.Bounds;
                var bmp = new Bitmap(bounds.Width, bounds.Height);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);
                }
                return bmp;
            }
            catch
            {
                return null;
            }
        }
        private static void CaptureAndPasteScreenshot()
        {
            var bmp = CaptureScreen();
            if (bmp != null)
            {
                Clipboard.SetImage(bmp);
                ActivateChatGPT();
                FocusChatInput();
                SendKeys.SendWait("^v");
                Console.WriteLine($"[Debug] Paste ScreenShot");
            }
        }
        private static void DumpAutomationTree()
        {
            IntPtr hWnd = GetChatGPTWindowHandle();
            if (hWnd == IntPtr.Zero) return;
            var root = AutomationElement.FromHandle(hWnd);
            if (root == null) return;

            var sb = new StringBuilder();
            void Recurse(AutomationElement el, int depth)
            {
                var c = el.Current;
                sb.AppendLine(new string(' ', depth * 2) +
                    $"[{c.ControlType.ProgrammaticName}] Name='{c.Name}', " +
                    $"AutomationId='{c.AutomationId}', ClassName='{c.ClassName}', " +
                    $"LocalizedControlType='{c.LocalizedControlType}', FrameworkId='{c.FrameworkId}', " +
                    $"HelpText='{c.HelpText}'");
                var children = el.FindAll(TreeScope.Children, Condition.TrueCondition);
                foreach (AutomationElement child in children)
                    Recurse(child, depth + 1);
            }
            Recurse(root, 0);
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automation_tree.txt");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            Console.WriteLine($"Automation tree dumped to {path}");
        }
        private static void ExecuteCommand(string command)
{
    // 複数行のコマンドを個別に実行
    var lines = command.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
    foreach (var line in lines)
    {
        var parts = line.Trim().Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) continue;
        var action = parts[0].ToLowerInvariant();
        var args = parts.Length > 1 ? parts[1] : string.Empty;

        try
        {
            switch (action)
            {
                case "open":
                    Process.Start(args);
                    break;
                case "type":
                    SendKeys.SendWait(args);
                    break;
                case "enter":
                    SendKeys.SendWait("{ENTER}");
                    break;
                case "click":
                    var coords = args.Split(',');
                    if (coords.Length == 2 &&
                        int.TryParse(coords[0], out int x) &&
                        int.TryParse(coords[1], out int y))
                    {
                        SetCursorPos(x, y);
                        mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
                        mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
                    }
                    break;
                case "wait":
                    if (int.TryParse(args, out int ms)) Thread.Sleep(ms);
                    break;
                case "screenshot":
                    CaptureAndPasteScreenshot();
                    break;
                case "prompt":
                    FocusChatInput();
                    SendKeys.SendWait(args);
                    SendKeys.SendWait("{ENTER}");
                    break;
                default:
                    Console.WriteLine($"Unknown command: {action}");
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Command execution error ({action}): {ex.Message}");
        }
    }
}
    }
}