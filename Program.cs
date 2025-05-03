using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

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
        [DllImport("user32.dll")] private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);
        [DllImport("user32.dll")] private static extern bool TranslateMessage([In] ref MSG lpMsg);
        [DllImport("user32.dll")] private static extern IntPtr DispatchMessage([In] ref MSG lpMsg);

        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 9000;
        private const uint MOD_CTRL  = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint VK_G      = 0x47; // Gキー

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

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [STAThread]
        static void Main()
        {
            // ホットキー登録: Ctrl+Shift+G
            if (!RegisterHotKey(IntPtr.Zero, HOTKEY_ID, MOD_CTRL | MOD_SHIFT, VK_G))
            {
                Console.WriteLine("ホットキーの登録に失敗しました。");
                return;
            }
            Console.WriteLine("Ctrl+Shift+G で ChatGPT をアクティブ化し、初期プロンプトを送信します。");

            // メッセージループ
            while (GetMessage(out MSG msg, IntPtr.Zero, 0, 0) != 0)
            {
                if (msg.message == WM_HOTKEY && (int)msg.wParam == HOTKEY_ID)
                {
                    ActivateChatGPT();
                    SendInitialPrompt();
                }
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }

            UnregisterHotKey(IntPtr.Zero, HOTKEY_ID);
        }

        private static void ActivateChatGPT()
        {
            // ChatGPT プロセスのウィンドウを取得
            IntPtr hWnd = GetChatGPTWindowHandle();
            if (hWnd == IntPtr.Zero)
            {
                // 見つからなければアプリ起動
                Process.Start(@"C:\Program Files\ChatGPT\ChatGPT.exe");
                Thread.Sleep(1000);
                hWnd = GetChatGPTWindowHandle();
                if (hWnd == IntPtr.Zero)
                    return;
            }

            // フォアグラウンド・アクティブ・フォーカス
            SetForegroundWindow(hWnd);
            SetActiveWindow(hWnd);
            SetFocus(hWnd);
            Thread.Sleep(200);
        }

        private static IntPtr GetChatGPTWindowHandle()
        {
            // プロセス名から見つける
            foreach (var proc in Process.GetProcessesByName("ChatGPT"))
            {
                if (proc.MainWindowHandle != IntPtr.Zero)
                    return proc.MainWindowHandle;
                // MainWindowHandle がない場合は列挙
                IntPtr found = IntPtr.Zero;
                EnumWindows((hWnd, lParam) =>
                {
                    GetWindowThreadProcessId(hWnd, out uint pid);
                    if (pid == proc.Id)
                    {
                        int length = GetWindowTextLength(hWnd);
                        if (length > 0)
                        {
                            var sb = new StringBuilder(length + 1);
                            GetWindowText(hWnd, sb, sb.Capacity);
                            if (sb.ToString().Contains("ChatGPT"))
                            {
                                found = hWnd;
                                return false; // 列挙停止
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
                if (found != IntPtr.Zero)
                    return found;
            }
            return IntPtr.Zero;
        }

        private static void SendInitialPrompt()
        {
            const string prompt = "あなたはWindows環境で動作する自動化アシスタントです。ユーザーの入力に従い、キーボード／マウス操作やファイル操作を自動化してください。準備ができたら、次のコマンドをお待ちします。";
            SendKeys.SendWait(prompt);
            // SendKeys.SendWait("{ENTER}");
        }
    }
}
