using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Comet
{
    public static class HotKeyManager
    {
        [DllImport("user32", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("kernel32.dll")]
        static extern uint GetLastError();

        private static UInt32 ERROR_HOTKEY_ALREADY_REGISTERED = 1409;

        private static int g_hotkeyId = 0;

        public static event EventHandler<HotKeyEventArgs> HotKeyPressed;

        public static int RegisterHotKey(Keys key, KeyModifiers modifiers)
        {
            int hotkeyId = Interlocked.Increment(ref g_hotkeyId);
            IntPtr hWnd = IntPtr.Zero;
            Boolean success = RegisterHotKey(hWnd, hotkeyId, (uint)modifiers, (uint)key);

            if (!success)
            {
                UInt32 error = GetLastError();
                if (error == ERROR_HOTKEY_ALREADY_REGISTERED)
                {
                    Console.WriteLine("Error: The specified hotkey combination " +
                        "is already registered.");
                }
                else
                {
                    Console.WriteLine("Erorr while registering global " +
                        "hotkey. Error code: {0}.", error);
                }
            }

            return hotkeyId;
        }
        public static void UnRegisterHotKey(int id)
        {
            var hWnd = IntPtr.Zero;
            UnregisterHotKey(hWnd, id);
        }

        public static void OnHotKeyPressed(HotKeyEventArgs e)
        {
            // If the HotKeyPressed event has at least one subscriber
            if (HotKeyManager.HotKeyPressed != null)
            {
                HotKeyManager.HotKeyPressed(null, e);
            }
        }
    }

    public class HotKeyMessageLoop : IMessageFilter
    {
        private const int WM_HOTKEY = 0x312;
        //  private const int WM_SETTINGCHANGE = 0x1A;
        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                var e = new HotKeyEventArgs(m.LParam);
                HotKeyManager.OnHotKeyPressed(e);
                return true;
            }

            return false;
        }
    }

    public class HotKeyEventArgs : EventArgs
    {
        public readonly Keys Key;
        public readonly KeyModifiers Modifiers;

        public HotKeyEventArgs(Keys key, KeyModifiers modifiers)
        {
            this.Key = key;
            this.Modifiers = modifiers;
        }

        public HotKeyEventArgs(IntPtr hotKeyParam)
        {
            uint param = (uint)hotKeyParam.ToInt64();
            Key = (Keys)((param & 0xffff0000) >> 16);
            Modifiers = (KeyModifiers)(param & 0x0000ffff);
        }
    }

    [Flags]
    public enum KeyModifiers
    {
        None = 0,
        Alt = 1,
        Control = 2,
        Shift = 4,
        Windows = 8,
        NoRepeat = 0x4000
    }
}
