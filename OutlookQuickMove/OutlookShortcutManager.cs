using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    /// <summary>
    /// Provides the Thunderbird-style "shortcut, type, arrows, Enter" entry points.
    ///
    /// The hook is scoped to Outlook's UI thread; it is not a system-wide keyboard hook and cannot
    /// see keystrokes in other applications. A WinForms timer defers command execution until after
    /// the keyboard callback returns, avoiding COM and modal-dialog work inside the hook itself.
    /// </summary>
    internal static class OutlookShortcutManager
    {
        private const int WhKeyboard = 2;
        private const int VkControl = 0x11;
        private const int VkShift = 0x10;
        private const int VkMenu = 0x12;
        private const int VkG = 0x47;
        private const int VkM = 0x4D;

        private static readonly object Gate = new object();
        private static KeyboardHookProc hookProc;
        private static IntPtr hookHandle;
        private static Timer dispatchTimer;
        private static ShortcutAction pendingAction;

        public static void Start()
        {
            lock (Gate)
            {
                if (hookHandle != IntPtr.Zero)
                {
                    return;
                }

                dispatchTimer = new Timer { Interval = 1 };
                dispatchTimer.Tick += DispatchTimerTick;
                hookProc = KeyboardHookCallback;
                hookHandle = SetWindowsHookEx(
                    WhKeyboard,
                    hookProc,
                    IntPtr.Zero,
                    GetCurrentThreadId());
                if (hookHandle == IntPtr.Zero)
                {
                    dispatchTimer.Dispose();
                    dispatchTimer = null;
                    hookProc = null;
                    QuickMoveLog.Write("keyboard shortcuts could not be registered.");
                    return;
                }

                QuickMoveLog.Write("keyboard shortcuts started: Ctrl+M=move, Ctrl+G=go to folder.");
            }
        }

        public static void Stop()
        {
            lock (Gate)
            {
                if (dispatchTimer != null)
                {
                    dispatchTimer.Stop();
                    dispatchTimer.Tick -= DispatchTimerTick;
                    dispatchTimer.Dispose();
                    dispatchTimer = null;
                }

                pendingAction = ShortcutAction.None;
                if (hookHandle != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(hookHandle);
                    hookHandle = IntPtr.Zero;
                }

                hookProc = null;
            }
        }

        private static IntPtr KeyboardHookCallback(int code, IntPtr virtualKey, IntPtr keyData)
        {
            if (code < 0)
            {
                return CallNextHookEx(hookHandle, code, virtualKey, keyData);
            }

            var key = virtualKey.ToInt32();
            var isShortcutKey = key == VkM || key == VkG;
            var controlDown = IsKeyDown(VkControl);
            var otherModifierDown = IsKeyDown(VkShift) || IsKeyDown(VkMenu);
            if (!isShortcutKey || !controlDown || otherModifierDown)
            {
                return CallNextHookEx(hookHandle, code, virtualKey, keyData);
            }

            // Bit 31 is clear on key-down; bit 30 is clear only for the initial press. Swallow
            // repeats as well, but schedule just one picker.
            var flags = keyData.ToInt64();
            var isKeyDown = (flags & (1L << 31)) == 0;
            var isInitialPress = (flags & (1L << 30)) == 0;
            if (isKeyDown && isInitialPress && pendingAction == ShortcutAction.None)
            {
                pendingAction = key == VkM ? ShortcutAction.Move : ShortcutAction.GoToFolder;
                if (dispatchTimer != null)
                {
                    dispatchTimer.Start();
                }
            }

            // Outlook uses Ctrl+M for Send/Receive. The user explicitly chose Ctrl+M for Quick
            // Move; F9 remains Outlook's standard Send/Receive shortcut.
            return new IntPtr(1);
        }

        private static void DispatchTimerTick(object sender, EventArgs e)
        {
            if (dispatchTimer != null)
            {
                dispatchTimer.Stop();
            }

            var action = pendingAction;
            pendingAction = ShortcutAction.None;
            if (action == ShortcutAction.Move)
            {
                QuickMoveRibbon.RunQuickMoveCommand();
            }
            else if (action == ShortcutAction.GoToFolder)
            {
                QuickMoveRibbon.RunGoToFolderCommand();
            }
        }

        private static bool IsKeyDown(int virtualKey)
        {
            return (GetKeyState(virtualKey) & 0x8000) != 0;
        }

        private enum ShortcutAction
        {
            None,
            Move,
            GoToFolder
        }

        private delegate IntPtr KeyboardHookProc(int code, IntPtr virtualKey, IntPtr keyData);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(
            int hookType,
            KeyboardHookProc callback,
            IntPtr module,
            uint threadId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(
            IntPtr hook,
            int code,
            IntPtr virtualKey,
            IntPtr keyData);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int virtualKey);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
    }
}
