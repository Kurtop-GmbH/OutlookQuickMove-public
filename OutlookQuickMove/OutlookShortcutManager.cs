using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

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

                QuickMoveLog.Write(
                    "keyboard shortcuts started: Ctrl+M/Shift+M=move, Ctrl+G/Shift+G=go to folder.");
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
            if (!isShortcutKey)
            {
                return CallNextHookEx(hookHandle, code, virtualKey, keyData);
            }

            var controlDown = IsKeyDown(VkControl);
            var shiftDown = IsKeyDown(VkShift);
            var menuDown = IsKeyDown(VkMenu);
            var shiftAliasAllowed = shiftDown
                && !controlDown
                && !menuDown
                && IsShiftShortcutContextAllowed();
            var action = ResolveShortcutAction(
                key,
                controlDown,
                shiftDown,
                menuDown,
                shiftAliasAllowed);
            if (action == ShortcutAction.None)
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
                pendingAction = action;
                if (dispatchTimer != null)
                {
                    dispatchTimer.Start();
                }
            }

            // Outlook uses Ctrl+M for Send/Receive. The user explicitly chose Ctrl+M for Quick
            // Move; F9 remains Outlook's standard Send/Receive shortcut. Shift aliases are only
            // swallowed in the guarded Explorer context, so normal uppercase letters keep working
            // in editors and search fields.
            return new IntPtr(1);
        }

        private static ShortcutAction ResolveShortcutAction(
            int key,
            bool controlDown,
            bool shiftDown,
            bool menuDown,
            bool shiftAliasAllowed)
        {
            if ((key != VkM && key != VkG)
                || menuDown
                || controlDown == shiftDown
                || (shiftDown && !shiftAliasAllowed))
            {
                return ShortcutAction.None;
            }

            return key == VkM ? ShortcutAction.Move : ShortcutAction.GoToFolder;
        }

        private static bool IsShiftShortcutContextAllowed()
        {
            // A managed add-in dialog (especially the picker search box) must receive ordinary
            // uppercase letters instead of recursively opening another picker.
            if (Form.ActiveForm != null)
            {
                return false;
            }

            var threadInfo = new GuiThreadInfo
            {
                Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo))
            };
            if (GetGUIThreadInfo(GetCurrentThreadId(), ref threadInfo)
                && IsTextInputWindow(threadInfo.CaretWindow))
            {
                return false;
            }

            Outlook.Explorer explorer = null;
            object inlineResponse = null;
            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                inlineResponse = explorer == null ? null : explorer.ActiveInlineResponse;
                return explorer != null
                    && inlineResponse == null
                    && GetExplorerWindowHandle(explorer) == GetForegroundWindow();
            }
            catch
            {
                // If Outlook cannot prove that the Explorer is active, preserve the typed letter.
                return false;
            }
            finally
            {
                ComUtil.Release(inlineResponse);
                ComUtil.Release(explorer);
            }
        }

        private static IntPtr GetExplorerWindowHandle(Outlook.Explorer explorer)
        {
            var oleWindow = explorer as IOleWindow;
            IntPtr windowHandle;
            return oleWindow != null && oleWindow.GetWindow(out windowHandle) == 0
                ? windowHandle
                : IntPtr.Zero;
        }

        private static bool IsTextInputWindow(IntPtr window)
        {
            if (window == IntPtr.Zero)
            {
                return false;
            }

            var className = new StringBuilder(256);
            GetClassName(window, className, className.Capacity);
            return className.ToString().IndexOf("EDIT", StringComparison.OrdinalIgnoreCase) >= 0;
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

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            public uint Size;
            public uint Flags;
            public IntPtr ActiveWindow;
            public IntPtr FocusWindow;
            public IntPtr CaptureWindow;
            public IntPtr MenuOwnerWindow;
            public IntPtr MoveSizeWindow;
            public IntPtr CaretWindow;
            public NativeRect CaretRectangle;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private delegate IntPtr KeyboardHookProc(int code, IntPtr virtualKey, IntPtr keyData);

        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            [PreserveSig]
            int GetWindow(out IntPtr windowHandle);

            [PreserveSig]
            int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        }

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

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo threadInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(
            IntPtr window,
            StringBuilder className,
            int maxCount);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
    }
}
