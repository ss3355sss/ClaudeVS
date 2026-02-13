namespace ClaudeVS
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Microsoft.VisualStudio.OLE.Interop;
    using Microsoft.VisualStudio;
    using System.ComponentModel.Design;

    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    ///
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    ///
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </summary>
    [Guid("f4c7b9e2-3a5d-6c8f-1b2e-4a9d7c5f3e8b")]
    public class ClaudeTerminal : ToolWindowPane, IOleCommandTarget, IVsWindowFrameNotify3
    {
        public ConPtyTerminal Terminal => (this.Content as ClaudeTerminalControl)?.ActiveTerminal;
        public ConPtyTerminalConnection TerminalConnection => (this.Content as ClaudeTerminalControl)?.ActiveConnection;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClaudeTerminal"/> class.
        /// </summary>
        public ClaudeTerminal() : base(null)
        {
            this.Caption = "ClaudeVS";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object as this is lifetime managed by the
            // shell so the copy instance will be reused.

            this.Content = new ClaudeTerminalControl(this);
            this.ToolBar = null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                (this.Content as ClaudeTerminalControl)?.DisposeAllTerminals();
            }
            base.Dispose(disposing);
        }

        public int OnShow(int fShow)
        {
            if (fShow == (int)__FRAMESHOW.FRAMESHOW_WinShown ||
                fShow == (int)__FRAMESHOW.FRAMESHOW_TabActivated ||
                fShow == (int)__FRAMESHOW.FRAMESHOW_WinRestored ||
                fShow == 12)
            {
                var control = this.Content as ClaudeTerminalControl;
                control?.FocusTerminal();
            }
            return VSConstants.S_OK;
        }

        public int OnMove(int x, int y, int w, int h)
        {
            return VSConstants.S_OK;
        }

        public int OnSize(int x, int y, int w, int h)
        {
            return VSConstants.S_OK;
        }

        public int OnDockableChange(int fDockable, int x, int y, int w, int h)
        {
            return VSConstants.S_OK;
        }

        public int OnClose(ref uint pgrfSaveOptions)
        {
            return VSConstants.S_OK;
        }

        int IOleCommandTarget.QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            var commandService = this.GetService(typeof(IMenuCommandService)) as IOleCommandTarget;
            if (commandService != null)
            {
                return commandService.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }
            return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_UNKNOWNGROUP;
        }

        int IOleCommandTarget.Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (pguidCmdGroup == Microsoft.VisualStudio.VSConstants.GUID_VSStandardCommandSet97)
            {
                if ((VSConstants.VSStd97CmdID)nCmdID == VSConstants.VSStd97CmdID.PaneActivateDocWindow)
                {
                    if (Terminal != null && Terminal.IsRunning)
                    {
						Terminal.WriteInput("\x1b");
                    }
					return (int)Microsoft.VisualStudio.VSConstants.S_OK;
                }
            }

            var commandService = this.GetService(typeof(IMenuCommandService)) as IOleCommandTarget;
            if (commandService != null)
            {
                return commandService.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }
            return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
        }

        [DllImport("Microsoft.Terminal.Control.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, PreserveSig = true)]
        private static extern void TerminalSendKeyEvent(IntPtr terminal, ushort vkey, ushort scanCode, ushort flags, bool keyDown);

        [DllImport("Microsoft.Terminal.Control.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, PreserveSig = true)]
        private static extern void TerminalSendCharEvent(IntPtr terminal, char ch, ushort scanCode, ushort flags);

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        private static extern int ToUnicode(uint virtualKeyCode, uint scanCode, byte[] keyboardState,
            [Out, MarshalAs(UnmanagedType.LPWStr, SizeParamIndex = 5)] StringBuilder receivingBuffer,
            int bufferSize, uint flags);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_CHAR = 0x0102;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const int VK_ESCAPE = 0x1B;
        private const int VK_CONTROL = 0x11;
        private const int VK_SHIFT = 0x10;
        private const int VK_MENU = 0x12;

        protected override bool PreProcessMessage(ref System.Windows.Forms.Message m)
        {
            if (m.Msg == WM_KEYDOWN || m.Msg == WM_KEYUP ||
                m.Msg == WM_SYSKEYDOWN || m.Msg == WM_SYSKEYUP ||
                m.Msg == WM_CHAR)
            {
                int vk = m.WParam.ToInt32() & 0xFFFF;

                if (m.Msg == WM_KEYDOWN || m.Msg == WM_SYSKEYDOWN)
                {
                    if (vk == VK_ESCAPE || IsRegisteredKeybinding(vk))
                    {
                        return base.PreProcessMessage(ref m);
                    }
                }

                var control = this.Content as ClaudeTerminalControl;
                IntPtr terminalHandle = control?.TerminalHandle ?? IntPtr.Zero;

                if (terminalHandle != IntPtr.Zero)
                {
                    ushort scanCode = (ushort)(((ulong)m.LParam.ToInt64() >> 16) & 0x00FFu);
                    ushort flags = (ushort)(((ulong)m.LParam.ToInt64() >> 16) & 0xFF00u);

                    if (m.Msg == WM_KEYDOWN || m.Msg == WM_SYSKEYDOWN)
                    {
                        TerminalSendKeyEvent(terminalHandle, (ushort)vk, scanCode, flags, true);

                        byte[] keyboardState = new byte[256];
                        GetKeyboardState(keyboardState);
                        var sb = new StringBuilder(5);
                        int result = ToUnicode((uint)vk, scanCode, keyboardState, sb, sb.Capacity, 0);
                        if (result < 0)
                        {
                            sb.Clear();
                            ToUnicode((uint)vk, scanCode, keyboardState, sb, sb.Capacity, 0);
                        }
                        else if (result > 0)
                        {
                            for (int i = 0; i < result; i++)
                            {
                                TerminalSendCharEvent(terminalHandle, sb[i], scanCode, flags);
                            }
                        }
                    }
                    else if (m.Msg == WM_KEYUP || m.Msg == WM_SYSKEYUP)
                    {
                        TerminalSendKeyEvent(terminalHandle, (ushort)vk, scanCode, flags, false);
                    }
                    else if (m.Msg == WM_CHAR)
                    {
                        TerminalSendCharEvent(terminalHandle, (char)vk, scanCode, flags);
                    }

                    return true;
                }
            }

            return base.PreProcessMessage(ref m);
        }

        private bool IsRegisteredKeybinding(int vk)
        {
            bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;

            if (ctrl && !alt)
            {
                if (vk == 'T' || vk == 'R' || vk == 'O' || vk == 'B' || vk == 'V')
                    return true;
            }

            if (alt && !ctrl)
            {
                if (vk == 'V' || vk == 'T' || vk == 'S' || vk == '1' || vk == '2' || vk == '3' || vk == '4')
                    return true;
            }

            return false;
        }
    }
}
