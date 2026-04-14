namespace ClaudeVS
{
	using System;
	using System.Diagnostics;
	using System.Runtime.InteropServices;
	using System.Windows;
	using System.Windows.Threading;
	using EnvDTE80;
	using Microsoft.VisualStudio.Shell;
	using Microsoft.VisualStudio.Shell.Interop;

	internal static class TerminalPasteHelper
	{
		[DllImport("user32.dll")]
		private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

		private const byte VK_CONTROL = 0x11;
		private const byte VK_V = 0x56;
		private const uint KEYEVENTF_KEYUP = 0x0002;

		public static void SendToTerminal(string text)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				Clipboard.SetText(text.Replace("\r\n", "\n"));

				DTE2 dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as DTE2;
				try { dte?.ExecuteCommand("View.Terminal"); }
				catch { }

				// View.Terminal is synchronous — pane is activated when it returns.
				// ApplicationIdle fires after all pending UI/focus processing completes.
				// Then a 150ms timer ensures the terminal input handler is ready.
				Application.Current.Dispatcher.BeginInvoke(
					new Action(() =>
					{
						var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
						timer.Tick += (s, e) =>
						{
							timer.Stop();
							SimulateCtrlV();
						};
						timer.Start();
					}),
					DispatcherPriority.ApplicationIdle
				);

				IVsStatusbar statusBar = Package.GetGlobalService(typeof(SVsStatusbar)) as IVsStatusbar;
				statusBar?.SetText("Sent to terminal");
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SendToTerminal: {ex}");
			}
		}

		private static void SimulateCtrlV()
		{
			keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
			keybd_event(VK_V, 0, 0, UIntPtr.Zero);
			keybd_event(VK_V, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
			keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
		}
	}
}
