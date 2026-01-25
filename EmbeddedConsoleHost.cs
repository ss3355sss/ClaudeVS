namespace ClaudeVS
{
	using System;
	using System.Diagnostics;
	using System.IO;
	using System.Runtime.InteropServices;
	using System.Text;
	using System.Threading;
	using System.Windows;
	using System.Windows.Interop;

	public class EmbeddedConsoleHost : HwndHost
	{
		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool Process32First(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool Process32Next(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

		private const uint TH32CS_SNAPPROCESS = 0x00000002;

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
		private struct PROCESSENTRY32
		{
			public uint dwSize;
			public uint cntUsage;
			public uint th32ProcessID;
			public IntPtr th32DefaultHeapID;
			public uint th32ModuleID;
			public uint cntThreads;
			public uint th32ParentProcessID;
			public int pcPriClassBase;
			public uint dwFlags;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
			public string szExeFile;
		}

		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr SetFocus(IntPtr hWnd);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll", SetLastError = true)]
		private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern uint GetCurrentThreadId();

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool FreeConsole();

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool AttachConsole(uint dwProcessId);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr GetStdHandle(int nStdHandle);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool WriteConsoleInput(IntPtr hConsoleInput, INPUT_RECORD[] lpBuffer, uint nLength, out uint lpNumberOfEventsWritten);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool GetConsoleScreenBufferInfoEx(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO_EX csbe);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool SetConsoleScreenBufferInfoEx(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO_EX csbe);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool GetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx);

		[DllImport("user32.dll")]
		private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

		private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

		[DllImport("ntdll.dll")]
		private static extern int NtQueryInformationProcess(IntPtr processHandle, int processInformationClass, ref PROCESS_BASIC_INFORMATION processInformation, int processInformationLength, out int returnLength);

		[StructLayout(LayoutKind.Sequential)]
		private struct PROCESS_BASIC_INFORMATION
		{
			public IntPtr Reserved1;
			public IntPtr PebBaseAddress;
			public IntPtr Reserved2_0;
			public IntPtr Reserved2_1;
			public IntPtr UniqueProcessId;
			public IntPtr InheritedFromUniqueProcessId;
		}

		[StructLayout(LayoutKind.Explicit)]
		private struct INPUT_RECORD
		{
			[FieldOffset(0)]
			public ushort EventType;
			[FieldOffset(4)]
			public KEY_EVENT_RECORD KeyEvent;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct KEY_EVENT_RECORD
		{
			public int bKeyDown;
			public ushort wRepeatCount;
			public ushort wVirtualKeyCode;
			public ushort wVirtualScanCode;
			public char UnicodeChar;
			public uint dwControlKeyState;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct COORD
		{
			public short X;
			public short Y;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct SMALL_RECT
		{
			public short Left;
			public short Top;
			public short Right;
			public short Bottom;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct CONSOLE_SCREEN_BUFFER_INFO_EX
		{
			public uint cbSize;
			public COORD dwSize;
			public COORD dwCursorPosition;
			public ushort wAttributes;
			public SMALL_RECT srWindow;
			public COORD dwMaximumWindowSize;
			public ushort wPopupAttributes;
			public bool bFullscreenSupported;
			[MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
			public uint[] ColorTable;
		}

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		private struct CONSOLE_FONT_INFOEX
		{
			public uint cbSize;
			public uint nFont;
			public COORD dwFontSize;
			public uint FontFamily;
			public uint FontWeight;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
			public string FaceName;
		}

		private const int GWL_STYLE = -16;
		private const int GWL_EXSTYLE = -20;

		private const int WS_CAPTION = 0x00C00000;
		private const int WS_THICKFRAME = 0x00040000;
		private const int WS_MINIMIZEBOX = 0x00020000;
		private const int WS_MAXIMIZEBOX = 0x00010000;
		private const int WS_SYSMENU = 0x00080000;
		private const int WS_CHILD = 0x40000000;
		private const int WS_VISIBLE = 0x10000000;
		private const int WS_BORDER = 0x00800000;
		private const int WS_DLGFRAME = 0x00400000;

		private const int WS_EX_WINDOWEDGE = 0x00000100;
		private const int WS_EX_CLIENTEDGE = 0x00000200;
		private const int WS_EX_DLGMODALFRAME = 0x00000001;
		private const int WS_EX_STATICEDGE = 0x00020000;

		private const int SW_SHOW = 5;
		private const int SW_HIDE = 0;

		private const int STD_INPUT_HANDLE = -10;
		private const int STD_OUTPUT_HANDLE = -11;
		private const ushort KEY_EVENT = 0x0001;

		private Process consoleProcess;
		private IntPtr consoleWindowHandle = IntPtr.Zero;
		private string workingDirectory;
		private string command;
		private short fontSize = 14;
		private bool isDarkTheme = true;
		private bool isDisposed = false;
		private bool isEmbedded = false;
		private HwndSource hwndSource;
		private System.Windows.Threading.DispatcherTimer fontRestoreTimer;

		public event EventHandler ProcessExited;

		public bool IsRunning => consoleProcess != null && !consoleProcess.HasExited;

		public EmbeddedConsoleHost()
		{
		}

		public void Configure(string workingDir, string cmd, short fontSz, bool darkTheme)
		{
			workingDirectory = workingDir;
			command = cmd;
			fontSize = fontSz;
			isDarkTheme = darkTheme;
		}

		public void Start()
		{
			if (hwndSource != null && !IsRunning)
			{
				StartConsoleProcess();
			}
		}

		protected override HandleRef BuildWindowCore(HandleRef hwndParent)
		{
			hwndSource = new HwndSource(new HwndSourceParameters("ConsoleHostWindow")
			{
				ParentWindow = hwndParent.Handle,
				WindowStyle = WS_CHILD | WS_VISIBLE,
				Width = (int)Math.Max(100, ActualWidth),
				Height = (int)Math.Max(100, ActualHeight)
			});

			return new HandleRef(this, hwndSource.Handle);
		}

		protected override void DestroyWindowCore(HandleRef hwnd)
		{
			Cleanup();
			hwndSource?.Dispose();
			hwndSource = null;
		}

		private void StartConsoleProcess()
		{
			try
			{
				string cliCommand = GetCliCommand();
				string workDir = workingDirectory ?? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

				string conhostPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "conhost.exe");

				var psi = new ProcessStartInfo
				{
					FileName = conhostPath,
					Arguments = $"\"{cliCommand}\"",
					WorkingDirectory = workDir,
					UseShellExecute = false,
					CreateNoWindow = false
				};

				consoleProcess = Process.Start(psi);
				consoleProcess.EnableRaisingEvents = true;
				consoleProcess.Exited += (s, e) =>
				{
					Dispatcher.BeginInvoke(new Action(() =>
					{
						ProcessExited?.Invoke(this, EventArgs.Empty);
					}));
				};

				ThreadPool.QueueUserWorkItem(_ => WaitAndEmbedConsole());
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in StartConsoleProcess: {ex}");
			}
		}

		private bool IsChildOfProcess(uint childPid, int parentPid)
		{
			try
			{
				using (var proc = Process.GetProcessById((int)childPid))
				{
					var pbi = new PROCESS_BASIC_INFORMATION();
					int status = NtQueryInformationProcess(proc.Handle, 0, ref pbi, Marshal.SizeOf(pbi), out _);
					if (status == 0)
					{
						return pbi.InheritedFromUniqueProcessId.ToInt32() == parentPid;
					}
				}
			}
			catch
			{
			}
			return false;
		}

		private uint GetChildProcessId(int parentPid)
		{
			IntPtr snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
			if (snapshot == IntPtr.Zero || snapshot == (IntPtr)(-1))
				return 0;

			try
			{
				var entry = new PROCESSENTRY32();
				entry.dwSize = (uint)Marshal.SizeOf(typeof(PROCESSENTRY32));

				if (Process32First(snapshot, ref entry))
				{
					do
					{
						if (entry.th32ParentProcessID == (uint)parentPid)
						{
							return entry.th32ProcessID;
						}
					} while (Process32Next(snapshot, ref entry));
				}
			}
			finally
			{
				CloseHandle(snapshot);
			}
			return 0;
		}

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool CloseHandle(IntPtr hObject);

		private void WaitAndEmbedConsole()
		{
			try
			{
				int attempts = 0;
				const int maxAttempts = 50;

				while (consoleWindowHandle == IntPtr.Zero && attempts < maxAttempts)
				{
					Thread.Sleep(20);
					attempts++;

					if (consoleProcess == null || consoleProcess.HasExited)
						return;

					int pid = consoleProcess.Id;
					EnumWindows((hWnd, lParam) =>
					{
						GetWindowThreadProcessId(hWnd, out uint windowPid);
						if (windowPid == pid || IsChildOfProcess(windowPid, pid))
						{
							StringBuilder className = new StringBuilder(256);
							GetClassName(hWnd, className, className.Capacity);
							if (className.ToString() == "ConsoleWindowClass")
							{
								consoleWindowHandle = hWnd;
								return false;
							}
						}
						return true;
					}, IntPtr.Zero);
				}

				if (consoleWindowHandle == IntPtr.Zero)
				{
					Debug.WriteLine("Failed to find console window");
					return;
				}

				Dispatcher.Invoke(new Action(() => EmbedConsoleWindow()));
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in WaitAndEmbedConsole: {ex}");
			}
		}

		private void EmbedConsoleWindow()
		{
			if (consoleWindowHandle == IntPtr.Zero || hwndSource == null)
				return;

			try
			{
				ShowWindow(consoleWindowHandle, SW_HIDE);

				int style = GetWindowLong(consoleWindowHandle, GWL_STYLE);
				style &= ~(WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU | WS_BORDER | WS_DLGFRAME);
				style |= WS_CHILD;
				SetWindowLong(consoleWindowHandle, GWL_STYLE, style);

				int exStyle = GetWindowLong(consoleWindowHandle, GWL_EXSTYLE);
				exStyle &= ~(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE | WS_EX_DLGMODALFRAME | WS_EX_STATICEDGE);
				SetWindowLong(consoleWindowHandle, GWL_EXSTYLE, exStyle);

				SetParent(consoleWindowHandle, hwndSource.Handle);

				int width = (int)ActualWidth;
				int height = (int)ActualHeight;
				if (width <= 0) width = 800;
				if (height <= 0) height = 600;

				MoveWindow(consoleWindowHandle, 0, 0, width, height, true);
				ShowWindow(consoleWindowHandle, SW_SHOW);

				isEmbedded = true;

				var settingsTimer = new System.Windows.Threading.DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(200)
				};
				settingsTimer.Tick += (s, args) =>
				{
					settingsTimer.Stop();
					ApplyConsoleSettings();
				};
				settingsTimer.Start();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in EmbedConsoleWindow: {ex}");
			}
		}

		private void ApplyConsoleSettings()
		{
			if (consoleProcess == null || consoleProcess.HasExited)
				return;

			try
			{
				FreeConsole();

				uint pidToAttach = (uint)consoleProcess.Id;
				uint childPid = GetChildProcessId(consoleProcess.Id);
				if (childPid != 0)
					pidToAttach = childPid;

				if (!AttachConsole(pidToAttach))
					return;

				IntPtr hOutput = GetStdHandle(STD_OUTPUT_HANDLE);
				if (hOutput == IntPtr.Zero || hOutput == (IntPtr)(-1))
				{
					FreeConsole();
					return;
				}

				var csbe = new CONSOLE_SCREEN_BUFFER_INFO_EX();
				csbe.cbSize = (uint)Marshal.SizeOf(csbe);
				csbe.ColorTable = new uint[16];

				if (GetConsoleScreenBufferInfoEx(hOutput, ref csbe))
				{
					if (isDarkTheme)
					{
						csbe.ColorTable[0] = 0x001E1E1E;
						csbe.ColorTable[7] = 0x00D4D4D4;
					}
					else
					{
						csbe.ColorTable[0] = 0x00FFFFFF;
						csbe.ColorTable[7] = 0x001E1E1E;
					}
					csbe.wAttributes = 0x07;
					SetConsoleScreenBufferInfoEx(hOutput, ref csbe);
				}

				var fontInfo = new CONSOLE_FONT_INFOEX();
				fontInfo.cbSize = (uint)Marshal.SizeOf(fontInfo);
				fontInfo.nFont = 0;
				fontInfo.dwFontSize.X = 0;
				fontInfo.dwFontSize.Y = fontSize;
				fontInfo.FontFamily = 54;
				fontInfo.FontWeight = 400;
				//fontInfo.FaceName = "Lucida Console";
				//fontInfo.FaceName = "Cascadia Mono";
				//fontInfo.FaceName = "Cascadia Code";
				//fontInfo.FaceName = "Cascadia";
				fontInfo.FaceName = "Consolas";
				SetCurrentConsoleFontEx(hOutput, false, ref fontInfo);

				FreeConsole();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ApplyConsoleSettings: {ex}");
				FreeConsole();
			}
		}

		private string GetCliCommand()
		{
			string cliCommand = string.IsNullOrWhiteSpace(command) ? "claude" : command.Trim();

			if (string.Equals(cliCommand, "claude", StringComparison.OrdinalIgnoreCase))
				return GetClaudeCliPath();
			else if (string.Equals(cliCommand, "copilot", StringComparison.OrdinalIgnoreCase))
				return GetCopilotCliPath();

			return cliCommand;
		}

		private string GetClaudeCliPath()
		{
			string npmPath = Path.Combine(
				Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
				"npm", "claude.cmd");
			if (File.Exists(npmPath))
				return npmPath;

			string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
			foreach (string dir in pathEnv.Split(Path.PathSeparator))
			{
				foreach (string name in new[] { "claude.cmd", "claude", "claude-code", "claude-code.cmd" })
				{
					string p = Path.Combine(dir, name);
					if (File.Exists(p)) return p;
				}
			}
			return "claude";
		}

		private string GetCopilotCliPath()
		{
			string npmPath = Path.Combine(
				Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
				"npm", "copilot.cmd");
			if (File.Exists(npmPath))
				return npmPath;

			string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
			foreach (string dir in pathEnv.Split(Path.PathSeparator))
			{
				foreach (string name in new[] { "copilot.cmd", "copilot" })
				{
					string p = Path.Combine(dir, name);
					if (File.Exists(p)) return p;
				}
			}
			return "copilot";
		}

		public void SendInput(string text)
		{
			if (consoleProcess == null || consoleProcess.HasExited)
				return;

			try
			{
				FreeConsole();

				uint pidToAttach = (uint)consoleProcess.Id;
				uint childPid = GetChildProcessId(consoleProcess.Id);
				if (childPid != 0)
					pidToAttach = childPid;

				if (!AttachConsole(pidToAttach))
					return;

				IntPtr hInput = GetStdHandle(STD_INPUT_HANDLE);
				if (hInput == IntPtr.Zero || hInput == (IntPtr)(-1))
				{
					FreeConsole();
					return;
				}

				var records = new INPUT_RECORD[text.Length * 2];
				int idx = 0;

				foreach (char c in text)
				{
					records[idx++] = new INPUT_RECORD
					{
						EventType = KEY_EVENT,
						KeyEvent = new KEY_EVENT_RECORD
						{
							bKeyDown = 1,
							wRepeatCount = 1,
							UnicodeChar = c
						}
					};
					records[idx++] = new INPUT_RECORD
					{
						EventType = KEY_EVENT,
						KeyEvent = new KEY_EVENT_RECORD
						{
							bKeyDown = 0,
							wRepeatCount = 1,
							UnicodeChar = c
						}
					};
				}

				WriteConsoleInput(hInput, records, (uint)records.Length, out _);
				FreeConsole();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SendInput: {ex}");
				FreeConsole();
			}
		}

		public void SendInputWithEnter(string text)
		{
			SendInput(text + "\r");
		}

		public void FocusConsole()
		{
			if (consoleWindowHandle == IntPtr.Zero)
				return;

			try
			{
				IntPtr foreground = GetForegroundWindow();
				uint foregroundThread = GetWindowThreadProcessId(foreground, out _);
				uint currentThread = GetCurrentThreadId();

				if (foregroundThread != currentThread)
				{
					AttachThreadInput(currentThread, foregroundThread, true);
					SetFocus(consoleWindowHandle);
					AttachThreadInput(currentThread, foregroundThread, false);
				}
				else
				{
					SetFocus(consoleWindowHandle);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in FocusConsole: {ex}");
			}
		}

		public void ResizeConsole(int width, int height)
		{
			if (consoleWindowHandle == IntPtr.Zero || !isEmbedded)
				return;

			MoveWindow(consoleWindowHandle, 0, 0, width, height, true);
		}

		public void SetTheme(bool darkTheme)
		{
			isDarkTheme = darkTheme;
			if (consoleProcess != null && !consoleProcess.HasExited)
			{
				ApplyConsoleSettings();
			}
		}

		public void SetFontSize(short newFontSize)
		{
			fontSize = newFontSize;
			if (consoleProcess != null && !consoleProcess.HasExited)
			{
				ApplyConsoleSettings();
			}
		}

		protected override IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
		{
			const int WM_SIZE = 0x0005;

			if (msg == WM_SIZE && consoleWindowHandle != IntPtr.Zero && isEmbedded)
			{
				int width = (int)(lParam.ToInt64() & 0xFFFF);
				int height = (int)((lParam.ToInt64() >> 16) & 0xFFFF);
				if (width > 0 && height > 0)
				{
					MoveWindow(consoleWindowHandle, 0, 0, width, height, true);
					ScheduleFontRestore();
				}
			}

			return base.WndProc(hwnd, msg, wParam, lParam, ref handled);
		}

		private void ScheduleFontRestore()
		{
			if (fontRestoreTimer == null)
			{
				fontRestoreTimer = new System.Windows.Threading.DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(300)
				};
				fontRestoreTimer.Tick += (s, args) =>
				{
					fontRestoreTimer.Stop();
					if (consoleProcess != null && !consoleProcess.HasExited && isEmbedded)
					{
						ApplyFontOnly();
					}
				};
			}

			fontRestoreTimer.Stop();
			fontRestoreTimer.Start();
		}

		private void ApplyFontOnly()
		{
			if (consoleProcess == null || consoleProcess.HasExited)
				return;

			try
			{
				FreeConsole();

				uint pidToAttach = (uint)consoleProcess.Id;
				uint childPid = GetChildProcessId(consoleProcess.Id);
				if (childPid != 0)
					pidToAttach = childPid;

				if (!AttachConsole(pidToAttach))
					return;

				IntPtr hOutput = GetStdHandle(STD_OUTPUT_HANDLE);
				if (hOutput == IntPtr.Zero || hOutput == (IntPtr)(-1))
				{
					FreeConsole();
					return;
				}

				var fontInfo = new CONSOLE_FONT_INFOEX();
				fontInfo.cbSize = (uint)Marshal.SizeOf(fontInfo);
				fontInfo.nFont = 0;
				fontInfo.dwFontSize.X = 0;
				fontInfo.dwFontSize.Y = fontSize;
				fontInfo.FontFamily = 54;
				fontInfo.FontWeight = 400;
				fontInfo.FaceName = "Consolas";
				SetCurrentConsoleFontEx(hOutput, false, ref fontInfo);

				FreeConsole();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ApplyFontOnly: {ex}");
				FreeConsole();
			}
		}

		private void Cleanup()
		{
			if (isDisposed)
				return;

			isDisposed = true;
			isEmbedded = false;

			try
			{
				fontRestoreTimer?.Stop();
				fontRestoreTimer = null;

				if (consoleProcess != null && !consoleProcess.HasExited)
				{
					try { consoleProcess.Kill(); } catch { }
				}
				consoleProcess?.Dispose();
				consoleProcess = null;
				consoleWindowHandle = IntPtr.Zero;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in Cleanup: {ex}");
			}
		}

		public void Restart(string workingDir, string cmd)
		{
			Cleanup();
			isDisposed = false;
			workingDirectory = workingDir;
			command = cmd;

			if (hwndSource != null)
				StartConsoleProcess();
		}
	}
}
