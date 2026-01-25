namespace ClaudeVS
{
	using System;
	using System.Diagnostics;
	using System.IO;
	using System.Reflection;
	using System.Runtime.InteropServices;
	using System.Windows.Threading;
	using System.Windows;
	using System.Windows.Controls;
	using System.Windows.Input;
	using System.Windows.Interop;
	using System.Windows.Media;
	using System.Windows.Media.Media3D;
	using System.Windows.Controls.Primitives;
	using EnvDTE;
	using EnvDTE80;
	using Microsoft.Terminal.Wpf;
	using Microsoft.VisualStudio.Shell;

	/// <summary>
	/// Interaction logic for ClaudeTerminalControl.xaml
	/// </summary>
	public partial class ClaudeTerminalControl : UserControl
	{
		[DllImport("user32.dll")]
		private static extern IntPtr SetFocus(IntPtr hWnd);

		[DllImport("Microsoft.Terminal.Control.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, PreserveSig = true)]
		[return: MarshalAs(UnmanagedType.I1)]
		private static extern bool TerminalIsSelectionActive(IntPtr terminal);

		[DllImport("Microsoft.Terminal.Control.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall, PreserveSig = true)]
		private static extern void TerminalUserScroll(IntPtr terminal, int viewTop);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool GetCursorPos(out POINT lpPoint);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll")]
		private static extern short GetAsyncKeyState(int vKey);

		private const int VK_LBUTTON = 0x01;
		private const uint WM_MOUSEWHEEL = 0x020A;
		private const int WHEEL_DELTA = 120;
		private const int MK_LBUTTON = 0x0001;

		[StructLayout(LayoutKind.Sequential)]
		private struct POINT
		{
			public int X;
			public int Y;
		}

		private IntPtr terminalHandle = IntPtr.Zero;
		private IntPtr terminalHwnd = IntPtr.Zero;
		private DispatcherTimer selectionScrollTimer;
		private object termContainerInstance = null;
		private ScrollBar terminalScrollbar = null;
		private MethodInfo userScrollMethod = null;

		private ClaudeTerminal claudeTerminal;
		private DTE2 dte;
		private SolutionEvents solutionEvents;
		private bool isInitialized;
		private string currentCommand = "claude";
		private bool needsResizeAfterOutput = false;
		private string currentSolutionPath = null;
		private short currentFontSize = 10;
		private string currentTheme = "System";
		private DispatcherTimer refreshTimer;
		private DateTime lastOutputTime = DateTime.MinValue;

		/// <summary>
		/// Initializes a new instance of the <see cref="ClaudeTerminalControl"/> class.
		/// </summary>
		public ClaudeTerminalControl(ToolWindowPane toolWindowPane = null)
		{
			this.claudeTerminal = toolWindowPane as ClaudeTerminal;
			this.InitializeComponent();
			this.Loaded += ClaudeTerminalControl_Loaded;
			this.Unloaded += ClaudeTerminalControl_Unloaded;
			this.SizeChanged += ClaudeTerminalControl_SizeChanged;
			this.IsVisibleChanged += ClaudeTerminalControl_IsVisibleChanged;
		}

		private void ClaudeTerminalControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
		{
			try
			{
				currentCommand = SettingsManager.GetLastCommand();
				currentFontSize = SettingsManager.GetFontSize();
				currentTheme = SettingsManager.GetTheme();

				// If terminal already exists and is running, don't reinitialize
				if (claudeTerminal?.Terminal != null)
				{
					// Refresh the terminal screen
					if (TerminalControl.ActualHeight > 0 && TerminalControl.ActualWidth > 0)
					{
						var size = new Size(TerminalControl.ActualWidth, TerminalControl.ActualHeight);
						TerminalControl.TriggerResize(size);
					}

					TerminalControl.Focus();
					return;
				}

				if (!isInitialized)
				{
					dte = GetDTE();
					if (dte != null && dte.Events != null)
					{
						solutionEvents = dte.Events.SolutionEvents;
						if (solutionEvents != null)
						{
							solutionEvents.Opened += SolutionEvents_Opened;
							solutionEvents.AfterClosing += SolutionEvents_AfterClosing;
						}
					}

					string projectDir = GetActiveProjectDirectory();
					if (!string.IsNullOrEmpty(projectDir))
					{
						currentSolutionPath = projectDir;
						InitializeConPtyTerminal();
					}

					isInitialized = true;
				}

				TerminalControl.Focus();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ClaudeTerminalControl_Loaded: {ex}");
			}
		}

		private void InitializeConPtyTerminal()
		{
			try
			{
				var conPtyTerminal = new ConPtyTerminal(rows: 30, columns: 120);
				conPtyTerminal.Command = currentCommand;

				string workingDir = GetActiveProjectDirectory();
				if (string.IsNullOrEmpty(workingDir))
				{
					workingDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
				}

				bool initialized = conPtyTerminal.Initialize(workingDir);

				if (!initialized)
				{
					conPtyTerminal?.Dispose();
					return;
				}

				var terminalConnection = new ConPtyTerminalConnection(conPtyTerminal);

				conPtyTerminal.OutputReceived += ConPtyTerminal_OutputReceived;

				claudeTerminal?.SetTerminalInstances(conPtyTerminal, terminalConnection);

				terminalConnection.WaitForConnectionReady();

				TerminalControl.Connection = terminalConnection;

				ExtractTerminalHandle();

				var theme = GetTerminalTheme();
				var bgColor = GetThemeBackgroundColor();
				TerminalControl.SetTheme(theme, "Consolas", currentFontSize, bgColor);
				TerminalControl.Background = new SolidColorBrush(bgColor);
				this.Background = new SolidColorBrush(bgColor);
				UpdateToolbarColors();

				UpdateTerminalMaxWidth();

				terminalConnection.Start();

				TerminalControl.SetTheme(theme, "Consolas", currentFontSize, bgColor);

				needsResizeAfterOutput = true;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in InitializeConPtyTerminal: {ex}");
			}
		}

		private void ReconnectTerminal()
		{
			try
			{
				if (claudeTerminal?.TerminalConnection != null)
				{
					TerminalControl.Connection = claudeTerminal.TerminalConnection;
					TerminalControl.InvalidateVisual();
					TerminalControl.UpdateLayout();
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ReconnectTerminal: {ex}");
			}
		}

		private void ClaudeTerminalControl_SizeChanged(object sender, System.Windows.SizeChangedEventArgs e)
		{
			try
			{
				var terminalConnection = claudeTerminal?.TerminalConnection;
				if (terminalConnection != null && TerminalControl.ActualHeight > 0 && TerminalControl.ActualWidth > 0)
				{
					double charHeight = currentFontSize * 1.2;

					uint columns = 120;
					uint rows = (uint)Math.Max(1, TerminalControl.ActualHeight / charHeight);

					terminalConnection.Resize(rows, columns);

					var size = new Size(TerminalControl.ActualWidth, TerminalControl.ActualHeight);
					TerminalControl.TriggerResize(size);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ClaudeTerminalControl_SizeChanged: {ex}");
			}
		}

		private void ClaudeTerminalControl_Unloaded(object sender, System.Windows.RoutedEventArgs e)
		{
			try
			{
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ClaudeTerminalControl_Unloaded: {ex}");
			}
		}

		private void SolutionEvents_Opened()
		{
			try
			{
				string projectDir = GetActiveProjectDirectory();
				if (!string.IsNullOrEmpty(projectDir))
				{
					if (projectDir == currentSolutionPath)
					{
						return;
					}

					currentSolutionPath = projectDir;
					RestartClaudeWithWorkingDirectory(projectDir);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SolutionEvents_Opened: {ex}");
			}
		}

		private void SolutionEvents_AfterClosing()
		{
			try
			{
				currentSolutionPath = null;
				StopClaude();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SolutionEvents_AfterClosing: {ex}");
			}
		}

		private void RestartClaudeWithWorkingDirectory(string workingDirectory)
		{
			try
			{
				StopClaude();
				InitializeConPtyTerminal();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in RestartClaudeWithWorkingDirectory: {ex}");
			}
		}

		private void StopClaude()
		{
			try
			{
				refreshTimer?.Stop();
				selectionScrollTimer?.Stop();

				TerminalControl.Connection = null;

				var terminal = claudeTerminal?.Terminal;
				if (terminal != null)
				{
					terminal.Dispose();
				}

				claudeTerminal?.SetTerminalInstances(null, null);
				isInitialized = false;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in StopClaude: {ex}");
			}
		}

		private DTE2 GetDTE()
		{
			try
			{
				if (claudeTerminal != null)
				{
					DTE2 result = claudeTerminal.GetService<EnvDTE.DTE, EnvDTE.DTE>() as DTE2;
					if (result != null)
						return result;
				}

				return (DTE2)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.18.0");
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in GetDTE: {ex}");
				return null;
			}
		}

		private string GetActiveProjectDirectory()
		{
			try
			{
				DTE2 localDte = dte ?? GetDTE();

				if (localDte == null || localDte.Solution == null)
				{
					return null;
				}

				if (localDte.Solution.IsOpen && !string.IsNullOrEmpty(localDte.Solution.FullName))
				{
					string solutionDir = Path.GetDirectoryName(localDte.Solution.FullName);
					return solutionDir;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in GetActiveProjectDirectory: {ex}");
			}

			return null;
		}

		public void SendToClaude(string message, bool bEnter)
		{
			try
			{
				var terminal = claudeTerminal?.Terminal;
				if (terminal != null)
					terminal.SendToClaude(message, bEnter);
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SendToClaude: {ex}");
			}
		}

		private DispatcherTimer focusTimer;

		public void FocusTerminal()
		{
			try
			{
				if (focusTimer != null)
				{
					focusTimer.Stop();
				}

				focusTimer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(100)
				};
				focusTimer.Tick += (s, e) =>
				{
					focusTimer.Stop();
					try
					{
						var hwndHost = FindVisualChild<HwndHost>(TerminalControl);
						if (hwndHost != null && hwndHost.Handle != IntPtr.Zero)
						{
							SetFocus(hwndHost.Handle);
						}
						TerminalControl.Focus();
						Keyboard.Focus(TerminalControl);
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"Exception in FocusTerminal timer: {ex}");
					}
				};
				focusTimer.Start();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in FocusTerminal: {ex}");
			}
		}

		private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
		{
			for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
			{
				var child = VisualTreeHelper.GetChild(parent, i);
				if (child is T result)
					return result;
				var descendant = FindVisualChild<T>(child);
				if (descendant != null)
					return descendant;
			}
			return null;
		}

		private void ClaudeTerminalControl_GotFocus(object sender, RoutedEventArgs e)
		{
			if (e.OriginalSource == this)
			{
				FocusTerminal();
				e.Handled = true;
			}
		}

		private void ClaudeTerminalControl_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
		{
		}

		private void ConPtyTerminal_OutputReceived(object sender, string e)
		{
			lastOutputTime = DateTime.UtcNow;

			var connection = claudeTerminal?.TerminalConnection;
			if (connection != null)
			{
				connection.IsPaused = IsTerminalSelectionActive();
			}

			if (needsResizeAfterOutput)
			{
				needsResizeAfterOutput = false;
				Dispatcher.BeginInvoke(new Action(() =>
				{
					try
					{
						ApplyFontSize();
						StartRefreshTimer();
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"Exception in ConPtyTerminal_OutputReceived resize handler: {ex}");
					}
				}), System.Windows.Threading.DispatcherPriority.Render);
			}
		}

		private void ExtractTerminalHandle()
		{
			try
			{
				var termContainerField = TerminalControl.GetType().GetField("termContainer", BindingFlags.NonPublic | BindingFlags.Instance);
				if (termContainerField != null)
				{
					var termContainer = termContainerField.GetValue(TerminalControl);
					if (termContainer != null)
					{
						termContainerInstance = termContainer;

						var terminalField = termContainer.GetType().GetField("terminal", BindingFlags.NonPublic | BindingFlags.Instance);
						if (terminalField != null)
						{
							terminalHandle = (IntPtr)terminalField.GetValue(termContainer);
						}

						var hwndField = termContainer.GetType().GetField("hwnd", BindingFlags.NonPublic | BindingFlags.Instance);
						if (hwndField != null)
						{
							terminalHwnd = (IntPtr)hwndField.GetValue(termContainer);
						}

						userScrollMethod = termContainer.GetType().GetMethod("UserScroll", BindingFlags.NonPublic | BindingFlags.Instance);
					}
				}

				var scrollbarField = TerminalControl.GetType().GetField("scrollbar", BindingFlags.NonPublic | BindingFlags.Instance);
				if (scrollbarField != null)
				{
					terminalScrollbar = scrollbarField.GetValue(TerminalControl) as ScrollBar;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ExtractTerminalHandle: {ex}");
			}
		}

		private bool IsTerminalSelectionActive()
		{
			try
			{
				if (terminalHandle != IntPtr.Zero)
				{
					return TerminalIsSelectionActive(terminalHandle);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in IsTerminalSelectionActive: {ex}");
			}
			return false;
		}

		private void StartRefreshTimer()
		{
			if (refreshTimer == null)
			{
				refreshTimer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(500)
				};
				refreshTimer.Tick += RefreshTimer_Tick;
			}
			refreshTimer.Start();

			StartSelectionScrollTimer();
		}

		private void StartSelectionScrollTimer()
		{
			if (selectionScrollTimer == null)
			{
				selectionScrollTimer = new DispatcherTimer(DispatcherPriority.Input)
				{
					Interval = TimeSpan.FromMilliseconds(16)
				};
				selectionScrollTimer.Tick += SelectionScrollTimer_Tick;
			}
			selectionScrollTimer.Start();
		}

		private bool IsLeftMouseButtonDown()
		{
			return (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
		}

		private bool wasMouseButtonDown = false;

		private void SelectionScrollTimer_Tick(object sender, EventArgs e)
		{
			try
			{
				if (terminalHwnd == IntPtr.Zero)
				{
					return;
				}

				bool isMouseDown = IsLeftMouseButtonDown();

				if (!isMouseDown)
				{
					wasMouseButtonDown = false;
					return;
				}

				wasMouseButtonDown = true;

				if (!GetCursorPos(out POINT cursorPos))
				{
					return;
				}

				var terminalPoint = TerminalControl.PointFromScreen(new Point(cursorPos.X, cursorPos.Y));
				double edgeMargin = 20;
				double terminalTop = edgeMargin;
				double terminalBottom = TerminalControl.ActualHeight - edgeMargin;

				bool isOutsideBounds = terminalPoint.Y < terminalTop || terminalPoint.Y > terminalBottom;
				if (!isOutsideBounds)
				{
					return;
				}

				int wheelDelta = 0;

				if (terminalPoint.Y < terminalTop)
				{
					double distance = terminalTop - terminalPoint.Y;
					int scrollSpeed = Math.Max(1, Math.Min(3, (int)(distance / 50) + 1));
					wheelDelta = (WHEEL_DELTA / 2) * scrollSpeed;
				}
				else if (terminalPoint.Y > terminalBottom)
				{
					double distance = terminalPoint.Y - terminalBottom;
					int scrollSpeed = Math.Max(1, Math.Min(3, (int)(distance / 50) + 1));
					wheelDelta = -(WHEEL_DELTA / 2) * scrollSpeed;
				}

				if (wheelDelta != 0)
				{
					int wParam = (wheelDelta << 16) | MK_LBUTTON;
					int lParam = (cursorPos.Y << 16) | (cursorPos.X & 0xFFFF);
					SendMessage(terminalHwnd, WM_MOUSEWHEEL, (IntPtr)wParam, (IntPtr)lParam);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SelectionScrollTimer_Tick: {ex}");
			}
		}

		private void RefreshTimer_Tick(object sender, EventArgs e)
		{
			try
			{
				bool isSelecting = IsTerminalSelectionActive();

				var connection = claudeTerminal?.TerminalConnection;
				if (connection != null)
				{
					connection.IsPaused = isSelecting;
				}

				if (!isSelecting &&
					(DateTime.UtcNow - lastOutputTime).TotalMilliseconds < 1000 &&
					TerminalControl.ActualHeight > 0 && TerminalControl.ActualWidth > 0)
				{
					var theme = GetTerminalTheme();
					var bgColor = GetThemeBackgroundColor();
					TerminalControl.SetTheme(theme, "Consolas", currentFontSize, bgColor);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in RefreshTimer_Tick: {ex}");
			}
		}

		private void ChangeCommandButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				var dialog = new CommandInputDialog(currentCommand);
				dialog.Owner = Application.Current?.MainWindow;
				if (dialog.ShowDialog() == true)
				{
					string newCommand = dialog.CommandName?.Trim();
					if (!string.IsNullOrWhiteSpace(newCommand) && !string.Equals(newCommand, currentCommand, StringComparison.OrdinalIgnoreCase))
					{
						currentCommand = newCommand;
						SettingsManager.SaveLastCommand(currentCommand);
						needsResizeAfterOutput = true;
						string projectDir = GetActiveProjectDirectory();
						if (!string.IsNullOrEmpty(projectDir))
						{
							currentSolutionPath = projectDir;
							RestartClaudeWithWorkingDirectory(projectDir);
						}
						else
						{
							currentSolutionPath = null;
							StopClaude();
							InitializeConPtyTerminal();
						}
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ChangeCommandButton_Click: {ex}");
			}
		}

		private void RestartAgentButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				needsResizeAfterOutput = true;
				string projectDir = GetActiveProjectDirectory();
				if (!string.IsNullOrEmpty(projectDir))
				{
					currentSolutionPath = projectDir;
					RestartClaudeWithWorkingDirectory(projectDir);
				}
				else
				{
					currentSolutionPath = null;
					StopClaude();
					InitializeConPtyTerminal();
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in RestartAgentButton_Click: {ex}");
			}
		}

		private void FontSizeButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				var dialog = new FontSizeDialog(currentFontSize);
				dialog.Owner = Application.Current?.MainWindow;
				if (dialog.ShowDialog() == true)
				{
					short newFontSize = dialog.SelectedFontSize;
					if (newFontSize != currentFontSize)
					{
						currentFontSize = newFontSize;
						SettingsManager.SaveFontSize(currentFontSize);
						ApplyFontSize();
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in FontSizeButton_Click: {ex}");
			}
		}

		private void ApplyFontSize()
		{
			try
			{
				var theme = GetTerminalTheme();
				var bgColor = GetThemeBackgroundColor();
				TerminalControl.SetTheme(theme, "Consolas", currentFontSize, bgColor);
				TerminalControl.Background = new SolidColorBrush(bgColor);
				this.Background = new SolidColorBrush(bgColor);
				UpdateToolbarColors();

				UpdateTerminalMaxWidth();

				if (TerminalControl.ActualHeight > 0 && TerminalControl.ActualWidth > 0)
				{
					var size = new Size(TerminalControl.ActualWidth, TerminalControl.ActualHeight);
					TerminalControl.TriggerResize(size);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ApplyFontSize: {ex}");
			}
		}

		private void UpdateTerminalMaxWidth()
		{
			// typically, it's +80 Width per font size but there's a trap at 16 where it's +3 compared to 14 and so requires + (3 * 80) instead of + (2 * 80)
			if (currentFontSize == 8)
				TerminalBorder.Width = 740.0;
			else if (currentFontSize == 9)
				TerminalBorder.Width = 820.0;
			else if (currentFontSize == 10)
				TerminalBorder.Width = 900.0;
			else if (currentFontSize == 11)
				TerminalBorder.Width = 980.0;
			else if (currentFontSize == 12)
				TerminalBorder.Width = 1060.0;
			else if (currentFontSize == 14)
				TerminalBorder.Width = 1220.0;
			else if (currentFontSize == 16)
				TerminalBorder.Width = 1460.0;
			else if (currentFontSize == 18)
				TerminalBorder.Width = 1620.0;
			else if (currentFontSize == 20)
				TerminalBorder.Width = 1780.0;
			else if (currentFontSize == 24)
				TerminalBorder.Width = 2100.0;
		}

		private void ThemeButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				var dialog = new ThemeDialog(currentTheme);
				dialog.Owner = Application.Current?.MainWindow;
				if (dialog.ShowDialog() == true)
				{
					string newTheme = dialog.SelectedTheme;
					if (newTheme != currentTheme)
					{
						currentTheme = newTheme;
						SettingsManager.SaveTheme(currentTheme);
						needsResizeAfterOutput = true;
						string projectDir = GetActiveProjectDirectory();
						if (!string.IsNullOrEmpty(projectDir))
						{
							currentSolutionPath = projectDir;
							RestartClaudeWithWorkingDirectory(projectDir);
						}
						else
						{
							currentSolutionPath = null;
							StopClaude();
							InitializeConPtyTerminal();
						}
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ThemeButton_Click: {ex}");
			}
		}

		private void ApplyTheme()
		{
			try
			{
				var theme = GetTerminalTheme();
				var bgColor = GetThemeBackgroundColor();
				TerminalControl.SetTheme(theme, "Consolas", currentFontSize, bgColor);
				TerminalControl.Background = new SolidColorBrush(bgColor);
				this.Background = new SolidColorBrush(bgColor);
				UpdateToolbarColors();

				if (TerminalControl.ActualHeight > 0 && TerminalControl.ActualWidth > 0)
				{
					var size = new Size(TerminalControl.ActualWidth, TerminalControl.ActualHeight);
					TerminalControl.TriggerResize(size);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ApplyTheme: {ex}");
			}
		}

		private Color GetThemeBackgroundColor()
		{
			string effectiveTheme = currentTheme;
			if (effectiveTheme == "System")
			{
				effectiveTheme = IsSystemDarkMode() ? "Dark" : "Light";
			}

			if (effectiveTheme == "Light")
			{
				return Color.FromRgb(0xFF, 0xFF, 0xFF);
			}
			else
			{
				return Color.FromRgb(0x1e, 0x1e, 0x1e);
			}
		}

		private void UpdateToolbarColors()
		{
			string effectiveTheme = currentTheme;
			if (effectiveTheme == "System")
			{
				effectiveTheme = IsSystemDarkMode() ? "Dark" : "Light";
			}

			SolidColorBrush toolbarBg, buttonBg, buttonFg, buttonBorder;

			if (effectiveTheme == "Light")
			{
				toolbarBg = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8));
				buttonBg = new SolidColorBrush(Color.FromRgb(0xF5, 0xF5, 0xF5));
				buttonFg = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));
				buttonBorder = new SolidColorBrush(Color.FromRgb(0xF5, 0xF5, 0xF5));

				MicButton.Style = (Style)FindResource("LightToggleButtonStyle");
				ChangeCommandButton.Style = (Style)FindResource("LightButtonStyle");
				RestartAgentButton.Style = (Style)FindResource("LightButtonStyle");
				ThemeButton.Style = (Style)FindResource("LightButtonStyle");
				FontSizeButton.Style = (Style)FindResource("LightButtonStyle");
			}
			else
			{
				toolbarBg = new SolidColorBrush(Color.FromRgb(0x0C, 0x0C, 0x0C));
				buttonBg = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));
				buttonFg = new SolidColorBrush(Color.FromRgb(0xD4, 0xD4, 0xD4));
				buttonBorder = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));

				MicButton.Style = (Style)FindResource("DarkToggleButtonStyle");
				ChangeCommandButton.Style = (Style)FindResource("DarkButtonStyle");
				RestartAgentButton.Style = (Style)FindResource("DarkButtonStyle");
				ThemeButton.Style = (Style)FindResource("DarkButtonStyle");
				FontSizeButton.Style = (Style)FindResource("DarkButtonStyle");
			}

			ToolbarBorder.Background = toolbarBg;

			MicButton.Background = buttonBg;
			MicButton.Foreground = buttonFg;
			MicButton.BorderBrush = buttonBorder;

			ChangeCommandButton.Background = buttonBg;
			ChangeCommandButton.Foreground = buttonFg;
			ChangeCommandButton.BorderBrush = buttonBorder;

			RestartAgentButton.Background = buttonBg;
			RestartAgentButton.Foreground = buttonFg;
			RestartAgentButton.BorderBrush = buttonBorder;

			ThemeButton.Background = buttonBg;
			ThemeButton.Foreground = buttonFg;
			ThemeButton.BorderBrush = buttonBorder;

			FontSizeButton.Background = buttonBg;
			FontSizeButton.Foreground = buttonFg;
			FontSizeButton.BorderBrush = buttonBorder;
		}

		private TerminalTheme GetTerminalTheme()
		{
			string effectiveTheme = currentTheme;
			if (effectiveTheme == "System")
			{
				effectiveTheme = IsSystemDarkMode() ? "Dark" : "Light";
			}

			if (effectiveTheme == "Light")
			{
				return new TerminalTheme
				{
					DefaultBackground = 0xFFFFFFFF,
					DefaultForeground = 0xFF000000,
					DefaultSelectionBackground = 0xFF0078D7,
					CursorStyle = CursorStyle.BlinkingBar,
					ColorTable = new uint[]
					{
						0xFFFFFFFF, 0xFFC50F1F, 0xFF13A10E, 0xFFC19C00,
						0xFF0037DA, 0xFF881798, 0xFF3A96DD, 0xFF000000,
						0xFFFFFFFF, 0xFFE74856, 0xFF16C60C, 0xFFF9F1A5,
						0xFF3B78FF, 0xFFB4009E, 0xFF61D6D6, 0xFF000000
					}
				};
			}
			else
			{
				return new TerminalTheme
				{
					DefaultBackground = 0xFF1e1e1e,
					DefaultForeground = 0xFFd4d4d4,
					DefaultSelectionBackground = 0xFF264F78,
					CursorStyle = CursorStyle.BlinkingBar,
					ColorTable = new uint[]
					{
						0xFF1e1e1e, 0xFFC50F1F, 0xFF13A10E, 0xFFC19C00,
						0xFF0037DA, 0xFF881798, 0xFF3A96DD, 0xFFCCCCCC,
						0xFF1e1e1e, 0xFFE74856, 0xFF16C60C, 0xFFF9F1A5,
						0xFF3B78FF, 0xFFB4009E, 0xFF61D6D6, 0xFFF2F2F2
					}
				};
			}
		}

		private bool IsSystemDarkMode()
		{
			try
			{
				using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
				{
					if (key != null)
					{
						var value = key.GetValue("AppsUseLightTheme");
						if (value is int intValue)
						{
							return intValue == 0;
						}
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in IsSystemDarkMode: {ex}");
			}
			return true;
		}

		private void MicButton_Checked(object sender, RoutedEventArgs e)
		{
			try
			{
				var speechCommand = SpeechCommand.Instance;
				if (speechCommand != null)
				{
					speechCommand.ListeningStateChanged -= OnListeningStateChanged;
					speechCommand.ListeningStateChanged += OnListeningStateChanged;
					speechCommand.StartListening();
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in MicButton_Checked: {ex}");
				MicButton.IsChecked = false;
			}
		}

		private void MicButton_Unchecked(object sender, RoutedEventArgs e)
		{
			try
			{
				var speechCommand = SpeechCommand.Instance;
				if (speechCommand != null && speechCommand.IsListening)
				{
					_ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
					{
						await speechCommand.StopListeningAsync();
					});
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in MicButton_Unchecked: {ex}");
			}
		}

		private void OnListeningStateChanged(bool isListening)
		{
			Dispatcher.BeginInvoke(new Action(() =>
			{
				MicButton.IsChecked = isListening;
			}));
		}

	}

}
