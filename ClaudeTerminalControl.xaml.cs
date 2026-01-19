namespace ClaudeVS
{
	using System;
	using System.Diagnostics;
	using System.IO;
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

		private ClaudeTerminal claudeTerminal;
		private DTE2 dte;
		private SolutionEvents solutionEvents;
		private bool isInitialized;
		private string currentCommand = "claude";
		private bool needsResizeAfterOutput = false;
		private string currentSolutionPath = null;
		private short currentFontSize = 10;
		private string currentTheme = "System";

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

				var theme = GetTerminalTheme();
				var bgColor = GetThemeBackgroundColor();
				TerminalControl.SetTheme(theme, "Consolas", currentFontSize, bgColor);
				TerminalControl.Background = new SolidColorBrush(bgColor);

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
			if (needsResizeAfterOutput)
			{
				needsResizeAfterOutput = false;
				Dispatcher.BeginInvoke(new Action(() =>
				{
					try
					{
						ApplyFontSize();
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"Exception in ConPtyTerminal_OutputReceived resize handler: {ex}");
					}
				}), System.Windows.Threading.DispatcherPriority.Render);
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
			// typically, it's +80 MaxWidth per font size but there's a trap at 16 where it's +3 compared to 14 and so requires + (3 * 80) instead of + (2 * 80)
			if (currentFontSize == 8)
				TerminalBorder.MaxWidth = 740.0;
			else if (currentFontSize == 9)
				TerminalBorder.MaxWidth = 820.0;
			else if (currentFontSize == 10)
				TerminalBorder.MaxWidth = 900.0;
			else if (currentFontSize == 11)
				TerminalBorder.MaxWidth = 980.0;
			else if (currentFontSize == 12)
				TerminalBorder.MaxWidth = 1060.0;
			else if (currentFontSize == 14)
				TerminalBorder.MaxWidth = 1220.0;
			else if (currentFontSize == 16)
				TerminalBorder.MaxWidth = 1460.0;
			else if (currentFontSize == 18)
				TerminalBorder.MaxWidth = 1620.0;
			else if (currentFontSize == 20)
				TerminalBorder.MaxWidth = 1780.0;
			else if (currentFontSize == 24)
				TerminalBorder.MaxWidth = 2100.0;
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
