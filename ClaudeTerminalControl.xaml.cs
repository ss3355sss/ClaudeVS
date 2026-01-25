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
	using System.Windows.Controls.Primitives;
	using EnvDTE;
	using EnvDTE80;
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

				if (ConsoleHost.IsRunning)
				{
					ConsoleHost.FocusConsole();
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
						InitializeConsoleHost();
					}

					isInitialized = true;
				}

				ConsoleHost.FocusConsole();
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ClaudeTerminalControl_Loaded: {ex}");
			}
		}

		private void InitializeConsoleHost()
		{
			try
			{
				string workingDir = GetActiveProjectDirectory();
				if (string.IsNullOrEmpty(workingDir))
				{
					workingDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
				}

				bool isDarkTheme = GetEffectiveTheme() == "Dark";
				ConsoleHost.Configure(workingDir, currentCommand, currentFontSize, isDarkTheme);
				ConsoleHost.Start();
				ConsoleHost.ProcessExited += ConsoleHost_ProcessExited;

				UpdateToolbarColors();
				var bgColor = GetThemeBackgroundColor();
				this.Background = new SolidColorBrush(bgColor);
				TerminalBorder.Background = new SolidColorBrush(bgColor);
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in InitializeTerminal: {ex}");
			}
		}

		private void ConsoleHost_ProcessExited(object sender, EventArgs e)
		{
			Debug.WriteLine("Terminal process exited");
		}

		private void ClaudeTerminalControl_SizeChanged(object sender, System.Windows.SizeChangedEventArgs e)
		{
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
				bool isDarkTheme = GetEffectiveTheme() == "Dark";
				ConsoleHost.Configure(workingDirectory, currentCommand, currentFontSize, isDarkTheme);
				ConsoleHost.Restart(workingDirectory, currentCommand);
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
				if (ConsoleHost != null && ConsoleHost.IsRunning)
				{
					if (bEnter)
						ConsoleHost.SendInputWithEnter(message);
					else
						ConsoleHost.SendInput(message);
				}
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
						ConsoleHost?.FocusConsole();
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
						string projectDir = GetActiveProjectDirectory();
						if (!string.IsNullOrEmpty(projectDir))
						{
							currentSolutionPath = projectDir;
							RestartClaudeWithWorkingDirectory(projectDir);
						}
						else
						{
							currentSolutionPath = null;
							string workingDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
							RestartClaudeWithWorkingDirectory(workingDir);
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
				string projectDir = GetActiveProjectDirectory();
				if (!string.IsNullOrEmpty(projectDir))
				{
					currentSolutionPath = projectDir;
					RestartClaudeWithWorkingDirectory(projectDir);
				}
				else
				{
					currentSolutionPath = null;
					string workingDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
					RestartClaudeWithWorkingDirectory(workingDir);
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
						ConsoleHost.SetFontSize(currentFontSize);
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in FontSizeButton_Click: {ex}");
			}
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

						UpdateToolbarColors();
						var bgColor = GetThemeBackgroundColor();
						this.Background = new SolidColorBrush(bgColor);
						TerminalBorder.Background = new SolidColorBrush(bgColor);

						bool isDark = GetEffectiveTheme() == "Dark";
						ConsoleHost.SetTheme(isDark);
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in ThemeButton_Click: {ex}");
			}
		}

		private string GetEffectiveTheme()
		{
			string effectiveTheme = currentTheme;
			if (effectiveTheme == "System")
			{
				effectiveTheme = IsSystemDarkMode() ? "Dark" : "Light";
			}
			return effectiveTheme;
		}

		private Color GetThemeBackgroundColor()
		{
			string effectiveTheme = GetEffectiveTheme();

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
			string effectiveTheme = GetEffectiveTheme();

			var bgColor = GetThemeBackgroundColor();
			SolidColorBrush toolbarBg = new SolidColorBrush(bgColor);
			SolidColorBrush buttonBg, buttonFg, buttonBorder;

			if (effectiveTheme == "Light")
			{
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
