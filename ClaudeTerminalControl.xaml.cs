namespace ClaudeVS
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Media;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.Terminal.Wpf;
    using Microsoft.VisualStudio.Shell;

    /// <summary>
    /// Interaction logic for ClaudeTerminalControl.xaml
    /// </summary>
    public partial class ClaudeTerminalControl : UserControl
    {
        private ClaudeTerminal claudeTerminal;
        private DTE2 dte;
        private SolutionEvents solutionEvents;
        private bool isInitialized;
        private string currentCommand = "claude";
        private bool needsResizeAfterOutput = false;
        private string currentSolutionPath = null;
        private short currentFontSize = 10;

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
        }

        private void ClaudeTerminalControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                currentCommand = SettingsManager.GetLastCommand();
                currentFontSize = SettingsManager.GetFontSize();

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

                var theme = new TerminalTheme
                {
                    DefaultBackground = 0xFF1e1e1e,
                    DefaultForeground = 0xFFd4d4d4,
                    DefaultSelectionBackground = 0xFF264F78,
                    CursorStyle = CursorStyle.BlinkingBar,
                    ColorTable = new uint[]
                    {
                        0xFF0C0C0C, 0xFFC50F1F, 0xFF13A10E, 0xFFC19C00,
                        0xFF0037DA, 0xFF881798, 0xFF3A96DD, 0xFFCCCCCC,
                        0xFF767676, 0xFFE74856, 0xFF16C60C, 0xFFF9F1A5,
                        0xFF3B78FF, 0xFFB4009E, 0xFF61D6D6, 0xFFF2F2F2
                    }
                };
                TerminalControl.SetTheme(theme, "Consolas", currentFontSize, Colors.Transparent);

                terminalConnection.WaitForConnectionReady();

                TerminalControl.Connection = terminalConnection;

                terminalConnection.Start();

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

        public void FocusTerminal()
        {
            try
            {
                TerminalControl.Focus();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in FocusTerminal: {ex}");
            }
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
                        if (TerminalControl.ActualHeight > 0 && TerminalControl.ActualWidth > 0)
                        {
                            var size = new Size(TerminalControl.ActualWidth, TerminalControl.ActualHeight);
                            TerminalControl.TriggerResize(size);
                        }
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
                var theme = new TerminalTheme
                {
                    DefaultBackground = 0xFF1e1e1e,
                    DefaultForeground = 0xFFd4d4d4,
                    DefaultSelectionBackground = 0xFF264F78,
                    CursorStyle = CursorStyle.BlinkingBar,
                    ColorTable = new uint[]
                    {
                        0xFF0C0C0C, 0xFFC50F1F, 0xFF13A10E, 0xFFC19C00,
                        0xFF0037DA, 0xFF881798, 0xFF3A96DD, 0xFFCCCCCC,
                        0xFF767676, 0xFFE74856, 0xFF16C60C, 0xFFF9F1A5,
                        0xFF3B78FF, 0xFFB4009E, 0xFF61D6D6, 0xFFF2F2F2
                    }
                };
                TerminalControl.SetTheme(theme, "Consolas", currentFontSize, Colors.Transparent);

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

    }

}
