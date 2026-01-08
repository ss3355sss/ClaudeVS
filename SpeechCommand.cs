namespace ClaudeVS
{
    using System;
    using System.ComponentModel.Design;
    using System.Diagnostics;
    using System.Security.Principal;
    using System.Windows;
    using System.Windows.Media;
    using Windows.Media.SpeechRecognition;
    using Microsoft.VisualStudio.PlatformUI;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Task = System.Threading.Tasks.Task;

    internal sealed class SpeechCommand
    {
        public const int SpeechCommandId = 0x010B;

        public static readonly Guid CommandSet = new Guid("a7c8e9d0-1234-5678-9abc-def012345678");

        private readonly AsyncPackage package;
        private SpeechRecognizer speechRecognizer;
        private bool isListening;
        private System.Collections.Generic.Dictionary<object, object> originalResources;

        private SpeechCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            if (commandService == null)
            {
                throw new ArgumentNullException(nameof(commandService));
            }

            var menuCommandID = new CommandID(CommandSet, SpeechCommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static SpeechCommand Instance { get; private set; }

        private IAsyncServiceProvider ServiceProvider => this.package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SpeechCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (isListening)
            {
                return;
            }

            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await StartListeningAsync();
            });
        }

        private static bool IsRunningAsAdmin()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private async Task StartListeningAsync()
        {
            if (isListening)
                return;

            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                if (IsRunningAsAdmin())
                {
                    SetStatusBarText("Speech unavailable (admin mode)");
                    MessageBox.Show("Speech recognition is not available when Visual Studio is running as Administrator.\n\nThis is a Windows limitation. Please restart Visual Studio without admin privileges to use speech input.", "Speech Recognition", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SetStatusBarText("Initializing speech...");

                if (speechRecognizer == null)
                {
                    speechRecognizer = await System.Threading.Tasks.Task.Run(() => new SpeechRecognizer());

                    var compilationResult = await speechRecognizer.CompileConstraintsAsync();
                    if (compilationResult.Status != SpeechRecognitionResultStatus.Success)
                    {
                        Debug.WriteLine($"SpeechCommand: Failed to compile constraints: {compilationResult.Status}");
                        SetStatusBarText($"Speech error: {compilationResult.Status}");
                        return;
                    }
                }

                isListening = true;
                SetStatusBarBackground(true);
                Console.Beep(700, 150);
                Console.Beep(800, 200);

                SetStatusBarText("Listening...");

                Debug.WriteLine("SpeechCommand: Started listening");

                var result = await speechRecognizer.RecognizeAsync();

                isListening = false;
                Console.Beep(600, 100);
                Console.Beep(600, 100);

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                SetStatusBarBackground(false);
                SetStatusBarText("Ready");

                if (result.Status == SpeechRecognitionResultStatus.Success && !string.IsNullOrEmpty(result.Text))
                {
                    Debug.WriteLine($"SpeechCommand: Recognized: {result.Text}");
                    SendToTerminal(result.Text);
                }
                else
                {
                    Debug.WriteLine($"SpeechCommand: Recognition failed or empty: {result.Status}");
                }
            }
            catch (Exception ex)
            {
                isListening = false;
                Trace.WriteLine($"[ClaudeVS] Speech recognition error: {ex.GetType().Name}: {ex.Message} (HResult: 0x{ex.HResult:X8})");
                Trace.WriteLine($"[ClaudeVS] Stack trace: {ex.StackTrace}");
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                SetStatusBarBackground(false);
                SetStatusBarText($"Speech error: {ex.Message}");
                MessageBox.Show($"Speech recognition error:\n\n{ex.GetType().Name}: {ex.Message}\n\nHResult: 0x{ex.HResult:X8}\n\n{ex.StackTrace}", "Speech Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SendToTerminal(string text)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ToolWindowPane window = this.package.FindToolWindow(typeof(ClaudeTerminal), 0, false);
            if (window == null)
            {
                Debug.WriteLine("SpeechCommand: Terminal window not found");
                return;
            }

            var terminalWindow = window as ClaudeTerminal;
            if (terminalWindow?.Terminal != null && terminalWindow.Terminal.IsRunning)
            {
                terminalWindow.Terminal.SendToClaude(text, true);
                Debug.WriteLine($"SpeechCommand: Sent text to terminal: {text}");
            }
            else
            {
                Debug.WriteLine("SpeechCommand: Terminal not running");
            }
        }

        private void SetStatusBarText(string text)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                IVsStatusbar statusBar = package.GetService<SVsStatusbar, IVsStatusbar>();
                if (statusBar != null)
                {
                    statusBar.SetText(text);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SpeechCommand: Failed to set status bar: {ex}");
            }
        }

        private void SetStatusBarBackground(bool recording)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var resources = Application.Current.Resources;
                Color color = Colors.LimeGreen;
				var colorBrush = new SolidColorBrush(color);

                var statusBarKeys = new object[]
                {
                    EnvironmentColors.StatusBarDefaultBrushKey,
                    EnvironmentColors.StatusBarDefaultColorKey,
                    EnvironmentColors.StatusBarBuildingBrushKey,
                    EnvironmentColors.StatusBarBuildingColorKey,
                    EnvironmentColors.StatusBarDebuggingBrushKey,
                    EnvironmentColors.StatusBarDebuggingColorKey,
                    EnvironmentColors.StatusBarNoSolutionBrushKey,
                    EnvironmentColors.StatusBarNoSolutionColorKey,
                };

                if (recording)
                {
                    originalResources = new System.Collections.Generic.Dictionary<object, object>();
                    foreach (var key in statusBarKeys)
                    {
                        if (resources.Contains(key))
                        {
                            originalResources[key] = resources[key];
                        }
                        if (key.ToString().Contains("Brush"))
                        {
                            resources[key] = colorBrush;
                        }
                        else
                        {
                            resources[key] = color;
                        }
                    }
                }
                else
                {
                    if (originalResources != null)
                    {
                        foreach (var kvp in originalResources)
                        {
                            resources[kvp.Key] = kvp.Value;
                        }
                        originalResources = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SpeechCommand: Failed to set status bar background: {ex}");
            }
        }

        public void Dispose()
        {
            if (speechRecognizer != null)
            {
                speechRecognizer.Dispose();
                speechRecognizer = null;
            }
        }
    }
}
