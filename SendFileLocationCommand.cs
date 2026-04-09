namespace ClaudeVS
{
    using System;
    using System.ComponentModel.Design;
    using System.Diagnostics;
    using System.IO;
    using System.Windows;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Task = System.Threading.Tasks.Task;

    internal sealed class SendFileLocationCommand
    {
        public const int CommandId = 0x0102;

        public static readonly Guid CommandSet = new Guid("a7c8e9d0-1234-5678-9abc-def012345678");

        private readonly AsyncPackage package;

        private SendFileLocationCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            if (package == null)
            {
                Debug.WriteLine("ArgumentNullException in SendFileLocationCommand constructor: package is null");
                throw new ArgumentNullException(nameof(package));
            }
            if (commandService == null)
            {
                Debug.WriteLine("ArgumentNullException in SendFileLocationCommand constructor: commandService is null");
                throw new ArgumentNullException(nameof(commandService));
            }
            this.package = package;

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static SendFileLocationCommand Instance
        {
            get;
            private set;
        }

        private IAsyncServiceProvider ServiceProvider
        {
            get { return this.package; }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SendFileLocationCommand(package, commandService);
        }

        public static string BuildLocationMessage(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dte == null || dte.ActiveDocument == null)
                return null;

            string filePath = dte.ActiveDocument.FullName;
            if (string.IsNullOrEmpty(filePath))
                return null;

            TextSelection selection = dte.ActiveDocument.Selection as TextSelection;
            int lineNumber = selection?.CurrentLine ?? 1;
            string selectedText = selection?.Text;

            string solutionDir = null;
            if (dte.Solution != null && !string.IsNullOrEmpty(dte.Solution.FullName))
                solutionDir = Path.GetDirectoryName(dte.Solution.FullName);

            string relativePath = filePath;
            if (!string.IsNullOrEmpty(solutionDir) && filePath.StartsWith(solutionDir, StringComparison.OrdinalIgnoreCase))
                relativePath = filePath.Substring(solutionDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            string message = $"@{relativePath} line {lineNumber}";
            if (!string.IsNullOrEmpty(selectedText))
                message += $"\n{selectedText}";
            message += "\n\n";

            return message;
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
                string message = BuildLocationMessage(dte);
                if (message == null)
                    return;

                Clipboard.SetText(message.Replace("\r\n", "\n"));

                try { dte.ExecuteCommand("View.Terminal"); }
                catch (Exception ex2) { Debug.WriteLine($"Failed to open VS terminal: {ex2.Message}"); }

                IVsStatusbar statusBar = Package.GetGlobalService(typeof(SVsStatusbar)) as IVsStatusbar;
                statusBar?.SetText("Copied to clipboard. Paste in terminal with Ctrl+V");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in SendFileLocationCommand Execute: {ex}");
            }
        }
    }
}
