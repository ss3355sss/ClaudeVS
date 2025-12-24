namespace ClaudeVS
{
    using System;
    using System.ComponentModel.Design;
    using System.Diagnostics;
    using System.Linq;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler for Agent Actions.
    /// </summary>
    internal sealed class AgentActionCommand
    {
        public const int AgentAction1Id = 0x0105;
        public const int AgentAction2Id = 0x0106;
        public const int AgentAction3Id = 0x0107;
        public const int AgentAction4Id = 0x0108;

        /// <summary>
        /// Command set GUID.
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a7c8e9d0-1234-5678-9abc-def012345678");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AgentActionCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table).
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private AgentActionCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            if (package == null)
            {
                Debug.WriteLine("ArgumentNullException in AgentActionCommand constructor: package is null");
                throw new ArgumentNullException(nameof(package));
            }
            if (commandService == null)
            {
                Debug.WriteLine("ArgumentNullException in AgentActionCommand constructor: commandService is null");
                throw new ArgumentNullException(nameof(commandService));
            }
            this.package = package;

            AddCommand(commandService, AgentAction1Id);
            AddCommand(commandService, AgentAction2Id);
            AddCommand(commandService, AgentAction3Id);
            AddCommand(commandService, AgentAction4Id);
        }

        private void AddCommand(OleMenuCommandService commandService, int commandId)
        {
            var menuCommandID = new CommandID(CommandSet, commandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AgentActionCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider
        {
            get { return this.package; }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <returns>A task representing the async work of command initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new AgentActionCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            var menuCommand = sender as MenuCommand;
            if (menuCommand == null) return;

            string charToSend = GetInputForCommand(menuCommand.CommandID.ID);

            if (string.IsNullOrEmpty(charToSend))
            {
                Debug.WriteLine($"AgentActionCommand: No valid input mapped for command {menuCommand.CommandID.ID}");
                return;
            }

            ToolWindowPane window = this.package.FindToolWindow(typeof(ClaudeTerminal), 0, false);
            if (window == null)
            {
                Debug.WriteLine($"AgentActionCommand triggered but terminal not found");
                return;
            }

            var terminalWindow = window as ClaudeTerminal;
            if (terminalWindow?.Terminal != null && terminalWindow.Terminal.IsRunning)
            {
                terminalWindow.Terminal.WriteInput(charToSend);
                Debug.WriteLine($"AgentActionCommand sent input to terminal");
            }
            else
            {
                 Debug.WriteLine($"AgentActionCommand triggered but terminal not running");
            }
        }

        private string GetInputForCommand(int commandId)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            try
            {
                DTE2 dte = this.package.GetService<DTE, DTE2>();
                if (dte == null) return null;

                string suffix;
                switch (commandId)
                {
                    case AgentAction1Id: suffix = "AgentAction1"; break;
                    case AgentAction2Id: suffix = "AgentAction2"; break;
                    case AgentAction3Id: suffix = "AgentAction3"; break;
                    case AgentAction4Id: suffix = "AgentAction4"; break;
                    default: return null;
                }

                // VS commands are typically named "PackageName.CommandName" or just "CommandName" depending on registration.
                // The LocCanonicalName is ".ClaudeVS.AgentAction1" (dot at start?)
                // Usually "ClaudeVS.AgentAction1" or "ClaudeVS.ClaudeVS.AgentAction1" if package name is ClaudeVS.
                // VSCT strings say: <LocCanonicalName>.ClaudeVS.AgentAction1</LocCanonicalName>
                // So DTE.Commands.Item("ClaudeVS.AgentAction1") should work.
                string commandName = $"ClaudeVS.{suffix}";
                Command cmd = null;
                try
                {
                    cmd = dte.Commands.Item(commandName);
                }
                catch
                {
                    Debug.WriteLine($"Could not find command {commandName}");
                }

                if (cmd == null) return null;

                object[] bindings = cmd.Bindings as object[];
                if (bindings == null || bindings.Length == 0) return null;

                // Grab the first binding (e.g., "Claude Terminal::Ctrl+T" or "Global::Ctrl+T")
                string binding = bindings[0] as string;
                if (string.IsNullOrEmpty(binding)) return null;

                // Extract keys: "Scope::Keys" -> "Keys"
                int sepIndex = binding.IndexOf("::");
                string keyStr = (sepIndex >= 0) ? binding.Substring(sepIndex + 2) : binding;

                return ConvertKeyBindingToInput(keyStr);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in GetInputForCommand: {ex}");
                return null;
            }
        }

        private string ConvertKeyBindingToInput(string keyStr)
        {
            // Expected format: "Ctrl+T", "Ctrl+Shift+R", "Alt+F4", etc.
            
            if (string.IsNullOrEmpty(keyStr)) return null;

            bool ctrl = keyStr.Contains("Ctrl");
            bool shift = keyStr.Contains("Shift");
            bool alt = keyStr.Contains("Alt");

            string[] parts = keyStr.Split('+');
            string key = parts.Last().Trim();

            if (key.Length == 1 && char.IsLetter(key[0]))
            {
                if (ctrl && !alt)
                {
                    // Ctrl+Letter (Shift doesn't change control code usually for A-Z)
                    char c = char.ToUpper(key[0]);
                    return ((char)(c - 'A' + 1)).ToString();
                }
            }

            return null;
        }
    }
}