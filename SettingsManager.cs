using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using System;

namespace ClaudeVS
{
    internal static class SettingsManager
    {
        private const string CollectionPath = "ClaudeVS";
        private const string LastCommandKey = "LastCommand";
        private const string DefaultCommand = "claude";

        public static string GetLastCommand()
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

                if (!userSettingsStore.CollectionExists(CollectionPath))
                {
                    return DefaultCommand;
                }

                if (!userSettingsStore.PropertyExists(CollectionPath, LastCommandKey))
                {
                    return DefaultCommand;
                }

                string command = userSettingsStore.GetString(CollectionPath, LastCommandKey);
                return string.IsNullOrWhiteSpace(command) ? DefaultCommand : command;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading last command: {ex.Message}");
                return DefaultCommand;
            }
        }

        public static void SaveLastCommand(string command)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

                if (!userSettingsStore.CollectionExists(CollectionPath))
                {
                    userSettingsStore.CreateCollection(CollectionPath);
                }

                userSettingsStore.SetString(CollectionPath, LastCommandKey, command);
                System.Diagnostics.Debug.WriteLine($"Saved last command: {command}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving last command: {ex.Message}");
            }
        }
    }
}
