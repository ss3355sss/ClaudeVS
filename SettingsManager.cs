using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using System;
using System.Diagnostics;

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
                Debug.WriteLine($"Exception in GetLastCommand: {ex}");
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in SaveLastCommand: {ex}");
            }
        }
    }
}
