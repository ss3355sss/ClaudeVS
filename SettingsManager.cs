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
        private const string FontSizeKey = "FontSize";
        private const short DefaultFontSize = 10;
        private const string ThemeKey = "Theme";
        private const string DefaultTheme = "System";

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

        public static short GetFontSize()
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

                if (!userSettingsStore.CollectionExists(CollectionPath))
                {
                    return DefaultFontSize;
                }

                if (!userSettingsStore.PropertyExists(CollectionPath, FontSizeKey))
                {
                    return DefaultFontSize;
                }

                int fontSize = userSettingsStore.GetInt32(CollectionPath, FontSizeKey);
                return (short)fontSize;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in GetFontSize: {ex}");
                return DefaultFontSize;
            }
        }

        public static void SaveFontSize(short fontSize)
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

                userSettingsStore.SetInt32(CollectionPath, FontSizeKey, fontSize);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in SaveFontSize: {ex}");
            }
        }

        public static string GetTheme()
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
                var userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

                if (!userSettingsStore.CollectionExists(CollectionPath))
                {
                    return DefaultTheme;
                }

                if (!userSettingsStore.PropertyExists(CollectionPath, ThemeKey))
                {
                    return DefaultTheme;
                }

                string theme = userSettingsStore.GetString(CollectionPath, ThemeKey);
                return string.IsNullOrWhiteSpace(theme) ? DefaultTheme : theme;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in GetTheme: {ex}");
                return DefaultTheme;
            }
        }

        public static void SaveTheme(string theme)
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

                userSettingsStore.SetString(CollectionPath, ThemeKey, theme);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in SaveTheme: {ex}");
            }
        }
    }
}
