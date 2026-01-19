using System.Windows;
using System.Windows.Controls;

namespace ClaudeVS
{
    public partial class ThemeDialog : Window
    {
        public string SelectedTheme { get; private set; }

        public ThemeDialog(string currentTheme)
        {
            InitializeComponent();
            SelectedTheme = currentTheme;
            SelectTheme(currentTheme);
        }

        private void SelectTheme(string theme)
        {
            foreach (ComboBoxItem item in ThemeCombo.Items)
            {
                if (item.Tag?.ToString() == theme)
                {
                    ThemeCombo.SelectedItem = item;
                    return;
                }
            }
            ThemeCombo.SelectedIndex = 0;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (ThemeCombo.SelectedItem is ComboBoxItem selectedItem)
            {
                SelectedTheme = selectedItem.Tag?.ToString() ?? "System";
            }
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
