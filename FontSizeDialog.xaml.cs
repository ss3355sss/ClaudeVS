using System.Windows;
using System.Windows.Controls;

namespace ClaudeVS
{
    public partial class FontSizeDialog : Window
    {
        public short SelectedFontSize { get; private set; }

        public FontSizeDialog(short currentFontSize)
        {
            InitializeComponent();
            SelectedFontSize = currentFontSize;
            SelectFontSize(currentFontSize);
        }

        private void SelectFontSize(short fontSize)
        {
            foreach (ComboBoxItem item in FontSizeCombo.Items)
            {
                if (item.Tag?.ToString() == fontSize.ToString())
                {
                    FontSizeCombo.SelectedItem = item;
                    return;
                }
            }
            FontSizeCombo.SelectedIndex = 2;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (FontSizeCombo.SelectedItem is ComboBoxItem selectedItem)
            {
                if (short.TryParse(selectedItem.Tag?.ToString(), out short size))
                {
                    SelectedFontSize = size;
                }
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
