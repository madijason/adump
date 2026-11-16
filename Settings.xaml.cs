using System;
using System.Windows;

namespace ADUMP
{
    public partial class Settings : Window
    {
        public Settings()
        {
            InitializeComponent();
            DomainNameTextBox.Text = Properties.Settings.Default.domainName;
            DelimiterBox.SelectedValue = Properties.Settings.Default.csvDelimiter;
            if (DelimiterBox.SelectedIndex < 0) DelimiterBox.SelectedIndex = 0;
        }

        private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
        {
            string domain = DomainNameTextBox.Text.Trim();
            if (!DomainConfiguration.IsValid(domain))
            {
                MessageBox.Show(this, "Entrez un nom de domaine DNS valide (ex. domaine.local).", "Configuration", MessageBoxButton.OK, MessageBoxImage.Warning);
                DomainNameTextBox.Focus();
                return;
            }
            try
            {
                Properties.Settings.Default.domainName = domain;
                Properties.Settings.Default.csvDelimiter = DelimiterBox.SelectedValue as string ?? "comma";
                Properties.Settings.Default.Save();
                DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Enregistrement impossible", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
