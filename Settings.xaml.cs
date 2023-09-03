using System.Windows;

namespace ADUMP
{
    public partial class Settings : Window
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(DomainNameTextBox.Text))
            {
                e.Cancel = true;
                MessageBox.Show("Veuillez entrer un nom de domaine avant de fermer la fenêtre ou d'enregistrer les modifications.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(DomainNameTextBox.Text))
            {
                MessageBox.Show("Veuillez entrer un nom de domaine avant de fermer la fenêtre ou d'enregistrer les modifications.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Properties.Settings.Default.domainName = DomainNameTextBox.Text;
                Properties.Settings.Default.Save();
                MessageBox.Show(Properties.Settings.Default.domainName);
            }
        }
    }
}