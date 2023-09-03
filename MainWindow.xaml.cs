using Microsoft.Win32;
using System;
using System.DirectoryServices;
using System.IO;
using System.Windows;

namespace ADUMP
{
    public partial class MainWindow : Window
    {
        private string domainName = Properties.Settings.Default.domainName;

        public MainWindow()
        {
            InitializeComponent();
            Title = "ADUMP - Active Directory User Management Program";
            if (Properties.Settings.Default.launched == false)
            {
                // Settings settingsWindow = new Settings();
                // settingsWindow.Show();
                // Properties.Settings.Default.launched = true;
                // Properties.Settings.Default.Save();
            }
        }

        // Fonction pour importer des utilisateurs
        private void ImportUsersButton_Click(object sender, RoutedEventArgs e)
        {
            String[] DC = domainName.Split(".");
            String DC1 = DC[0];
            String DC2 = DC[1];
            byte userCount = 0;

            // Demander à l'utilisateur de sélectionner un fichier CSV
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            if (openFileDialog.ShowDialog() == true)
            {
                string csvFilePath = openFileDialog.FileName;

                // Lire les données du fichier CSV
                // Lire les données du fichier CSV
                using (StreamReader reader = new StreamReader(csvFilePath))
                {
                    // Ignorer la première ligne (en-têtes de colonnes)
                    reader.ReadLine();

                    // Lire les données ligne par ligne
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] values = line.Split(',');

                        // Récupérer les valeurs des champs
                        string lastName = values[0];
                        string firstName = values[1];
                        string username = values[2];
                        string password = values[3];
                        string email = values[4];
                        string className = values[5];
                        string profilePath = values[6];
                        string homeDirectory = values[7];

                        // Vérifier si le dossier "className" existe
                        string classFolderPath = $"LDAP://CN={className},CN=Users,DC={DC1},DC={DC2}";
                        if (!DirectoryEntry.Exists(classFolderPath))
                        {
                            // Créer le dossier "className"
                            DirectoryEntry usersFolder = new DirectoryEntry($"LDAP://CN=Users,DC={DC1},DC={DC2}");
                            DirectoryEntry classFolder = usersFolder.Children.Add($"CN={className}", "container");
                            classFolder.CommitChanges();
                        }

                        // Vérifier si un groupe ayant le nom className existe
                        DirectoryEntry root = new DirectoryEntry($"LDAP://CN=Users,DC={DC1},DC={DC2}");
                        DirectorySearcher searcher = new DirectorySearcher(root);
                        searcher.Filter = $"(&(objectClass=group)(cn={className}))";
                        SearchResult result = searcher.FindOne();
                        if (result == null)
                        {
                            // Créer le groupe "className"
                            DirectoryEntry classFolder = new DirectoryEntry(classFolderPath);
                            DirectoryEntry group = classFolder.Children.Add($"CN={className}", "group");
                            group.CommitChanges();
                        }
                        result = searcher.FindOne();

                        // Ajouter l'utilisateur au dossier "className"
                        string ouPath = $"LDAP://CN={className},CN=Users,DC={DC1},DC={DC2}";
                        try
                        {
                            DirectoryEntry ou = new DirectoryEntry(ouPath);
                            DirectoryEntry user = ou.Children.Add($"CN={username}", "user");
                            user.Properties["samAccountName"].Value = username;
                            user.Properties["userPrincipalName"].Value = $"{username}@{domainName}";
                            user.Properties["givenName"].Value = firstName;
                            user.Properties["sn"].Value = lastName;
                            user.Properties["mail"].Value = email;

                            // Définir le dossier de base en local et le chemin du profil
                            user.Properties["homeDirectory"].Value = homeDirectory;
                            user.Properties["profilePath"].Value = profilePath;

                            /* DirectoryManager */
                            if (!Directory.Exists(homeDirectory))
                            {
                                Directory.CreateDirectory(homeDirectory);
                            }
                            else
                            {

                            }
                            if (!Directory.Exists(profilePath))
                            {
                                Directory.CreateDirectory(profilePath);
                            }
                            else
                            {

                            }
                            /* DirectoryManager */

                            user.CommitChanges();

                            // Définir le mot de passe de l'utilisateur
                            user.Invoke("SetPassword", password);
                            user.CommitChanges();

                            // Activer l'utilisateur
                            int val = (int)user.Properties["userAccountControl"].Value;
                            user.Properties["userAccountControl"].Value = val & ~0x2; // Désactiver le bit UF_ACCOUNTDISABLE
                            user.CommitChanges();

                            // Empêcher l'utilisateur de changer son mot de passe
                            val = (int)user.Properties["userAccountControl"].Value;
                            user.Properties["userAccountControl"].Value = val | 0x40; // Activer le bit UF_PASSWD_CANT_CHANGE
                            user.CommitChanges();

                            // Ajouter l'utilisateur comme membre du groupe ayant son className
                            DirectoryEntry group = new DirectoryEntry(result.Path);
                            group.Invoke("Add", new object[] { user.Path });
                            userCount++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
            MessageBox.Show($"{userCount} utilisateurs crées.");
        }

        private void ShowUsersButton_Click(object sender, RoutedEventArgs e)
        {
            ListUser listUser = new ListUser();
            listUser.Show();
        }

        private void ConfigurationButton_Click(object sender, RoutedEventArgs e)
        {
            Settings settingsWindow = new Settings();
            settingsWindow.Show();
        }
    }
}
