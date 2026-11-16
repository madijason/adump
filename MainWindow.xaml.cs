using Microsoft.Win32;
using Microsoft.VisualBasic.FileIO;
using System;
using System.DirectoryServices;
using System.IO;
using System.Windows;

namespace ADUMP
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            RefreshDomain();
        }

        private void RefreshDomain()
        {
            string domain = Properties.Settings.Default.domainName.Trim();
            DomainStatus.Text = DomainConfiguration.IsValid(domain) ? domain : "Aucun domaine configuré";
        }

        // Fonction pour importer des utilisateurs
        private void ImportUsersButton_Click(object sender, RoutedEventArgs e)
        {
            string domainName = Properties.Settings.Default.domainName.Trim();
            if (!DomainConfiguration.IsValid(domainName))
            {
                MessageBox.Show(this, "Configurez d’abord le domaine dans les paramètres.", "Import CSV", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string domainDn = DomainConfiguration.DistinguishedName(domainName);
            int userCount = 0;

            // Demander à l'utilisateur de sélectionner un fichier CSV
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            if (openFileDialog.ShowDialog(this) == true)
            {
                string csvFilePath = openFileDialog.FileName;
                if (MessageBox.Show(this, $"Importer les comptes dans {domainName} ? Cette opération crée des comptes et dossiers.", "Confirmer l’import", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes) return;
                try
                {

                // Lire les données du fichier CSV
                // Lire les données du fichier CSV
                using (var reader = new TextFieldParser(csvFilePath))
                {
                    reader.TextFieldType = FieldType.Delimited;
                    reader.SetDelimiters(DomainConfiguration.Delimiter(Properties.Settings.Default.csvDelimiter));
                    reader.HasFieldsEnclosedInQuotes = true;
                    reader.TrimWhiteSpace = false;
                    if (!reader.EndOfData) reader.ReadFields();

                    // Lire les données ligne par ligne
                    while (!reader.EndOfData)
                    {
                        string[]? values = reader.ReadFields();
                        if (values is null || values.Length != 8)
                            throw new FormatException($"Ligne CSV invalide : 8 colonnes attendues (près de la ligne {reader.LineNumber}).");

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
                        string classFolderPath = $"LDAP://CN={className},CN=Users,{domainDn}";
                        if (!DirectoryEntry.Exists(classFolderPath))
                        {
                            // Créer le dossier "className"
                            DirectoryEntry usersFolder = new DirectoryEntry($"LDAP://CN=Users,{domainDn}");
                            DirectoryEntry classFolder = usersFolder.Children.Add($"CN={className}", "container");
                            classFolder.CommitChanges();
                        }

                        // Vérifier si un groupe ayant le nom className existe
                        DirectoryEntry root = new DirectoryEntry($"LDAP://CN=Users,{domainDn}");
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
                        string ouPath = $"LDAP://CN={className},CN=Users,{domainDn}";
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
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "Import interrompu", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                MessageBox.Show(this, $"{userCount} utilisateur(s) créé(s).", "Bilan de l’import", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ShowUsersButton_Click(object sender, RoutedEventArgs e)
        {
            var listUser = new ListUser { Owner = this };
            listUser.ShowDialog();
        }

        private void ConfigurationButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new Settings { Owner = this };
            settingsWindow.ShowDialog();
            RefreshDomain();
        }
    }
}
