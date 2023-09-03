using System.DirectoryServices;
using System.Collections.ObjectModel;
using System.Windows;
using System;
using System.Text.RegularExpressions;

namespace ADUMP
{
    public partial class ListUser : Window
    {
        public ListUser()
        {
            InitializeComponent();
            LoadUsers();
        }

        private void LoadUsers()
        {
            try {
                // Récupération du nom de domaine et détermination des valeurs de DC1 et DC2
                string domainName = Properties.Settings.Default.domainName;
                String[] DC = domainName.Split(".");
                String DC1 = DC[0];
                String DC2 = DC[1];

                // Construction du chemin d'accès LDAP à partir des valeurs de DC1 et DC2
                string ldapPath = $"LDAP://CN=Users,DC={DC1},DC={DC2}";
                DirectoryEntry directoryEntry = new DirectoryEntry(ldapPath);
                DirectorySearcher directorySearcher = new DirectorySearcher(directoryEntry);
                directorySearcher.Filter = "(objectClass=user)";
                directorySearcher.SearchScope = SearchScope.Subtree;
                SearchResultCollection searchResults = directorySearcher.FindAll();

                // Création d'une collection d'objets User à partir des résultats de la recherche
                ObservableCollection<User> users = new ObservableCollection<User>();
                foreach (SearchResult searchResult in searchResults)
                {
                    DirectoryEntry userEntry = searchResult.GetDirectoryEntry();
                    User user = new User();
                    user.LastName = (string)userEntry.Properties["sn"].Value;
                    user.FirstName = (string)userEntry.Properties["givenName"].Value;
                    user.Username = (string)userEntry.Properties["sAMAccountName"].Value;
                    user.Email = (string)userEntry.Properties["mail"].Value;

                    // Recherche du premier groupe dans le même dossier que l'utilisateur
                    string userFolder = userEntry.Parent.Path;
                    DirectoryEntry folderEntry = new DirectoryEntry(userFolder);
                    DirectorySearcher folderSearcher = new DirectorySearcher(folderEntry);
                    folderSearcher.Filter = "(objectClass=group)";
                    folderSearcher.SearchScope = SearchScope.OneLevel;
                    SearchResult groupResult = folderSearcher.FindOne();
                    if (groupResult != null)
                    {
                        // Affectation du nom du groupe à la propriété "Class" de l'objet User
                        DirectoryEntry groupEntry = groupResult.GetDirectoryEntry();
                        user.Class = (string)groupEntry.Properties["cn"].Value;
                    }
                    else
                    {
                        user.Class = "";
                    }

                    users.Add(user);
                }
                dgUsers.ItemsSource = users;
            } catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }

    public class User
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Username { get; set; }
        public string Email { get; set; }
        public string Class { get; set; }
    }
}
