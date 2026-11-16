using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.DirectoryServices;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ADUMP
{
    public partial class ListUser : Window
    {
        private ICollectionView? view;
        private bool loading;

        public ListUser()
        {
            InitializeComponent();
            Loaded += async (_, _) => await LoadUsersAsync();
        }

        private async Task LoadUsersAsync()
        {
            if (loading) return;
            loading = true;
            RefreshButton.IsEnabled = false;
            EmptyState.Visibility = Visibility.Collapsed;
            ListStatus.Text = "Chargement de l’annuaire…";
            try
            {
                string domain = Properties.Settings.Default.domainName.Trim();
                var users = await Task.Run(() => ReadUsers(domain));
                view = CollectionViewSource.GetDefaultView(users);
                view.Filter = MatchesSearch;
                dgUsers.ItemsSource = view;
                UpdateStatus();
            }
            catch (Exception ex)
            {
                dgUsers.ItemsSource = null;
                view = null;
                ListStatus.Text = "Impossible de charger les utilisateurs.";
                EmptyState.Text = "Vérifiez le domaine et votre connexion.";
                EmptyState.Visibility = Visibility.Visible;
                if (IsLoaded) MessageBox.Show(this, ex.Message, "Active Directory", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                loading = false;
                RefreshButton.IsEnabled = true;
            }
        }

        private static List<User> ReadUsers(string domain)
        {
            if (!DomainConfiguration.IsValid(domain))
                throw new InvalidOperationException("Configurez d’abord un nom de domaine DNS valide.");
            string ldapPath = "LDAP://CN=Users," + DomainConfiguration.DistinguishedName(domain);
            using var root = new DirectoryEntry(ldapPath);
            using var search = new DirectorySearcher(root)
            {
                Filter = "(&(objectCategory=person)(objectClass=user))",
                SearchScope = SearchScope.Subtree,
                PageSize = 500
            };
            using var results = search.FindAll();
            var users = new List<User>();
            var classes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (SearchResult result in results)
            {
                using var entry = result.GetDirectoryEntry();
                using var parent = entry.Parent;
                if (!classes.TryGetValue(parent.Path, out var className))
                {
                    using var folderSearch = new DirectorySearcher(parent)
                    {
                        Filter = "(objectClass=group)", SearchScope = SearchScope.OneLevel
                    };
                    var group = folderSearch.FindOne();
                    className = group is null || !group.Properties.Contains("cn")
                        ? "" : Convert.ToString(group.Properties["cn"][0]) ?? "";
                    classes[parent.Path] = className;
                }
                users.Add(new User
                {
                    LastName = Convert.ToString(entry.Properties["sn"].Value) ?? "",
                    FirstName = Convert.ToString(entry.Properties["givenName"].Value) ?? "",
                    Username = Convert.ToString(entry.Properties["sAMAccountName"].Value) ?? "",
                    Email = Convert.ToString(entry.Properties["mail"].Value) ?? "",
                    Class = className
                });
            }
            return users;
        }

        private bool MatchesSearch(object item)
        {
            if (item is not User user) return false;
            string query = SearchBox.Text.Trim();
            return query.Length == 0 || new[] { user.LastName, user.FirstName, user.Username, user.Email, user.Class }
                .Any(value => value.Contains(query, StringComparison.CurrentCultureIgnoreCase));
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (view is null) return;
            view.Refresh();
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            if (view is null) return;
            int count = 0;
            foreach (var item in view) count++;
            ListStatus.Text = $"{count} utilisateur(s) affiché(s)";
            EmptyState.Text = "Aucun utilisateur ne correspond à votre recherche.";
            EmptyState.Visibility = count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e) => await LoadUsersAsync();
    }

    public class User
    {
        public string LastName { get; set; } = "";
        public string FirstName { get; set; } = "";
        public string Username { get; set; } = "";
        public string Email { get; set; } = "";
        public string Class { get; set; } = "";
    }
}
