# ADUMP

Application Windows WPF en C# pour la gestion de comptes Active Directory.

## Fonctionnalités

- Tableau de bord et interface Fluent Windows.
- Thème clair/sombre suivant les paramètres du système.
- Import de comptes depuis un fichier CSV (virgule, point-virgule ou tabulation).
- Consultation, recherche et actualisation des utilisateurs du conteneur Users.
- Configuration persistante du domaine et du délimiteur.

## Prérequis et compilation

Windows, SDK .NET 10 et accès autorisé à un environnement Active Directory. Depuis la racine :

```powershell
dotnet restore ADUMP.sln
dotnet build ADUMP.sln -c Release
dotnet run --project ADUMP.csproj
```

Le thème Fluent utilise l’API WPF ThemeMode=System (.NET 9+), sans bibliothèque UI tierce.

## Utilisation

1. Ouvrir les paramètres et enregistrer le nom DNS du domaine.
2. Consulter les utilisateurs ou préparer un CSV contenant une ligne d'en-tête puis huit colonnes dans cet ordre : nom, prénom, identifiant, mot de passe, e-mail, classe, chemin de profil, répertoire personnel.
3. Sélectionner le délimiteur correspondant au fichier dans les paramètres. Vérifier les données et confirmer avant de lancer l’import.

L'application utilise le contexte Windows courant. Tester les imports dans un domaine de laboratoire : ils modifient réellement l’annuaire et peuvent créer des dossiers. Un import n’est pas transactionnel. Protéger les fichiers CSV contenant des mots de passe.

## Vérifications manuelles

- Ouvrir les trois fenêtres et changer le thème Windows clair/sombre.
- Enregistrer un domaine, rouvrir les paramètres et vérifier sa persistance.
- Ouvrir l’annuaire avec un compte autorisé ; rechercher, effacer la recherche, puis actualiser.
- Tester un import dans un domaine de laboratoire, avec un CSV valide et un fichier malformé ; vérifier aussi l’annulation du dialogue.

## Licence

Voir [LICENSE.txt](LICENSE.txt).
