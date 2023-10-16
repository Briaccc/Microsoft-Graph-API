Documentation du Code
Ce code Python se connecte à Microsoft Graph API pour interagir avec les informations de comptes et de messagerie d'un utilisateur Office 365. Il offre trois fonctionnalités principales :

Lire les informations d'un utilisateur
Cette option permet de récupérer les informations du profil d'un utilisateur à partir de son adresse e-mail.

Lire les 5 derniers courriels d'un utilisateur
Cette option recherche un utilisateur par adresse e-mail, puis affiche les détails des 5 derniers courriels reçus par cet utilisateur.

Télécharger le courriel en format MIME et les pièces jointes
Cette option permet de télécharger un courriel en format MIME à partir de son ID. Les pièces jointes du courriel sont également téléchargées, le cas échéant.

Configuration
Avant d'exécuter ce code, vous devez effectuer certaines étapes de configuration :

Remplacez les valeurs des variables client_id, client_secret, et tenant_id par les informations d'identification de votre application dans Azure Active Directory.

Assurez-vous d'avoir installé les dépendances requises en exécutant pip install msal requests beautifulsoup4.

Assurez-vous que l'application a les autorisations nécessaires pour accéder aux ressources Microsoft Graph API.

Utilisation
Exécutez le code Python.

Un menu s'affiche, vous permettant de choisir parmi les options disponibles en entrant le numéro correspondant.

Suivez les invites pour entrer les informations nécessaires.

Notes
Assurez-vous que l'application a les autorisations appropriées pour accéder aux données Office 365, sinon les appels API échoueront.

Les informations de profil de l'utilisateur et les courriels sont récupérés à partir de Microsoft Graph API.

Les pièces jointes des courriels sont téléchargées localement.

Assurez-vous d'utiliser ce code conformément aux lois sur la confidentialité et les règles de conformité.

N'hésitez pas à personnaliser davantage ce README en fonction des besoins de votre projet ou à ajouter des informations supplémentaires pour aider les utilisateurs à comprendre et à utiliser votre code.
