import json
import requests
from msal import ConfidentialClientApplication
import os
import zipfile
import email
from email import policy
from bs4 import BeautifulSoup
import base64



client_id = "APP_ID"
client_secret = "SECRET"
tenant_id = "TENANT_ID"

msal_authority = f"https://login.microsoftonline.com/{tenant_id}"

msal_scope = ["https://graph.microsoft.com/.default"]

msal_app = ConfidentialClientApplication(
    client_id=client_id,
    client_credential=client_secret,
    authority=msal_authority,
)

result = msal_app.acquire_token_silent(
    scopes=msal_scope,
    account=None,
)

if not result:
    result = msal_app.acquire_token_for_client(scopes=msal_scope)

if "access_token" in result:
    access_token = result["access_token"]
else:
    raise Exception("No Access Token found")

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

# ----------------------------------------------------------------------------------------------------

                
#OPTION 1
#OPTION 1
#OPTION 1
#OPTION 1


# ----------------------------------------------------------------------------------------------------

def get_user_profile_by_email(email):
    # Construire l'URL pour obtenir le profil de l'utilisateur par courriel
    user_profile_url = f"https://graph.microsoft.com/v1.0/users/{email}"

    # Obtenez le jeton d'accès comme vous l'avez fait dans votre code initial

    # Faites une requête GET pour obtenir les informations du profil
    response = requests.get(
        url=user_profile_url,
        headers=headers,
    )

    if response.status_code == 200:
        user_profile = response.json()
        print("Informations du profil de l'utilisateur:")
        print(json.dumps(user_profile, indent=4))
    else:
        print(f"Impossible d'obtenir les informations du profil de l'utilisateur. Code de statut : {response.status_code}")

# ----------------------------------------------------------------------------------------------------

                
#OPTION 2
#OPTION 2
#OPTION 2
#OPTION 2

# ----------------------------------------------------------------------------------------------------


def get_last_5_emails_by_user(email):
    # Rechercher l'ID de l'utilisateur par courriel
    user_search_url = f"https://graph.microsoft.com/v1.0/users?$filter=mail eq '{email}'"

    # Obtenir le jeton d'accès comme dans votre code initial

    # Faire une requête GET pour rechercher l'ID de l'utilisateur
    response = requests.get(
        url=user_search_url,
        headers=headers,
    )

    if response.status_code == 200:
        user_data = response.json()
        if user_data.get("value"):
            user_id = user_data["value"][0]["id"]
            # Obtenir les 5 derniers courriels reçus de l'utilisateur
            user_emails_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages?$select=sender,subject,receivedDateTime,id&$top=5&$orderby=receivedDateTime desc"  # Ajout de 'id'

            # Faire une requête GET pour obtenir les courriels
            response = requests.get(
                url=user_emails_url,
                headers=headers,
            )

            if response.status_code == 200:
                emails = response.json()
                print("Les 5 derniers courriels reçus de l'utilisateur:")
                for email in emails.get("value", []):
                    # Assurez-vous que la structure de l'objet 'email' correspond à celle de la réponse JSON
                    print(f"ID du courriel : {email.get('id', 'N/A')}")  # Ajout de l'ID
                    print(f"Objet : {email.get('subject', 'N/A')}")
                    print(f"Reçu le : {email.get('receivedDateTime', 'N/A')}\n")
            else:
                print(f"Impossible d'obtenir les courriels de l'utilisateur. Code de statut : {response.status_code}")
        else:
            print(f"Utilisateur avec l'adresse e-mail {email} non trouvé.")
    else:
        print(f"Impossible de rechercher l'utilisateur. Code de statut : {response.status_code}")


# ----------------------------------------------------------------------------------------------------

# OPTION 3
# OPTION 3
# OPTION 3
# OPTION 3

# ----------------------------------------------------------------------------------------------------

# Fonction pour télécharger un message en format MIME
def download_message_mime(recipient_email, message_id):
    # Construisez l'URL pour obtenir le message
    email_search_url = f"https://graph.microsoft.com/v1.0/users/{recipient_email}/messages/{message_id}/$value"
    
    # Faites une requête GET pour obtenir le message
    response = requests.get(
        url=email_search_url,
        headers=headers,
    )

    if response.status_code == 200:
        message_mime_content = response.content
        
        # Créez un message MIME à partir du contenu du message
        message = email.message_from_bytes(message_mime_content, policy=policy.default)
        
        # Enregistrez le message dans un fichier texte nommé 'courriel.txt'
        with open('courriel.txt', 'w', encoding='utf-8') as file:
            file.write(message.as_string())
            
        print("Le message a été enregistré dans le fichier 'courriel.txt'.")

        # Téléchargez les pièces jointes s'il y en a
        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue  # Évitez les parties multipart
            if part.get('Content-Disposition') is None:
                continue  # Ignorez les parties sans disposition de contenu

            # Obtenez le nom de la pièce jointe
            attachment_filename = part.get_filename()
            
            if attachment_filename:
                # Enregistrez la pièce jointe localement
                with open(attachment_filename, 'wb') as attachment_file:
                    attachment_file.write(part.get_payload(decode=True))
                print(f"Pièce jointe téléchargée : {attachment_filename}")
    else:
        print(f"Impossible d'obtenir les informations du message. Code de statut : {response.status_code}")

# ----------------------------------------------------------------------------------------------------

# OPTION 4
# OPTION 4
# OPTION 4
# OPTION 4

# ----------------------------------------------------------------------------------------------------



def get_alert_details(alert_id):
    alert_url = f"https://graph.microsoft.com/v1.0/security/alerts/{alert_id}"

    response = requests.get(
        url=alert_url,
        headers=headers,
    )

    if response.status_code == 200:
        alert_details = response.json()

        sender_email = 'N/A'

        print("Détails de l'alerte:")
        print(f"Titre de l'alerte: {alert_details.get('title', 'N/A')}")
        print(f"Description: {alert_details.get('description', 'N/A')}")
        print(f"Categorie: {alert_details.get('category', 'N/A')}")
        print(f"Statut: {alert_details.get('status', 'N/A')}")
        print(f"Sévérité: {alert_details.get('severity', 'N/A')}")
        print(f"Date de création: {alert_details.get('createdDateTime', 'N/A')}")
        print(f"URL du fournisseur: {alert_details.get('sourceMaterials', ['N/A'])[0]}")
        
        print("\nUtilisateurs associés:")
        for user_state in alert_details.get('userStates', []):
            print(f"  Nom du compte: {user_state.get('accountName', 'N/A')}")
            print(f"  Domaine: {user_state.get('domainName', 'N/A')}")
            print(f"  Rôle de l'email: {user_state.get('emailRole', 'N/A')}")
            print(f"  Adresse e-mail: {user_state.get('userPrincipalName', 'N/A')}")

            if user_state.get('emailRole') == 'sender':
                sender_email = user_state.get('userPrincipalName', 'N/A')

            print()

        recipient_email = user_state.get('userPrincipalName', 'N/A')

        user_search_url = f"https://graph.microsoft.com/v1.0/users?$filter=mail eq '{recipient_email}'"


        response = requests.get(
            url=user_search_url,
            headers=headers,
        )

        if response.status_code == 200:
            user_data = response.json()
            if user_data.get("value"):
                user_id = user_data["value"][0]["id"]
                
                user_emails_url = f"https://graph.microsoft.com/v1.0/users/Briac@6mk6l2.onmicrosoft.com/messages?&$filter=from/emailAddress/address eq 'MSSecurity-noreply@microsoft.com'"

                response = requests.get(
                    url=user_emails_url,
                    headers=headers,
                )



                if response.status_code == 200:
                    emails = response.json()
                    print(f"Le courriel reçu {recipient_email} provenant de l'expéditeur {sender_email}:")
                    for email in emails.get("value", []):
                        print(f"ID du courriel : {email.get('id', 'N/A')}")  # Ajout de l'ID
                        print(f"Objet : {email.get('subject', 'N/A')}")
                        print(f"Reçu le : {email.get('receivedDateTime', 'N/A')}\n")
                else:
                    print(f"Impossible d'obtenir les courriels de l'utilisateur. Code de statut : {response.status_code}")
            else:
                print(f"Utilisateur avec l'adresse e-mail {recipient_email} non trouvé.")
        else:
            print(f"Impossible de rechercher l'utilisateur. Code de statut : {response.status_code}")

    else:
        print(f"Impossible d'obtenir les informations du message. Code de statut : {response.status_code}")



while True:
    print("Menu:")
    print("0. Quitter")
    print("1. Lire les informations d'un utilisateur")
    print("2. Lire les 5 derniers courriels d'un utilisateur")
    print("3. Télécharger le courriel en format MIME et les pièces jointes")
    print("4. Obtenir les détails d'une alerte")

    choice = input("Sélectionnez une option (0/1/2/3/4) : ")

    if choice == '0':
        break  # Quitter le programme
    elif choice == '1':
        email = input("Entrez le courriel de l'utilisateur : ")
        get_user_profile_by_email(email)
    elif choice == '2':
        email = input("Entrez le courriel de l'utilisateur : ")
        get_last_5_emails_by_user(email)
    elif choice == '3':
        recipient_email = input("Entrez l'adresse électronique du destinataire du message : ")
        message_id = input("Entrez le ID du message : ")
        download_message_mime(recipient_email, message_id)
    elif choice == '4':
        # Exécutez la fonction avec l'ID de l'alerte
        alert_id = input("Entrez l'ID de l'alerte : ")
        get_alert_details(alert_id)
    else:
        print("Option invalide. Sélectionnez 0, 1, 2, 3, 4.")
