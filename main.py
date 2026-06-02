import json

from google_auth_oauthlib.flow import InstalledAppFlow
from openpyxl import Workbook
from tkinter import filedialog
import tkinter as tk
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build

def creer_donnee_brut():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])

    chemin_output = "/".join(path.split("/")[:-1])
    nom_output = path.split("/")[-1].split(".")[0]
    with open(path, 'r') as file:
        content = file.read()
    obj = json.loads(content)

    NumberofLines = 1


    wb = Workbook()
    ws = wb.active
    ws.append(["Nom", "prix", "moyen de paiement", "date", "id_client", "type de transaction", "commentaire"])

    for transaction in obj["fiscal_receipts"]:
        list = transaction["line_data"]
        pay_data = transaction["payment_data"]
        if transaction["customer"] != None:
            client_id = transaction["customer"]["bci_id"]
        else:
            client_id = "Non renseigné"
        transaction_type = transaction["type"]
        raison_annulation = transaction["cancellation_reason"]

        i=0
        price = 0
        dict = {}
        payment_method = pay_data[0]["payment_method_label"]
        payment_date = pay_data[0]["created"]
        for i in range(len(list)):
            nom_offre = ""
            nom_prestation = list[i]["name_line_1"]
            item_price = list[i]["item_price"]
            if "-" in nom_prestation :
                split = nom_prestation.split("-")
                for e in split[0:-1]:
                    nom_offre += e

                if not nom_offre in dict: #Si on trouve une nouvelle offre dans la commande, on l'ajoute au dict
                    dict[nom_offre] = [False, 0, [], 0]
                    #[Est ce-que c'est bien une offre, le nombre de fois que l'offre a été prise,la liste des prestations différentes, prix de l'offre, ]

                if not(split[-1] in dict[nom_offre][2]):
                    dict[nom_offre][2].append(split[-1])
                    dict[nom_offre][3] += float(item_price)
                    if (split[-1] == " Accueil (5m)") or (split[-1] == " Accuiel (5m)") or (split[-1] == " accueil (5m)"):  # si c'est le produit accueil, on valide que c'est bien une offre, et on ajoute le prix de accueil à l'offre
                        dict[nom_offre][0] = True
                        dict[nom_offre][1] += 1
                elif (split[-1] in dict[nom_offre][2]) and ((split[-1] == " Accueil") or (split[-1] == " Accuiel") or (split[-1] == " accueil")):
                    dict[nom_offre][1] += 1
            else:
                ws.append([nom_prestation, item_price, payment_method, payment_date, client_id, transaction_type, raison_annulation])
                NumberofLines += 1

        for key in dict.keys():
            if dict[key][0]==True:
                nom_prestation = key
                item_price = dict[key][3]
                for i in range(dict[key][1]):
                    ws.append([nom_prestation, item_price, payment_method, payment_date, client_id, transaction_type, raison_annulation])
                    NumberofLines+=1
    ws.auto_filter.ref = f"A1:G{ws.max_row}"

    wb.save(chemin_output + "/" + nom_output + ".xlsx")
    return chemin_output + "/" + nom_output + ".xlsx", nom_output, NumberofLines

def importer_sur_le_drive(path, nom):
    creds = get_credentials()

    # Define File Metadata
    file_metadata = {
        'name': nom,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': "1XiSdGCxVyTb5xMxWU5oLBmqO1biS9fBn"
    }

    # Prepare the media payload
    media = MediaFileUpload(
        path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )

    # Execute Upload
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        uploaded_file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()

        print("Uploaded to Google Drive")
        print(f"File URL: {uploaded_file.get('webViewLink')}")
        return uploaded_file.get('id')

    except Exception as e:
        print(f"Erreur: {e}")

def get_credentials():
    SCOPES = [
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/spreadsheets'
    ]

    # Look for the saved login session (token.json)
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        creds = None

    # If it doesn't exist (or has expired), use the Google Cloud file to log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save the session to token.json for next time
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def creerTableauDyn(id, sheetId, NumberofLines):
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # On demande à l'API de nous donner les métadonnées des feuilles du fichier
    spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=id).execute()
    sheets = spreadsheet_metadata.get('sheets', [])

    # on récupère l'id de la première feuille
    source_sheet_id = sheets[0]['properties']['sheetId']

    SPREADSHEET_ID = id
    SOURCE_SHEET_ID = source_sheet_id
    TARGET_SHEET_ID = sheetId

    # 2. Construct the batchUpdate Payload
    body = {
        "requests": [
            {
                "updateCells": {
                    "start": {
                        "sheetId": TARGET_SHEET_ID,
                        "rowIndex": 0,
                        "columnIndex": 0
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "pivotTable": {
                                        "source": {
                                            "sheetId": SOURCE_SHEET_ID,
                                            "startRowIndex": 0,
                                            "endRowIndex": NumberofLines,
                                            "startColumnIndex": 0,
                                            "endColumnIndex": 6
                                        },
                                        # ROW LAYOUT: Group by "Nom" (Column Index 0)
                                        "rows": [
                                            {
                                                "sourceColumnOffset": 0,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # COLUMN LAYOUT: Group by "Transaction" (Column Index 5)
                                        "columns": [
                                            {
                                                "sourceColumnOffset": 5,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # VALUES/AGGREGATION: Sum of "prix" (Column Index 1)
                                        "values": [
                                            {
                                                "summarizeFunction": "SUM",
                                                "sourceColumnOffset": 1,
                                                "name": "Total Sales"  # Custom header name
                                            }
                                        ],
                                        "valueLayout": "HORIZONTAL"
                                    }
                                }
                            ]
                        }
                    ],
                    "fields": "pivotTable"
                }
            }
        ]
    }

    # Execute the API Request
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=body
    ).execute()

    print("Sheets Pivot Table created")

def nvOnglet(id):
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # On commence par créer un nouvel onglet pour le Pivot
    request_body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": "Pivot Summary"
                    }
                }
            }
        ]
    }

    # On exécute la création de l'onglet
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=id,
        body=request_body
    ).execute()

    # On récupère l'ID généré automatiquement par Google pour ce nouvel onglet
    return response['replies'][0]['addSheet']['properties']['sheetId']

#Pour une future version
def update_sheet(nom):
    root = tk.Tk()
    root.withdraw()
    file_p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
    path = os.path(file_p)



if __name__ == "__main__" :
    path, nom, NumberofLines = creer_donnee_brut()
    id = importer_sur_le_drive(path, nom)
    sheetId = nvOnglet(id)
    creerTableauDyn(id, sheetId, NumberofLines)


