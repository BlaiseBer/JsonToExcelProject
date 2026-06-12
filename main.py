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
from groq import Groq

def creer_donnee_brut(path,categories):
    nom_output = os.path.splitext(os.path.basename(path))[0]
    l=[]
    with open(path, 'r') as file:
        content = file.read()
    obj = json.loads(content)

    NumberOfLines = 1

    for transaction in obj["fiscal_receipts"]:
        list = transaction["line_data"]
        pay_data = transaction["payment_data"]
        if transaction["customer"] is not None:
            client_id = transaction["customer"]["bci_id"]
        else:
            client_id = "Non renseigné"
        transaction_type = transaction["type"]
        raison_annulation = transaction["cancellation_reason"]

        dict = {}
        payment_method = pay_data[0]["payment_method_label"]
        payment_date = pay_data[0]["created"]
        for i in range(len(list)):
            nom_offre = ""
            nom_prestation = list[i]["name_line_1"]
            item_price = float(list[i]["item_price"])
            if ("- Accueil (5m)" in nom_prestation) or ("- accueil (5m)" in nom_prestation) or ("- Accuiel (5m)" in nom_prestation):
                split = nom_prestation.split("-")
                for e in split[0:-1]:
                    nom_offre += e
                if not nom_offre in dict:
                    dict[nom_offre] = [0, [split[-1]], item_price]
            elif not("-" in nom_prestation):
                l.append([nom_prestation, item_price, payment_method, payment_date, client_id, transaction_type, raison_annulation])

                NumberOfLines += 1

        for i in range(len(list)):
            nom_offre = ""
            nom_prestation = list[i]["name_line_1"]
            item_price = float(list[i]["item_price"])
            if "-" in nom_prestation:
                est_une_offre = False
                split = nom_prestation.split("-")
                nom_offre = split[0]
                for e in split[0:-1]:
                    if nom_offre in dict.keys():
                        est_une_offre = True
                        break
                    nom_offre += "-" + e
                if est_une_offre:
                    if (split[-1] == " Accueil (5m)") or (split[-1] == " Accuiel (5m)") or (split[-1] == " accueil (5m)"):
                        dict[nom_offre][0]+=1
                    elif not(split[-1] in dict[nom_offre][1]):
                        dict[nom_offre][1].append(split[-1])
                        dict[nom_offre][2]+= item_price
        for key in dict.keys():
            nom_prestation = key
            item_price = dict[key][2]
            for i in range(dict[key][0]):
                l.append([nom_prestation, item_price, payment_method, payment_date, client_id, transaction_type, raison_annulation])
                NumberOfLines+=1

    fullcsv = open("fullData.csv", "a+")
    csv = open(nom_output + ".csv", "w+")
    for x in l:
        s=x[0] + ";"
        s+= add_categ(x[0], categories) + ";"
        for i in range(1, len(x)):
            s+=str(x[i]).replace(',','')+";"
        fullcsv.write(s[:-1]+"\n")
        csv.write(s[:-1] + "\n")
    return nom_output

def groq(description,categories):
    # Le prompt système dicte les règles absolues à l'IA
    with open("grokapikey.txt","r") as f:
        contenu = f.readlines()
    os.environ["GROQ_API_KEY"] = contenu[0]
    client = Groq()
    prompt_systeme = f"""Tu es un expert en classification pour un centre de bien-être.
    Tu dois classer la prestation fournie par l'utilisateur dans EXACTEMENT UNE des catégories suivantes :
    {', '.join(categories.keys())}.
    
    Si tu as l'impression que le nom de la prestation est étrange il s'agit probablement d'une offre temporaire, 
    En particulier n'attribue la catégorie "Espace Détente (Spa)" que si le nom de la prestation contient le mot "détente"

    RÈGLE ABSOLUE : Tu dois répondre UNIQUEMENT par le nom exact de la catégorie. 
    Aucun autre mot, aucune salutation, aucune ponctuation supplémentaire.

"""

    try:
        # On interroge le modèle Llama 3 (le plus intelligent, version 70 milliards de paramètres)
        reponse = client.chat.completions.create(
            messages=[
                {"role": "system", "content": prompt_systeme},
                {"role": "user", "content": description}
            ],
            model="llama-3.1-8b-instant",
            temperature=1  # Température à 0 = l'IA ne fait pas preuve de créativité, elle est factuelle et prévisible
        )
        return reponse.choices[0].message.content.strip()

    except Exception as e:
        print(f"Erreur API : {e}")
        return "Erreur"

def add_categ (prest,categories):

    def aux(prest, categories):
        result=[]
        for i in range(10):
            gc=groq(prest,categories)
            while not(gc in categories.keys()):
                gc=groq(prest,categories)
                print('error')
                print(gc)
            print(gc)
            result.append(gc)
        nb=0
        indice=0
        for i,x in enumerate(result):
            if result.count(x)>=nb :
                indice=i
        categories[result[indice]].append(prest)
        print("_"*100+str(result[indice]))
        with open("SavedCateg.txt", "w+") as text:
            for x in categories:
                text.write(f"{x}$")
                for j in range(len(categories[x])-1):
                    text.write(f"{categories[x][j]}£")
                text.write(f"{categories[x][-1]}\n")
        return result[indice]

    IsCategorized = ''
    for categ in categories.keys():
        if prest in categories[categ]:
            IsCategorized = categ
            break

    if IsCategorized == '':
        return aux(prest, categories)
    else:
        return IsCategorized

def into_wb(nom):
    NumberOfLines = 0
    with open(nom + ".csv", "r+") as csv:
        lignes = [ligne.replace('\n',"").split(";") for ligne in csv.readlines()]
    wb = Workbook()
    ws = wb.active
    ws.append(["Nom","Catégorie", "Prix", "Moyen de paiement", "Date", "Id_client", "Type de transaction", "Commentaire"])
    for x in lignes:
        NumberOfLines+=1
        L = x[:2] + [float(x[2])] + x[3:]
        ws.append(L)
    ws.auto_filter.ref = f"A1:G{ws.max_row}"
    wb.save(nom + ".xlsx")
    return NumberOfLines

def importer_sur_le_drive(nom):
    creds = get_credentials()

    # Define File Metadata
    file_metadata = {
        'name': nom,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': ["1XiSdGCxVyTb5xMxWU5oLBmqO1biS9fBn"]
    }

    # Prepare the media payload
    media = MediaFileUpload(
        nom + ".xlsx",
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

def nvOnglet(id, nom):
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # On commence par créer un nouvel onglet pour le Pivot
    request_body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": nom
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

def creer_analyse(tabId, NumberOfLines):
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # On demande à l'API de nous donner les métadonnées des feuilles du fichier
    spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=tabId).execute()
    sheets = spreadsheet_metadata.get('sheets', [])

    # on récupère l'id de la première feuille
    source_sheet_id = sheets[0]['properties']['sheetId']

    sheetId1 = nvOnglet(tabId, "Paramètres")
    sheetId2 = nvOnglet(tabId, "Total de gains par prestations")
    sheetId3 = nvOnglet(tabId, "Nombre de prestations par mots-clés")
    sheetId4 = nvOnglet(tabId, "Total de gains par catégorie")

    creer_parametres(tabId, sheetId1)
    tableau_dyn_total_gains_prestations(creds, tabId, source_sheet_id, sheetId2, NumberOfLines)
    tableau_dyn_nombre_prestation_mots_cles(creds, tabId, source_sheet_id, sheetId3, NumberOfLines)
    tableau_dyn_total_gains_categorie(creds, tabId, source_sheet_id, sheetId4, NumberOfLines)

def creer_parametres(id, sheetId):
    creds = get_credentials()

    ligne1 = ["Recherche par mots-clés", "", "", "", "", "", "", "", "", "", ""]
    ligne2 = ["Mots-clés", "Vide", "Vide", "Vide", "Vide", "Vide", "Vide", "Vide", "Vide", "Vide", "Vide"]

    body = {'values': [ligne1, ligne2]}

    try:
        service = build('sheets', 'v4', credentials=creds)

        result = service.spreadsheets().values().update(
            spreadsheetId=id,
            range = "Paramètres!B1:L2",
            valueInputOption="RAW",
            body=body
        ).execute()

    except Exception as e:
        print(f"Une erreur est survenue : {e}")

def tableau_dyn_total_gains_prestations(creds, tabId,source_sheetId, sheetId, NumberOfLines):
    service = build('sheets', 'v4', credentials=creds)

    body = {
        "requests": [
            {
                "updateCells": {
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 0,
                        "columnIndex": 0
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "pivotTable": {
                                        "source": {
                                            "sheetId": source_sheetId,
                                            "startRowIndex": 0,
                                            "endRowIndex": NumberOfLines,
                                            "startColumnIndex": 0,
                                            "endColumnIndex": 8
                                        },
                                        # ROW LAYOUT: Group by "Nom"
                                        "rows": [
                                            {
                                                "sourceColumnOffset": 0,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # COLUMN LAYOUT: Group by "Transaction"
                                        "columns": [
                                            {
                                                "sourceColumnOffset": 6,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # VALUES/AGGREGATION: Sum of "prix"
                                        "values": [
                                            {
                                                "summarizeFunction": "SUM",
                                                "sourceColumnOffset": 2,
                                                "name": "Ventes totales"
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

    service.spreadsheets().batchUpdate(
        spreadsheetId=tabId,
        body=body
    ).execute()

    print("Tableau croisé dynamique total gains prestations créer")

def tableau_dyn_nombre_prestation_mots_cles(creds, tabId,source_sheetId, sheetId, NumberOfLines):

    service = build('sheets', 'v4', credentials=creds)

    body = {
        "requests": [
            {
                "updateCells": {
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 0,
                        "columnIndex": 0
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "pivotTable": {
                                        "source": {
                                            "sheetId": source_sheetId,
                                            "startRowIndex": 0,
                                            "endRowIndex": NumberOfLines,
                                            "startColumnIndex": 0,
                                            "endColumnIndex": 8
                                        },
                                        # ROW LAYOUT: Group by "Type de prestation"
                                        "rows": [
                                            {
                                                "sourceColumnOffset": 8,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # COLUMN LAYOUT: Group by "moyen de paiement"
                                        "columns": [
                                            {
                                                "sourceColumnOffset": 3,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # VALUES/AGGREGATION: Count of "Type de prestation"
                                        "values": [
                                            {
                                                "summarizeFunction": "COUNT",
                                                "sourceColumnOffset": 8,
                                                "name": "Nombre de prestations par mots-clés"
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

    service.spreadsheets().batchUpdate(
        spreadsheetId=tabId,
        body=body
    ).execute()

    print("Tableau croisé dynamique nombre prestations mots-clés créer")

def tableau_dyn_total_gains_categorie(creds, tabId, source_sheetId, sheetId, NumberOfLines):

    service = build('sheets', 'v4', credentials=creds)

    body = {
        "requests": [
            {
                "updateCells": {
                    "start": {
                        "sheetId": sheetId,
                        "rowIndex": 0,
                        "columnIndex": 0
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "pivotTable": {
                                        "source": {
                                            "sheetId": source_sheetId,
                                            "startRowIndex": 0,
                                            "endRowIndex": NumberOfLines,
                                            "startColumnIndex": 0,
                                            "endColumnIndex": 8
                                        },
                                        # ROW LAYOUT: Group by "catégorie"
                                        "rows": [
                                            {
                                                "sourceColumnOffset": 1,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # COLUMN LAYOUT: Group by "moyen de paiement"
                                        "columns": [
                                            {
                                                "sourceColumnOffset": 3,
                                                "sortOrder": "ASCENDING",
                                                "showTotals": True
                                            }
                                        ],
                                        # VALUES/AGGREGATION: Sum of "price"
                                        "values": [
                                            {
                                                "summarizeFunction": "SUM",
                                                "sourceColumnOffset": 2,
                                                "name": "total gains par catégorie"
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

    service.spreadsheets().batchUpdate(
        spreadsheetId=tabId,
        body=body
    ).execute()

    print("Tableau croisé dynamique total gains par catégorie créer")


if __name__ == "__main__" :

# On récupère les prestations déjà associés à des catégories
    categories={}
    with open('SavedCateg.txt', 'r+') as file:
        fileContent = [line.replace('\n',"") for line in file.readlines()]
        for line in fileContent:
            categories[line.split('$')[0]]=[]
            for prest in line.split('$')[1].split("£"):
                categories[line.split('$')[0]].append(prest)

    root = tk.Tk()
    root.withdraw()
    paths = filedialog.askopenfilenames(filetypes=[("JSON", "*.json")])

    for path in paths:
        #On crée le tableau associé aux données du json en entrée
        nom = creer_donnee_brut(path,categories)
        NumberOfLines = into_wb(nom)
        id1 = importer_sur_le_drive(nom)
        creer_analyse(id1, NumberOfLines)

        os.remove(nom+".xlsx")
        os.remove(nom+".csv")

        #On crée la data complète
        into_wb("fullData")
        id2 = importer_sur_le_drive("fullData")
        creer_analyse(id2, NumberOfLines)
        print("importation de " + nom + " terminée")

        os.remove("fullData.xlsx")
