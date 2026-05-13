import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

# Credenciales desde el secreto de GitHub
creds_json = os.environ["GDRIVE_CREDENTIALS"]
folder_id = os.environ["GDRIVE_FOLDER_ID"]

creds_dict = json.loads(creds_json)
creds = service_account.Credentials.from_service_account_info(
    creds_dict,
    scopes=["https://www.googleapis.com/auth/drive.readonly"]
)

service = build("drive", "v3", credentials=creds)

# Buscar el Excel más reciente en la carpeta
results = service.files().list(
    q=f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false",
    orderBy="modifiedTime desc",
    pageSize=1,
    fields="files(id, name)"
).execute()

files = results.get("files", [])
if not files:
    print("ERROR: No hay archivo Excel en la carpeta de Drive")
    exit(1)

file = files[0]
print(f"Descargando: {file['name']}")

request = service.files().get_media(fileId=file["id"])
fh = io.FileIO(f"archivos/{file['name']}", "wb")
downloader = MediaIoBaseDownload(fh, request)
done = False
while not done:
    _, done = downloader.next_chunk()

print(f"Descargado exitosamente: {file['name']}")
