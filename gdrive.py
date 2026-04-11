import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "google_service_account.json"

def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=GDRIVE_SCOPES,
    )
    return build("drive", "v3", credentials=creds)


def upload_to_gdrive(local_path: str, folder_id: str):
    service = get_drive_service()

    file_metadata = {
        "name": os.path.basename(local_path),
        "parents": [folder_id],
    }

    media = MediaFileUpload(local_path, resumable=True)

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True,
    ).execute()

    return file.get("id")
