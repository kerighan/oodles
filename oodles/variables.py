from googleapiclient.discovery import build
from google.oauth2 import service_account
from httplib2 import Http
import json
import os

# environment variables
scopes = ("https://www.googleapis.com/auth/presentations",
          "https://www.googleapis.com/auth/drive",
          "https://www.googleapis.com/auth/spreadsheets")
GOOGLE_APPLICATION_CREDENTIALS = os.environ["GOOGLE_APPLICATION_CREDENTIALS"]
credentials = service_account.Credentials.from_service_account_file(GOOGLE_APPLICATION_CREDENTIALS, scopes=scopes)
DRIVE = build("drive", "v3", credentials=credentials)
SLIDES = build("slides", "v1", credentials=credentials)
SHEETS = build("sheets", "v4", credentials=credentials)
BUCKET = os.environ.get("OODLES_BUCKET", "gs://data-studies/img")


with open(GOOGLE_APPLICATION_CREDENTIALS, "r") as f:
    service_email = json.load(f)["client_email"]
