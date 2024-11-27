from googleapiclient.discovery import build, Resource
from google.oauth2 import service_account
from google.cloud import storage
import json
import os

# environment variables
scopes = (
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
)


class Config:
    def __init__(self, credentials_path: str = None):
        if "GOOGLE_DOC_CREDENTIALS" not in os.environ:
            return

        credentials_path = os.environ.get("GOOGLE_DOC_CREDENTIALS", credentials_path)
        # if file does not exist, return
        if not os.path.exists(credentials_path):
            return

        credentials = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=scopes
        )
        self.DRIVE = build("drive", "v3", credentials=credentials)
        self.SLIDES = build("slides", "v1", credentials=credentials)
        self.SHEETS = build("sheets", "v4", credentials=credentials)
        self.DOCS = build("docs", "v1", credentials=credentials)
        self.BUCKET = os.environ.get("OODLES_BUCKET", "gs://data-studies/img")
        self.STORAGE_CLIENT = storage.Client.from_service_account_json(credentials_path)

        with open(credentials_path, "r") as f:
            service_email = json.load(f)["client_email"]

        self.service_email = service_email

    def init(
        self,
        DRIVE: Resource,
        SLIDES: Resource,
        SHEETS: Resource,
        DOCS: Resource,
        BUCKET: str,
        STORAGE_CLIENT,
        service_email: str,
    ):
        self.DRIVE = DRIVE
        self.SLIDES = SLIDES
        self.SHEETS = SHEETS
        self.DOCS = DOCS
        self.BUCKET = BUCKET
        self.STORAGE_CLIENT = STORAGE_CLIENT
        self.service_email = service_email


config = Config()


def init(credentials_path: str):
    credentials = service_account.Credentials.from_service_account_file(
        credentials_path, scopes=scopes
    )
    DRIVE = build("drive", "v3", credentials=credentials)
    SLIDES = build("slides", "v1", credentials=credentials)
    SHEETS = build("sheets", "v4", credentials=credentials)
    DOCS = build("docs", "v1", credentials=credentials)
    BUCKET = os.environ.get("OODLES_BUCKET", "gs://data-studies/img")
    STORAGE_CLIENT = storage.Client.from_service_account_json(credentials_path)

    with open(credentials_path, "r") as f:
        service_email = json.load(f)["client_email"]

    config.init(DRIVE, SLIDES, SHEETS, DOCS, BUCKET, STORAGE_CLIENT, service_email)
