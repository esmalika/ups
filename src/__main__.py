from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pygsheets as pygsheet
import time
import os
import pickle
import numpy as np
from google.auth.transport.requests import Request

# from dcms import DCMS
from pseudo_dcmsius import DCMS
from src.generator import Generator


nowTime = time.time()
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]
creds = None
DISPERSE_CLIENT_SECRET_PATH = os.getenv("DISPERSE_CLIENT_SECRET_PATH")
requestList = []
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.

if os.path.exists(DISPERSE_CLIENT_SECRET_PATH + "/token.pickle"):
    with open(DISPERSE_CLIENT_SECRET_PATH + "/token.pickle", "rb") as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            DISPERSE_CLIENT_SECRET_PATH + "/client_secret.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(DISPERSE_CLIENT_SECRET_PATH + "/token.pickle", "wb") as token:
        pickle.dump(creds, token)

serviceSheet = build("sheets", "v4", credentials=creds)
authorizing = pygsheet.authorize(
    client_secret=DISPERSE_CLIENT_SECRET_PATH + "/client_secret.json",
    credentials_directory=DISPERSE_CLIENT_SECRET_PATH,
    local=True,
)
gene = Generator(DCMS, os, authorizing, requestList, serviceSheet)
