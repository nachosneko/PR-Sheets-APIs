import re
import discord
from discord.ext import commands
from discord import app_commands
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly",
          "https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

# Default values.
DEFAULT_PR_NAME = "Artist"
DEFAULT_SPREADSHEET_ID = "1x_3eFPDudXxkYWTvsuWZJKU5rSdMQEUbE40fT4lmJCU"
PR_NAME = DEFAULT_PR_NAME
SAMPLE_SPREADSHEET_ID = DEFAULT_SPREADSHEET_ID
SAMPLE_RANGE_NAME = "Sheet1!A2:E"
admin_user_ids = 

intents = discord.Intents.all()
intents.reactions = True
intents.messages = True

bot = commands.Bot(command_prefix='¤', intents=intents)

def create_spreadsheet_copy(drive_service, spreadsheet_id):
    try:
        # Copy the spreadsheet using Google Drive API
        drive_response = drive_service.files().copy(fileId=spreadsheet_id).execute()
        new_spreadsheet_id = drive_response['id']
        return new_spreadsheet_id
    except HttpError as err:
        print(f"Error creating copy: {err}")
        return None

@bot.event
async def on_ready():
    print(f'Logged in as {bot.user.name}')

@bot.tree.command(name="copy_and_set")
@app_commands.describe(
    new_pr_name="input pr name",
    new_spreadsheet_id="input spreadsheet id (after /d/ in links)",
    )
async def copy_and_set(interaction: discord.Interaction, new_pr_name: str, new_spreadsheet_id: str):
    global PR_NAME, SAMPLE_SPREADSHEET_ID
    PR_NAME = new_pr_name
    SAMPLE_SPREADSHEET_ID = new_spreadsheet_id

    creds = None

    if interaction.user.id not in admin_user_ids:
        await interaction.response.send_message(f"only admins can change the clips")
    else:
        await change_female(new_female_clip, new_correct_female)

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Build Google Sheets and Google Drive services
        sheets_service = build("sheets", "v4", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)

        # Call the Sheets API
        sheet = sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
        values = result.get("values", [])

        if not values:
            await interaction.send("No data found.")
            return

        # Create a copy of the spreadsheet
        new_spreadsheet_id = create_spreadsheet_copy(drive_service, SAMPLE_SPREADSHEET_ID)
        if new_spreadsheet_id:
            await interaction.send(f"New Spreadsheet: [{PR_NAME} PR](https://docs.google.com/spreadsheets/d/{new_spreadsheet_id})")

    except HttpError as err:
        await interaction.send(f"Error accessing spreadsheet: {err}")

@bot.command(name='set_spreadsheet_id')
async def set_spreadsheet_id(ctx, new_spreadsheet_id):
    global SAMPLE_SPREADSHEET_ID
    SAMPLE_SPREADSHEET_ID = new_spreadsheet_id
    await ctx.send(f"Spreadsheet ID set to: {new_spreadsheet_id}")

@bot.command(name='set_pr_name')
async def set_pr_name(ctx, new_pr_name):
    global PR_NAME
    PR_NAME = new_pr_name
    await ctx.send(f"PR Name set to: {new_pr_name}")

bot.run('MTE2OTAxMzc5Mjg0MDA5Nzk1Mw.G5Azm9.qXN_I3czPqLw788TFZBqoU2olTGhNdhCITuhho')
