import os.path

import typing
from typing import Optional

import discord
from discord.ext import commands
from discord import app_commands

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
DEFAULT_SPREADSHEET_ID = "SpreadsheetID"
PR_NAME = DEFAULT_PR_NAME
SAMPLE_SPREADSHEET_ID = DEFAULT_SPREADSHEET_ID
SAMPLE_RANGE_NAME = "Sheet1!A2:E"
admin_user_ids = [
    195534572581158913,
] 

intents = discord.Intents.all()
intents.reactions = True
intents.messages = True

bot = commands.Bot(command_prefix='¤', intents=intents)

def create_spreadsheet_copy(drive_service, spreadsheet_id, new_title):
    try:
        # Copy the spreadsheet using Google Drive API
        drive_response = drive_service.files().copy(fileId=spreadsheet_id, body={'name': new_title}).execute()
        new_spreadsheet_id = drive_response['id']
        return new_spreadsheet_id
    except HttpError as err:
        print(f"error creating copy: {err}")
        return None

async def set_spreadsheet_id(interaction, new_spreadsheet_id):
    global SAMPLE_SPREADSHEET_ID
    SAMPLE_SPREADSHEET_ID = new_spreadsheet_id

async def set_pr_name(interaction, new_pr_name):
    global PR_NAME
    PR_NAME = new_pr_name

@bot.event
async def on_ready():
    print(f'logged in as {bot.user.name}')
    try:
        synced = await bot.tree.sync()
        print(f"synced {len(synced)} command(s)")
    except Exception as e:
        print(e)

@bot.tree.command(name="copy_and_set")
@app_commands.describe(
    new_pr_name="input pr name",
    new_spreadsheet_id="input spreadsheet id (after /d/ in links)",
    ping_users="mention users in DMs (optional)"
)
async def copy_and_set(
    interaction: discord.Interaction, 
    new_pr_name: str, 
    new_spreadsheet_id: str,
    ping_users: typing.Optional[str] = None  # Use string for mentions
):
    global PR_NAME, SAMPLE_SPREADSHEET_ID
    PR_NAME = new_pr_name
    SAMPLE_SPREADSHEET_ID = new_spreadsheet_id

    # Acknowledge the interaction immediately
    await interaction.response.send_message(content="processing...", ephemeral=True)

    creds = None

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
            await interaction.followup.send(content="no data found.")
            return
        
        if interaction.user.id not in admin_user_ids:
            await interaction.followup.send(content="no authorization.")
        else:
            # Parse mentions from the command string
            mentions = [mention for mention in ping_users.split() if mention.startswith("<@")]
            
            # Create a copy of the spreadsheet for each mentioned user
            for mention in mentions:
                user_id = int(mention.strip("<@!>"))
                member = bot.get_user(user_id)
                if not member:
                    await interaction.followup.send(content=f"error: user not found for mention {mention}.")
                    continue

                # Create a copy of the spreadsheet (time-consuming operation)
                new_spreadsheet_id = await bot.loop.run_in_executor(None, create_spreadsheet_copy, drive_service, SAMPLE_SPREADSHEET_ID, f"{PR_NAME} - {member.name}")

                if new_spreadsheet_id:
                    message = f"sent a copy for {member.mention}: ([{PR_NAME} Party Rank](https://docs.google.com/spreadsheets/d/{new_spreadsheet_id}))"
                    await interaction.followup.send(content=message)

                    # Send a DM to the mentioned user
                    dm_message = f"hello!! here is your {PR_NAME} sheet: [Link](https://docs.google.com/spreadsheets/d/{new_spreadsheet_id})"
                    await member.send(content=dm_message)
                else:
                    await interaction.followup.send(content=f"error occurred during processing for {member.mention}.")

    except HttpError as err:
        await interaction.followup.send(content=f"error accessing spreadsheet: {err}")


bot.run('BOT_TOKEN')
