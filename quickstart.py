import os.path
import numpy as np
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

from paginator import EmbedPaginatorSession

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
        sheet_id = drive_response['id']

        # Copy permissions from the original spreadsheet to the new one
        copy_permissions(drive_service, spreadsheet_id, sheet_id)

        return sheet_id
    except HttpError as err:
        print(f"error creating copy: {err}")
        return None

def copy_permissions(service, original_spreadsheet_id, sheet_id):
    try:
        # Get the permissions from the original spreadsheet
        permissions = service.permissions().list(fileId=original_spreadsheet_id).execute()

        # Copy each permission to the new spreadsheet
        for permission in permissions.get('permissions', []):
            # Exclude non-writable fields from the request body
            permission_body = {key: permission[key] for key in permission.keys() if key in ['role', 'type']}
            
            # Add emailAddress field for 'user' or 'group' type
            if permission.get('type') in ['user', 'group']:
                permission_body['emailAddress'] = permission.get('emailAddress', 'placeholder@example.com')

            # Exclude owner permissions
            if permission.get('role') != 'owner':
                service.permissions().create(
                    fileId=sheet_id,
                    body=permission_body
                ).execute()

        print("Permissions copied successfully.")
    except Exception as e:
        print(f"Error copying permissions: {e}")

async def set_spreadsheet_id(sheet_id):
    global SAMPLE_SPREADSHEET_ID
    SAMPLE_SPREADSHEET_ID = sheet_id

async def set_pr_name(pr_name):
    global PR_NAME
    PR_NAME = pr_name

@bot.event
async def on_ready():
    print(f'logged in as {bot.user.name}')
    try:
        synced = await bot.tree.sync()
        print(f"synced {len(synced)} command(s)")
    except Exception as e:
        print(e)

@bot.tree.command(name="copy_and_send")
@app_commands.describe(
    pr_name="input sheet(s) name",
    sheet_id="input spreadsheet id",
    send_to=" mention user(s) to DM the sheet(s) to"
)
async def copy_and_send(
    interaction: discord.Interaction, 
    pr_name: str, 
    sheet_id: str,
    send_to: typing.Optional[str] = None  # Use string for mentions
):
    global PR_NAME, SAMPLE_SPREADSHEET_ID
    PR_NAME = pr_name
    SAMPLE_SPREADSHEET_ID = sheet_id

    # Acknowledge the interaction immediately
    await interaction.response.send_message(content="processing...")

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
            mentions = [mention for mention in send_to.split() if mention.startswith("<@")]
            
            embed = discord.Embed(
                title=f"completed sheet(s) copy for {PR_NAME}",
                description=f"original spreadsheet: [link](https://docs.google.com/spreadsheets/d/{SAMPLE_SPREADSHEET_ID})",
                color=0x00ff00,
            )
            # Create a copy of the spreadsheet for each mentioned user
            for mention in mentions:
                user_id = int(mention.strip("<@!>"))
                member = bot.get_user(user_id)
                if not member:
                    await interaction.followup.send(content=f"error: user not found for mention {mention}.")
                    continue

                # Create a copy of the spreadsheet (time-consuming operation)
                user_sheet_id = await bot.loop.run_in_executor(None, create_spreadsheet_copy, drive_service, SAMPLE_SPREADSHEET_ID, f"{PR_NAME} - {member.name}")

                if user_sheet_id:
                    embed.add_field(
                        name=f"{member.display_name}'s copy",
                        value=f"[{PR_NAME} party rank](https://docs.google.com/spreadsheets/d/{user_sheet_id})\n```{user_sheet_id}```",
                        inline=False
                    )
                    
                    # Send a DM to the mentioned user
                    dm_message = f"hello!! here is your {PR_NAME} sheet\nhttps://docs.google.com/spreadsheets/d/{user_sheet_id}"
                    await member.send(content=dm_message)
                else:
                    await interaction.followup.send(content=f"error occurred during processing for {member.mention}.")

            await interaction.followup.send(embed=embed)
    except HttpError as err:
        await interaction.followup.send(content=f"error accessing spreadsheet: {err}")

bot.run('BOT_TOKEN')
