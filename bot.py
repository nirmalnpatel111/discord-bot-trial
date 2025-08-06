from dotenv import load_dotenv
import discord
import asyncio
from datetime import datetime, timedelta, timezone
import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import webserver

load_dotenv()

intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
client = discord.Client(intents=intents)

SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/spreadsheets'
]

# --- CONFIG ---
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")
LOCATION_TAGS = {
    '-ieee': 'IEEE',
    '-mcgill': 'McGill',
    '-home': 'Home',
    '-conco': 'Concordia'
}
# ---------------

user_sessions = {}  # user_id: {event_id, task, start_time, location}

def get_google_service(api_name, version):
    creds = None
    if os.path.exists('token.pkl'):
        with open('token.pkl', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.pkl', 'wb') as token:
            pickle.dump(creds, token)
    return build(api_name, version, credentials=creds)

@client.event
async def on_ready():
    print(f'✅ Logged in as {client.user}')

@client.event
async def on_message(message):
    if message.author == client.user:
        return

    user_id = message.author.id
    username = message.author.name
    msg = message.content.strip().lower()
    tokens = msg.split()

    if tokens and tokens[0] == "start":
        if user_id in user_sessions:
            await message.channel.send("❗ You already have an active session.")
            return

        # Check for location
        location = None
        for token in tokens[1:]:
            if token in LOCATION_TAGS:
                location = LOCATION_TAGS[token]

        if not location:
            await message.channel.send("⚠️ Please specify a location using one of the following: `-ieee`, `-mcgill`, `-home`, `-conco`\nExample: `start -ieee`")
            return

        start_time = datetime.now(timezone.utc)
        end_time = start_time + timedelta(minutes=15)
        summary = f"{username} is working @{location}"

        calendar_service = get_google_service('calendar', 'v3')
        event = {
            'summary': summary,
            'start': {'dateTime': start_time.isoformat(), 'timeZone': 'UTC'},
            'end': {'dateTime': end_time.isoformat(), 'timeZone': 'UTC'},
        }
        created_event = calendar_service.events().insert(calendarId='primary', body=event).execute()
        event_id = created_event['id']

        task = asyncio.create_task(update_event_periodically(user_id, username, event_id, location))
        user_sessions[user_id] = {
            'event_id': event_id,
            'task': task,
            'start_time': start_time,
            'location': location
        }

        await message.channel.send(f"📅 Started: `{summary}`\n⏱ Dynamic 15-min updates enabled.\n🕛 Auto-stop at 11:59 PM (UTC).")

    elif msg == "stop":
        if user_id not in user_sessions:
            await message.channel.send("⚠️ No active session to stop.")
            return

        session = user_sessions.pop(user_id)
        session['task'].cancel()

        end_time = datetime.now(timezone.utc)
        calendar_service = get_google_service('calendar', 'v3')
        calendar_service.events().patch(
            calendarId='primary',
            eventId=session['event_id'],
            body={'end': {'dateTime': end_time.isoformat(), 'timeZone': 'UTC'}}
        ).execute()

        # Log to Google Sheet
        log_to_sheet(
            name=username,
            start=session['start_time'],
            end=end_time,
            location=session['location']
        )

        await message.channel.send("✅ Session stopped and calendar + sheet updated.")

async def update_event_periodically(user_id, username, event_id, location):
    try:
        calendar_service = get_google_service('calendar', 'v3')
        while True:
            await asyncio.sleep(15 * 60)
            now = datetime.now(timezone.utc)

            if now.hour == 23 and now.minute >= 59:
                break

            calendar_service.events().patch(
                calendarId='primary',
                eventId=event_id,
                body={'end': {'dateTime': now.isoformat(), 'timeZone': 'UTC'}}
            ).execute()
            print(f"⏱ Updated {username}'s event end time to {now}")

        print(f"🛑 Auto-stopped {username}'s session at midnight.")
        if user_id in user_sessions:
            user_sessions.pop(user_id)

    except asyncio.CancelledError:
        print(f"🛑 Cancelled task for {username}")
        return

def log_to_sheet(name, start, end, location):
    sheet_service = get_google_service('sheets', 'v4')

    # Use ISO-like string formatting for clarity
    values = [[
        start.strftime('%Y-%m-%d'),               # Date
        name,                                     # Name
        start.strftime('%Y-%m-%d %H:%M:%S'),      # Worked From
        end.strftime('%Y-%m-%d %H:%M:%S'),        # Worked Till
        location                                  # Place
    ]]

    body = {
        'values': values
    }

    result = sheet_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:E",
        valueInputOption="USER_ENTERED",  # Still allows formulas
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

    print(f"📄 Logged to sheet: {result.get('updates').get('updatedRange')}")


webserver.keep_alive()
client.run(os.getenv("DISCORD_BOT_TOKEN"))