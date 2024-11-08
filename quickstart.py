import os.path
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", "https://www.googleapis.com/auth/calendar.readonly"]


def main():
  """Uses the Gmail API to read and search for third-party automated maintenance emails and then creates Google Calendar events using the Google Calendar API.
  """
  
  # AUTHORIZATION
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # Call the Gmail API
    service_gmail = build("gmail", "v1", credentials=creds)
    results = service_gmail.users().labels().list(userId="me").execute()
    labels = results.get("labels", [])
    
     # Perform search query
    sentFrom = ""
    label = "sent"
    subject = ""
    body = "Dear"
  
    f = f'from:"{sentFrom}"' if sentFrom else ''
    l = f'label:"{label}"' if label else ''
    s = f'subject:"{subject}"' if subject else ''
    searchQuery = f'{f} {l} {s} {body}'
    print("\n\nYour search query: " + searchQuery)
    messages = service_gmail.users().messages().list(userId="me", q=searchQuery, maxResults=500).execute()
    
    if 'messages' not in messages or (messages['resultSizeEstimate'] <= 0):
      print("No messages found")
      return 
    print("-----------------------------------------------------------------------------------------------------------------------------\n")
    print(f"Quantity: {messages['resultSizeEstimate']} \t\t\t\t\t Max: 500")
    for message in messages['messages']:
      print(message)
    print("\n-----------------------------------------------------------------------------------------------------------------------------\n\n")
    
    
    # For each correct email create an event
    # Go through events and merge if needed (events have unique id's)
    

    # if not labels:
    #   print("No labels found.")
    #   return
    # print("Labels:")
    # for label in labels:
    #   print(label["name"])

    
    # Call the Calendar API
    service_gcal = build("calendar", "v3", credentials=creds)
    now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
    print("Getting the upcoming 10 events")
    events_result = (
        service_gcal.events()
        .list(
            calendarId="primary",
            timeMin=now,
            maxResults=10,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])

    # if not events:
    #   print("No upcoming events found.")
    #   return

    # Prints the start and name of the next 10 events
    # for event in events:
    #   start = event["start"].get("dateTime", event["start"].get("date"))
    #   print(start, event["summary"])

  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()
