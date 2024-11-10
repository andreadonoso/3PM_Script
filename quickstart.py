import os.path
import datetime
import base64
import email as email_parser
from bs4 import BeautifulSoup

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", "https://www.googleapis.com/auth/calendar.readonly"]


def main():
  """Uses the Gmail API to read and search for third-party automated maintenance emails and then 
     creates Google Calendar events using the Google Calendar API based on the email information.
  """
  
  # AUTHORIZATION
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  
  creds = None
  
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
     # PERFORM SEARCH QUERY
    sentFrom = ""
    label = ""
    subject = ""
    body = ""
  
    f = f'from:"{sentFrom}"' if sentFrom else ''
    l = f'label:"{label}"' if label else ''
    s = f'subject:"{subject}"' if subject else ''
    searchQuery = f'{f} {l} {s} {body}'
    
    # Call the Gmail API
    service_gmail = build("gmail", "v1", credentials=creds)
    queryRes = service_gmail.users().messages().list(userId="me", q=searchQuery, maxResults=100).execute()
    
    # Print search query results
    if 'messages' not in queryRes or (queryRes['resultSizeEstimate'] <= 0):
      print("No emails found")
      return 
    
    print("\n\n-----------------------------------------------------------------------------------------------------------------------------\n\n")
    print(f"Query:\t {searchQuery.strip()}")
    print(f"Results: {queryRes['resultSizeEstimate']}\n")
    
    messages = queryRes['messages']
    for count,message in enumerate(messages, start=1):
      print(f"\n{count})- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\n")
      fullMsg = service_gmail.users().messages().get(userId="me", id=message['id'], format="full").execute()
      headers = fullMsg['payload']['headers']
      resSubject = next((header['value'] for header in headers if header['name'] == 'Subject'), "No Subject")
      print(f"SUBJECT: {resSubject}\n")

      # Filter for visible text
      # Check if the email has multiple parts (i.e., plain text and HTML)
      visible_text = None
      if 'parts' in fullMsg['payload']:
          for part in fullMsg['payload']['parts']:
              if part['mimeType'] == 'text/plain':
                  print("TYPE:\t MP text/plain")
                  # Plain text content - decode to UTF-8
                  body_data = part['body']['data']
                  visible_text = base64.urlsafe_b64decode(body_data).decode('utf-8')
                  break
              elif part['mimeType'] == 'text/html':
                  print("TYPE:\t MP text/html")
                  # HTML content - decode and parse with BeautifulSoup
                  body_data = part['body']['data']
                  html_content = base64.urlsafe_b64decode(body_data).decode('utf-8')
                  soup = BeautifulSoup(html_content, 'html.parser')
                  visible_text = soup.get_text(separator="\n", strip=True)
                  break
      else:
          # If the email does not have multiple parts, check if it's HTML or plain text
          if fullMsg['payload']['mimeType'] == 'text/plain':
              print("TYPE:\t text/plain")
              visible_text = base64.urlsafe_b64decode(fullMsg['payload']['body']['data']).decode('utf-8')
          elif fullMsg['payload']['mimeType'] == 'text/html':
              print("TYPE:\t text/html")
              html_content = base64.urlsafe_b64decode(fullMsg['payload']['body']['data']).decode('utf-8')
              soup = BeautifulSoup(html_content, 'html.parser')
              visible_text = soup.get_text(separator="\n", strip=True)

      # Print the extracted visible text
      if visible_text:
          print("\nBODY:\n")
          print(f"{visible_text}")
      else:
          print("No visible text found.")
          
    print("\n-----------------------------------------------------------------------------------------------------------------------------\n\n")
    
    # Determine correct dates from subjecta & that we don't have duplicate maintenance id's/(date-time && place)
    # Create a Google calendar event based on the email's body
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
