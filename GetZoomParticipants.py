import os
import sys
import argparse
import requests
import csv
import base64
import json
from datetime import datetime
import pytz


ZOOM_TOKEN_URL = "https://zoom.us/oauth/token"
ZOOM_PARTICIPANTS_URL = "https://api.zoom.us/v2/metrics/meetings/{}/participants"

def get_access_token(client_id, client_secret, account_id):
    headers = {
        "Authorization": "Basic " + base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    }
    data = {
        "grant_type": "account_credentials",
        "account_id": account_id
    }
    response = requests.post(ZOOM_TOKEN_URL, headers=headers, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

def get_all_participants(meeting_id, access_token):
    participants = []
    seen = set()
    page_size = 150
    next_page_token = None
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    while True:
        params = {
            "type": "live",
            "page_size": page_size,
        }
        if next_page_token:
            params["next_page_token"] = next_page_token

        url = ZOOM_PARTICIPANTS_URL.format(meeting_id)
        response = requests.get(url, headers=headers, params=params)

        if response.status_code != 200:
            handle_http_error(response)

        data = response.json()

        for p in data.get("participants", []):
            # Create a unique key based on fields that identify a participant
            key = (p.get("participant_user_id"), p.get("email"), p.get("user_name"))
            if key not in seen:
                seen.add(key)
                participants.append(p)

        next_page_token = data.get("next_page_token")

        if not next_page_token:
            break

    return participants

def iso_to_local(iso_str):
    if not iso_str:  # Handle blank values (e.g. leave_time still empty)
        return ""
    dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))  # parse ISO8601
    local_tz = datetime.now().astimezone().tzinfo  # system local timezone
    return dt.astimezone(local_tz).strftime("%Y-%m-%d %H:%M:%S")

def write_csv(participants, output_path):
    fields = ["status", "join_time", "leave_time", "user_name", "email", 
              "participant_user_id", "pc_name", "client", "browser_name", "device_name"]
    with open(output_path, "w", newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fields)
        writer.writeheader()
        for p in participants:
            row = {k: p.get(k, "") for k in fields}
            row["join_time"] = iso_to_local(row["join_time"])
            row["leave_time"] = iso_to_local(row["leave_time"])
            writer.writerow(row)

def handle_http_error(response):
    code = response.status_code
    msg = response.text
    if code in [400, 401, 403, 404, 429]:
        print(f"Error {code}: {msg}")
        if code == 429:
            print("Rate limit exceeded. Consider retrying later.")
        sys.exit(1)
    else:
        response.raise_for_status()

def main():
    parser = argparse.ArgumentParser(
        description="""
Fetch Zoom meeting participants using Server-to-Server OAuth.

Required Environment Variables:
  ZOOM_CLIENT_ID     Your Zoom OAuth app's client ID
  ZOOM_CLIENT_SECRET Your Zoom OAuth app's client secret
  ZOOM_ACCOUNT_ID    Your Zoom account ID

Example:
  python GetZoomParticipants.py 123456789 output.csv --overwrite

Use --overwrite to allow replacing output.csv if it exists.
        """,
        formatter_class=argparse.RawTextHelpFormatter
    )

    parser.add_argument("meeting_id", help="Zoom Meeting ID (must be a live meeting)")
    parser.add_argument("output_file", help="Path to write the participant list (CSV)")
    parser.add_argument("--overwrite", action="store_true", help="Allow overwriting the output file")

    args = parser.parse_args()

    client_id = os.getenv("ZOOM_CLIENT_ID")
    client_secret = os.getenv("ZOOM_CLIENT_SECRET")
    account_id = os.getenv("ZOOM_ACCOUNT_ID")

    if not all([client_id, client_secret, account_id]):
        print("Error: Missing one or more required environment variables:\n"
              "  ZOOM_CLIENT_ID\n  ZOOM_CLIENT_SECRET\n  ZOOM_ACCOUNT_ID")
        sys.exit(1)

    if os.path.exists(args.output_file) and not args.overwrite:
        print(f"Error: Output file '{args.output_file}' already exists.\n"
              "Use --overwrite if you want to replace it.")
        sys.exit(1)

    try:
        access_token = get_access_token(client_id, client_secret, account_id)
        participants = get_all_participants(args.meeting_id, access_token)

        print(f"Fetched {len(participants)} participants from meeting {args.meeting_id}.")
        write_csv(participants, args.output_file)
        print(f"Participants written to: {args.output_file}")

    except requests.exceptions.HTTPError as e:
        print(f"HTTP error: {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
