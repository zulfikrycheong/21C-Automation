import base64
from datetime import datetime
import json
import os
import requests
import streamlit as st

# Configuration for GitHub Repo Sync (The GBA Wireless Adapter)
# Store your GitHub Personal Access Token (with repo / contents write access) in st.secrets["github_token"]
# Store your repo details in st.secrets["github_repo"] (e.g., "your-username/chambersos")
GITHUB_API_URL = "https://api.github.com/repos"


def get_sync_credentials():
  try:
    token = st.secrets["github_token"]
    repo = st.secrets["github_repo"]  # format: "owner/repo"
    return token, repo
  except Exception:
    return None, None


def pull_active_session():
  """Pulls the latest scanning queue state from the repo JSON file."""
  token, repo = get_sync_credentials()
  if not token or not repo:
    return None, None  # <-- Fixed: returns two values so unpacking never crashes

  url = f"{GITHUB_API_URL}/{repo}/contents/carton_sync.json"
  headers = {
      "Authorization": f"Bearer {token}",
      "Accept": "application/vnd.github+json",
  }

  try:
    res = requests.get(url, headers=headers, timeout=5)
    if res.status_code == 200:
      data = res.json()
      decoded_content = base64.b64decode(data["content"]).decode("utf-8")
      return json.loads(decoded_content), data["sha"]
  except Exception as e:
    print(f"Sync pull error: {e}")

  return None, None


def push_active_session(carton_no, department, queue):
  """Pushes updated active carton state to the repo JSON file so all devices see it live."""
  token, repo = get_sync_credentials()
  if not token or not repo:
    return False

  url = f"{GITHUB_API_URL}/{repo}/contents/carton_sync.json"
  headers = {
      "Authorization": f"Bearer {token}",
      "Accept": "application/vnd.github+json",
  }

  # First, get current SHA to prevent edit conflicts
  _, sha = pull_active_session()

  payload_data = {
      "carton_no": carton_no,
      "department": department,
      "queue": queue,
      "last_updated": datetime.utcnow().isoformat(),
  }

  json_str = json.dumps(payload_data, indent=2)
  encoded_content = base64.b64encode(json_str.encode("utf-8")).decode("utf-8")

  body = {
      "message": f"Sync active carton session: {carton_no} ({len(queue)} items)",
      "content": encoded_content,
  }
  if sha:
    body["sha"] = sha

  try:
    res = requests.put(url, headers=headers, json=body, timeout=5)
    return res.status_code in [200, 201]
  except Exception as e:
    print(f"Sync push error: {e}")
    return False
