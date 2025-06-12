import requests
from datetime import datetime

HEADERS = None  # To be set from main script
JIRA_API_ID = None  # To be set from main script


def get_all_paginated_features(url):
    all_features = []
    page = 1
    while url:
        print(f"📅 Fetching features page {page}: {url}")
        response = requests.get(url, headers=HEADERS)
        if response.status_code != 200:
            print(f"❌ Failed to fetch features: {response.text}")
            break
        data = response.json()
        features = data.get("data", [])
        print(f"🔹 Retrieved {len(features)} features on page {page}")
        all_features.extend(features)
        url = data.get("links", {}).get("next")
        page += 1
    print(f"📆 Total features retrieved: {len(all_features)}")
    return all_features


def get_feature_ids_by_status_id(status_id):
    url = f"https://api.productboard.com/features?status.id={status_id}"
    return set(f["id"] for f in get_all_paginated_features(url))


def get_feature_details(fid):
    url = f"https://api.productboard.com/features/{fid}"
    response = requests.get(url, headers=HEADERS)
    print(f"📄 Feature {fid} details status: {response.status_code}")
    return response.json()


def get_requirements_link(fid):
    custom_field_id = "52ae58e7-6417-4898-956b-bd74d4e87502"
    url = f"https://api.productboard.com/hierarchy-entities/custom-fields-values/value?customField.id={custom_field_id}&hierarchyEntity.id={fid}"
    response = requests.get(url, headers=HEADERS)
    return response.json().get("data", {}).get("value", None)


def get_jira_details(fid):
    url = f"https://api.productboard.com/jira-integrations/{JIRA_API_ID}/connections/{fid}"
    response = requests.get(url, headers=HEADERS)
    if response.status_code != 200:
        print(f"⚠️ Jira link not found for {fid}: {response.status_code}")
        return None, None
    jira_data = response.json()
    issue_key = jira_data.get("data", {}).get("connection", {}).get("issueKey")
    if issue_key:
        return issue_key, f"https://jira.egnyte-it.com/browse/{issue_key}"
    return None, None


def get_placeholder_by_idx(slide, idx):
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == idx:
            if shape.has_text_frame and not shape.text_frame.text.strip():
                shape.text_frame.text = " "
            return shape
    return None
