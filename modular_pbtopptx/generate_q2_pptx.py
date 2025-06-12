import argparse
import keyring
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from .image_utils import extract_image_urls, fetch_image, insert_image_with_aspect_ratio
from .text_utils import safe_get, fill_empty_text_if_needed, add_run, clean_html_and_format_text
from .data_fetch import get_all_paginated_features, get_feature_ids_by_status_id, get_feature_details, get_requirements_link, get_jira_details, get_placeholder_by_idx
from .pptx_utils import safe_clear_and_add_text, create_pptx

TITLE_IDX = 0
DESCRIPTION_IDX = 1
IMAGE_IDX = 10
PB_LINK_IDX = 11
ID_IDX = 12
REQUIREMENTS_IDX = 13

INCLUDED_STATUS_IDS = [
    #"402b0df5-1554-44bf-bcc4-8a11d3ca0a65",  # Candidate
    "b944953f-2a93-4e85-b036-9bef92432588", #Planned
    "d3dad2ee-2692-4e06-a29c-4ec314f807a0", #In Progress
    "d6c9c0da-411a-41a9-8343-a8a5f75036e3", #EG.EG
    "93d3820a-af3f-43e3-9122-32e302d7efd1", #Beta
    "65a03095-72c0-42f2-952a-ab1c30abdae4", #Limited Release 
    "88c3f59c-2cb0-4438-8840-d00f85fce6d1" # Released
]

EXCLUDED_STATUS_IDS = [
    "402b0df5-1554-44bf-bcc4-8a11d3ca0a65",  # Candidate
    "310a6c38-3719-4bff-a26c-b5d520ea1298",  # New Idea
]
JIRA_API_ID = "155d80cb-8b4c-4bff-a26c-b5d520ea1298"

api_token = keyring.get_password("productboard-api", "default")
if not api_token:
    raise RuntimeError("API token not found in keyring.")

HEADERS = {
    "accept": "application/json",
    "X-Version": "1",
    "authorization": f"Bearer {api_token}"
}

TIMEFRAME_START = datetime.fromisoformat("2025-04-01")
TIMEFRAME_END = datetime.fromisoformat("2025-06-30")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--owner_email", default=None, help="Filter by owner email")
    parser.add_argument("--timeframe_start", default=None, help="Start date for timeframe filter (YYYY-MM-DD)")
    parser.add_argument("--timeframe_end", default=None, help="End date for timeframe filter (YYYY-MM-DD)")
    args = parser.parse_args()

    global TIMEFRAME_START, TIMEFRAME_END
    if args.timeframe_start:
        TIMEFRAME_START = datetime.fromisoformat(args.timeframe_start)
    if args.timeframe_end:
        TIMEFRAME_END = datetime.fromisoformat(args.timeframe_end)

    # 1) Fetch only the statuses you care about
    print("🔍 Fetching only specified statuses…")
    feature_ids = set()
    for status_id in INCLUDED_STATUS_IDS:
        print(f"🔎 Fetching features with status {status_id}")
        feature_ids |= get_feature_ids_by_status_id(status_id)

    # 2) (Optional) preserve the API’s original ordering
    print("📋 Retrieving all features to preserve order…")
    all_features = get_all_paginated_features("https://api.productboard.com/features")
    remaining_ids = [f["id"] for f in all_features if f["id"] in feature_ids]
    print(f"✅ Will process {len(remaining_ids)} features\n")

    # 3) Build initiative lookups
    print("📂 Fetching initiatives…")
    initiatives = get_all_paginated_features("https://api.productboard.com/initiatives")
    initiative_map = {i["id"]: i["name"] for i in initiatives}

    print("🔗 Building feature→initiative map…")
    feature_to_initiative = {}
    for iid in initiative_map:
        links = get_all_paginated_features(
            f"https://api.productboard.com/initiatives/{iid}/links/features"
        )
        for link in links:
            fid = link.get("id")
            if fid:
                feature_to_initiative[fid] = iid

    # 4) Fetch full feature details in parallel, filter by timeframe + owner
    features = []
    print("🧵 Fetching feature details in parallel…")
    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = {
            executor.submit(get_feature_details, fid): fid
            for fid in remaining_ids
        }
        for future in as_completed(futures):
            fid = futures[future]
            try:
                details = future.result().get("data", {})
                tf = details.get("timeframe", {})
                start, end = tf.get("startDate"), tf.get("endDate")
                owner = details.get("owner") or {}
                email = owner.get("email", "")

                # Skip if no valid dates
                if not (start and end):
                    continue

                s_dt = datetime.fromisoformat(start)
                e_dt = datetime.fromisoformat(end)
                if not (s_dt <= TIMEFRAME_END and e_dt >= TIMEFRAME_START):
                    continue

                # Skip if owner_email filter is set
                if args.owner_email and args.owner_email != email:
                    continue

                # Gather links
                req_link = get_requirements_link(fid)
                jira_key, jira_url = get_jira_details(fid)
                initiative = initiative_map.get(
                    feature_to_initiative.get(fid), "Uncategorized"
                )

                features.append({
                    "id": fid,
                    "title": details.get("name", ""),
                    "description": details.get("description", ""),
                    "requirements_link": req_link,
                    "html_link": details.get("links", {}).get("html", ""),
                    "initiative": initiative,
                    "jira_key": jira_key,
                    "jira_url": jira_url
                })

            except Exception as e:
                print(f"❌ Error fetching feature {fid}: {e}")

    print(f"\n📊 Final features to generate slides for: {len(features)}\n")

    # 5) Create the PPTX
    create_pptx(features)


if __name__ == "__main__":
    main()
