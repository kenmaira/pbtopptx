import os
import requests

def list_releases(api_key):
    # Base URL for the Productboard API
    url = "https://api.productboard.com/releases"

    # Headers with API Key
    headers = {
        "Authorization": f"Bearer {os.getenv('PRODUCTBOARD_API_TOKEN')}",
        "Content-Type": "application/json",
        "X-Version": "1"
    }

    # List to store all releases
    all_releases = []

    while url:  # Continue as long as there's a `next` URL
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(f"Error: {response.status_code} - {response.text}")
            break

        data = response.json()

        # Add releases from the current page to the list
        all_releases.extend(data.get("data", []))

        # Update URL to the next page if it exists
        url = data.get("links", {}).get("next")

    return all_releases

if __name__ == "__main__":
    YOUR_API_KEY = "Authorization"

    releases = list_releases("Authorization")

    if releases:
        print(f"Total releases found: {len(releases)}\n")
        for release in releases:
            print(f"Release ID: {release.get('id')}")
            print(f"Name: {release.get('name')}")
            print(f"Description: {release.get('description', 'No description')}\n")
    else:
        print("No releases found.")
