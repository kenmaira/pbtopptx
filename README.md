# Productboard Q2 PPTX Generator

This script generates a PowerPoint deck of Productboard features for a given timeframe and set of statuses, grouped by initiative.

## Overview

- **Fetches** only features with specific status IDs (whitelisted).
- **Filters** by timeframe (default Q2 2025: 2025-04-01 to 2025-06-30).
- **Optionally** filters by owner email.
- **Generates** slides grouped by initiative in a corporate PowerPoint template.

## Prerequisites

- Python 3.8+
- A corporate PowerPoint template with IDX containers set for Title, Description, Images, Productboard Link, Jira ID, and Requirements link on the master slides at `templates/corporate_template.pptx`, and at least two slides in the master slide.

   - Default is set to:
      ```TITLE_IDX = 0
      DESCRIPTION_IDX = 1
      IMAGE_IDX = 10
      PB_LINK_IDX = 11
      ID_IDX = 12
      REQUIREMENTS_IDX = 13
      ```

### Productboard API Token

To authenticate, the script uses a Productboard API token stored securely in your keyring.

1. **Generate a token**  
   - Log into Productboard  
   - Navigate to **Your Profile** → **API Tokens**  
   - Create a new token with **read** permissions for features, initiatives, and custom fields  
   - Copy the token value

2. **Store in keyring**  
   ```bash
   pip install keyring  # if you haven’t already
   keyring set productboard-api default
   # When prompted, paste your Productboard token
   ```

3. **Verify storage**  
   ```bash
   keyring get productboard-api default
   # Should echo your token back
   ```

The script will automatically retrieve your token at runtime using `keyring.get_password("productboard-api", "default")`.

### Python dependencies

Install via pip:

```bash
pip install -r requirements.txt
```

## Configuration

1. **Status IDs**  
   In the script, define `INCLUDED_STATUS_IDS` with the UUIDs of the statuses you want:

   ```python
   INCLUDED_STATUS_IDS = [
       "11111111-2222-3333-4444-555555555555",  # e.g. Planned
       "66666666-7777-8888-9999-000000000000",  # e.g. In Review
   ]
   ```

2. **Timeframe**  
   Adjust `TIMEFRAME_START` and `TIMEFRAME_END` as needed:

   ```python
   TIMEFRAME_START = datetime.fromisoformat("2025-04-01")
   TIMEFRAME_END   = datetime.fromisoformat("2025-06-30")
   ```

3. **Template Path**  
   Ensure your corporate template is at `templates/corporate_template.pptx`. Rename or update the path if different.

## Usage

```bash
python generate_q2_pptx.py [--owner_email someone@example.com]
```

- `--owner_email`: (optional) only include features owned by this email.

Example:

```bash
python generate_q2_pptx.py --owner_email alice@company.com
```

After running, an output file `output_presentation_YYYY-MM-DD_HH-MM-SS.pptx` will be created.

## How It Works

1. **Fetch IDs**  
   Calls `/features?status.id={status_id}` for each `INCLUDED_STATUS_IDS` and unions the results.
2. **Preserve Order**  
   Retrieves all features once to capture API ordering, then filters to the whitelisted IDs.
3. **Initiative Mapping**  
   Fetches `/initiatives` and their linked features to group slides.
4. **Detail Fetch**  
   In parallel, calls `/features/{id}` to get descriptions, timeframe, owner, and links.
5. **Filter**  
   Filters by timeframe overlap and optional owner email.
6. **Slide Generation**  
   Uses `python-pptx` to generate a deck:
   - Cover slides per initiative using the second slide on the slide master 
   - Feature slides with title, formatted description, images, Productboard/JIRA links, requirements link using the first slide in the slide master

## Troubleshooting

- **Empty placeholders**: If slides corrupt, ensure `fill_empty_text_if_needed()` logic is intact.
- **Missing images**: Check that image URLs are valid Productboard S3 links.
- **Authentication**: If the script errors finding the API token, re-run the keyring set command.

## Contributing

Feel free to open issues or submit pull requests to:

- Add more filtering options (by tags, owner roles, etc.)
- Support different output templates or formats
- Improve HTML-to-PPTX formatting rules

