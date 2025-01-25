import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image
from datetime import datetime
import argparse
import re
import os

# Headers for API calls
headers = {
    "accept": "application/json",
    "X-Version": "1",
    "authorization": "Bearer {os.getenv('PRODUCTBOARD_API_TOKEN')}"
}

# Step 1: Retrieve all features in a release
def get_feature_uuids(release_id):
    url = f"https://api.productboard.com/feature-release-assignments?release.id={release_id}"
    response = requests.get(url, headers=headers)
    print(f"API response for feature UUIDs: {response.status_code}")  # Debug: Check API response status
    features = response.json().get("data", [])
    feature_ids = [feature['feature']['id'] for feature in features]
    print(f"Feature UUIDs retrieved: {feature_ids}")  # Debug: Check the feature IDs
    return feature_ids

# Step 2: Get feature details
def get_feature_details(feature_id):
    url = f"https://api.productboard.com/features/{feature_id}"
    response = requests.get(url, headers=headers)
    print(f"API response for feature details: {response.status_code}")  # Debug: Check API response status
    return response.json()

# Step 3: Get custom field (requirements link)
def get_requirements_link(feature_id):
    custom_field_id = "52ae58e7-6417-4898-956b-bd74d4e87502"
    url = f"https://api.productboard.com/hierarchy-entities/custom-fields-values/value?customField.id={custom_field_id}&hierarchyEntity.id={feature_id}"
    response = requests.get(url, headers=headers)
    print(f"API response for requirements link: {response.status_code}")  # Debug: Check API response status
    return response.json().get("data", {}).get("value", None)

# Helper function to extract image URLs from the HTML
def extract_image_urls(description_html):
    img_tags = re.findall(r'<img [^>]*src="([^"]+)"[^>]*>', description_html)
    description_html = re.sub(r'<img [^>]*src="[^"]+"[^>]*>', '', description_html)
    return img_tags, description_html

# Helper function to clean HTML and apply formatting in PowerPoint
def clean_html_and_format_text(description_html, text_frame):
    description_html = description_html.replace('<br />', '\n').replace('<br>', '\n')
    description_html = description_html.replace('<p>', '').replace('</p>', '\n')
    paragraphs = description_html.split('\n')

    first_paragraph = True

    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if not paragraph:
            continue

        if first_paragraph and len(text_frame.paragraphs) > 0:
            p = text_frame.paragraphs[0]  # Use the first existing paragraph
            first_paragraph = False
        else:
            p = text_frame.add_paragraph()  # Add new paragraph only after the first

        bold_parts = re.split(r'(<b>|</b>|<strong>|</strong>)', paragraph)

        is_bold = False
        for part in bold_parts:
            run = p.add_run()

            if part in ['<b>', '<strong>']:
                is_bold = True
            elif part in ['</b>', '</strong>']:
                is_bold = False
            else:
                run.text = part.strip()
                if is_bold:
                    run.font.bold = True

            run.font.size = Pt(12)

#Insert image with maintained aspect ratio
def insert_image_with_aspect_ratio(slide, placeholder, img):
    # Get the dimensions of the placeholder
    ph_width = placeholder.width
    ph_height = placeholder.height

    # Get the dimensions of the image
    img_width, img_height = img.size

    # Calculate the aspect ratios
    ph_aspect = ph_width / ph_height
    img_aspect = img_width / img_height

    # Determine how to scale the image based on aspect ratio comparison
    if img_aspect > ph_aspect:
        # Image is wider than the placeholder (scale by width)
        new_width = ph_width
        new_height = int(ph_width / img_aspect)
    else:
        # Image is taller than the placeholder (scale by height)
        new_height = ph_height
        new_width = int(ph_height * img_aspect)

    # Calculate the position to center the image within the placeholder
    left = placeholder.left + int((ph_width - new_width) / 2)
    top = placeholder.top + int((ph_height - new_height) / 2)

    # Insert the resized image into the slide at the calculated position and size
    img_stream = BytesIO()
    img.save(img_stream, format="PNG")
    img_stream.seek(0)

    # Insert the picture into the slide
    picture = slide.shapes.add_picture(img_stream, left, top, width=new_width, height=new_height)

    print("Image inserted with proper aspect ratio and centered.")


# Step 4: Create PowerPoint slides
def create_pptx(features_data):
    prs = Presentation("corporate_template.pptx")
    
    for feature in features_data:
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        print(f"Added slide for feature: {feature['title']}")

        title = slide.placeholders[0]
        title.text = feature['title']
        print(f"Set title: {feature['title']}")

        description_html = feature.get("description", "")
        images_links, cleaned_description_html = extract_image_urls(description_html)

        if images_links:
            for idx, img_url in enumerate(images_links[:4]):  # Max 4 images per slide
                try:
                    print(f"Fetching image {idx + 1} from URL: {img_url}")  # Debug: Print image URL
                    img_response = requests.get(img_url, stream=True)

                    # Verify the response status and Content-Type
                    if img_response.status_code != 200:
                        print(f"Failed to fetch image {idx + 1}. HTTP status code: {img_response.status_code}")
                        continue
                    content_type = img_response.headers.get("Content-Type", "")
                    if not content_type.startswith("image/"):
                        print(f"URL does not point to a valid image: {img_url}")
                        continue

                    # Try opening the image
                    img = Image.open(BytesIO(img_response.content))
                    img_placeholder = slide.placeholders[10]  # Use the picture placeholder

                    # Insert image with aspect ratio preserved
                    insert_image_with_aspect_ratio(slide, img_placeholder, img)
                    print(f"Added image {idx + 1}")
                except Exception as e:
                    print(f"Error processing image {idx + 1} from URL {img_url}: {e}")
        else:
            print("No images found for this feature.")
            
        description_placeholder = slide.placeholders[1]
        clean_html_and_format_text(cleaned_description_html, description_placeholder.text_frame)
        print(f"Set description")

        try:
            requirements_link = feature.get("requirements_link", "Missing requirements")
            requirements_placeholder = slide.placeholders[11]
            p = requirements_placeholder.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = "Link to requirements"
            r = run.hyperlink
            r.address = requirements_link
            print(f"Set requirements link: {requirements_link}")
        except KeyError:
            print("No placeholder with index 11 on this slide, skipping the requirements link.")

        textbox = slide.shapes.add_textbox(left=Inches(0.25), top=Inches(0.25), width=Inches(2), height=Inches(0.5))
        tb = textbox.text_frame
        pb_paragraph = tb.add_paragraph()

        run = pb_paragraph.add_run()
        run.text = "PB Link"
        run.font.size = Pt(11)
        run.font.name = "Avenir"
        run.hyperlink.address = feature.get("html_link", "")
        print(f"Set Productboard link")

    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    output_filename = f"output_presentation_{timestamp}.pptx"
    prs.save(output_filename)
    print(f"Presentation saved as {output_filename}")

# Main execution
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate PowerPoint slides from Productboard release features.")
    parser.add_argument("release_id", help="The release ID for which to generate the PowerPoint slides.")
    parser.add_argument("--owner_email", help="The email of the owner to filter features by (optional).", default=None)
    args = parser.parse_args()

    release_id = args.release_id
    owner_email = args.owner_email

    # Retrieve feature IDs
    feature_ids = get_feature_uuids(release_id)

    features_data = []
    for feature_id in feature_ids:
        feature_details = get_feature_details(feature_id)
        owner_email_from_feature = feature_details.get("data", {}).get("owner", {}).get("email", "")

        if not owner_email or owner_email_from_feature == owner_email:
            requirements_link = get_requirements_link(feature_id)

            feature_data = {
                "title": feature_details.get("data", {}).get("name", ""),
                "description": feature_details.get("data", {}).get("description", ""),
                "images": [],
                "requirements_link": requirements_link,
                "html_link": feature_details.get("data", {}).get("links", {}).get("html", "")
            }
            features_data.append(feature_data)
        else:
            print(f"Skipping feature {feature_id} as it is not assigned to {owner_email}")

    create_pptx(features_data)
