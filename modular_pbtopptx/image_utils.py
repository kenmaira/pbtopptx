from PIL import Image
from io import BytesIO
import re
import requests

# Image handling utilities for modular_pbtopptx

def extract_image_urls(description_html):
    img_tags = re.findall(r'<img [^>]*src="([^"]+)"[^>]*>', description_html)
    cleaned_html = re.sub(r'<img [^>]*src="[^"]+"[^>]*>', '', description_html)
    return img_tags, cleaned_html

def fetch_image(url):
    try:
        if "pb-files.s3.amazonaws.com" not in url:
            print(f"Skipping non-Productboard image URL: {url}")
            return None
        resp = requests.get(url)
        if resp.status_code != 200:
            print(f"Failed to fetch image from {url}. Status code: {resp.status_code}")
            return None
        content_type = resp.headers.get("Content-Type")
        if not content_type or not content_type.startswith("image/"):
            print(f"URL is not an image: {url}")
            return None
        img = Image.open(BytesIO(resp.content))
        img.verify()
        return Image.open(BytesIO(resp.content))
    except Exception as e:
        print(f"Error fetching or processing image from {url}: {e}")
        return None

def insert_image_with_aspect_ratio(slide, placeholder, img):
    ph_width = placeholder.width
    ph_height = placeholder.height
    img_width, img_height = img.size
    ph_aspect = ph_width / ph_height
    img_aspect = img_width / img_height
    if img_aspect > ph_aspect:
        new_width = ph_width
        new_height = int(ph_width / img_aspect)
    else:
        new_height = ph_height
        new_width = int(ph_height * img_aspect)
    left = placeholder.left + int((ph_width - new_width) / 2)
    top = placeholder.top + int((ph_height - new_height) / 2)
    img_stream = BytesIO()
    img.save(img_stream, format="PNG")
    img_stream.seek(0)
    slide.shapes.add_picture(img_stream, left, top, width=new_width, height=new_height)
    print("✅ Inserted image with aspect ratio and centered.")
