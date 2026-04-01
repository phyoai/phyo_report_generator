"""
Campaign Report Generator - Flask Backend
Analyzes campaign screenshots using Anthropic Claude and populates PowerPoint template
"""

from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from anthropic import Anthropic
import os
import base64
import io
import json
import mimetypes
import re
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import copy
import time
from dotenv import load_dotenv
import pytesseract
from PIL import Image
import shutil

# Configure Tesseract path for Windows
import os
if os.name == 'nt':  # Windows
    pytesseract.pytesseract.pytesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

load_dotenv(override=True)  # Load environment variables from .env file (override=True forces reload)

app = Flask(__name__)
CORS(app)  # Enable CORS for React frontend

# Free tools for video/social metrics (no API keys needed)
try:
    import yt_dlp
    HAS_YT_DLP = True
except:
    HAS_YT_DLP = False

# Initialize client - supports both Anthropic and OpenRouter
if os.environ.get("OPENROUTER_API_KEY"):
    # Use OpenRouter with OpenAI-compatible format
    from openai import OpenAI as OpenAIClient
    client = OpenAIClient(
        api_key=os.environ.get("OPENROUTER_API_KEY"),
        base_url="https://openrouter.ai/api/v1",
        default_headers={
            "HTTP-Referer": "http://localhost:5000",
            "X-Title": "Campaign Report Generator"
        }
    )
    USE_OPENROUTER = True
else:
    # Use direct Anthropic API
    client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    USE_OPENROUTER = False

# Configuration
# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "Campaign_Report_PyroMedia.pptx")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "generated_reports")
os.makedirs(OUTPUT_DIR, exist_ok=True)

FETCH_CACHE_TTL_SECONDS = 300
_FETCH_CACHE = {}


class InMemoryUpload(io.BytesIO):
    def __init__(self, data, filename="", content_type=""):
        super().__init__(data)
        self.filename = filename
        self.content_type = content_type


def make_in_memory_upload(upload):
    upload_bytes = upload.read()
    upload.seek(0)
    return InMemoryUpload(
        upload_bytes,
        filename=getattr(upload, "filename", "") or "",
        content_type=getattr(upload, "content_type", "") or "",
    )


def clone_in_memory_upload(upload):
    upload.seek(0)
    return InMemoryUpload(
        upload.getvalue(),
        filename=getattr(upload, "filename", "") or "",
        content_type=getattr(upload, "content_type", "") or "",
    )


def clone_uploads(images):
    return [clone_in_memory_upload(image) for image in images]


def _cache_get(cache_key):
    cached = _FETCH_CACHE.get(cache_key)
    if not cached:
        return None
    cached_at, value = cached
    if time.time() - cached_at > FETCH_CACHE_TTL_SECONDS:
        _FETCH_CACHE.pop(cache_key, None)
        return None
    return copy.deepcopy(value)


def _cache_set(cache_key, value):
    _FETCH_CACHE[cache_key] = (time.time(), copy.deepcopy(value))


OPENROUTER_PROMPT_TEMPLATES = {
    "master": """You are an influencer marketing analyst.

Create a professional campaign performance report using the data below.

Campaign Details:
- Campaign Name: {campaign_name}
- Influencer: {influencer}
- Budget: ${budget}
- Duration: {duration}
- Platform: {platform}

Input Metrics:
- Views: {views}
- Likes: {likes}
- Comments: {comments}
- Total Engagement: {total_engagement}
- CPV: {cpv}
- CPE: {cpe}
- Engagement Rate: {engagement_rate}

Important Rules:
- Shares and Saves are not available, do not fabricate real values
- You may estimate them logically if needed but clearly mark them as estimated
- Keep report concise but professional
- Give actionable insights

Output Format (STRICT):
1. Campaign Summary (short paragraph)
2. Key Metrics (bullet points)
3. Performance Analysis
4. Recommendations
5. Final JSON (very important)

JSON Format:
{{
  "views": number,
  "likes": number,
  "comments": number,
  "shares": "N/A or estimated",
  "saves": "N/A or estimated",
  "total_engagement": number,
  "budget": number,
  "cpv": number,
  "cpe": number,
  "engagement_rate": number
}}""",
    "fast": """Generate a short influencer campaign report.

Data:
Views: {views}
Likes: {likes}
Comments: {comments}
Budget: {budget}

Calculate:
- Total Engagement
- CPV
- CPE
- Engagement Rate

Return:
1. Summary
2. Insights
3. JSON output""",
    "advanced": """Act as a senior marketing analyst.

Analyze this influencer campaign deeply.

Campaign:
{campaign_name}
Budget: ${budget}

Metrics:
Views: {views}
Likes: {likes}
Comments: {comments}

Tasks:
1. Calculate engagement metrics
2. Evaluate performance (low / average / high)
3. Identify strengths and weaknesses
4. Suggest improvements for next campaign
5. Compare engagement efficiency vs cost

Return:
- Performance Score (out of 10)
- Insights
- Optimization strategy
- JSON metrics""",
    "json_only": """Calculate campaign metrics and return ONLY JSON.

Input:
Views: {views}
Likes: {likes}
Comments: {comments}
Budget: {budget}

Output JSON:
{{
  "views": number,
  "likes": number,
  "comments": number,
  "total_engagement": number,
  "cpv": number,
  "cpe": number,
  "engagement_rate": number
}}""",
    "multi_platform": """Create a campaign report for:

Campaign: {campaign_name}
Budget: ${budget}

Instagram Metrics:
Views: {ig_views}
Likes: {ig_likes}
Comments: {ig_comments}

YouTube Metrics:
Views: {yt_views}
Likes: {yt_likes}
Comments: {yt_comments}

Tasks:
- Compare platform performance
- Identify which platform performed better
- Provide insights
- Return combined JSON

JSON:
{{
  "instagram": {{...}},
  "youtube": {{...}},
  "best_platform": "Instagram/YouTube"
}}""",
}


def has_tesseract():
    """Return True when the Tesseract executable is available."""
    configured_path = getattr(pytesseract.pytesseract, "tesseract_cmd", "")
    return bool(
        (configured_path and os.path.exists(configured_path))
        or shutil.which("tesseract")
    )


def detect_image_media_type(image_file, image_bytes):
    """Best-effort detection of an uploaded image's MIME type."""
    content_type = getattr(image_file, "content_type", "") or ""
    if content_type.startswith("image/"):
        return content_type

    filename = getattr(image_file, "filename", "") or ""
    guessed_type, _ = mimetypes.guess_type(filename)
    if guessed_type and guessed_type.startswith("image/"):
        return guessed_type

    try:
        pil_image = Image.open(io.BytesIO(image_bytes))
        image_format = (pil_image.format or "").lower()
        format_map = {
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "png": "image/png",
            "webp": "image/webp",
            "gif": "image/gif",
            "bmp": "image/bmp",
        }
        return format_map.get(image_format, "image/jpeg")
    except Exception:
        return "image/jpeg"


class SafePromptDict(dict):
    def __missing__(self, key):
        return "N/A"


def build_openrouter_prompt(template_name="master", data=None):
    """Format one of the built-in OpenRouter prompt templates."""
    template = OPENROUTER_PROMPT_TEMPLATES.get(template_name)
    if not template:
        raise ValueError(f"Unknown prompt template: {template_name}")

    payload = SafePromptDict({
        "campaign_name": "Summer Fashion Campaign",
        "influencer": "Influencer X",
        "budget": 5000,
        "duration": "June 1 to June 30",
        "platform": "Instagram Reel",
        "views": "N/A",
        "likes": "N/A",
        "comments": "N/A",
        "total_engagement": "N/A",
        "cpv": "N/A",
        "cpe": "N/A",
        "engagement_rate": "N/A",
        "ig_views": "N/A",
        "ig_likes": "N/A",
        "ig_comments": "N/A",
        "yt_views": "N/A",
        "yt_likes": "N/A",
        "yt_comments": "N/A",
    })
    if data:
        payload.update(data)
    return template.format_map(payload)


def extract_metrics_from_image_ocr(image_file):
    """Extract metrics from image using Tesseract OCR"""
    import re

    try:
        if not has_tesseract():
            print("[WARNING]  OCR skipped: Tesseract is not installed or not available in PATH")
            return {}

        print("[SEARCH] Running OCR on image...")

        # Read image file
        image_bytes = image_file.read()
        image_file.seek(0)

        # Convert to PIL Image
        img = Image.open(io.BytesIO(image_bytes))

        # Run Tesseract OCR
        extracted_text = pytesseract.image_to_string(img)

        print(f"[NOTE] OCR extracted text (first 300 chars):\n{extracted_text[:300]}...\n")

        # Parse metrics from extracted text
        return extract_metrics_from_text(extracted_text)

    except Exception as e:
        print(f"[WARNING]  OCR failed: {str(e)}")
        return {}


def extract_metrics_from_images_ocr(images):
    merged_metrics = {}
    if not images:
        return merged_metrics

    max_workers = min(4, len(images)) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(extract_metrics_from_image_ocr, image) for image in images]
        for future in futures:
            image_metrics = future.result() or {}
            for key, value in image_metrics.items():
                if value and value != "N/A" and not merged_metrics.get(key):
                    merged_metrics[key] = value
    return merged_metrics


def fetch_youtube_metrics(url_or_text):
    """
    Fetch YouTube metrics using yt-dlp (FREE - no API key needed).
    Handles both video URLs and channel/username URLs (@handle or /channel/).
    Returns dict with keys matching the report's youtube schema.
    """
    if not HAS_YT_DLP:
        print("[WARNING]  yt-dlp not installed")
        return {}

    import yt_dlp
    import re

    # Accept either a raw URL or a blob of text containing a URL
    youtube_url = url_or_text.strip()
    if not youtube_url.startswith("http"):
        urls = re.findall(r'(https?://(?:www\.)?(?:youtube\.com|youtu\.be)[^\s]*)', url_or_text)
        if not urls:
            return {}
        youtube_url = urls[0]

    print(f"[VIDEO] Fetching YouTube metrics from: {youtube_url}")

    # Detect whether this is a channel/user URL or a direct video URL
    is_channel = any(x in youtube_url for x in [
        "/@", "/channel/", "/user/", "/c/",
        "youtube.com/@", "youtube.com/channel",
    ])

    quiet_opts = {"quiet": True, "no_warnings": True}

    try:
        if is_channel:
            # --- Channel URL: fetch the most recent upload and use its metrics ---
            channel_opts = {
                **quiet_opts,
                "extract_flat": "in_playlist",   # list videos without downloading
                "playlistend": 1,                 # only the latest video
            }
            with yt_dlp.YoutubeDL(channel_opts) as ydl:
                channel_info = ydl.extract_info(youtube_url, download=False)

            # Get the first video URL from the channel playlist
            entries = channel_info.get("entries") or []
            if not entries:
                print("[WARNING]  No videos found on channel")
                return {}

            first_entry = entries[0]
            video_id  = first_entry.get("id") or first_entry.get("url", "")
            video_url = f"https://www.youtube.com/watch?v={video_id}" if len(video_id) == 11 else first_entry.get("url", "")
            creator_name = channel_info.get("uploader") or channel_info.get("channel") or ""
            print(f"[VIDEO] Latest video: {video_url}")
        else:
            video_url    = youtube_url
            creator_name = ""

        # --- Fetch full video metadata ---
        with yt_dlp.YoutubeDL(quiet_opts) as ydl:
            info = ydl.extract_info(video_url, download=False)

        metrics = {}

        if info.get("view_count"):
            metrics["views"] = str(info["view_count"])
        if info.get("like_count"):
            metrics["likes"] = str(info["like_count"])
        if info.get("comment_count"):
            metrics["comments"] = str(info["comment_count"])

        # Duration in seconds → format as "Xm Ys" (closest public proxy for watch time)
        duration = info.get("duration")
        if duration:
            mins, secs = divmod(int(duration), 60)
            metrics["watchTime"] = f"{mins}m {secs}s"

        # Top-level metadata (used outside youtube dict)
        metrics["_title"]        = info.get("title", "")
        metrics["_creator_name"] = creator_name or info.get("uploader") or info.get("channel") or ""

        print(f"[OK] YouTube metrics: {metrics}")
        return metrics

    except Exception as e:
        print(f"[WARNING]  YouTube fetch error: {str(e)}")
        return {}


def extract_metrics_from_text(text):
    """Extract campaign metrics from user's text input - AGGRESSIVE extraction"""
    import re

    data = {
        "clicks": "N/A",
        "views": "N/A",
        "likes": "N/A",
        "shares": "N/A",
        "saves": "N/A",
        "total_engagement": "N/A",
        "budget": "N/A",
        "budget_currency": "",
        "cpv": "N/A",
        "cpe": "N/A",
        "cpc": "N/A",
        "cpc_goal": "",
        "cpv_goal": "",
        "cpc_calculation": "",
        "cpv_calculation": "",
        "engagement_rate": "N/A"
    }

    text_lower = text.lower()

    # Extract Views (multiple patterns)
    views_patterns = [
        r'no\.?\s+of\s+views?[\s:]*([\d,\.]+[kmb]?)',
        r'views?[\s:]*[\$]?([\d,\.]+[kmb]?)',
        r'(\d+[,\d]*)\s+views?',
        r'view count[\s:]*(\d+[,\d]*)',
    ]
    for pattern in views_patterns:
        views_match = re.search(pattern, text_lower)
        if views_match:
            data["views"] = views_match.group(1).replace(',', '').replace('k', '000').replace('m', '000000')
            break

    clicks_patterns = [
        r'clicks?[\s:]*([\d,\.]+[kmb]?)',
        r'total\s+clicks?[\s:]*([\d,\.]+[kmb]?)',
        r'link\s+clicks?[\s:]*([\d,\.]+[kmb]?)',
    ]
    for pattern in clicks_patterns:
        clicks_match = re.search(pattern, text_lower)
        if clicks_match:
            data["clicks"] = clicks_match.group(1).replace(',', '').replace('k', '000').replace('m', '000000')
            break

    # Extract Likes (multiple patterns)
    likes_patterns = [
        r'likes?[\s:]*[\$]?([\d,\.]+[kmb]?)',
        r'(\d+[,\d]*)\s+likes?',
        r'like count[\s:]*(\d+[,\d]*)',
    ]
    for pattern in likes_patterns:
        likes_match = re.search(pattern, text_lower)
        if likes_match:
            data["likes"] = likes_match.group(1).replace(',', '').replace('k', '000').replace('m', '000000')
            break

    # Extract Shares
    shares_patterns = [
        r'shares?[\s:]*[\$]?([\d,\.]+[kmb]?)',
        r'(\d+[,\d]*)\s+shares?',
    ]
    for pattern in shares_patterns:
        shares_match = re.search(pattern, text_lower)
        if shares_match:
            data["shares"] = shares_match.group(1).replace(',', '').replace('k', '000').replace('m', '000000')
            break

    # Extract Saves
    saves_patterns = [
        r'saves?[\s:]*[\$]?([\d,\.]+[kmb]?)',
        r'(\d+[,\d]*)\s+saves?',
    ]
    for pattern in saves_patterns:
        saves_match = re.search(pattern, text_lower)
        if saves_match:
            data["saves"] = saves_match.group(1).replace(',', '').replace('k', '000').replace('m', '000000')
            break

    # Extract Budget (multiple patterns)
    budget_patterns = [
        (r'budget[\s:]*([\d,]+\.?\d*)\s*inr', 'INR'),
        (r'budget[\s:]*inr\s*([\d,]+\.?\d*)', 'INR'),
        (r'budget[\s:]*rs\.?\s*([\d,]+\.?\d*)', 'INR'),
        (r'budget[\s:]*₹\s*([\d,]+\.?\d*)', 'INR'),
        (r'budget[\s:]*[\$]([\d,]+\.?\d*)', 'USD'),
        (r'\$([\d,]+)\s+budget', 'USD'),
        (r'budget.*?([\d,]+)', ''),
    ]
    for pattern, currency in budget_patterns:
        budget_match = re.search(pattern, text_lower)
        if budget_match:
            data["budget"] = budget_match.group(1).replace(',', '')
            data["budget_currency"] = currency
            break

    # Extract CPV (Cost Per View)
    cpv_match = re.search(r'(?:cpv|cost\s+per\s+view)[\s:]*?(?:rs\.?|inr|\$)?\s*([\d,]+(?:\.\d+)?)', text_lower)
    if cpv_match:
        data["cpv"] = cpv_match.group(1).replace(',', '')

    # Extract CPC (Cost Per Click)
    cpc_match = re.search(r'(?:cpc|cost\s+per\s+click)[\s:]*?(?:rs\.?|inr|\$)?\s*([\d,]+(?:\.\d+)?)', text_lower)
    if cpc_match:
        data["cpc"] = cpc_match.group(1).replace(',', '')

    # Extract CPE (Cost Per Engagement)
    cpe_match = re.search(r'(?:cpe|cost\s+per\s+engagement)[\s:]*?(?:rs\.?|inr|\$)?\s*([\d,]+(?:\.\d+)?)', text_lower)
    if cpe_match:
        data["cpe"] = cpe_match.group(1).replace(',', '')

    # Extract explicit Total Engagement if provided.
    total_engagement_match = re.search(r'total\s+engagement[\s:]*([\d,\.]+[kmb]?)', text_lower)
    if total_engagement_match:
        data["total_engagement"] = total_engagement_match.group(1).replace(',', '').replace('k', '000').replace('m', '000000')

    # Capture non-numeric goal/intent lines for CPC / CPV if present.
    for metric_key, goal_key in [("cpc", "cpc_goal"), ("cpv", "cpv_goal")]:
        goal_match = re.search(rf'(?im)^\s*{metric_key}\s*:\s*([^\n]*[A-Za-z][^\n]*)$', text)
        if goal_match:
            data[goal_key] = goal_match.group(1).strip()

    calc_section_match = re.search(r'(?is)calculation\s*:\s*(.+)', text)
    if calc_section_match:
        calc_section = calc_section_match.group(1)
        cpc_calc_match = re.search(r'(?im)^\s*cpc\s*:\s*([^\n]+)$', calc_section)
        cpv_calc_match = re.search(r'(?im)^\s*cpv\s*:\s*([^\n]+)$', calc_section)
        if cpc_calc_match:
            data["cpc_calculation"] = cpc_calc_match.group(1).strip()
        if cpv_calc_match:
            data["cpv_calculation"] = cpv_calc_match.group(1).strip()

    # Extract Engagement Rate
    eng_rate_patterns = [
        r'engagement\s+rate[\s:]*(\d+\.?\d*)%?',
        r'engagement rate.*?(\d+\.?\d*)%',
    ]
    for pattern in eng_rate_patterns:
        eng_rate_match = re.search(pattern, text_lower)
        if eng_rate_match:
            data["engagement_rate"] = eng_rate_match.group(1)
            break

    # Calculate total engagement if not provided
    if data["total_engagement"] == "N/A":
        engagement_sum = 0
        count = 0
        for metric in ["likes", "shares", "saves"]:
            if data[metric] != "N/A":
                try:
                    engagement_sum += int(float(data[metric]))
                    count += 1
                except:
                    pass
        if count > 0:
            data["total_engagement"] = str(int(engagement_sum))

    print(f"[OK] Text metrics extracted: {data}")
    return data


def extract_prompt_context(prompt):
    """Extract campaign identity fields from the user's text prompt."""
    import re

    context = {
        "campaignName": "",
        "brand": "",
        "creator": "",
        "startDate": "",
        "endDate": "",
        "deliverables": "",
        "financial": {},
    }

    if not prompt:
        return context

    def first(patterns, text, flags=re.IGNORECASE):
        for pattern in patterns:
            match = re.search(pattern, text, flags)
            if match:
                return match.group(1).strip()
        return ""

    campaign_name = first([
        r'report\s+for\s+([^\n,\.]+?campaign)(?:\s+with\b|\s*,|\s*$)',
        r'campaign\s+name\s*[:\-]\s*([^\n,\.]+)',
        r'\bcampaign\s*:\s*([^\n,\.]+)',
        r'for\s+([^\n,\.]+?campaign)(?:\s+with\b|\s*,|\s*$)',
    ], prompt)
    if campaign_name:
        context["campaignName"] = campaign_name[:80]

    brand = first([r'\bbrand\s*[:\-]\s*([^\n,\.]+)'], prompt)
    if brand:
        context["brand"] = brand[:40]
    else:
        repeated_name_matches = re.findall(r'(?im)^\s*([A-Z][A-Za-z0-9&.+-]{2,})\s*$', prompt)
        repeated_names = []
        for name in repeated_name_matches:
            if repeated_name_matches.count(name) >= 2 and name.lower() not in {"cpc", "cpv", "budget", "campaign"}:
                repeated_names.append(name)
        if repeated_names:
            context["brand"] = repeated_names[0][:40]

    creator = first([
        r'(?:influencer|creator)\s*[:\-]\s*([^\n,\.]{2,})',
        r'with\s+((?:influencer|creator)\s+[^\n,\.]+?)(?:\s*,|\s+budget|\s+ran|\s+from|$)',
        r'with\s+(@[A-Za-z0-9_.]+)',
    ], prompt)
    if creator and len(creator.strip()) > 1:
        context["creator"] = creator[:60]

    range_match = re.search(
        r'ran\s+([A-Za-z]+)\s+(\d+)\s*(?:-|to|through|until)\s*(?:([A-Za-z]+)\s+)?(\d+)',
        prompt,
        re.IGNORECASE,
    )
    if range_match:
        month_start = range_match.group(1)
        day_start = range_match.group(2)
        month_end = range_match.group(3) or month_start
        day_end = range_match.group(4)
        context["startDate"] = f"{month_start} {day_start}"
        context["endDate"] = f"{month_end} {day_end}"
    else:
        start_date = first([r'(?:ran|from|start(?:ed)?)\s+([A-Za-z]+\s+\d+)'], prompt)
        end_date = first([r'(?:to|through|until|end(?:ed)?)\s+([A-Za-z]+\s+\d+)'], prompt)
        if start_date:
            context["startDate"] = start_date
        if end_date:
            context["endDate"] = end_date

    budget = first([
        r'budget\s*\$?([\d,]+(?:\.\d+)?)',
        r'\$([\d,]+(?:\.\d+)?)\s+budget',
    ], prompt)
    if budget:
        context["financial"]["totalBudget"] = budget.replace(",", "")
    if re.search(r'\b(inr|rs\.?|₹)\b', prompt, re.IGNORECASE):
        context["financial"]["budgetCurrency"] = "INR"
    elif "$" in prompt:
        context["financial"]["budgetCurrency"] = "USD"

    cpv = first([r'\bcpv\s*[:\-]?\s*\$?([\d.]+)'], prompt)
    if cpv:
        context["financial"]["cpv"] = cpv

    cpe = first([r'\bcpe\s*[:\-]?\s*\$?([\d.]+)'], prompt)
    if cpe:
        context["financial"]["cpe"] = cpe

    cpc = first([r'\bcpc\s*[:\-]?\s*\$?([\d.]+)'], prompt)
    if cpc:
        context["financial"]["cpc"] = cpc

    for metric_key, fin_key in [("cpc", "cpcGoal"), ("cpv", "cpvGoal")]:
        goal_match = re.search(rf'(?im)^\s*{metric_key}\s*:\s*([^\n]*[A-Za-z][^\n]*)$', prompt)
        if goal_match:
            context["financial"][fin_key] = goal_match.group(1).strip()

    calc_section_match = re.search(r'(?is)calculation\s*:\s*(.+)', prompt)
    if calc_section_match:
        calc_section = calc_section_match.group(1)
        for metric_key, fin_key in [("cpc", "cpcCalculation"), ("cpv", "cpvCalculation")]:
            calc_match = re.search(rf'(?im)^\s*{metric_key}\s*:\s*([^\n]+)$', calc_section)
            if calc_match:
                context["financial"][fin_key] = calc_match.group(1).strip()

    mentioned_platforms = []
    if re.search(r'\binstagram\b', prompt, re.IGNORECASE):
        mentioned_platforms.append("Instagram")
    if re.search(r'\byoutube\b', prompt, re.IGNORECASE):
        mentioned_platforms.append("YouTube")
    if mentioned_platforms:
        context["deliverables"] = " + ".join(mentioned_platforms)

    return context


def apply_prompt_context(extracted_data, prompt_context, force_identity=False):
    """Apply prompt-derived campaign context onto extracted data."""
    if not isinstance(extracted_data, dict):
        extracted_data = create_default_data()

    if not prompt_context:
        return extracted_data

    for key in ["campaignName", "creator", "startDate", "endDate"]:
        value = prompt_context.get(key)
        if value and (force_identity or not extracted_data.get(key)):
            extracted_data[key] = value

    # Brand should only come from the prompt when the prompt explicitly provides it.
    brand = prompt_context.get("brand")
    if brand and (force_identity or not extracted_data.get("brand") or extracted_data.get("brand") == "Unknown Brand"):
        extracted_data["brand"] = brand

    deliverables = prompt_context.get("deliverables")
    if deliverables and not extracted_data.get("deliverables"):
        extracted_data["deliverables"] = deliverables

    for fin_key in ["totalBudget", "budgetCurrency", "cpv", "cpe", "cpc", "cpcGoal", "cpvGoal", "cpcCalculation", "cpvCalculation", "roi"]:
        fin_value = prompt_context.get("financial", {}).get(fin_key)
        if fin_value and not extracted_data.get("financial", {}).get(fin_key):
            extracted_data.setdefault("financial", {})[fin_key] = fin_value

    return extracted_data


def has_meaningful_metrics(extracted_data):
    """Return True if any Instagram or YouTube metric fields contain values."""
    if not isinstance(extracted_data, dict):
        return False

    def _has_value(value):
        if value is None:
            return False
        if isinstance(value, str):
            return value.strip() not in ("", "0", "0.0", "0%")
        return bool(value)

    metric_fields = {
        "instagram": ["views", "likes", "comments", "shares", "saves", "reach", "impressions", "engagementRate"],
        "youtube": ["views", "likes", "comments", "shares", "watchTime", "ctr"],
    }
    for section, keys in metric_fields.items():
        values = extracted_data.get(section, {})
        if any(_has_value(values.get(key)) for key in keys):
            return True
    return False


def create_default_data():
    """Return PyroMedia style data structure (matches React component)"""
    return {
        "campaignName": "Campaign Report",
        "brand": "Unknown Brand",
        "agency": "PyroMedia",
        "agencyContact": "",
        "agencyWebsite": "",
        "creator": "",
        "startDate": "",
        "endDate": "",
        "deliverables": "",
        "postImage": "",
        "videoUrl": "",
        "instagram": {
            "views": "",
            "likes": "",
            "comments": "",
            "shares": "",
            "saves": "",
            "reach": "",
            "impressions": "",
            "engagementRate": ""
        },
        "youtube": {
            "views": "",
            "likes": "",
            "comments": "",
            "shares": "",
            "watchTime": "",
            "ctr": ""
        },
        "financial": {
            "totalBudget": "",
            "budgetCurrency": "",
            "cpv": "",
            "cpe": "",
            "cpc": "",
            "cpcGoal": "",
            "cpvGoal": "",
            "cpcCalculation": "",
            "cpvCalculation": "",
            "roi": ""
        },
        "performance": {
            "totalViews": "",
            "totalLikes": "",
            "totalComments": "",
            "totalShares": "",
            "totalSaves": "",
            "totalReach": "",
            "totalClicks": "",
            "totalInteractions": "",
            "totalEngagement": "",
            "keyLearnings": "Campaign data extraction complete."
        },
        "creators": []
    }


def validate_and_clean_data(data):
    """Ensure extracted data has all required fields and proper structure"""
    if not isinstance(data, dict):
        return create_default_data()

    defaults = create_default_data()
    cleaned_data = data.copy()

    # Validate and initialize nested objects
    for key in ["instagram", "youtube", "financial", "performance"]:
        if key not in cleaned_data or not isinstance(cleaned_data[key], dict):
            cleaned_data[key] = defaults[key].copy()
        else:
            # Ensure all sub-keys from defaults exist and convert None values
            for sub_key in defaults[key]:
                if sub_key not in cleaned_data[key]:
                    cleaned_data[key][sub_key] = defaults[key][sub_key]
                elif cleaned_data[key][sub_key] is None:
                    cleaned_data[key][sub_key] = ""

    # Ensure creators is a list
    if not isinstance(cleaned_data.get("creators"), list):
        cleaned_data["creators"] = []

    # Ensure top-level fields
    for key in ["campaignName", "brand", "agency", "creator", "startDate", "endDate", "deliverables", "agencyContact", "agencyWebsite"]:
        if key not in cleaned_data:
            cleaned_data[key] = defaults.get(key, "")
        # Convert None to empty string
        if cleaned_data[key] is None:
            cleaned_data[key] = ""

    return cleaned_data


def calculate_metrics(data):
    """Calculate CPV and CPE if budget and views/engagement are available"""
    
    def safe_float(value):
        """Convert string to float, handling various formats"""
        if value == "N/A" or value is None or value == "":
            return None
        try:
            # Remove commas and convert to float
            return float(str(value).replace(',', '').strip())
        except (ValueError, AttributeError):
            return None
    
    def format_number(value):
        """Format number with 2 decimal places"""
        if value is None:
            return "N/A"
        return f"{value:.2f}"
    
    def format_integer(value):
        """Format number as integer"""
        if value is None:
            return "N/A"
        return str(int(value))
    
    # Calculate for overall campaign
    overall = data.get('overall_campaign', {})
    budget = safe_float(overall.get('budget'))
    views = safe_float(overall.get('views'))
    likes = safe_float(overall.get('likes'))
    shares = safe_float(overall.get('shares'))
    saves = safe_float(overall.get('saves'))
    engagement = safe_float(overall.get('total_engagement'))
    
    # Calculate total engagement if it's N/A but we have individual metrics
    if engagement is None:
        engagement_sum = 0
        count = 0
        for metric in [likes, shares, saves]:
            if metric is not None:
                engagement_sum += metric
                count += 1
        
        if count > 0:
            engagement = engagement_sum
            overall['total_engagement'] = format_integer(engagement)
    
    # Calculate CPV
    if budget and views and views > 0:
        cpv = budget / views
        overall['cpv'] = format_number(cpv)
    
    # Calculate CPE
    if budget and engagement and engagement > 0:
        cpe = budget / engagement
        overall['cpe'] = format_number(cpe)
    
    # Calculate for creator data
    creator = data.get('creator_data', {})
    
    # If creator data is N/A or not provided, copy from overall campaign
    creator_views = safe_float(creator.get('views'))
    creator_likes = safe_float(creator.get('likes'))
    creator_shares = safe_float(creator.get('shares'))
    creator_saves = safe_float(creator.get('saves'))
    creator_engagement = safe_float(creator.get('total_engagement'))
    creator_budget = safe_float(creator.get('budget'))
    
    # Copy overall data to creator if creator data is N/A
    if creator_views is None and views is not None:
        creator['views'] = overall.get('views')
        creator_views = views
    
    if creator_likes is None and likes is not None:
        creator['likes'] = overall.get('likes')
        creator_likes = likes
    
    if creator_shares is None and shares is not None:
        creator['shares'] = overall.get('shares')
        creator_shares = shares
    
    if creator_saves is None and saves is not None:
        creator['saves'] = overall.get('saves')
        creator_saves = saves
    
    if creator_budget is None and budget is not None:
        creator['budget'] = overall.get('budget')
        creator_budget = budget
    
    # Calculate creator total engagement if it's N/A
    if creator_engagement is None:
        engagement_sum = 0
        count = 0
        for metric in [creator_likes, creator_shares, creator_saves]:
            if metric is not None:
                engagement_sum += metric
                count += 1
        
        if count > 0:
            creator_engagement = engagement_sum
            creator['total_engagement'] = format_integer(creator_engagement)
    
    # Calculate creator CPV
    if creator_budget and creator_views and creator_views > 0:
        creator_cpv = creator_budget / creator_views
        creator['cpv'] = format_number(creator_cpv)
    
    # Calculate creator CPE
    if creator_budget and creator_engagement and creator_engagement > 0:
        creator_cpe = creator_budget / creator_engagement
        creator['cpe'] = format_number(creator_cpe)
    
    return data


def analyze_images_with_gpt4(images, user_prompt):
    """Analyze campaign screenshots using Claude Vision or OpenRouter"""
    api_name = "OpenRouter" if USE_OPENROUTER else "Claude"
    print(f"[IMAGE] Analyzing {len(images)} images with {api_name} Vision...")

    # Prepare images in the correct format
    image_contents = []
    for image in images:
        image_bytes = image.read()
        image.seek(0)  # Reset file pointer for later use
        image_b64 = base64.b64encode(image_bytes).decode('utf-8')
        media_type = detect_image_media_type(image, image_bytes)

        if USE_OPENROUTER:
            # OpenAI-compatible format for OpenRouter
            image_contents.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:{media_type};base64,{image_b64}"
                }
            })
        else:
            # Anthropic format
            image_contents.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": media_type,
                    "data": image_b64
                }
            })
    
    # Build a clear extraction prompt
    analysis_prompt = f"""You are a precise data extraction assistant. Your job is to read EXACT numbers visible in these screenshots.

Campaign context from user: "{user_prompt}"

CRITICAL RULES:
1. Read the EXACT numbers shown on screen — do NOT estimate or calculate averages
2. These are screenshots of a specific Instagram post/Reel or YouTube video — extract the EXACT metrics shown
3. Numbers may appear as "3,245" or "3.2K" or "3245" — convert all to plain integers (e.g. "3245")
4. Look for numbers next to labels: Views, Plays, Likes, Comments, Shares, Saves, Reach, Impressions, Followers, Engagement Rate, Watch Time, CTR
5. If a number is NOT visible in any image, use "" — never guess

WHAT TO LOOK FOR IN INSTAGRAM SCREENSHOTS:
- Post/Reel plays or views count (shown as "X plays" or "X views")
- Likes count (heart icon number)
- Comments count (speech bubble number)
- Shares count (arrow icon number)
- Saves count (bookmark icon number)
- Reach (unique accounts reached)
- Impressions (total times shown)
- Engagement rate (shown as X%)
- Account followers count
- Username / handle (@username)

WHAT TO LOOK FOR IN YOUTUBE SCREENSHOTS:
- View count
- Like count
- Comment count
- Watch time (hours/minutes)
- CTR (click-through rate %)
- Subscribers count

Return ONLY valid JSON in exactly this format — no markdown, no extra text:
{{
  "campaignName": "exact campaign name from image or user context",
  "brand": "brand name visible in image",
  "agency": "agency name if visible",
  "creator": "exact @username or creator name from image",
  "startDate": "date if visible",
  "endDate": "date if visible",
  "deliverables": "type of content eg Instagram Reel, YouTube Video",
  "instagram": {{
    "views": "EXACT number from image",
    "likes": "EXACT number from image",
    "comments": "EXACT number from image",
    "shares": "EXACT number from image",
    "saves": "EXACT number from image",
    "reach": "EXACT number from image",
    "impressions": "EXACT number from image",
    "engagementRate": "EXACT percentage from image eg 1.82%"
  }},
  "youtube": {{
    "views": "EXACT number from image",
    "likes": "EXACT number from image",
    "comments": "EXACT number from image",
    "shares": "EXACT number from image",
    "watchTime": "EXACT watch time from image",
    "ctr": "EXACT CTR percentage from image"
  }},
  "financial": {{
    "totalBudget": "from user text if provided",
    "cpv": "if visible",
    "cpe": "if visible",
    "cpc": "if visible",
    "cpcGoal": "goal/intent text for CPC if visible",
    "cpvGoal": "goal/intent text for CPV if visible",
    "cpcCalculation": "CPC calculation text if visible",
    "cpvCalculation": "CPV calculation text if visible",
    "roi": "if visible"
  }},
  "performance": {{
    "totalViews": "sum of all views across all posts/images",
    "totalLikes": "sum of all likes",
    "totalComments": "sum of all comments",
    "totalShares": "sum of all shares",
    "totalSaves": "sum of all saves",
    "totalReach": "total reach",
    "totalInteractions": "likes + comments + shares + saves",
    "totalEngagement": "engagement rate or total engagement number",
    "keyLearnings": "2 sentences summarising what the campaign achieved based on the numbers"
  }},
  "creators": [
    {{
      "name": "creator name or @username",
      "platform": "Instagram or YouTube",
      "views": "EXACT number",
      "likes": "EXACT number",
      "comments": "EXACT number",
      "shares": "EXACT number",
      "saves": "EXACT number",
      "reach": "EXACT number",
      "engagementRate": "EXACT percentage",
      "watchTime": "if YouTube",
      "interactions": "likes + comments + shares + saves"
    }}
  ]
}}"""

    # Create messages for Claude API
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": analysis_prompt},
                *image_contents
            ]
        }
    ]

    try:
        # Call Vision API (Claude or OpenRouter)
        model_id = os.environ.get("MODEL_ID", "claude-3-5-sonnet-20241022")

        # Fallback model list in case primary model fails
        FALLBACK_MODELS = [
            "meta-llama/llama-3.2-11b-vision-instruct:free",
            "qwen/qwen-2-vl-7b-instruct:free",
            "nvidia/llama-3.1-nemotron-nano-8b-v1:free",
            "nvidia/nemotron-nano-12b-v2-vl",
        ]

        if USE_OPENROUTER:
            # Try primary model, then fallbacks if it fails
            models_to_try = [model_id] + [m for m in FALLBACK_MODELS if m != model_id]
            response = None
            used_model = None

            for try_model in models_to_try:
                try:
                    print(f"[NOTE] Trying model: {try_model}")
                    response = client.chat.completions.create(
                        model=try_model,
                        messages=[
                            {
                                "role": "system",
                                "content": "You are a data extraction assistant. Read text and numbers from images. Return ONLY valid JSON, no markdown."
                            },
                            {
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": analysis_prompt},
                                    *image_contents
                                ]
                            }
                        ],
                        max_tokens=4096,
                        temperature=0.1
                    )
                    used_model = try_model
                    print(f"[OK] Using model: {used_model}")
                    break
                except Exception as model_err:
                    print(f"[WARNING] Model {try_model} failed: {str(model_err)[:100]}")
                    continue

            if response is None:
                print("[ERROR] All models failed")
                return create_default_data()

            # Handle both regular responses and reasoning-based responses
            message = response.choices[0].message
            response_text = message.content

            # If content is None but reasoning exists, try to extract JSON from reasoning
            if not response_text and hasattr(message, 'reasoning') and message.reasoning:
                print(f"[NOTE] Extracting JSON from reasoning field...")
                response_text = message.reasoning
                if isinstance(response_text, list) and len(response_text) > 0:
                    response_text = response_text[0].get('text', '') if isinstance(response_text[0], dict) else str(response_text[0])
        else:
            # Anthropic API call
            response = client.messages.create(
                model=model_id,
                max_tokens=2000,
                messages=messages,
                system="You are a data extraction assistant specialized in reading text and numbers from marketing analytics screenshots. You extract only visible text and metrics, never identifying people."
            )
            response_text = response.content[0].text

        # Verify response is not None
        if not response_text:
            print(f"[ERROR] {api_name} returned empty response")
            print(f"Full response object: {response}")
            return create_default_data()

        print(f"[OK] {api_name} Response received ({len(response_text)} chars total)")
        print(f"[DEBUG] Full response:\n{response_text}\n[/DEBUG]")
        
        # Clean up response (remove markdown code blocks if present)
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0]
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0]

        # Extract JSON from response - find first { to last }
        response_text = response_text.strip()
        start = response_text.find('{')

        if start == -1:
            print("[ERROR] No JSON found in response")
            print(f"Full response (first 500 chars): {response_text[:500]}")
            return create_default_data()

        # Try to find the closing brace, accounting for nested structures
        brace_count = 0
        end = start
        for i in range(start, len(response_text)):
            if response_text[i] == '{':
                brace_count += 1
            elif response_text[i] == '}':
                brace_count -= 1
                if brace_count == 0:
                    end = i + 1
                    break

        # Handle truncated JSON (no matching closing brace)
        if end == start:
            print("[WARNING] JSON appears truncated - attempting repair...")
            json_str = response_text[start:]
            open_braces = json_str.count('{') - json_str.count('}')
            open_brackets = json_str.count('[') - json_str.count(']')
            json_str = json_str.rstrip(',').rstrip('"').rstrip(' ')
            # Close any open string by checking if last char is inside a string
            if open_brackets > 0:
                json_str += '"' + ']' * open_brackets
            json_str += '}' * open_braces
            print(f"[NOTE] Added {open_braces} closing braces, {open_brackets} brackets")
        else:
            json_str = response_text[start:end]

        print(f"[DATA] Extracted JSON ({len(json_str)} chars)")

        try:
            extracted_data = json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"[ERROR] JSON Parse Error (attempt 1): {str(e)}")
            # Try to fix common JSON issues (escape sequences, newlines)
            json_str_fixed = json_str.replace('\\n', ' ').replace('\n', ' ')
            try:
                extracted_data = json.loads(json_str_fixed)
                print("[OK] Fixed JSON with escape sequence cleanup")
            except:
                # Last resort: extract key fields using regex
                print("[NOTE] Attempting regex field extraction from partial JSON...")
                import re
                extracted_data = create_default_data()
                # Extract visible string fields from partial JSON
                for field in ['campaignName', 'brand', 'agency', 'creator', 'startDate', 'endDate', 'deliverables']:
                    m = re.search(rf'"{field}"\s*:\s*"([^"]*)"', json_str)
                    if m and m.group(1):
                        extracted_data[field] = m.group(1)
                # Extract instagram metrics
                for field in ['views', 'likes', 'comments', 'shares', 'saves', 'reach', 'impressions', 'engagementRate']:
                    m = re.search(rf'(?:instagram[^}}]+)?"{field}"\s*:\s*"([^"]*)"', json_str)
                    if m and m.group(1):
                        extracted_data['instagram'][field] = m.group(1)
                print(f"[NOTE] Regex extraction found: campaignName={extracted_data.get('campaignName')}, brand={extracted_data.get('brand')}")
                if extracted_data.get('brand') == 'Unknown Brand':
                    return create_default_data()
        
        # Strip "0" string values from Vision AI output — these mean "not found", not actual zeros
        def _strip_zeros(d):
            if not isinstance(d, dict):
                return
            for k, v in list(d.items()):
                if isinstance(v, str) and v.strip() in ("0", "0.0", "0%"):
                    d[k] = ""
                elif isinstance(v, dict):
                    _strip_zeros(v)
                elif isinstance(v, list):
                    for item in v:
                        _strip_zeros(item)

        _strip_zeros(extracted_data)
        print("[NOTE] Stripped '0' placeholder values from Vision AI output")

        # Fill missing top-level fields from user prompt text
        import re as _re

        def _first(patterns, text, flags=_re.IGNORECASE):
            for p in patterns:
                m = _re.search(p, text, flags)
                if m:
                    return m.group(1).strip()
            return ""

        if not extracted_data.get("campaignName") or extracted_data["campaignName"] in ("", "Campaign Report"):
            val = _first([
                # explicit label: "campaign name: X"
                r'campaign\s+name\s*[:\-]\s*([^\n,\.]+)',
                # explicit label: "campaign: X"
                r'\bcampaign\s*:\s*([^\n,\.]+)',
                # "for X campaign" — capture the descriptive words BEFORE "campaign"
                r'for\s+((?:\w+\s+){1,5}?campaign)\b',
                # "for X campaign with ..." — same but with trailing context
                r'report\s+for\s+([^\n,\.]+?campaign[^\n,\.]*?)(?:\s+with\b|\s*,|\s*$)',
            ], user_prompt)
            if val:
                extracted_data["campaignName"] = val.strip()[:60]

        if not extracted_data.get("brand") or extracted_data["brand"] in ("", "Unknown Brand"):
            val = _first([r'\bbrand[:\s]+([^\n,\.]+)'], user_prompt)
            if val:
                extracted_data["brand"] = val[:40]

        if not extracted_data.get("creator"):
            val = _first([
                # explicit labels
                r'influencer\s*[:\-]\s*([^\n,\.]+)',
                r'creator\s*[:\-]\s*([^\n,\.]+)',
                # "with Influencer X" or "with Creator X" — name follows the keyword
                r'with\s+((?:influencer|creator)\s+\S+(?:\s+\S+)?)',
                # "with @handle"
                r'with\s+(@\w+)',
            ], user_prompt)
            if val and len(val.strip()) > 1:
                extracted_data["creator"] = val.strip()[:40]

        if not extracted_data.get("startDate"):
            val = _first([r'(?:ran|from|start(?:ed)?)[:\s]+([A-Za-z]+\s+\d+)'], user_prompt)
            if val:
                extracted_data["startDate"] = val

        if not extracted_data.get("endDate"):
            val = _first([r'(?:to|-|through|until|end(?:ed)?)[:\s]+([A-Za-z]+\s+\d+)'], user_prompt)
            if val:
                extracted_data["endDate"] = val

        # Ensure creators array is populated
        if not extracted_data.get("creators") or len(extracted_data["creators"]) == 0:
            print("[NOTE] Populating creators array from extracted data...")
            creator_name = extracted_data.get("creator", "Creator")
            ig = extracted_data.get("instagram", {})
            yt = extracted_data.get("youtube", {})

            extracted_data["creators"] = [
                {
                    "name": creator_name,
                    "platform": "Instagram" if ig.get("views") else "YouTube",
                    "views": ig.get("views", ""),
                    "likes": ig.get("likes", ""),
                    "comments": ig.get("comments", ""),
                    "shares": ig.get("shares", ""),
                    "saves": ig.get("saves", ""),
                    "reach": ig.get("reach", ""),
                    "engagementRate": ig.get("engagementRate", ""),
                    "watchTime": yt.get("watchTime", ""),
                    "interactions": extracted_data.get("performance", {}).get("totalInteractions", ""),
                }
            ]

        print("[OK] Successfully extracted data from images")
        print(f"[DEBUG] Parsed data: campaignName={extracted_data.get('campaignName')}, brand={extracted_data.get('brand')}, ig_views={extracted_data.get('instagram',{}).get('views')}, creators={len(extracted_data.get('creators',[]))}")
        # Ensure extracted data is properly validated
        return validate_and_clean_data(extracted_data)

    except json.JSONDecodeError as e:
        print(f"[ERROR] JSON Parse Error: {str(e)}")
        print(f"Attempted to parse: {json_str if 'json_str' in locals() else response_text}")
        return create_default_data()

    except Exception as e:
        print(f"[ERROR] Error analyzing images with {api_name}: {str(e)}")
        import traceback
        traceback.print_exc()
        return create_default_data()


def populate_powerpoint(data, images):
    """Populate PowerPoint template with extracted data and images"""
    print("[FILE] Creating PowerPoint presentation...")
    
    try:
        # Load template
        prs = Presentation(TEMPLATE_PATH)
        print(f"[OK] Template loaded: {len(prs.slides)} slides")

        def _download_image_stream(url):
            if not url:
                return None
            try:
                import requests
                headers = {
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/123.0.0.0 Safari/537.36"
                    ),
                    "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
                    "Referer": "https://www.instagram.com/",
                }
                resp = requests.get(url, headers=headers, timeout=30)
                resp.raise_for_status()
                content = resp.content
                content_type = (resp.headers.get("Content-Type") or "").lower()

                # PowerPoint is happiest with PNG/JPEG; convert unsupported formats like WEBP.
                try:
                    img = Image.open(io.BytesIO(content))
                    out = io.BytesIO()
                    save_format = "PNG"
                    if img.mode in ("RGBA", "LA", "P"):
                        img = img.convert("RGBA")
                    else:
                        img = img.convert("RGB")
                        save_format = "JPEG"
                    img.save(out, format=save_format)
                    out.seek(0)
                    print(f"[OK] Downloaded fallback image for PPT ({content_type or 'unknown'})")
                    return out
                except Exception:
                    raw = io.BytesIO(content)
                    raw.seek(0)
                    print(f"[OK] Downloaded raw fallback image for PPT ({content_type or 'unknown'})")
                    return raw
            except Exception as e:
                print(f"[WARNING] Could not download image from URL: {e}")
                return None

        def _uploaded_image_stream(image_obj):
            try:
                img_bytes = image_obj.read()
                image_obj.seek(0)
                return io.BytesIO(img_bytes)
            except Exception as e:
                print(f"[WARNING] Could not read uploaded image: {e}")
                return None

        def _add_picture_with_fallback(slide, img_stream, placeholder=None, default_box=None, rotation=0):
            if not img_stream:
                return False
            try:
                img_stream.seek(0)
                if placeholder:
                    picture = slide.shapes.add_picture(
                        img_stream,
                        placeholder['left'],
                        placeholder['top'],
                        width=placeholder['width'],
                        height=placeholder['height']
                    )
                    picture.rotation = rotation or placeholder.get('rotation', 0)
                    return True
                if default_box:
                    picture = slide.shapes.add_picture(
                        img_stream,
                        default_box['left'],
                        default_box['top'],
                        width=default_box['width'],
                        height=default_box['height']
                    )
                    picture.rotation = default_box.get('rotation', 0)
                    return True
            except Exception as e:
                print(f"[WARNING] Could not add fallback picture to slide: {e}")
            return False

        def _first_present(*values):
            for value in values:
                if value is None:
                    continue
                if isinstance(value, str):
                    if not value.strip():
                        continue
                    return value
                return value
            return ""

        def _prefer_interactions(total_engagement_value, total_interactions_value):
            interaction_value = _first_present(total_interactions_value)
            engagement_value = _first_present(total_engagement_value)
            if interaction_value not in ("", None):
                return interaction_value
            return engagement_value

        image_classifications = data.get('image_classifications', []) or []
        type_to_indices = {}
        for classification in image_classifications:
            idx = classification.get('image_index')
            img_type = classification.get('type')
            if isinstance(idx, int) and img_type:
                type_to_indices.setdefault(img_type, []).append(idx)

        used_uploaded_indices = set()

        def _select_uploaded_streams(preferred_types=None, limit=1, fallback_any=False):
            selected = []
            selected_indices = []
            preferred_types = preferred_types or []

            for img_type in preferred_types:
                for idx in type_to_indices.get(img_type, []):
                    if idx in used_uploaded_indices or idx >= len(images):
                        continue
                    stream = _uploaded_image_stream(images[idx])
                    if stream:
                        selected.append(stream)
                        selected_indices.append(idx)
                        used_uploaded_indices.add(idx)
                        if len(selected) >= limit:
                            return selected, selected_indices

            if fallback_any and len(selected) < limit:
                for idx, image_obj in enumerate(images):
                    if idx in used_uploaded_indices:
                        continue
                    stream = _uploaded_image_stream(image_obj)
                    if stream:
                        selected.append(stream)
                        selected_indices.append(idx)
                        used_uploaded_indices.add(idx)
                        if len(selected) >= limit:
                            break

            return selected, selected_indices
        
        # SLIDE 1: Insert Brand Logo
        if len(prs.slides) > 0:
            slide_1 = prs.slides[0]
            
            # Find and replace "BRAND LOGO HERE" text box with brand logo image
            logo_images = [
                img for idx, img in enumerate(images) 
                if any(c.get('image_index') == idx and c.get('type') == 'brand_logo' 
                      for c in data.get('image_classifications', []))
            ]
            
            if logo_images:
                print("[DESIGN] Adding brand logo to Slide 1")
                
                # Find the "BRAND LOGO HERE" text box and get its position
                logo_placeholder = None
                for shape in slide_1.shapes:
                    if hasattr(shape, 'text') and 'BRAND LOGO HERE' in shape.text.upper():
                        logo_placeholder = shape
                        break
                
                if logo_placeholder:
                    # Get position and size from placeholder
                    left = logo_placeholder.left
                    top = logo_placeholder.top
                    width = logo_placeholder.width
                    height = logo_placeholder.height
                    
                    # Delete the placeholder text box
                    sp = logo_placeholder.element
                    sp.getparent().remove(sp)
                    
                    # Insert brand logo image at the same position
                    img_bytes = logo_images[0].read()
                    logo_images[0].seek(0)
                    
                    with io.BytesIO(img_bytes) as img_stream:
                        slide_1.shapes.add_picture(
                            img_stream,
                            left,
                            top,
                            width=width,
                            height=height
                        )
                    
                    print("[OK] Brand logo inserted successfully on Slide 1")
                else:
                    print("[WARNING]  'BRAND LOGO HERE' placeholder not found on Slide 1")
            else:
                print("[WARNING]  No brand logo image classified in uploaded images")
        
        # SLIDE 3: Campaign Name with proper styling
        if len(prs.slides) > 2:
            slide_3 = prs.slides[2]
            
            for shape in slide_3.shapes:
                if hasattr(shape, 'text') and '[Campaign Name Here]' in shape.text:
                    if hasattr(shape, 'text_frame'):
                        text_frame = shape.text_frame
                        text_frame.clear()  # Clear existing text
                        
                        # Add campaign name with proper formatting
                        p = text_frame.paragraphs[0]
                        campaign_name = data.get('campaignName') or data.get('campaign_name') or 'Campaign Report'
                        p.text = campaign_name
                        p.alignment = PP_ALIGN.CENTER

                        # Set font properties
                        for run in p.runs:
                            run.font.name = 'YouTube Sans'
                            run.font.size = Pt(53)
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0x2B, 0x3E, 0x5C)  # #2b3e5c

                        print(f"[OK] Updated Slide 3 with styled campaign name: {campaign_name}")
        
        # SLIDE 4: Overall Campaign Report with bold data
        if len(prs.slides) > 3:
            slide_4 = prs.slides[3]
            # Support both old and new data formats
            overall = data.get('overall_campaign', {})
            instagram = data.get('instagram', {})
            financial = data.get('financial', {})
            performance = data.get('performance', {})
            budget_display = financial.get('totalBudget') or overall.get('budget') or ''
            if budget_display and financial.get('budgetCurrency'):
                budget_display = f"{budget_display} {financial.get('budgetCurrency')}"

            for shape in slide_4.shapes:
                if hasattr(shape, 'text_frame') and hasattr(shape, 'text') and 'No. of Views:' in shape.text:
                    text_frame = shape.text_frame
                    text_frame.clear()

                    # Define metrics with labels and values (support both old and new formats)
                    metrics = [
                        ("No. of Views: ", _first_present(performance.get('totalViews'), instagram.get('views'), overall.get('views'))),
                        ("Likes: ", _first_present(performance.get('totalLikes'), instagram.get('likes'), overall.get('likes'))),
                        ("Shares: ", _first_present(performance.get('totalShares'), instagram.get('shares'), overall.get('shares'))),
                        ("Saves: ", _first_present(performance.get('totalSaves'), instagram.get('saves'), overall.get('saves'))),
                        ("Total Engagement: ", _prefer_interactions(performance.get('totalEngagement'), performance.get('totalInteractions')) or overall.get('total_engagement') or ''),
                        ("Budget: ", budget_display),
                        ("CPC: ", _first_present(financial.get('cpc'), overall.get('cpc'))),
                        ("CPV: ", _first_present(financial.get('cpv'), overall.get('cpv'))),
                        ("CPE: ", _first_present(financial.get('cpe'), overall.get('cpe')))
                    ]
                    
                    # Add each metric with bold labels and normal values
                    for i, (label, value) in enumerate(metrics):
                        if i == 0:
                            p = text_frame.paragraphs[0]
                        else:
                            p = text_frame.add_paragraph()
                        
                        # Add label (bold)
                        run = p.add_run()
                        run.text = label
                        run.font.bold = True
                        
                        # Add value (normal)
                        run = p.add_run()
                        run.text = str(value)
                        run.font.bold = False

                    extras = [
                        ("CPC Goal: ", financial.get('cpcGoal') or ''),
                        ("CPV Goal: ", financial.get('cpvGoal') or ''),
                        ("CPC Calculation: ", financial.get('cpcCalculation') or ''),
                        ("CPV Calculation: ", financial.get('cpvCalculation') or ''),
                    ]
                    for label, value in extras:
                        if value:
                            p = text_frame.add_paragraph()
                            run = p.add_run()
                            run.text = label
                            run.font.bold = True
                            run = p.add_run()
                            run.text = str(value)
                            run.font.bold = False
                    
                    # Add learnings (keep only first 2 sentences)
                    learnings_text = overall.get('learnings', 'No learnings available.')
                    # Split by periods and keep first 2 sentences
                    sentences = [s.strip() + '.' for s in learnings_text.split('.') if s.strip()]
                    learnings_text = ' '.join(sentences[:2]) if sentences else learnings_text
                    
                    # Add empty line before learnings
                    p = text_frame.add_paragraph()
                    p.text = ""
                    
                    # Add learnings paragraph (normal text, not bold)
                    p = text_frame.add_paragraph()
                    
                    # Add "Learnings:" label (bold)
                    run = p.add_run()
                    run.text = "Learnings: "
                    run.font.bold = True
                    
                    # Add learnings text (normal, not bold)
                    run = p.add_run()
                    run.text = learnings_text
                    run.font.bold = False
                    
                    print("[OK] Updated Slide 4 with bold formatting for overall campaign metrics")
                    break
            
            # Add campaign photos (up to 3) - Replace template placeholders with rotation
            campaign_image_streams, _ = _select_uploaded_streams(
                preferred_types=['campaign_photo'],
                limit=3,
                fallback_any=False,
            )
            if not campaign_image_streams:
                fallback_campaign_image = _download_image_stream(data.get("postImage"))
                if fallback_campaign_image:
                    campaign_image_streams.append(fallback_campaign_image)
            
            if campaign_image_streams:
                print(f"[IMAGE] Adding {len(campaign_image_streams[:3])} campaign photos to Slide 4")
                
                # Find and delete text placeholder "Campaign Photos Here at Least 3 photos"
                text_placeholder = None
                shapes_to_check = list(slide_4.shapes)  # Create a list to avoid iteration issues
                for shape in shapes_to_check:
                    if hasattr(shape, 'text') and 'Campaign Photos Here' in shape.text:
                        text_placeholder = shape
                        break
                
                if text_placeholder:
                    sp = text_placeholder.element
                    sp.getparent().remove(sp)
                    print("[OK] Removed 'Campaign Photos Here' text placeholder")
                
                # Find AUTO_SHAPE placeholders (shapes for 3 photos with rotation)
                # Collect ALL placeholder info BEFORE deleting any
                shape_placeholders = []
                shapes_to_check = list(slide_4.shapes)  # Fresh list after text deletion
                for shape in shapes_to_check:
                    if shape.shape_type == 1:  # AUTO_SHAPE
                        # Check if it's in the right area (right side of slide)
                        if shape.left > 3000000 and shape.top > 1000000 and shape.top < 4000000:
                            shape_placeholders.append({
                                'shape': shape,
                                'left': shape.left,
                                'top': shape.top,
                                'width': shape.width,
                                'height': shape.height,
                                'rotation': shape.rotation
                            })
                
                # Sort by left position to maintain order (left to right)
                shape_placeholders.sort(key=lambda x: x['left'])
                
                print(f"Found {len(shape_placeholders)} photo placeholders")
                
                # Now delete all placeholder shapes
                for placeholder in shape_placeholders:
                    sp = placeholder['shape'].element
                    sp.getparent().remove(sp)
                
                # Now add all images with their rotation
                for idx, (img_stream, placeholder) in enumerate(zip(campaign_image_streams[:3], shape_placeholders[:3])):
                    _add_picture_with_fallback(slide_4, img_stream, placeholder=placeholder, rotation=placeholder['rotation'])
                    
                    print(f"[OK] Added campaign photo {idx + 1} with rotation {placeholder['rotation']:.2f}°")

                if not shape_placeholders and campaign_image_streams:
                    default_campaign_box = {
                        'left': 3900000,
                        'top': 1000000,
                        'width': 5100000,
                        'height': 3100000,
                        'rotation': 0,
                    }
                    if _add_picture_with_fallback(slide_4, campaign_image_streams[0], default_box=default_campaign_box):
                        print("[OK] Added fallback campaign image using default Slide 4 bounds")
        
        # SLIDE 5: Creator x Brand with bold data
        if len(prs.slides) > 4:
            slide_5 = prs.slides[4]
            # Support both old and new data formats
            creator = data.get('creator_data', {})
            creators = data.get('creators', [])
            # Get first creator from array if available, otherwise use creator_data
            if creators and len(creators) > 0:
                creator = creators[0]
                creator_name = creator.get('name', 'Creator')
            else:
                creator_name = creator.get('creator_name', data.get('creator', 'Creator'))
            creator_budget_display = creator.get('budget') or data.get('financial', {}).get('totalBudget') or ''
            budget_currency = data.get('financial', {}).get('budgetCurrency') or ''
            if creator_budget_display and budget_currency:
                creator_budget_display = f"{creator_budget_display} {budget_currency}"

            brand_name = data.get('brand') or data.get('brand_name', 'Brand')
            creator_title = f"{creator_name} x {brand_name}"

            for shape in slide_5.shapes:
                if hasattr(shape, 'text') and 'Creator x Brand Name' in shape.text:
                    if hasattr(shape, 'text_frame'):
                        text_frame = shape.text_frame
                        text_frame.clear()

                        # Add title with proper formatting
                        p = text_frame.paragraphs[0]
                        p.text = creator_title

                        # Set font properties
                        for run in p.runs:
                            run.font.name = 'YouTube Sans'
                            run.font.size = Pt(25)
                            run.font.bold = True
                            run.font.italic = True
                            run.font.color.rgb = RGBColor(0x2B, 0x3E, 0x5C)  # #2b3e5c

                        print(f"[OK] Updated Slide 5 title with styling: {creator_title}")
                    else:
                        shape.text = creator_title
                        print(f"[OK] Updated Slide 5 title: {creator_title}")

                if hasattr(shape, 'text_frame') and hasattr(shape, 'text') and 'No. of Views:' in shape.text and 'Learnings:' in shape.text:
                    text_frame = shape.text_frame
                    text_frame.clear()

                    # Define metrics with labels and values (support both old and new formats)
                    financial_data = data.get('financial', {})
                    metrics = [
                        ("No. of Views: ", _first_present(creator.get('views'))),
                        ("Likes: ", _first_present(creator.get('likes'))),
                        ("Shares: ", _first_present(creator.get('shares'))),
                        ("Saves: ", _first_present(creator.get('saves'))),
                        ("Total Engagement: ", _prefer_interactions(creator.get('total_engagement'), creator.get('interactions'))),
                        ("Budget: ", creator_budget_display),
                        ("CPC: ", _first_present(creator.get('cpc'), financial_data.get('cpc'))),
                        ("CPV: ", _first_present(creator.get('cpv'), financial_data.get('cpv'))),
                        ("CPE: ", _first_present(creator.get('cpe'), financial_data.get('cpe')))
                    ]
                    
                    # Add each metric with bold labels and normal values
                    for i, (label, value) in enumerate(metrics):
                        if i == 0:
                            p = text_frame.paragraphs[0]
                        else:
                            p = text_frame.add_paragraph()
                        
                        # Add label (bold)
                        run = p.add_run()
                        run.text = label
                        run.font.bold = True
                        
                        # Add value (normal)
                        run = p.add_run()
                        run.text = str(value)
                        run.font.bold = False

                    extras = [
                        ("CPC Goal: ", creator.get('cpcGoal') or data.get('financial', {}).get('cpcGoal') or ''),
                        ("CPV Goal: ", creator.get('cpvGoal') or data.get('financial', {}).get('cpvGoal') or ''),
                        ("CPC Calculation: ", creator.get('cpcCalculation') or data.get('financial', {}).get('cpcCalculation') or ''),
                        ("CPV Calculation: ", creator.get('cpvCalculation') or data.get('financial', {}).get('cpvCalculation') or ''),
                    ]
                    for label, value in extras:
                        if value:
                            p = text_frame.add_paragraph()
                            run = p.add_run()
                            run.text = label
                            run.font.bold = True
                            run = p.add_run()
                            run.text = str(value)
                            run.font.bold = False
                    
                    # Add learnings (keep only first 2 sentences)
                    learnings_text = creator.get('learnings', 'No learnings available.')
                    # Split by periods and keep first 2 sentences
                    sentences = [s.strip() + '.' for s in learnings_text.split('.') if s.strip()]
                    learnings_text = ' '.join(sentences[:2]) if sentences else learnings_text
                    
                    # Add empty line before learnings
                    p = text_frame.add_paragraph()
                    p.text = ""
                    
                    # Add learnings paragraph (normal text, not bold)
                    p = text_frame.add_paragraph()
                    
                    # Add "Learnings:" label (bold)
                    run = p.add_run()
                    run.text = "Learnings: "
                    run.font.bold = True
                    
                    # Add learnings text (normal, not bold)
                    run = p.add_run()
                    run.text = learnings_text
                    run.font.bold = False
                    
                    print("[OK] Updated Slide 5 with bold formatting for creator metrics")
            
            # Add creator content image - Replace template placeholder
            creator_streams, _ = _select_uploaded_streams(
                preferred_types=['creator_content'],
                limit=1,
                fallback_any=True,
            )
            creator_image_stream = creator_streams[0] if creator_streams else _download_image_stream(data.get("postImage"))
            
            if creator_image_stream:
                print("[IMAGE] Adding creator content image to Slide 5")
                
                # Find and delete text placeholder "Creator content picture here"
                text_placeholder = None
                for shape in slide_5.shapes:
                    if hasattr(shape, 'text') and 'Creator content picture here' in shape.text:
                        text_placeholder = shape
                        break
                
                if text_placeholder:
                    sp = text_placeholder.element
                    sp.getparent().remove(sp)
                    print("[OK] Removed 'Creator content picture here' text placeholder")
                
                # Find the AUTO_SHAPE placeholder for creator content (Shape 2)
                # It's positioned in the middle-left area
                creator_placeholder = None
                for shape in slide_5.shapes:
                    if shape.shape_type == 1:  # AUTO_SHAPE
                        # Check if it's the creator content placeholder (left side, middle height)
                        if 3900000 < shape.left < 4500000 and 1000000 < shape.top < 2000000:
                            creator_placeholder = {
                                'shape': shape,
                                'left': shape.left,
                                'top': shape.top,
                                'width': shape.width,
                                'height': shape.height
                            }
                            break
                
                if creator_placeholder:
                    # Delete the placeholder shape
                    sp = creator_placeholder['shape'].element
                    sp.getparent().remove(sp)
                    
                    # Insert image at the same position and size
                    _add_picture_with_fallback(slide_5, creator_image_stream, placeholder=creator_placeholder)
                    
                    print("[OK] Replaced creator content placeholder with actual image")
                else:
                    default_creator_box = {
                        'left': 3958173,
                        'top': 1105651,
                        'width': 2559000,
                        'height': 2932200,
                        'rotation': 0,
                    }
                    if _add_picture_with_fallback(slide_5, creator_image_stream, default_box=default_creator_box):
                        print("[OK] Added fallback creator image using default Slide 5 bounds")
            
            # Add insights screenshot - Replace template placeholder
            insight_streams, _ = _select_uploaded_streams(
                preferred_types=['insights_screenshot', 'campaign_dashboard'],
                limit=1,
                fallback_any=True,
            )
            
            if insight_streams:
                print("[IMAGE] Adding insights screenshot to Slide 5")
                
                # Find and delete text placeholder "Creator's content insights here"
                text_placeholder = None
                for shape in slide_5.shapes:
                    if hasattr(shape, 'text') and 'content insights here' in shape.text:
                        text_placeholder = shape
                        break
                
                if text_placeholder:
                    sp = text_placeholder.element
                    sp.getparent().remove(sp)
                    print("[OK] Removed 'Creator's content insights here' text placeholder")
                
                # Find the AUTO_SHAPE placeholder for insights (Shape 3)
                # It's positioned on the right side
                insight_placeholder = None
                for shape in slide_5.shapes:
                    if shape.shape_type == 1:  # AUTO_SHAPE
                        # Check if it's the insights placeholder (right side)
                        if shape.left > 6000000:
                            insight_placeholder = {
                                'shape': shape,
                                'left': shape.left,
                                'top': shape.top,
                                'width': shape.width,
                                'height': shape.height
                            }
                            break
                
                if insight_placeholder:
                    # Delete the placeholder shape
                    sp = insight_placeholder['shape'].element
                    sp.getparent().remove(sp)
                    
                    # Insert image at the same position and size
                    insight_streams[0].seek(0)
                    slide_5.shapes.add_picture(
                        insight_streams[0],
                        insight_placeholder['left'],
                        insight_placeholder['top'],
                        width=insight_placeholder['width'],
                        height=insight_placeholder['height']
                    )
                    
                    print("[OK] Replaced insights placeholder with actual screenshot")
        
        # Save the presentation
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        # Handle both old format (campaign_name) and new format (campaignName)
        campaign_name = data.get('campaignName') or data.get('campaign_name') or 'Campaign Report'
        normalized_campaign_name = re.sub(r'[^A-Za-z0-9]+', '_', str(campaign_name)).strip('_')
        safe_campaign_name = re.sub(r'_+', '_', normalized_campaign_name) or "campaign_report"
        filename = f"PyroMedia_Report_{safe_campaign_name}_{timestamp}.pptx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        prs.save(output_path)
        print(f"[OK] Report saved: {output_path}")
        
        return output_path
        
    except Exception as e:
        print(f"[ERROR] Error populating PowerPoint: {str(e)}")
        import traceback
        traceback.print_exc()
        raise


def upload_to_google_drive(file_path, filename):
    """Upload PPTX to Google Drive and convert to Google Slides - PRODUCTION READY (OAuth)"""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        import pickle
        
        SCOPES = ['https://www.googleapis.com/auth/drive.file']
        
        creds = None
        
        # Try to load from token.pickle file
        if os.path.exists('token.pickle'):
            print("📂 Loading credentials from token.pickle...")
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        else:
            print("[ERROR] ERROR: token.pickle not found!")
            print("Please generate token.pickle by running generate_token.py locally")
            return None
        
        # Refresh token if expired (NO BROWSER NEEDED)
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing expired token...")
            creds.refresh(Request())
            
            # Save refreshed token
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
            print("[OK] Token refreshed successfully")
        
        # Check if credentials are valid
        if not creds or not creds.valid:
            print("[ERROR] ERROR: Invalid credentials!")
            print("Please regenerate token.pickle by running generate_token.py locally")
            return None
        
        # Build Drive API service
        service = build('drive', 'v3', credentials=creds)
        
        # Upload file with conversion to Google Slides
        file_metadata = {
            'name': filename.replace('.pptx', ''),
            'mimeType': 'application/vnd.google-apps.presentation'
        }
        
        media = MediaFileUpload(
            file_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            resumable=True
        )
        
        print(f"📤 Uploading to Google Drive...")
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        file_link = file.get('webViewLink')
        print(f"[OK] Uploaded to Google Drive: {file_link}")
        
        # Make file accessible to anyone with link
        try:
            permission = {
                'type': 'anyone',
                'role': 'writer',
                'allowFileDiscovery': False
            }
            service.permissions().create(
                fileId=file.get('id'),
                body=permission
            ).execute()
            print(f"[OK] File is now accessible with link")
        except Exception as perm_error:
            print(f"[WARNING]  Could not set link sharing: {perm_error}")
        
        return file_link
        
    except Exception as e:
        print(f"[ERROR] Error uploading to Google Drive: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


@app.route('/')
def home():
    """Home endpoint"""
    return jsonify({
        "message": "Campaign Report Generator API",
        "status": "running",
        "version": "1.0",
        "endpoints": {
            "generate_report": "/api/generate-report (POST)",
            "download": "/api/download/<filename> (GET)",
            "prompt_presets": "/api/prompt-presets (GET)",
            "build_prompt": "/api/build-prompt (POST)",
        }
    })


@app.route('/api/prompt-presets', methods=['GET'])
def get_prompt_presets():
    """Return the built-in OpenRouter prompt presets."""
    return jsonify({
        "success": True,
        "default_template": "master",
        "templates": OPENROUTER_PROMPT_TEMPLATES,
    })


@app.route('/api/build-prompt', methods=['POST'])
def build_prompt():
    """Build a formatted prompt from one of the OpenRouter prompt presets."""
    body = request.get_json(silent=True) or {}
    template_name = body.get("template", "master")
    data = body.get("data", {}) or {}

    try:
        prompt = build_openrouter_prompt(template_name, data)
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400

    return jsonify({
        "success": True,
        "template": template_name,
        "prompt": prompt,
    })


@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    """Main endpoint to generate campaign report"""
    try:
        warnings = []
        # Get form data
        prompt = request.form.get('prompt', '')
        raw_images = request.files.getlist('images')
        images = [make_in_memory_upload(image) for image in raw_images]
        instagram_post_url = request.form.get('instagram_post_url', '').strip()
        instagram_profile_url = request.form.get('instagram_profile_url', '').strip()
        instagram_username = request.form.get('instagram_username', '').strip()
        youtube_post_url   = request.form.get('youtube_post_url',  '').strip()

        # Validate: need at least one input
        if not prompt and not images and not instagram_post_url and not instagram_profile_url and not instagram_username and not youtube_post_url:
            return jsonify({"success": False, "error": "Provide a prompt, images, an Instagram username/profile URL, or a post URL"}), 400

        print(f"\n{'='*60}")
        print(f"[START] NEW REQUEST")
        print(f"{'='*60}")
        print(f"Prompt: {prompt}")
        print(f"Images: {len(images)} uploaded")
        if instagram_post_url: print(f"Instagram URL: {instagram_post_url}")
        if instagram_profile_url: print(f"Instagram Profile URL: {instagram_profile_url}")
        if instagram_username: print(f"Instagram Username: {instagram_username}")
        if youtube_post_url:   print(f"YouTube URL:   {youtube_post_url}")

        # Step 1: Extract metrics from text prompt
        print("[NOTE] Extracting metrics from text...")
        text_metrics = extract_metrics_from_text(prompt) if prompt else {}
        prompt_context = extract_prompt_context(prompt) if prompt else {}

        # Step 2: Run slow independent fetches in parallel.
        instagram_post = {}
        instagram_profile = {}
        youtube_metrics = {}
        ocr_metrics = {}
        extracted_data = create_default_data()

        yt_url = youtube_post_url or None
        if not yt_url and prompt and ("youtube.com" in prompt or "youtu.be" in prompt):
            yt_url = prompt
        profile_ref = instagram_profile_url or instagram_username
        profile_username = (
            profile_ref.rstrip("/").split("/")[-1]
            if profile_ref and "instagram.com" in profile_ref.lower()
            else profile_ref
        )

        future_map = {}
        max_workers = 1 + int(bool(instagram_post_url)) + int(bool(profile_username)) + int(bool(HAS_YT_DLP and yt_url)) + int(bool(images)) + int(bool(images))
        with ThreadPoolExecutor(max_workers=max(1, min(6, max_workers))) as executor:
            if instagram_post_url:
                print(f"[INSTAGRAM POST] Fetching exact metrics from: {instagram_post_url}")
                future_map["instagram_post"] = executor.submit(fetch_instagram_post_data, instagram_post_url)

            if profile_username:
                print(f"[INSTAGRAM PROFILE] Fetching profile data for: {profile_ref}")
                future_map["instagram_profile"] = executor.submit(fetch_instagram_data, profile_username)

            if HAS_YT_DLP and yt_url:
                print("[VIDEO] Fetching YouTube metrics...")
                future_map["youtube_metrics"] = executor.submit(fetch_youtube_metrics, yt_url)

            if images:
                future_map["ocr_metrics"] = executor.submit(extract_metrics_from_images_ocr, clone_uploads(images))
                future_map["vision_extract"] = executor.submit(analyze_images_with_gpt4, clone_uploads(images), prompt)

            if "instagram_post" in future_map:
                post_data = future_map["instagram_post"].result()
                if post_data:
                    instagram_post = post_data
                    print(f"[OK] Post metrics fetched — likes: {post_data['instagram'].get('likes')}, views: {post_data['instagram'].get('views')}")
                else:
                    warnings.append(
                        "Bright Data did not return usable Instagram metrics for this URL, so the backend skipped Instagram URL metrics and continued with screenshots, prompt data, and other available sources."
                    )

            if "instagram_profile" in future_map:
                instagram_profile = future_map["instagram_profile"].result()
                if not instagram_profile:
                    warnings.append(
                        "Bright Data did not return usable Instagram profile data, so the backend skipped profile metrics."
                    )

            if "youtube_metrics" in future_map:
                youtube_metrics = future_map["youtube_metrics"].result() or {}
                if not youtube_metrics:
                    warnings.append(
                        "YouTube URL fetch did not return usable metrics, so the backend skipped it and continued with the remaining sources."
                    )

            if "ocr_metrics" in future_map:
                ocr_metrics = future_map["ocr_metrics"].result() or {}

            if "vision_extract" in future_map:
                extracted_data = future_map["vision_extract"].result() or create_default_data()

        if images and not has_meaningful_metrics(extracted_data) and not ocr_metrics:
            warnings.append(
                "The uploaded images did not expose readable metrics, so the backend ignored them for metrics and used only prompt text and other available sources."
            )

        # Treat the user's prompt as the source of truth for campaign identity fields.
        extracted_data = apply_prompt_context(extracted_data, prompt_context, force_identity=True)

        # Step 4b: Fill top-level fields from prompt text if Vision AI left them empty
        import re as _re2

        def _extract_from_prompt(patterns, text):
            for p in patterns:
                m = _re2.search(p, text, _re2.IGNORECASE)
                if m:
                    return m.group(1).strip()
            return ""

        if prompt:
            if not extracted_data.get("campaignName") or extracted_data["campaignName"] in ("", "Campaign Report"):
                val = _extract_from_prompt([
                    r'report\s+for\s+([^\n,\.]+)',
                    r'campaign(?:\s+name)?[:\s]+([^\n,\.]+)',
                ], prompt)
                if val:
                    extracted_data["campaignName"] = val.strip()[:60]

            # Only extract brand if explicitly stated with "brand:" prefix — never guess from campaign name
            if not extracted_data.get("brand") or extracted_data["brand"] in ("", "Unknown Brand"):
                val = _extract_from_prompt([r'\bbrand[:\s]+([^\n,\.]+)'], prompt)
                if val:
                    extracted_data["brand"] = val.strip()[:40]

            if not extracted_data.get("creator"):
                val = _extract_from_prompt([
                    r'(?:influencer|creator)\s*[:\s]\s*([^\n,\.]{2,})',          # "influencer: Name"
                    r'with\s+((?:influencer|creator)\s+[^\n,\.]+?)(?:\s*,|\s+budget|\s+ran|$)',  # "with Influencer X"
                ], prompt)
                # Reject single-letter or very short matches (e.g. "X")
                if val and len(val.strip()) > 1:
                    extracted_data["creator"] = val.strip()[:40]

            # Budget → financial
            if not extracted_data.get("financial", {}).get("totalBudget"):
                bval = _extract_from_prompt([r'budget\s*\$?([\d,]+)'], prompt)
                if bval:
                    extracted_data.setdefault("financial", {})["totalBudget"] = bval.replace(",", "")

            # Dates
            if not extracted_data.get("startDate") or not extracted_data.get("endDate"):
                # Match "June 1-30" or "June 1 - June 30" or "June 1 to June 30"
                range_match = _re2.search(
                    r'ran\s+([A-Za-z]+)\s+(\d+)\s*[-–to]+\s*(?:([A-Za-z]+)\s+)?(\d+)',
                    prompt, _re2.IGNORECASE
                )
                if range_match:
                    month_start = range_match.group(1)
                    day_start   = range_match.group(2)
                    month_end   = range_match.group(3) or month_start
                    day_end     = range_match.group(4)
                    if not extracted_data.get("startDate"):
                        extracted_data["startDate"] = f"{month_start} {day_start}"
                    if not extracted_data.get("endDate"):
                        extracted_data["endDate"] = f"{month_end} {day_end}"
                else:
                    if not extracted_data.get("startDate"):
                        val = _extract_from_prompt([r'ran\s+([A-Za-z]+\s+\d+)'], prompt)
                        if val:
                            extracted_data["startDate"] = val

        extracted_data = apply_prompt_context(extracted_data, prompt_context, force_identity=True)

        # Step 5: Merge all metrics from prompt, images, and URLs.
        # Priority (highest → lowest):
        #   1. Bright Data / direct URL metrics with non-zero values
        #   2. Vision / OCR image metrics
        #   3. Prompt text metrics
        #   4. YouTube API fills YouTube fields
        print("[MERGE] Merging metrics from prompt, images, and URLs...")

        extracted_data.setdefault("instagram", {})
        extracted_data.setdefault("youtube", {})

        def fill_empty(target_dict, source_dict):
            """Copy from source into target only where target value is empty/missing."""
            for key, value in source_dict.items():
                if _has_metric_value(value) and not _has_metric_value(target_dict.get(key)):
                    target_dict[key] = value

        def fill_prefer_source(target_dict, source_dict):
            """Copy source into target when the source value is meaningful."""
            for key, value in source_dict.items():
                if _has_metric_value(value):
                    target_dict[key] = value

        def _has_metric_value(value):
            if value is None:
                return False
            if isinstance(value, str):
                return value.strip() not in ("", "0", "0.0", "0%")
            return bool(value)

        prompt_instagram = {}
        prompt_to_ig = {
            "views": "views",
            "likes": "likes",
            "shares": "shares",
            "saves": "saves",
            "engagement_rate": "engagementRate",
        }
        for text_key, ig_key in prompt_to_ig.items():
            value = text_metrics.get(text_key)
            if _has_metric_value(value):
                prompt_instagram[ig_key] = value

        image_instagram = {}
        ocr_to_ig = {
            "views": "views",
            "likes": "likes",
            "shares": "shares",
            "saves": "saves",
        }
        for ocr_key, ig_key in ocr_to_ig.items():
            value = ocr_metrics.get(ocr_key)
            if _has_metric_value(value):
                image_instagram[ig_key] = value

        # Keep useful metrics found from screenshots / OCR / prompt.
        fill_empty(extracted_data["instagram"], prompt_instagram)
        fill_empty(extracted_data["instagram"], image_instagram)

        # Keep financial totals from prompt/OCR text.
        if _has_metric_value(text_metrics.get("total_engagement")) and not _has_metric_value(extracted_data.get("performance", {}).get("totalEngagement")):
            extracted_data.setdefault("performance", {})["totalEngagement"] = text_metrics.get("total_engagement")
        if _has_metric_value(ocr_metrics.get("total_engagement")) and not _has_metric_value(extracted_data.get("performance", {}).get("totalEngagement")):
            extracted_data.setdefault("performance", {})["totalEngagement"] = ocr_metrics.get("total_engagement")

        # Financial details from text go into financial
        fin_key_map = {
            "budget": "totalBudget",
            "budget_currency": "budgetCurrency",
            "cpv": "cpv",
            "cpe": "cpe",
            "cpc": "cpc",
            "cpc_goal": "cpcGoal",
            "cpv_goal": "cpvGoal",
            "cpc_calculation": "cpcCalculation",
            "cpv_calculation": "cpvCalculation",
        }
        for txt_key, fin_key in fin_key_map.items():
            val = text_metrics.get(txt_key) or ocr_metrics.get(txt_key)
            if val and val != "N/A" and not extracted_data.get("financial", {}).get(fin_key):
                extracted_data.setdefault("financial", {})[fin_key] = val

        clicks_value = text_metrics.get("clicks") or ocr_metrics.get("clicks")
        if _has_metric_value(clicks_value) and not _has_metric_value(extracted_data.get("performance", {}).get("totalClicks")):
            extracted_data.setdefault("performance", {})["totalClicks"] = clicks_value

        # Instagram post URL — exact metrics for this specific post/reel.
        if instagram_post:
            ig_post_metrics = instagram_post.get("instagram", {})
            fill_prefer_source(extracted_data["instagram"], ig_post_metrics)
            if not extracted_data.get("creator"):
                extracted_data["creator"] = (
                    instagram_post.get("fullName") or instagram_post.get("username", "")
                )
            if not extracted_data.get("deliverables"):
                extracted_data["deliverables"] = instagram_post.get("mediaType", "")
            if instagram_post.get("postImage"):
                extracted_data["postImage"] = instagram_post.get("postImage")
            if instagram_post.get("videoUrl"):
                extracted_data["videoUrl"] = instagram_post.get("videoUrl")
            extracted_data["instagramSource"] = instagram_post.get("source", "instagram_url")

        # Instagram profile dataset — fill remaining Instagram fields only.
        if instagram_profile:
            fill_empty(extracted_data["instagram"], instagram_profile.get("instagram", {}))
            if not extracted_data.get("creator"):
                extracted_data["creator"] = (
                    instagram_profile.get("fullName") or instagram_profile.get("username", "")
                )
            extracted_data["instagramSource"] = extracted_data.get("instagramSource") or instagram_profile.get("source", "brightdata_profile_dataset")

        # YouTube API — fill empty youtube fields
        youtube_schema_keys = {"views", "likes", "comments", "shares", "watchTime", "ctr"}
        for key, value in youtube_metrics.items():
            if not value or value == "N/A":
                continue
            if key == "_title":
                if not extracted_data.get("campaignName"):
                    extracted_data["campaignName"] = value
            elif key == "_creator_name":
                if not extracted_data.get("creator"):
                    extracted_data["creator"] = value
            elif key in youtube_schema_keys:
                if not extracted_data["youtube"].get(key):
                    extracted_data["youtube"][key] = value
        
        # Step 5b: Propagate instagram → performance totals and creators
        ig = extracted_data.get("instagram", {})
        perf = extracted_data.setdefault("performance", {})

        # Map final instagram fields → performance totals
        ig_to_perf = {
            "views":          "totalViews",
            "likes":          "totalLikes",
            "comments":       "totalComments",
            "shares":         "totalShares",
            "saves":          "totalSaves",
            "reach":          "totalReach",
        }
        for ig_key, perf_key in ig_to_perf.items():
            if _has_metric_value(ig.get(ig_key)):
                perf[perf_key] = ig[ig_key]

        # Calculate totalInteractions = likes + comments + shares + saves
        def _safe_int(v):
            """Return int value only if v is a non-empty, non-zero string."""
            if not v or str(v).strip() in ("", "0"):
                return 0
            try:
                return int(float(str(v).replace('%', '').replace(',', '')))
            except Exception:
                return 0

        if not perf.get("totalInteractions"):
            total_interact = sum(_safe_int(ig.get(k)) for k in ["likes", "comments", "shares", "saves"])
            if total_interact > 0:
                perf["totalInteractions"] = str(total_interact)

        def _safe_float(v):
            if v in (None, "", "N/A"):
                return 0.0
            try:
                return float(str(v).replace('%', '').replace(',', '').strip())
            except Exception:
                return 0.0

        # Formula-based financial calculations
        financial = extracted_data.setdefault("financial", {})
        budget_val = _safe_float(financial.get("totalBudget"))
        views_val = _safe_float(perf.get("totalViews") or ig.get("views"))
        interactions_val = _safe_float(perf.get("totalInteractions"))
        clicks_val = _safe_float(perf.get("totalClicks"))
        if not financial.get("budgetCurrency"):
            financial["budgetCurrency"] = "INR"

        # Derive clicks from YouTube CTR and views when explicit clicks are missing.
        if clicks_val <= 0:
            yt_views = _safe_float(extracted_data.get("youtube", {}).get("views"))
            yt_ctr = _safe_float(extracted_data.get("youtube", {}).get("ctr"))
            if yt_views > 0 and yt_ctr > 0:
                clicks_val = round(yt_views * yt_ctr / 100.0, 2)
                perf["totalClicks"] = str(int(clicks_val)) if clicks_val.is_integer() else str(clicks_val)

        def _format_formula_metric(value):
            rounded = round(value, 2)
            return str(int(rounded)) if float(rounded).is_integer() else f"{rounded:.2f}"

        # Always prefer exact formula-derived numeric values in final output.
        if budget_val > 0 and views_val > 0:
            financial["cpv"] = _format_formula_metric(budget_val / views_val)

        if budget_val > 0 and interactions_val > 0:
            financial["cpe"] = _format_formula_metric(budget_val / interactions_val)

        if budget_val > 0 and clicks_val > 0:
            financial["cpc"] = _format_formula_metric(budget_val / clicks_val)

        # Propagate into creators list — dedupe and sync final instagram values
        creator_name = extracted_data.get("creator", "")
        creators = extracted_data.get("creators", [])
        if not creators and creator_name:
            creators = [{"name": creator_name}]
            extracted_data["creators"] = creators
        deduped_creators = []
        seen_creator_keys = set()
        for c in creators:
            if not isinstance(c, dict):
                continue
            name = (c.get("name") or creator_name or "Creator").strip()
            creator_key = name.lower()
            if creator_key in seen_creator_keys:
                continue
            seen_creator_keys.add(creator_key)
            normalized_creator = dict(c)
            normalized_creator["name"] = name
            for ig_key in ["views", "likes", "comments", "shares", "saves", "reach", "engagementRate"]:
                ig_val = ig.get(ig_key, "")
                if _has_metric_value(ig_val):
                    normalized_creator[ig_key] = ig_val
            total = sum(_safe_int(normalized_creator.get(k)) for k in ["likes", "comments", "shares", "saves"])
            normalized_creator["interactions"] = str(total) if total > 0 else normalized_creator.get("interactions", "")
            normalized_creator["platform"] = normalized_creator.get("platform") or "Instagram"
            deduped_creators.append(normalized_creator)
        extracted_data["creators"] = deduped_creators

        # Step 6: Populate PowerPoint template
        output_path = populate_powerpoint(extracted_data, images)

        # Step 7: Upload to Google Drive (optional)
        google_slides_link = None
        try:
            google_slides_link = upload_to_google_drive(output_path, os.path.basename(output_path))
        except Exception as e:
            print(f"[WARNING]  Google Drive upload failed: {str(e)}")

        # Validate and clean extracted data
        extracted_data = validate_and_clean_data(extracted_data)

        # Debug: Print what's being returned
        print(f"[OK] Returning extracted data:")
        print(f"   - campaignName: {extracted_data.get('campaignName')}")
        print(f"   - brand: {extracted_data.get('brand')}")
        print(f"   - instagram views: {extracted_data.get('instagram', {}).get('views')}")
        print(f"   - postImage present: {bool(extracted_data.get('postImage'))}")
        print(f"   - videoUrl present: {bool(extracted_data.get('videoUrl'))}")
        print(f"   - creators: {len(extracted_data.get('creators', []))} creator(s)")
        print(f"   - Keys: {list(extracted_data.keys())}")
        # Build warnings list for the frontend
        ig = extracted_data.get("instagram", {})
        has_ig_metrics = any(ig.get(k) for k in ["views", "likes", "comments", "reach"])
        if not has_ig_metrics:
            warnings.append("No Instagram metrics found from Bright Data. Provide a valid Instagram profile URL/username or post/reel URL.")

        return jsonify({
            "success": True,
            "message": "Report generated successfully!",
            "filename": os.path.basename(output_path),
            "google_slides_link": google_slides_link,
            "extracted_data": extracted_data,
            "warnings": warnings,
        })
        
    except Exception as e:
        print(f"[ERROR] Error in generate_report: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


def _safe_int(value):
    if value in (None, "", False):
        return 0
    if isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    import re
    digits = re.sub(r"[^\d]", "", str(value))
    return int(digits) if digits else 0


def _dig(obj, *path):
    cur = obj
    for key in path:
        if isinstance(cur, dict):
            cur = cur.get(key)
        elif isinstance(cur, list) and isinstance(key, int) and 0 <= key < len(cur):
            cur = cur[key]
        else:
            return None
    return cur


def _first_value(record, *paths):
    for path in paths:
        if not isinstance(path, tuple):
            path = (path,)  
        value = _dig(record, *path)
        if value not in (None, ""):
            return value
    return None


def _find_first_media_url(obj, media_kind):
    """Recursively search nested payloads for a likely image/video URL."""
    if media_kind not in ("image", "video"):
        return ""

    preferred_key_hints = {
        "image": ("display", "image", "thumbnail", "cover", "poster", "src"),
        "video": ("video", "playback", "stream", "mp4"),
    }
    banned_key_hints = {
        "image": ("profile_pic", "avatar", "icon"),
        "video": (),
    }
    preferred_exts = {
        "image": (".jpg", ".jpeg", ".png", ".webp", ".gif"),
        "video": (".mp4", ".mov", ".m3u8"),
    }

    def walk(value, parent_key=""):
        if isinstance(value, dict):
            for key, child in value.items():
                found = walk(child, str(key).lower())
                if found:
                    return found
        elif isinstance(value, list):
            for child in value:
                found = walk(child, parent_key)
                if found:
                    return found
        elif isinstance(value, str):
            lower = value.lower().strip()
            if lower.startswith("http"):
                key_ok = any(h in parent_key for h in preferred_key_hints[media_kind]) if parent_key else False
                ext_ok = any(ext in lower for ext in preferred_exts[media_kind])
                banned = any(h in parent_key for h in banned_key_hints[media_kind])
                if not banned and (key_ok or ext_ok):
                    return value
        return ""

    return walk(obj)

  
def _extract_brightdata_records(payload):
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for key in ["data", "results", "items", "output"]:
            value = payload.get(key)
            if isinstance(value, list):
                return value
        if payload:
            return [payload]
    return []


def _fetch_brightdata_records(urls, dataset_id, max_attempts=12, poll_delay=3):
    import requests as req
    import time

    brightdata_token = (
        os.environ.get("BRIGHTDATA_API_TOKEN", "").strip()
        or os.environ.get("BRIGHTDATA_API_KEY", "").strip()
    )
    if not brightdata_token:
        print("[WARNING] Bright Data token is not configured")
        return []

    if not dataset_id:
        print("[WARNING] Bright Data dataset id is not configured")
        return []

    resp = req.post(
        f"https://api.brightdata.com/datasets/v3/scrape?dataset_id={dataset_id}&notify=false&include_errors=true",
        headers={
            "Authorization": f"Bearer {brightdata_token}",
            "Content-Type": "application/json",
        },
        json={"input": [{"url": url} for url in urls]},
        timeout=90,
    )
    resp.raise_for_status()
    payload = resp.json()
    records = _extract_brightdata_records(payload)
    if records:
        return records

    snapshot_id = ""
    if isinstance(payload, dict):
        snapshot_id = payload.get("snapshot_id") or payload.get("id") or ""
    if not snapshot_id:
        return []

    print(f"[BRIGHTDATA] Snapshot created: {snapshot_id}")
    for attempt in range(max_attempts):
        progress_resp = req.get(
            f"https://api.brightdata.com/datasets/v3/progress/{snapshot_id}",
            headers={"Authorization": f"Bearer {brightdata_token}"},
            timeout=30,
        )
        progress_resp.raise_for_status()
        progress_data = progress_resp.json()
        status = progress_data.get("status", "")
        print(f"[BRIGHTDATA] Snapshot status: {status} (attempt {attempt + 1}/{max_attempts})")

        if status == "ready":
            snapshot_resp = req.get(
                f"https://api.brightdata.com/datasets/v3/snapshot/{snapshot_id}",
                headers={"Authorization": f"Bearer {brightdata_token}"},
                params={"format": "json"},
                timeout=60,
            )
            snapshot_resp.raise_for_status()
            return _extract_brightdata_records(snapshot_resp.json())

        if status == "failed":
            print(f"[WARNING] Bright Data snapshot failed for {snapshot_id}")
            return []

        time.sleep(poll_delay)

    return []


def fetch_instagram_post_data(post_url):
    """Fetch exact metrics for a specific Instagram post/Reel from Bright Data."""
    import re
    cache_key = ("instagram_post", (post_url or "").strip())
    cached = _cache_get(cache_key)
    if cached is not None:
        print(f"[CACHE] Using cached Instagram post data for: {post_url}")
        return cached

    media_type = "Instagram Reel" if "/reel/" in post_url else "Instagram Post"
    match = re.search(r'/(?:p|reel|tv)/([A-Za-z0-9_-]+)/?', post_url)
    if not match:
        print(f"[ERROR] Cannot parse shortcode from: {post_url}")
        return None

    shortcode = match.group(1)
    dataset_id = os.environ.get("BRIGHTDATA_INSTAGRAM_POST_DATASET_ID", "gd_lyclm20il4r5helnj").strip()
    merged = {
        "mediaType": media_type,
        "source": "",
        "username": "",
        "fullName": "",
        "brandUsername": "",
        "postImage": "",
        "videoUrl": "",
        "raw_data": {},
        "instagram": {
            "views": "",
            "likes": "",
            "comments": "",
            "shares": "",
            "saves": "",
            "reach": "",
            "impressions": "",
            "engagementRate": "",
        },
    }

    def _fill_post_field(key, value):
        if value not in (None, "") and not merged.get(key):
            merged[key] = value

    def _fill_instagram_field(key, value):
        if value not in (None, "") and not merged["instagram"].get(key):
            merged["instagram"][key] = str(value)

    # Bright Data is the single source of truth for Instagram post metrics.
    try:
        records = _fetch_brightdata_records([post_url], dataset_id)
    except Exception as e:
        print(f"[WARNING] Bright Data dataset fetch failed: {e}")
        records = []

    for record in records:
        if not isinstance(record, dict):
            continue

        record_shortcode = (
            _first_value(record, "shortcode", ("post", "shortcode"), ("node", "shortcode"))
            or ""
        )
        record_url = (
            _first_value(record, "url", "post_url", ("post", "url"), ("node", "url"))
            or ""
        )
        if record_shortcode and record_shortcode != shortcode:
            continue
        if record_url and shortcode not in str(record_url):
            continue

        likes = _safe_int(_first_value(record, "like_count", "likes", ("post", "like_count"), ("node", "like_count")))
        comments = _safe_int(_first_value(record, "comment_count", "comments_count", "comments", ("post", "comment_count"), ("node", "comment_count")))
        views = _safe_int(_first_value(record, "video_play_count", "play_count", "video_view_count", "view_count", "views", ("post", "video_play_count"), ("post", "play_count"), ("post", "video_view_count"), ("node", "video_view_count")))
        shares = _safe_int(_first_value(record, "share_count", "shares", ("post", "share_count")))
        saves = _safe_int(_first_value(record, "save_count", "saves", ("post", "save_count")))
        reach = _safe_int(_first_value(record, "reach", ("owner", "edge_followed_by", "count"), ("user", "follower_count")))
        impressions = _safe_int(_first_value(record, "impressions", ("post", "impressions")))
        username = _first_value(record, ("owner", "username"), ("user", "username"), "username", ("post", "owner_username")) or ""
        full_name = _first_value(record, ("owner", "full_name"), ("user", "full_name"), "full_name", ("post", "owner_full_name")) or ""
        engagement_rate = _first_value(record, "engagement_rate", ("post", "engagement_rate")) or ""
        post_image = _first_value(
            record,
            "display_url",
            "displayUrl",
            "image_url",
            "thumbnail_src",
            "thumbnail",
            "thumbnail_url",
            ("post", "display_url"),
            ("post", "thumbnail_src"),
            ("node", "display_url"),
            ("node", "thumbnail_src"),
            ("image_versions2", "candidates", 0, "url"),
            ("post", "image_versions2", "candidates", 0, "url"),
            ("display_resources", 0, "src"),
            ("node", "display_resources", 0, "src"),
            ("additional_candidates", "first_frame", "url"),
            ("post", "additional_candidates", "first_frame", "url"),
            ("carousel_media", 0, "image_versions2", "candidates", 0, "url"),
        ) or _find_first_media_url(record, "image") or ""
        video_url = _first_value(
            record,
            "video_url",
            "videoUrl",
            ("post", "video_url"),
            ("node", "video_url"),
            ("video_versions", 0, "url"),
            ("post", "video_versions", 0, "url"),
            ("carousel_media", 0, "video_versions", 0, "url"),
        ) or _find_first_media_url(record, "video") or ""

        _fill_instagram_field("views", views)
        _fill_instagram_field("likes", likes)
        _fill_instagram_field("comments", comments)
        _fill_instagram_field("shares", shares)
        _fill_instagram_field("saves", saves)
        _fill_instagram_field("reach", reach)
        _fill_instagram_field("impressions", impressions)
        if engagement_rate:
            _fill_instagram_field("engagementRate", engagement_rate)
        _fill_post_field("username", username)
        _fill_post_field("fullName", full_name)
        _fill_post_field("postImage", post_image)
        _fill_post_field("videoUrl", video_url)
        merged["source"] = merged["source"] or "brightdata_post_dataset"
        merged["raw_data"]["brightdata_post_dataset"] = record
        print(f"[OK] Bright Data post data found for shortcode {shortcode}")
        break

    # Compute a lightweight public engagement rate when we have enough public data.
    if not merged["instagram"].get("engagementRate"):
        try:
            views = _safe_int(merged["instagram"].get("views"))
            likes = _safe_int(merged["instagram"].get("likes"))
            comments = _safe_int(merged["instagram"].get("comments"))
            if views > 0 and (likes or comments):
                merged["instagram"]["engagementRate"] = f"{round((likes + comments) / views * 100, 2)}%"
        except Exception:
            pass

    has_metrics = any(merged["instagram"].get(k) for k in ["views", "likes", "comments", "shares", "saves", "reach", "impressions"])
    if not has_metrics:
        print("[WARNING] No usable Instagram post metrics returned from Bright Data")
        return None

    _cache_set(cache_key, merged)
    return merged


def fetch_instagram_data(username):
    """Fetch Instagram profile data using only the Bright Data profile dataset."""
    username = (username or "").strip().lstrip("@")
    if not username:
        return None
    cache_key = ("instagram_profile", username.lower())
    cached = _cache_get(cache_key)
    if cached is not None:
        print(f"[CACHE] Using cached Instagram profile data for: @{username}")
        return cached

    dataset_id = os.environ.get("BRIGHTDATA_INSTAGRAM_PROFILE_DATASET_ID", "gd_l1vikfch901nx3by4").strip()
    profile_url = f"https://www.instagram.com/{username}/"

    try:
        records = _fetch_brightdata_records([profile_url], dataset_id)
    except Exception as e:
        print(f"[WARNING] Bright Data dataset fetch failed: {e}")
        return None

    for record in records:
        if not isinstance(record, dict):
            continue

        record_username = (_first_value(record, "username", ("user", "username"), ("owner", "username")) or "").strip().lstrip("@")
        record_url = str(_first_value(record, "url", "profile_url", ("user", "url")) or "")
        if record_username and record_username.lower() != username.lower():
            continue
        if record_url and f"/{username.lower()}" not in record_url.lower():
            continue

        followers = _safe_int(_first_value(record, "followers", "followers_count", ("user", "followers_count"), ("owner", "edge_followed_by", "count")))
        following = _safe_int(_first_value(record, "following", "following_count", ("user", "following_count"), ("owner", "edge_follow", "count")))
        post_count = _safe_int(_first_value(record, "posts_count", "posts", ("user", "media_count")))
        full_name = _first_value(record, "profile_name", "full_name", ("user", "full_name")) or username
        bio = _first_value(record, "biography", "bio", ("user", "biography")) or ""
        is_verified = bool(_first_value(record, "is_verified", ("user", "is_verified")))
        profile_pic = _first_value(record, "profile_image_link", "profile_pic_url", ("user", "profile_pic_url")) or ""

        posts = _first_value(record, "posts") or []
        if not isinstance(posts, list):
            posts = []

        likes_list = [_safe_int(post.get("likes")) for post in posts if isinstance(post, dict) and post.get("likes") not in (None, "")]
        comments_list = [_safe_int(post.get("comments")) for post in posts if isinstance(post, dict) and post.get("comments") not in (None, "")]
        views_list = [_safe_int(post.get("views")) for post in posts if isinstance(post, dict) and post.get("views") not in (None, "")]

        def avg(values):
            return int(sum(values) / len(values)) if values else 0

        avg_likes = avg(likes_list)
        avg_comments = avg(comments_list)
        avg_views = avg(views_list)
        engagement_rate = ""
        if followers > 0 and (avg_likes or avg_comments):
            engagement_rate = f"{round((avg_likes + avg_comments) / followers * 100, 2)}%"

        print(f"[OK] Bright Data profile data found for @{username}")
        result = {
            "source": "brightdata_profile_dataset",
            "username": record_username or username,
            "fullName": full_name,
            "bio": bio,
            "isVerified": is_verified,
            "profilePic": profile_pic,
            "followersCount": followers,
            "followingCount": following,
            "postCount": post_count,
            "instagram": {
                "views": str(avg_views) if avg_views else "",
                "likes": str(avg_likes) if avg_likes else "",
                "comments": str(avg_comments) if avg_comments else "",
                "shares": "",
                "saves": "",
                "reach": str(followers) if followers else "",
                "impressions": "",
                "engagementRate": engagement_rate,
            },
            "raw_data": record,
        }
        _cache_set(cache_key, result)
        return result

    print(f"[WARNING] Bright Data returned no usable Instagram profile data for @{username}")
    return None


@app.route('/api/get-instagram-data', methods=['GET', 'POST'])
def get_instagram_data():
    """
    GET  /api/get-instagram-data?username=cristiano
    GET  /api/get-instagram-data?url=https://www.instagram.com/cristiano/
    POST /api/get-instagram-data  body: {"username": "cristiano"} or {"url": "..."}
    Returns Instagram profile data from the Bright Data profile dataset.
    """
    if request.method == 'POST':
        body     = request.get_json(silent=True) or {}
        username = body.get('username', '').strip()
        profile_url = body.get('url', '').strip()
    else:
        username = request.args.get('username', '').strip()
        profile_url = request.args.get('url', '').strip()

    if not username and not profile_url:
        return jsonify({"success": False, "error": "username or url is required"}), 400

    if profile_url and "instagram.com" in profile_url.lower():
        username = profile_url.rstrip("/").split("/")[-1]

    username = username.lstrip('@')

    print(f"[INSTAGRAM] Fetching Bright Data profile data for @{username} ...")
    data = fetch_instagram_data(username)

    if data is None:
        return jsonify({
            "success": False,
            "error":   (
                "Could not fetch Instagram data from Bright Data. "
                "The account may be private, missing, or unavailable from the dataset."
            )
        }), 502

    return jsonify({"success": True, "data": data})


@app.route('/api/proxy-media')
def proxy_media():
    media_url = request.args.get('url', '').strip()
    if not media_url or not media_url.startswith(('http://', 'https://')):
        return jsonify({"success": False, "error": "valid url is required"}), 400

    try:
        import requests
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            ),
            "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            "Referer": "https://www.instagram.com/",
        }
        resp = requests.get(media_url, headers=headers, timeout=30)
        resp.raise_for_status()
        content_type = resp.headers.get("Content-Type", "application/octet-stream")
        return Response(resp.content, content_type=content_type)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 502


@app.route('/api/download/<filename>')
def download_report(filename):
    """Download generated PowerPoint report"""
    try:
        file_path = os.path.join(OUTPUT_DIR, filename)
        
        if not os.path.exists(file_path):
            return jsonify({
                "success": False,
                "error": "File not found"
            }), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        print(f"[ERROR] Error downloading file: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


if __name__ == '__main__':
    # Check if template exists
    if not os.path.exists(TEMPLATE_PATH):
        print(f"[WARNING]  WARNING: Template file not found: {TEMPLATE_PATH}")
        print("Please place your template file in the same directory as app.py")
    
    # Check if API key is set
    if not os.environ.get("OPENROUTER_API_KEY") and not os.environ.get("ANTHROPIC_API_KEY"):
        print("[WARNING]  WARNING: No API key found!")
        print("Please set OPENROUTER_API_KEY or ANTHROPIC_API_KEY in .env")
    
    print("\n[START] Starting Flask server...")
    print(f"[DIR] Output directory: {OUTPUT_DIR}")
    print(f"[FILE] Template: {TEMPLATE_PATH}")
    
    app.run(debug=True, port=5000, host='0.0.0.0')
