import requests
import feedparser
import pandas as pd
import os
import datetime
import openpyxl
import re
import html

request_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

def clean_html(raw_html):
    # raw_html: string that may contain raw HTML tags and entities
    if not raw_html:
        return ""
    # Remove HTML tags using regular expression
    cleaned = re.sub(r"<[^>]+>", "", str(raw_html))
    # Unescape HTML entities like &nbsp; and &amp;
    cleaned = html.unescape(cleaned)
    # Strip unnecessary leading and trailing whitespace
    return cleaned.strip()

def fetch_api_source(api_url, source_name, language):
    # api_url: string containing the Quintype JSON API endpoint URL
    # source_name: string containing the display name of the newspaper
    # language: string indicating the reporting language (Tamil or English)
    print(f"Fetching API news from: {source_name}")
    response = requests.get(api_url, headers=request_headers, timeout=15)
    print(f"Received response status {response.status_code} from {source_name}")
    data = response.json()
    items = data.get("items", [])
    print(f"Total items found for {source_name}: {len(items)}")
    
    collected_records = []
    for item in items:
        story = item.get("story", item)
        headline = clean_html(story.get("headline", ""))
        summary = clean_html(story.get("subheadline", ""))
        url = story.get("url", "")
        
        raw_time = story.get("published-at")
        if isinstance(raw_time, int):
            published_time = datetime.datetime.fromtimestamp(raw_time / 1000).strftime("%Y-%m-%d %H:%M:%S")
        else:
            published_time = str(raw_time) if raw_time else ""
            
        record = {
            "Source": source_name,
            "Language": language,
            "Headline": headline,
            "Summary": summary,
            "URL": url,
            "Published Time": published_time
        }
        collected_records.append(record)
    return collected_records

def fetch_rss_source(rss_url, source_name, language):
    # rss_url: string containing the RSS feed XML URL
    # source_name: string containing the display name of the newspaper
    # language: string indicating the reporting language (Tamil or English)
    print(f"Fetching RSS feed from: {source_name}")
    response = requests.get(rss_url, headers=request_headers, timeout=15)
    print(f"Received response status {response.status_code} from {source_name}")
    feed = feedparser.parse(response.content)
    entries = feed.entries
    print(f"Total feed entries found for {source_name}: {len(entries)}")
    
    collected_records = []
    for entry in entries:
        headline = clean_html(entry.get("title", ""))
        summary = clean_html(entry.get("summary", ""))
        url = entry.get("link", "")
        published_time = entry.get("published", entry.get("updated", ""))
        
        record = {
            "Source": source_name,
            "Language": language,
            "Headline": headline,
            "Summary": summary,
            "URL": url,
            "Published Time": published_time
        }
        collected_records.append(record)
    return collected_records

api_sources = [
    {
        "name": "Daily Thanthi",
        "url": "https://www.dailythanthi.com/api/v1/collections/tamilnadu-news",
        "language": "Tamil"
    },
    {
        "name": "Dinamani",
        "url": "https://www.dinamani.com/api/v1/collections/tamilnadu",
        "language": "Tamil"
    },
    {
        "name": "Maalai Malar",
        "url": "https://www.maalaimalar.com/api/v1/collections/tamilnadu-news",
        "language": "Tamil"
    },
    {
        "name": "The New Indian Express",
        "url": "https://www.newindianexpress.com/api/v1/collections/tamil-nadu-states",
        "language": "English"
    },
    {
        "name": "DT Next",
        "url": "https://www.dtnext.in/api/v1/collections/newstamilnadu-news",
        "language": "English"
    }
]

rss_sources = [
    {
        "name": "Dinamalar",
        "url": "https://news.google.com/rss/search?q=site:dinamalar.com&hl=ta&gl=IN&ceid=IN:ta",
        "language": "Tamil"
    },
    {
        "name": "The Hindu",
        "url": "https://www.thehindu.com/news/national/tamil-nadu/feeder/default.rss",
        "language": "English"
    }
]

all_news_records = []

# Fetch from all API sources
for src in api_sources:
    records = fetch_api_source(src["url"], src["name"], src["language"])
    all_news_records.extend(records)
    print(f"Accumulated total records: {len(all_news_records)}")

# Fetch from all RSS sources
for src in rss_sources:
    records = fetch_rss_source(src["url"], src["name"], src["language"])
    all_news_records.extend(records)
    print(f"Accumulated total records: {len(all_news_records)}")

print(f"Total articles harvested across all sources: {len(all_news_records)}")

# Create pandas DataFrame
news_df = pd.DataFrame(all_news_records)

# Define target output Excel file path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
output_excel_path = os.path.join(BASE_DIR, "tamil_nadu_news.xlsx")

# Write DataFrame to Excel with openpyxl styling
with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
    news_df.to_excel(writer, index=False, sheet_name="Tamil Nadu News")
    worksheet = writer.sheets["Tamil Nadu News"]
    
    # Auto-adjust column widths for readability
    for column_cells in worksheet.columns:
        col_name = column_cells[0].value
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        col_letter = column_cells[0].column_letter
        
        # Give comfortable margins but cap broad columns like URL and summary
        if col_name in ["Summary", "Headline"]:
            worksheet.column_dimensions[col_letter].width = min(max(max_length, 25), 60)
        elif col_name == "URL":
            worksheet.column_dimensions[col_letter].width = 45
        else:
            worksheet.column_dimensions[col_letter].width = max(max_length + 4, 15)

print(f"Successfully created and styled Excel file at: {output_excel_path}")

# Detailed Explanation:
# 1. Imports and Setup: The script imports requests for HTTP requests, feedparser for RSS parsing, pandas for data structuring and Excel output, datetime for timestamp conversions, and re & html for cleaning HTML tags and entities.
# 2. Function clean_html: Takes a raw text string, strips away any HTML markup tags using regex, decodes HTML entities into normal text characters, and trims trailing whitespace.
# 3. Function fetch_api_source: Takes an API endpoint URL, a source name, and a language string. It makes an HTTP GET request with a browser User-Agent header, extracts story items from the Quintype JSON response, cleans headline and summary text, standardizes timestamp formats, and returns a list of dictionary records.
# 4. Function fetch_rss_source: Takes an RSS XML URL, a source name, and a language string. It downloads the XML content using requests, parses entries via feedparser, cleans HTML markup from titles and summaries, and maps clean records.
# 5. Source Definitions: Quintype API sources (Daily Thanthi, Dinamani, Maalai Malar, The New Indian Express, DT Next) and RSS sources (Dinamalar, The Hindu) are organized in straightforward list-of-dictionary structures.
# 6. Data Extraction Loop: Iterates through all API and RSS sources sequentially, appending each batch of news stories to the master all_news_records list while printing diagnostic log messages.
# 7. Excel Creation and Formatting: The unified list is converted into a pandas DataFrame. Using pd.ExcelWriter and the openpyxl engine, the data is saved to /home/ganesh/Downloads/News_Agent/tamil_nadu_news.xlsx. Column widths are automatically adjusted to ensure titles, summaries, and URLs are properly displayed and readable without manual resizing.
