import requests
import feedparser
import pandas as pd
import os
import datetime
import time
import random
import openpyxl
import re
import html

start_dt = datetime.datetime(2026, 11, 1, 0, 0, 0)
end_dt = datetime.datetime(2026, 11, 17, 23, 59, 59)

start_timestamp_ms = int(start_dt.timestamp() * 1000)
end_timestamp_ms = int(end_dt.timestamp() * 1000)

date_range = f"{start_dt.strftime('%Y_%m_%d')}_to_{end_dt.strftime('%Y_%m_%d')}"

tsv_checkpoint_path = f"/home/ganesh/Downloads/News_Agent/tamil_nadu_news_archive_{date_range}.tsv"
final_excel_path = f"/home/ganesh/Downloads/News_Agent/tamil_nadu_news_archive_{date_range}.xlsx"

print(tsv_checkpoint_path)
print(final_excel_path)

request_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

print(f"Target Date Window: {start_dt.strftime('%Y-%m-%d')} to {end_dt.strftime('%Y-%m-%d')}")
print(f"Continuous Checkpoint TSV: {tsv_checkpoint_path}")
print(f"Final Excel Output: {final_excel_path}")

def clean_html(raw_html):
    # raw_html: string containing raw HTML text with markup tags
    if not raw_html:
        return ""
    cleaned = re.sub(r"<[^>]+>", "", str(raw_html))
    cleaned = html.unescape(cleaned)
    return cleaned.strip()

def append_to_tsv(new_records, target_tsv):
    # new_records: list of dictionary objects holding harvested news stories
    # target_tsv: string filepath where records are saved continuously in TSV format
    if not new_records:
        return
    batch_df = pd.DataFrame(new_records)
    file_exists = os.path.exists(target_tsv)
    batch_df.to_csv(target_tsv, mode="a", header=not file_exists, index=False, sep="\t", encoding="utf-8")
    print(f"Saved checkpoint: {len(new_records)} new records appended to TSV.")

def fetch_url_with_retry(target_url, headers, timeout=40, max_retries=5, backoff_seconds=3):
    # target_url: string containing the HTTP/HTTPS endpoint to fetch
    # headers: dictionary containing HTTP browser headers
    # timeout: integer seconds to wait before timing out a request
    # max_retries: integer number of times to retry on transient failure
    # backoff_seconds: integer seconds to pause between retry attempts
    for attempt_idx in range(1, max_retries + 1):
        try:
            resp = requests.get(target_url, headers=headers, timeout=timeout)
            if resp.status_code == 200:
                return resp
            print(f"HTTP {resp.status_code} on attempt {attempt_idx}/{max_retries}. Retrying in {backoff_seconds}s...")
        except requests.exceptions.RequestException as err:
            print(f"Connection warning on attempt {attempt_idx}/{max_retries}: {err}. Retrying in {backoff_seconds}s...")
        sleep_duration = backoff_seconds * attempt_idx + random.uniform(0.5, 1.5)
        time.sleep(sleep_duration)
    print(f"Server unreachable after {max_retries} attempts: {target_url}")
    return None

seen_urls = set()
completed_sources = set()
if os.path.exists(tsv_checkpoint_path):
    print("Found existing continuous checkpoint TSV. Loading saved records with pandas...")
    existing_df = pd.read_csv(tsv_checkpoint_path, sep="\t")
    for existing_url in existing_df.get("URL", []):
        if pd.notna(existing_url):
            seen_urls.add(str(existing_url).strip())
    print(f"Resumed with {len(seen_urls)} previously saved unique URLs.")

    # todetect newspapers that have already finished fetching all historical news
    completion_cutoff = (start_dt + datetime.timedelta(days=7)).strftime("%Y-%m-%d")
    if "Source" in existing_df.columns and "Published Time" in existing_df.columns:
        for source_name_val, group in existing_df.groupby("Source"):
            oldest_time = str(group["Published Time"].min())
            if oldest_time <= completion_cutoff:
                completed_sources.add(source_name_val)
                print(f"Source '{source_name_val}' already completed (oldest: {oldest_time}). Will skip.")
else:
    print("No previous checkpoint found. Starting fresh collection.")

rss_sources = [
    {
        "name": "Daily Thanthi",
        "site_query": "site:dailythanthi.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Dinamani",
        "site_query": "site:dinamani.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Maalai Malar",
        "site_query": "site:maalaimalar.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Dinamalar",
        "site_query": "site:dinamalar.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Hindu Tamil Thisai",
        "site_query": "site:hindutamil.in",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Dinakaran",
        "site_query": "site:dinakaran.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Puthiyathalaimurai",
        "site_query": "site:puthiyathalaimurai.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "Vikatan",
        "site_query": "site:vikatan.com",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "BBC Tamil",
        "site_query": "site:bbc.com/tamil",
        "lang_code": "ta",
        "language": "Tamil"
    },
    {
        "name": "The Hindu",
        "site_query": "site:thehindu.com/news/national/tamil-nadu",
        "lang_code": "en",
        "language": "English"
    },
    {
        "name": "The New Indian Express",
        "site_query": 'site:newindianexpress.com "Tamil Nadu"',
        "lang_code": "en",
        "language": "English"
    },
    {
        "name": "DT Next",
        "site_query": "site:dtnext.in",
        "lang_code": "en",
        "language": "English"
    },
    {
        "name": "The Times of India (Chennai)",
        "site_query": "site:timesofindia.indiatimes.com/city/chennai",
        "lang_code": "en",
        "language": "English"
    },
    {
        "name": "The Indian Express",
        "site_query": 'site:indianexpress.com "Tamil Nadu"',
        "lang_code": "en",
        "language": "English"
    },
    {
        "name": "The News Minute",
        "site_query": "site:thenewsminute.com/tamil-nadu",
        "lang_code": "en",
        "language": "English"
    }
]

print("Starting historical collection for regional news archives...")

#generate 14-day intervals across configured date window
date_intervals = []
curr_date = start_dt
while curr_date < end_dt:
    next_date = min(curr_date + datetime.timedelta(days=14), end_dt)
    date_intervals.append((curr_date.strftime("%Y-%m-%d"), next_date.strftime("%Y-%m-%d")))
    curr_date = next_date + datetime.timedelta(days=1)

for src in rss_sources:
    source_name = src["name"]
    site_query = src["site_query"]
    lang_code = src["lang_code"]
    language = src["language"]

    if source_name in completed_sources:
        print(f"Skipping {source_name} as it was completely harvested in a previous run.")
        continue

    print(f"Harvesting historical RSS windows for: {source_name}")

    for interval_start, interval_end in date_intervals:
        rss_url = f"https://news.google.com/rss/search?q={site_query}+after:{interval_start}+before:{interval_end}&hl={lang_code}&gl=IN&ceid=IN:{lang_code}"
        print(f"[{source_name}] Requesting window {interval_start} to {interval_end}...")

        response = fetch_url_with_retry(rss_url, request_headers, timeout=40, max_retries=5, backoff_seconds=3)
        if response is None:
            print(f"Skipping window {interval_start} to {interval_end} for {source_name} due to repeated timeouts.")
            continue

        feed = feedparser.parse(response.content)
        entries = feed.entries
        print(f"[{source_name}] Found {len(entries)} entries in window {interval_start} to {interval_end}")

        new_batch_records = []
        for entry in entries:
            raw_url = str(entry.get("link", "")).strip()
            if not raw_url or raw_url in seen_urls:
                continue

            headline = clean_html(entry.get("title", ""))
            summary = clean_html(entry.get("summary", ""))
            published_time = str(entry.get("published", entry.get("updated", "")))

            record = {
                "Source": source_name,
                "Language": language,
                "Headline": headline,
                "Summary": summary,
                "URL": raw_url,
                "Published Time": published_time
            }
            new_batch_records.append(record)
            seen_urls.add(raw_url)

        if new_batch_records:
            append_to_tsv(new_batch_records, tsv_checkpoint_path)

        print(f"[{source_name}] Window done. Total unique URLs captured: {len(seen_urls)}")
        #time.sleep(0.8)
        time.sleep(random.uniform(1.2, 2.6))


print("Finalizing Excel archive file...")
if os.path.exists(tsv_checkpoint_path):
    all_news_df = pd.read_csv(tsv_checkpoint_path, sep="\t")
    print(f"Total rows in final dataset: {len(all_news_df)}")

    excel_sheet_name = f"TN News {start_dt.year}-{end_dt.year}"
    with pd.ExcelWriter(final_excel_path, engine="openpyxl") as writer:
        all_news_df.to_excel(writer, index=False, sheet_name=excel_sheet_name)
        worksheet = writer.sheets[excel_sheet_name]

        for column_cells in worksheet.columns:
            header_val = column_cells[0].value
            max_len = max(len(str(c.value or "")) for c in column_cells)
            col_letter = column_cells[0].column_letter

            if header_val in ("Headline", "Summary"):
                worksheet.column_dimensions[col_letter].width = min(max(max_len, 25), 65)
            elif header_val == "URL":
                worksheet.column_dimensions[col_letter].width = 45
            else:
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

    print(f"Successfully generated final formatted Excel file: {final_excel_path}")
else:
    print("No TSV data found to export to Excel.")