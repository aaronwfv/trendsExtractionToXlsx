import pandas as pd
import json
import re
import os
import requests
import yaml
import time
from datetime import datetime

RE_STRIP = re.compile(r'</?[^>]+>', re.IGNORECASE)

def strip_tags(html_content):
    if not html_content:
        return ''
    return re.sub(RE_STRIP, ' ', html_content).replace('\n', ' ').replace('\r', ' ').replace('  ', ' ').strip()

def get_human_date(unix_ms):
    try:
        return datetime.utcfromtimestamp(int(unix_ms) / 1000).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ''

def extract_url(article):
    if 'canonicalUrl' in article and article['canonicalUrl']:
        return article['canonicalUrl']
    alternates = article.get('alternate', [])
    if alternates and isinstance(alternates, list):
        return alternates[0].get('href', '')
    return ''

def extract_content(article):
    if article.get('fullContent'):
        return strip_tags(article['fullContent'])
    if article.get('content') and article['content'].get('content'):
        return strip_tags(article['content']['content'])
    if article.get('summary') and article['summary'].get('content'):
        return strip_tags(article['summary']['content'])
    return ''

def extract_type(article, regulatory_ids, expert_id, conference_id):
    # Check businessEvents for regulatory or conference, else expert, else blank
    types = []
    for be in article.get('businessEvents', []):
        if be.get('id') in regulatory_ids:
            types.append('Regulatory')
        if be.get('id') == conference_id:
            types.append('Conferences')
    # Check for expert mention (contentType)
    for topic in article.get('commonTopics', []):
        if topic.get('id') == expert_id:
            types.append('Expert Mentions')
    return ', '.join(set(types)) if types else ''

def sanitize_sheet_name(name):
    # Excel sheet names: max 31 chars, no []:*?/\\
    name = re.sub(r'[\[\]:*?/\\]', '', name)
    return name[:31]

def load_config(config_path='config.yaml'):
    with open(config_path, 'r') as f:
        config = yaml.safe_load(f)
    api_key = config['feedly']['api_key']
    num_trends = config['feedly'].get('num_trends', 1)
    days_ago = config['feedly'].get('days_ago', 90)
    nlp_id = config['feedly'].get('nlp_id', 'nlp/f/topic/1874')
    return api_key, num_trends, days_ago, nlp_id

def get_trends(api_key, nlp_id):
    url = "https://api.feedly.com/v3/ml/trend-discovery/trends?count=20&sort=-growth"
    payload = {
        "filters": [
            {"type": "growth", "values": ["Exploding", "Surging", "Growing"]},
            {"type": "size", "values": ["Mainstream", "Known", "Niche"]},
            {"type": "industry", "values": []}
        ],
        "period": "Last30Days",
        "searchLayers": [{"parts": [{"id": nlp_id}]}]
    }
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    response = requests.post(url, json=payload, headers=headers)
    response.raise_for_status()
    return response.json()

def get_aliases_from_trend(trend):
    return list(trend.get('aliases', []))

def search_articles(api_key, aliases, newer_than=None):
    url = "https://api.feedly.com/v3/search/contents"
    parts = [{"text": alias} for alias in aliases]
    layers = [
        {
            "parts": parts,
            "salience": "mention",
            "type": "matches"
        },
        {
            "negate": True,
            "parts": [{"text": "site:reddit.com"}],
            "type": "matches",
            "salience": "mention"
        }
    ]
    payload = {
        "layers": layers,
        "source": {"items": [
            {
                "id": "discovery:all-topics",
                "label": "All Feedly Sources",
                "tier": "tier3",
                "type": "publicationBucket",
                "description": "Millions of news sites, blogs, trade magazines, subreddits, newsletters, and more"
            }
        ]}
    }
    params = {"count": 100}
    if newer_than is not None:
        params["newerThan"] = newer_than
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    all_items = []
    continuation = None
    last_continuation = None
    page = 1
    while True:
        req_payload = payload.copy()
        req_params = params.copy()
        if continuation:
            req_params["continuation"] = continuation
        try:
            response = requests.post(url, params=req_params, json=req_payload, headers=headers)
            response.raise_for_status()
            data = response.json()
        except Exception as e:
            print(f"  ! Error during Feedly search API call: {e}")
            break
        all_items.extend(data.get("items", []))
        last_continuation = continuation
        continuation = data.get("continuation")
        if not continuation or continuation == last_continuation:
            break
        page += 1
        time.sleep(1)
    return all_items

def main():
    print("Loading config and trends...")
    api_key, num_trends, days_ago, nlp_id = load_config()
    trends_json = get_trends(api_key, nlp_id)
    trends = trends_json.get('trends', [])
    if not trends:
        print("No trends found. Exiting.")
        return
    newer_than = -days_ago * 24 * 60 * 60 * 1000
    print(f"Searching for articles newer than {days_ago} days ago (newerThan: {newer_than})...")
    summary_rows = []
    trend_data = []
    regulatory_ids = [
        'nlp/f/businessEvent/regulatory-changes',
        'nlp/f/businessEvent/regulatory-approvals'
    ]
    expert_id = 'nlp/f/topic/6001'
    conference_id = 'nlp/f/businessEvent/participation-in-an-event'
    for i, trend in enumerate(trends[:num_trends]):
        aliases = get_aliases_from_trend(trend)
        label = trend.get('label', '').strip()
        # If no aliases, use label as the only search term
        if not aliases:
            print(f"No aliases found for trend #{i+1}. Using label '{label}' as search term.")
            aliases = [label] if label else []
        if not aliases:
            print(f"No label or aliases for trend #{i+1}. Skipping.")
            print("Trend JSON:", json.dumps(trend, indent=2))
            continue
        aliases_str = ", ".join(aliases)
        print(f"\n=== Processing Trend #{i+1}: {aliases[0]} ===")
        # Fetch all articles for this trend
        articles = search_articles(api_key, aliases, newer_than=newer_than)
        if not articles:
            print(f"No articles found for trend #{i+1} ({aliases[0]}). Skipping this trend in summary and tabs.")
            continue
        # Prepare summary row
        summary_rows.append([
            aliases_str,
            len(articles),  # Total Articles
            sum(1 for a in articles if extract_type(a, regulatory_ids, expert_id, conference_id).find('Regulatory') != -1),
            sum(1 for a in articles if extract_type(a, regulatory_ids, expert_id, conference_id).find('Expert Mentions') != -1),
            sum(1 for a in articles if extract_type(a, regulatory_ids, expert_id, conference_id).find('Conferences') != -1)
        ])
        # Prepare detailed article data for this trend
        rows = []
        for article in articles:
            rows.append({
                'id': article.get('id', ''),
                'title': article.get('title', ''),
                'crawled': get_human_date(article.get('crawled', '')),
                'url': extract_url(article),
                'content': extract_content(article),
                'type': extract_type(article, regulatory_ids, expert_id, conference_id)
            })
        # Use label for sheet name if aliases was empty
        sheet_name = sanitize_sheet_name(aliases[0]) if aliases else sanitize_sheet_name(label)
        trend_data.append((sheet_name, rows))
        time.sleep(2)
    # Write to Excel
    print("\nWriting article details to article_details.xlsx...")
    with pd.ExcelWriter('article_details.xlsx', engine='openpyxl') as writer:
        # Write summary tab
        summary_df = pd.DataFrame(summary_rows, columns=["Aliases", "Total Articles", "Regulatory Articles", "Expert Mentions", "Conferences"])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        # Write each trend tab
        for sheet, rows in trend_data:
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet, index=False)
    print('Done! Output written to article_details.xlsx')

if __name__ == '__main__':
    main() 