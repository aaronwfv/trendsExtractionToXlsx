# Feedly Trends to Excel Exporter

This script automates the process of fetching trending topics from Feedly, searching for related articles, and exporting the results to an Excel file with a summary and detailed tabs for each trend.

## What the Script Does
- Connects to the Feedly API to retrieve trending topics ("trends") based on your configuration.
- For each trend, searches for related articles in Feedly.
- Extracts key information from each article (ID, title, date, URL, cleaned content, and type).
- Outputs a single Excel file (`article_details.xlsx`) with:
  - A **Summary** tab: overview of each trend and article counts by type.
  - One tab per trend: detailed article data for that trend.
- Skips trends that return no articles.

## Configuration: `config.yaml`
Edit the `config.yaml` file to control the script's behavior:

```yaml
feedly:
  api_key: "YOUR_FEEDLY_API_KEY"
  num_trends: 10         # Number of top trends to process (e.g., 10)
  days_ago: 30           # How many days back to search for articles (e.g., 30)
  nlp_id: nlp/f/topic/1874  # NLP topic ID for the dashboard (used in trend discovery)
```
- **api_key**: Your Feedly API key (required).
- **num_trends**: How many top trends to process (default: 20).
- **days_ago**: How far back to search for articles, in days (default: 30).
- **nlp_id**: The NLP topic ID used for trend discovery. Change this to target a different dashboard in Feedly.

## How to Run
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Edit `config.yaml` with your Feedly API key and desired settings.
3. Run the script:
   ```bash
   python articles_to_xlsx.py
   ```
4. Open `article_details.xlsx` to review the results.

## Output Structure
- **Summary tab**: Each row is a trend, with columns for aliases, total articles, and counts for regulatory, expert, and conference articles.
- **Trend tabs**: Each tab is named after the trend (first alias or label). Columns include:
  - `id`: Article ID
  - `title`: Article title
  - `crawled`: Date the article was indexed (human-readable)
  - `url`: Canonical URL or alternate link
  - `content`: Cleaned article content (HTML stripped)
  - `type`: Regulatory, Expert Mentions, Conferences, or blank

## Notes
- Only trends with at least one article are included in the output.
- The script automatically handles pagination and HTML cleaning.
- If a trend has no aliases, its label is used as the search term and tab name.
- The `nlp_id` config option allows you to target different Feedly dashboards for trend discovery.

## Troubleshooting
- If you see empty cells in Excel, try enabling text wrap or expanding the row height.
- If you encounter API errors, check your API key and network connection.

---

**Questions or issues?** Contact your Feedly Customer Success Manager. 