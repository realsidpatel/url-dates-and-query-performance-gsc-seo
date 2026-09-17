# url-dates-and-query-performance-gsc-seo
The program extracts the published dates of the provided URLs in its list, and the GSC program will fetch the top 10 queries and organise them by impression metrics in descending order.  It will also mark common queries if it is appearing across multiple URLs.

## Run this first in running in Google Colab:
```python
!pip -q install requests beautifulsoup4 openpyxl
```

## The following Python program requires a URL list, and it does the following:
- Checks when the page was created
- Checks when the page was updated
- If 404, it will not fetch any dates
- It has a defined date range in the code: 20 August 2025 to 20 August 2026
- It then checks the top 10 queries by impressions for which the URL is ranking
- If a common query is found across multiple URLs, it marks them red
- Formatting of red is not applicable in CSV output
- It provides 2 download options back-to-back: XLSX and CSV

## Notes:
- No file is dependent on the others.
- I ran this program in Google Colab. So each file is basically a block.
- The flow is auth.py > workfile.py
- Make sure to update the URL list. To update it, use ChatGPT to do it quickly.
