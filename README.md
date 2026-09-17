# url-dates-and-query-performance-gsc-seo
The program extract the published dates of the provided URL in its lists and the GSC program will fetch top 10 queries and organise them by impression metrics in decending order.  It will also mark common queries if it is appearing across multiple URLs.


## The following Python program required URL list and it does the following:
- Checks when page was created
- Checks when page was updated
- If 404, it will not fetch any dates
- It has defined date range in code: 20 August 2025 to 20 August 2026
- It then checks top 10 queries by impression for which the URL is ranking for
- If a common query is found across multiple URL, it marks them red
- Formatting of red is not applicable in CSV output
- It provides 2 download options back to back: XLSX and CSV
