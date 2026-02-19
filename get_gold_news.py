"""Fetch recent gold news for this week and save to JSON.

Usage: python get_gold_news.py
"""
import feedparser
import datetime
import time
import json
from urllib.parse import urljoin

# Google News RSS search for 'gold price' limited to last 7 days
RSS_URL = (
    "https://news.google.com/rss/search?q=gold+price+when:7d&hl=en-US&gl=US&ceid=US:en"
)

def fetch_news(rss_url=RSS_URL):
    feed = feedparser.parse(rss_url)
    results = []
    now = datetime.datetime.utcnow()
    week_ago = now - datetime.timedelta(days=7)

    for entry in feed.entries:
        # published_parsed is a time.struct_time if available
        pub_parsed = getattr(entry, "published_parsed", None)
        if pub_parsed:
            pub_dt = datetime.datetime.utcfromtimestamp(time.mktime(pub_parsed))
        else:
            pub_dt = None

        # Build result
        link = entry.get("link")
        # Some Google News links are relative; make absolute if needed
        if link and link.startswith("./"):
            link = urljoin("https://news.google.com/", link)

        results.append(
            {
                "title": entry.get("title"),
                "link": link,
                "summary": entry.get("summary"),
                "published": pub_dt.isoformat() if pub_dt else None,
                "source": entry.get("source", {}).get("title") if entry.get("source") else None,
            }
        )

    return results


def main():
    news = fetch_news()
    out_file = "gold_news_this_week.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(news, f, ensure_ascii=False, indent=2)

    print(f"Saved {len(news)} items to {out_file}")
    for i, item in enumerate(news[:10], 1):
        print(f"{i}. {item['title']}")
        print(f"   {item['link']}")


if __name__ == "__main__":
    main()
