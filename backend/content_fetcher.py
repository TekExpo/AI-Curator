import re
import httpx
import logging
import feedparser
import urllib.parse
from datetime import datetime
from typing import List, Dict

logger = logging.getLogger(__name__)

class ContentFetcher:
    """Fetches content from selected tech RSS sources"""

    GENERIC_TECH_FEEDS = {
        "Dev.to": "https://dev.to/api/articles?tag={topic}&per_page=10",
        "Medium Tech": "https://medium.com/feed/tag/",
    }
    
    @staticmethod
    async def fetch_from_learn_api(topic: str, skip: int = 0, limit: int = 10) -> List[Dict]:
        """Fetch from Microsoft Learn API using multiple strategies with pagination support"""
        articles = []

        # Strategy 1: Direct Learn API search
        try:
            clean_topic = topic.strip().strip('"').strip("'")
            # Microsoft Learn API uses 'skip' for pagination and 'top' for items per request
            url = f"https://learn.microsoft.com/api/search?search={clean_topic.replace(' ', '%20')}&top={limit}&skip={skip}&locale=en-us&target=search"
            logger.info(f"[Learn API] Fetching: {topic}")

            async with httpx.AsyncClient(timeout=15.0, follow_redirects=True) as client:
                response = await client.get(url, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                })

                if response.status_code == 200:
                    try:
                        data = response.json()
                        if "results" in data:
                            results = data.get("results", [])
                            logger.info(f"[Learn API] Got {len(results)} raw results (skip={skip}, limit={limit})")
                            for idx, item in enumerate(results, 1):
                                if item.get("title"):
                                    articles.append({
                                        "title": item.get("title", ""),
                                        "url": item.get("url", ""),
                                        "summary": item.get("description", item.get("summary", "")),
                                        "source": "Microsoft Learn",
                                        "topic": topic,
                                        "tags": item.get("category", ""),
                                        "published": datetime.utcnow().isoformat()
                                    })
                            if articles:
                                logger.info(f"[Learn API] Found {len(articles)} articles for: {topic}")
                                return articles
                    except Exception as parse_err:
                        logger.warning(f"[Learn API] Failed to parse response: {parse_err}")

        except Exception as e:
            logger.warning(f"[Learn API] Error fetching from {topic}: {e}")

        logger.info(f"[Learn] No articles found for: {topic}")
        return articles

    @staticmethod
    async def fetch_from_rss(feed_url: str, topic: str, source: str) -> List[Dict]:
        """Fetch from RSS feeds"""
        try:
            logger.info(f"Fetching from {source}: {topic}")
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}
            async with httpx.AsyncClient(timeout=10.0, follow_redirects=True, headers=headers) as client:
                response = await client.get(feed_url)
                if response.status_code == 200:
                    feed = feedparser.parse(response.content)
                    articles = []
                    
                    matched_entries = []
                    clean_topic = topic.strip().strip('"').strip("'").lower()
                    match_pattern = re.compile(rf'\b{re.escape(clean_topic)}\b', re.IGNORECASE)
                    
                    for entry in feed.entries:
                        title = entry.get("title", "")
                        summary = entry.get("summary", "")
                        
                        entry_tags = []
                        if hasattr(entry, "tags"):
                            entry_tags = [t.term for t in entry.tags if hasattr(t, "term")]
                        
                        is_match = match_pattern.search(title) or match_pattern.search(summary) or \
                                   any(match_pattern.search(tag) for tag in entry_tags)
                        
                        if is_match:
                            matched_entries.append(entry)

                    for entry in matched_entries[:5]: 
                        original_link = entry.get("link", "")
                        url_topic = clean_topic.replace(" ", "-")
                        
                        if "?" in original_link:
                            db_safe_link = f"{original_link}&_topic_match={url_topic}"
                        else:
                            db_safe_link = f"{original_link}?_topic_match={url_topic}"

                        tags_list = []
                        if hasattr(entry, "tags"):
                            for t in entry.tags:
                                if "term" in t:
                                    tags_list.append(t.term)
                        tags_str = ", ".join(tags_list)
                        
                        articles.append({
                            "title": entry.get("title", ""),
                            "url": db_safe_link,
                            "summary": entry.get("summary", ""),
                            "source": source,
                            "topic": topic,
                            "tags": tags_str,
                            "published": entry.get("published", datetime.utcnow().isoformat())
                        })
                    logger.info(f"✓ Found {len(articles)} articles from {source}")
                    return articles
                else:
                    logger.warning(f"HTTP {response.status_code} fetching from {source} URL: {feed_url}")
        except Exception as e:
            logger.warning(f"Error fetching from {source}: {e}")

        return []

    @staticmethod
    async def fetch_from_devto(topic: str) -> List[Dict]:
        """Fetch from Dev.to API"""
        try:
            clean_topic = topic.strip().strip('"').strip("'").lower().replace(" ", "-")
            url = f"https://dev.to/api/articles?tag={clean_topic}&per_page=10"
            logger.info(f"Fetching from Dev.to: {topic}")
            
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
            async with httpx.AsyncClient(timeout=10.0, headers=headers) as client:
                response = await client.get(url)
                if response.status_code == 200:
                    articles = []
                    try:
                        data = response.json()
                        if isinstance(data, list):
                            for item in data[:5]:
                                if item.get("title"):
                                    articles.append({
                                        "title": item.get("title", ""),
                                        "url": item.get("url", ""),
                                        "summary": item.get("description", item.get("body_markdown", "")[:200]),
                                        "source": "Dev.to",
                                        "topic": topic,
                                        "tags": ", ".join(item.get("tag_list", [])) if item.get("tag_list") else "",
                                        "published": item.get("published_at", datetime.utcnow().isoformat())
                                    })
                            if articles:
                                logger.info(f"✓ Found {len(articles)} articles from Dev.to")
                                return articles
                    except Exception as parse_err:
                        logger.warning(f"Failed to parse Dev.to response: {parse_err}")
                else:
                    logger.warning(f"HTTP {response.status_code} fetching from Dev.to")
        except Exception as e:
            logger.warning(f"Error fetching from Dev.to: {e}")
        
        return []

    @staticmethod
    async def fetch_articles_for_topic(topic: str, skip: int = 0, limit: int = 10) -> List[Dict]:
        """Fetch articles for a given topic with pagination support"""
        clean_topic = topic.strip().strip('"').strip("'")
        logger.info(f"======== Fetching (skip={skip}, limit={limit}) for: {clean_topic} ========")
        all_articles = []

        # Fetch at least three articles from each source
        per_source_min = 3
        per_source_articles = []

        # Microsoft Learn API
        try:
            learn_articles = await ContentFetcher.fetch_from_learn_api(topic, skip=0, limit=per_source_min)
            per_source_articles.append(learn_articles)
        except Exception as e:
            logger.warning(f"Error fetching from Learn API: {e}")

        # Other sources
        for source_name, feed_url in ContentFetcher.GENERIC_TECH_FEEDS.items():
            try:
                if source_name == "Dev.to":
                    articles = await ContentFetcher.fetch_from_devto(topic)
                    logger.info(f"Dev.to returned {len(articles)} articles for topic '{topic}'")
                elif "medium.com/feed/tag" in feed_url:
                    tag_slug = clean_topic.lower().replace(" ", "-")
                    # if source_name == "Medium Tech":
                    target_url = f"https://medium.com/feed/tag/{tag_slug}"
                    # else:
                    #     target_url = feed_url
                    articles = await ContentFetcher.fetch_from_rss(target_url, topic, source_name)
                else:
                    articles = await ContentFetcher.fetch_from_rss(feed_url, topic, source_name)
                per_source_articles.append(articles[:per_source_min])
            except Exception as e:
                logger.warning(f"Error fetching from {source_name}: {e}")


        
        # Combine per-source articles ensuring at least per_source_min per source
        combined_articles = []
        seen_urls = set()
        for source_list in per_source_articles:
            for article in source_list:
                url = article.get("url", "")
                if url and url not in seen_urls:
                    seen_urls.add(url)
                    combined_articles.append(article)

        # If combined less than limit, fetch more from all sources to fill
        if len(combined_articles) < limit:
            remaining = limit - len(combined_articles)
            # Fetch more from Learn API
            try:
                more_learn = await ContentFetcher.fetch_from_learn_api(topic, skip=per_source_min, limit=remaining)
                for article in more_learn:
                    url = article.get("url", "")
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        combined_articles.append(article)
                        remaining -= 1
                        if remaining <= 0:
                            break
            except Exception as e:
                logger.warning(f"Error fetching additional Learn API articles: {e}")

            # Fetch more from other sources
            for source_name, feed_url in ContentFetcher.GENERIC_TECH_FEEDS.items():
                if remaining <= 0:
                    break
                try:
                    if source_name == "Dev.to":
                        more_articles = await ContentFetcher.fetch_from_devto(topic)
                    # elif source_name == "LinkedIn Learning":
                    #     more_articles = await ContentFetcher.fetch_from_linkedin_learning(topic)
                    elif "medium.com/feed/tag" in feed_url:
                        tag_slug = clean_topic.lower().replace(" ", "-")
                        if source_name == "Medium Tech":
                            target_url = f"https://medium.com/feed/tag/{tag_slug}"
                        else:
                            target_url = feed_url
                        more_articles = await ContentFetcher.fetch_from_rss(target_url, topic, source_name)
                    else:
                        more_articles = await ContentFetcher.fetch_from_rss(feed_url, topic, source_name)
                    for article in more_articles:
                        url = article.get("url", "")
                        if url and url not in seen_urls:
                            seen_urls.add(url)
                            combined_articles.append(article)
                            remaining -= 1
                            if remaining <= 0:
                                break
                except Exception as e:
                    logger.warning(f"Error fetching additional articles from {source_name}: {e}")

        logger.info(f"======== Total: {len(combined_articles)} unique articles for {topic} ========")
        return combined_articles
