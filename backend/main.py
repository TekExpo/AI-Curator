import os
import asyncio
import logging

from dotenv import load_dotenv
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

# local imports
from schemas import ArticlesResponse, ArticleResponse, SuggestTopicsResponse
from content_fetcher import ContentFetcher
from summarizer import ArticleSummarizer

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

load_dotenv()
logger.info("✓ Environment loaded")

# Initialize OpenAI client once at startup
summarizer = ArticleSummarizer()
logger.info("✓ ArticleSummarizer initialized globally")

# Initialize app
app = FastAPI(title="AI-Powered Knowledge Curator API", version="2.0.0")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



@app.get("/suggest-topics", response_model=SuggestTopicsResponse)
async def suggest_topics(query: str = Query(..., description="Free-text user query for topic suggestions")):
    """
    Suggest relevant topics based on a free-text query using LLM.
    Returns up to 10 topics.
    """
    if not summarizer.available:
        return {"topics": [], "articles": []}
        
    prompt = f"""
    Suggest 10 concise, relevant technology topics based on the following user query. Return only a comma-separated list of topics, no explanations or extra text.

    User Query: {query}
    """
    try:
        response = summarizer.client.chat.completions.create(
            model=summarizer.model,
            messages=[
                {"role": "system", "content": "You are an expert assistant that suggests relevant technology topics."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=100
        )
        topics_text = response.choices[0].message.content
        topics = [t.strip() for t in topics_text.split(",") if t.strip()][:10]
        return {"topics": topics, "articles": []}
    except Exception as e:
        logger.error(f"✗ Error suggesting topics: {e}")
        return {"topics": [], "articles": []}

@app.get("/articles", response_model=ArticlesResponse)
async def get_articles(
    topics: str = Query(..., description="Comma-separated list of topics"),
    skip: int = Query(0, ge=0, description="Number of articles to skip (for pagination)"),
    limit: int = Query(10, ge=1, le=20, description="Max number of articles to return"),
):
    """
    Retrieve up to 20 latest articles for the given topics (comma-separated),
    ensuring at least one article per topic per source if available.
    Supports pagination via skip/limit.
    Fetches fresh articles on every call.
    """
    # Clean topics immediately (strip quotes, whitespace)
    topic_list = [t.strip().strip('"').strip("'") for t in topics.split(",") if t.strip()]

    # Calculate per-topic limit
    per_topic_limit = max(1, limit // len(topic_list)) if topic_list else limit

    # Fetch articles for all topics in parallel with per-topic limit
    fetch_tasks = [ContentFetcher.fetch_articles_for_topic(topic, skip=0, limit=per_topic_limit) for topic in topic_list]
    fetched_lists = await asyncio.gather(*fetch_tasks)

    # Summarize and tag articles per topic
    all_articles = []
    for topic, fetched in zip(topic_list, fetched_lists):
        if fetched:
            # Group articles by source
            source_groups = {}
            for article in fetched:
                src = article.get("source", "Unknown")
                source_groups.setdefault(src, []).append(article)

            # Round-robin pick articles from each source to fill per_topic_limit
            limited_fetched = []
            idx = 0
            while len(limited_fetched) < per_topic_limit:
                added_any = False
                for src, articles in source_groups.items():
                    if idx < len(articles):
                        limited_fetched.append(articles[idx])
                        added_any = True
                        if len(limited_fetched) >= per_topic_limit:
                            break
                if not added_any:
                    break
                idx += 1

            # Ensure each article is tagged with the correct topic
            for article in limited_fetched:
                article["topic"] = topic

            # Summarize only the limited articles
            summarized = await summarizer.batch_summarize_async(limited_fetched, target_topic=topic)
            all_articles.extend(summarized)

    # Deduplicate by URL to ensure each article appears only once in the aggregated feed
    seen = set()
    unique_articles = []
    for article in all_articles:
        # Check base URL to avoid duplicates even with different topic match params
        url = article.get("url", "")
        base_url = url.split("?")[0] if "?" in url else url
        if base_url and base_url not in seen:
            seen.add(base_url)
            unique_articles.append(article)

    # Sort all articles by published date (desc)
    unique_articles.sort(key=lambda x: x.get("published", ""), reverse=True)

    # Pagination
    paged = unique_articles[skip:skip+limit]

    # Format response
    articles = [
        ArticleResponse(
            id=i+1,
            title=a.get("title", "Untitled"),
            summary=a.get("summary", "No summary available"),
            url=a.get("url", ""),
            source=a.get("source", "Unknown"),
            published=a.get("published", ""),
            topic=a.get("topic", ""),
            tags=a.get("tags")
        )
        for i, a in enumerate(paged)
    ]
    return ArticlesResponse(articles=articles)

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    logger.info("=" * 60)
    logger.info("Starting FastAPI server...")
    logger.info(f"Access documentation at: http://0.0.0.0:{port}/docs")
    logger.info("=" * 60)
    try:
        uvicorn.run(app, host="0.0.0.0", port=port)
    except Exception as e:
        logger.error(f"✗ Server error: {e}")
        raise
