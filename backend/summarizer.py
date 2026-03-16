import os
import json
import asyncio
import logging
from typing import List, Dict

from openai import AzureOpenAI, OpenAI

logger = logging.getLogger(__name__)

class ArticleSummarizer:
    """Generates AI-powered summaries and extracts high-quality metadata for articles"""

    def __init__(self):
        # Support both Azure OpenAI and regular OpenAI
        self.use_azure = os.getenv("AZURE_OPENAI_KEY") is not None
        self.client = None
        self.model = None
        self.available = False

        try:
            if self.use_azure:
                logger.info("Initializing Azure OpenAI...")
                self.client = AzureOpenAI(
                    api_key=os.getenv("AZURE_OPENAI_KEY"),
                    api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2023-05-15"),
                    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
                )
                self.model = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-35-turbo")
                logger.info(f"✓ Azure OpenAI initialized (model: {self.model})")
                self.available = True
            else:
                api_key = os.getenv("OPENAI_API_KEY")
                if not api_key:
                    logger.warning("⚠ No OPENAI_API_KEY found - summarization disabled")
                    self.available = False
                else:
                    logger.info("Initializing OpenAI...")
                    self.client = OpenAI(api_key=api_key)
                    self.model = "gpt-3.5-turbo"
                    logger.info(f"✓ OpenAI initialized (model: {self.model})")
                    self.available = True
        except Exception as e:
            logger.error(f"✗ Failed to initialize AI client: {e}")
            self.available = False

    async def _summarize_batch(self, articles_batch: List[Dict], semaphore: asyncio.Semaphore, target_topic: str) -> List[Dict]:
        """Internal helper to process a batch of articles through the LLM"""
        async with semaphore:
            logger.debug(f"Summarizing batch of {len(articles_batch)} articles...")
            
            # Build combined prompt for batch
            combined_prompt = f"Please analyze the following {len(articles_batch)} technical articles for relevance to the target topic: '{target_topic}'.\n\n"
            for i, article in enumerate(articles_batch, 1):
                combined_prompt += f"Article {i} Title: {article.get('title', '')}\n"
                combined_prompt += f"Article {i} Content: {article.get('summary', '')}\n\n"

            combined_prompt += (
                "Return EXACTLY one JSON object with a single key called 'summaries'.\n"
                f"The value of 'summaries' must be an array of exactly {len(articles_batch)} objects.\n"
                "Each object must include all of the following fields:\n"
                "- 'is_relevant': boolean (false if the article does NOT meaningfully address the target topic, is purely marketing fluff, or is unrelated to the technical focus).\n"
                "- 'summary': a concise 2-3 sentence summary focused on actionable developer insights and key technical takeaways.\n"
                "- 'tags': a comma-separated list of 3-5 specific, technical tags. Avoid generic tags such as 'Documentation' or 'News'.\n"
                "If an article is not relevant, set 'is_relevant' to false, and return an empty string for 'summary' and 'tags'.\n"
                "Return ONLY the JSON object, with no extra text, markdown, or explanation.\n"
            )

            try:
                response = await asyncio.to_thread(self.client.chat.completions.create,
                    model=self.model,
                    messages=[
                        {"role": "system", "content": f"You are a technical editor. Output a valid JSON object with a 'summaries' array containing exactly {len(articles_batch)} items."},
                        {"role": "user", "content": combined_prompt}
                    ],
                    response_format={ "type": "json_object" } if self.use_azure or (not self.use_azure and "gpt-4" in self.model) else None,
                    temperature=0.3,
                    max_tokens=750
                )
                
                response_text = response.choices[0].message.content.strip()

                # Clean markdown formatting if present
                if response_text.startswith("```json"): response_text = response_text[7:]
                if response_text.startswith("```"): response_text = response_text[3:]
                if response_text.endswith("```"): response_text = response_text[:-3]
                
                data = json.loads(response_text.strip())
                
                # Extract results array from common wrapper keys
                results = []
                if isinstance(data, list):
                    results = data
                elif isinstance(data, dict):
                    results = data.get("summaries", data.get("results", data.get("articles", [])))
                    if not results and all(str(i) in data for i in range(len(articles_batch))):
                        results = [data[str(i)] for i in range(len(articles_batch))]
                        logger.info(results)

                if not isinstance(results, list) or len(results) != len(articles_batch):
                    logger.error(f"Batch processing mismatch: expected {len(articles_batch)}, got {len(results)}")
                    return articles_batch

                for article, result in zip(articles_batch, results):
                    if not result.get("is_relevant", True):
                        article["summary"] = "(Irrelevant article)"
                        article["tags"] = ""
                    else:
                        article["summary"] = result.get("summary", article.get("summary", ""))
                        article["tags"] = result.get("tags", "")
                return articles_batch

            except Exception as e:
                logger.error(f"✗ Error batch summarizing articles: {e}")
                return articles_batch

    async def batch_summarize_async(self, articles: List[Dict], target_topic: str) -> List[Dict]:
        """Public method to summarize multiple articles asynchronously using batching"""
        if not self.available:
            return articles
            
        logger.info(f"Summarizing {len(articles)} articles for topic: {target_topic}")
        semaphore = asyncio.Semaphore(3) # Limit concurrency
        
        # Split articles into batches of 3
        batch_size = 3
        batches = [articles[i:i + batch_size] for i in range(0, len(articles), batch_size)]
        
        tasks = [self._summarize_batch(batch, semaphore, target_topic) for batch in batches]
        summarized_batches = await asyncio.gather(*tasks)
        
        # Flatten and return
        summarized = [article for batch in summarized_batches for article in batch]
        logger.info(f"✓ Completed summarizing {len(summarized)} articles")
        return summarized

    def summarize_article(self, title: str, content: str, target_topic: str, max_bullet_points: int = 3) -> dict:
        """Legacy synchronous method for single article summarization (wraps call to client)"""
        if not self.available:
            return {"summary": "• Summarization not available", "tags": "", "irrelevant": False}

        try:

            prompt = f"""
                        You are an AI assistant that analyzes technical articles.

                        Your task is to determine whether the article is relevant to the target topic and produce a concise technical summary.

                        TARGET TOPIC:
                        {target_topic}

                        ARTICLE:
                        Title: {title}

                        Content:
                        {content}

                        INSTRUCTIONS:
                        1. Determine if the article is meaningfully related to the target topic.
                        2. If relevant, summarize the key technical insights.
                        3. Focus on concrete concepts such as technologies, frameworks, tools, architectures, APIs, or implementation practices.
                        4. Avoid marketing language and generic statements.
                        5. Use clear, concise technical language.

                        OUTPUT REQUIREMENTS:
                        Return ONLY a valid JSON object with the following structure:

                        {{
                        "is_relevant": true | false,
                        "summary": [
                            "• bullet point 1",
                            "• bullet point 2",
                            "• bullet point 3"
                        ],
                        "tags": ["tag1", "tag2", "tag3"]
                        }}

                        RULES:
                        - "summary" MUST contain exactly {max_bullet_points} bullet points.
                        - Each bullet must start with "• ".
                        - Each bullet must be under 25 words.
                        - "tags" must contain 3–5 short technical keywords.
                        - If the article is NOT relevant, return:
                        - "is_relevant": false
                        - "summary": []
                        - "tags": []
                        - Do NOT include explanations, markdown, or text outside the JSON object.
                    """
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Return valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                response_format={ "type": "json_object" } if self.use_azure or (not self.use_azure and "gpt-4" in self.model) else None,
                temperature=0.3,
                max_tokens=250
            )

            result = json.loads(response.choices[0].message.content)
            if not result.get("is_relevant", True):
                return {"irrelevant": True}

            return {
                "summary": result.get("summary", ""),
                "tags": result.get("tags", ""),
                "irrelevant": False
            }

        except Exception as e:
            logger.error(f"✗ Single summary error: {e}")
            return {"summary": "• AI summary unavailable", "tags": "", "irrelevant": False}

    def batch_summarize(self, articles: List[Dict], target_topic: str) -> List[Dict]:
        """Legacy synchronous batch method"""
        logger.info(f"Summarizing {len(articles)} articles (Sync)...")
        summarized = []
        for article in articles:
            res = self.summarize_article(article.get("title", ""), article.get("summary", ""), target_topic)
            if not res.get("irrelevant"):
                article["summary"] = res.get("summary")
                article["tags"] = res.get("tags")
                summarized.append(article)
        return summarized
