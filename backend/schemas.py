from pydantic import BaseModel
from typing import List

class ArticleResponse(BaseModel):
    id: int
    title: str
    summary: str
    url: str
    source: str
    published: str
    topic: str
    tags: str = None

class ArticlesResponse(BaseModel):
    articles: List[ArticleResponse]

class SuggestTopicsResponse(ArticlesResponse):
    topics: List[str]