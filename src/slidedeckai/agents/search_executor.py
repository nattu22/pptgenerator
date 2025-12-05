# slidedeckai/agents/search_executor.py - SMART & CLEAN
"""
Provides web search capabilities using LLMs to simulate search results.

This module is responsible for retrieving factual, quantitative data for
the presentation content. It prompts an LLM to simulate a web search
and return formatted facts.
"""
import logging
import re
from typing import List, Dict, Union
from openai import OpenAI

logger = logging.getLogger(__name__)


class WebSearchExecutor:
    """
    Executes simulated web searches to retrieve quantitative data.

    Ideally, this would connect to a real search engine API. Currently, it
    uses an LLM to generate plausible (or halluncinated, if not careful)
    factual data based on its training data.
    """
    
    def __init__(self, api_key: str):
        """
        Initialize the WebSearchExecutor.

        Args:
            api_key (str): OpenAI API key.
        """
        self.client = OpenAI(api_key=api_key)
        # Use GPT-5 family for web search extraction to maximize factual recall
        # Note: runtime may require provider model mapping; this is the logical model selection.
        # This will likely default to gpt-4o-mini in a real environment if gpt-5-mini doesn't exist yet.
        self.model = "gpt-4o-mini" # Adjusted to a known model for reliability
    
    def execute_searches(self, queries: Union[List[str], str]) -> Dict[str, List[str]]:
        """
        Execute multiple searches in parallel and return extracted facts.

        Args:
            queries (Union[List[str], str]): A single query string or a list of query strings.

        Returns:
            Dict[str, List[str]]: A dictionary mapping each query to a list of extracted facts.
        """
        from concurrent.futures import ThreadPoolExecutor, as_completed

        results = {}

        # Normalize input to list
        if isinstance(queries, str):
            queries = [queries]
        elif queries is None:
            return {}

        # Run searches in parallel to speed up IO-bound LLM calls
        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_query = {executor.submit(self._search_with_gpt, q): q for q in queries}
            for future in as_completed(future_to_query):
                q = future_to_query[future]
                try:
                    facts = future.result()
                    results[q] = facts
                    logger.info(f"  ✓ {q}: {len(facts)} facts")
                except Exception as e:
                    logger.error(f"  ✗ {q} failed: {e}")
                    results[q] = [f"Data for {q}: Contact financial analyst"]

        return results

    def _search_with_gpt(self, query: str) -> List[str]:
        """
        Simulate a web search using GPT to find quantitative facts.

        Args:
            query (str): The search query.

        Returns:
            List[str]: A list of formatted fact strings.
        """
        
        prompt = f"""Find 3-5 QUANTITATIVE facts for: {query}

MUST include:
- Specific numbers with units ($, %, M, B)
- Timeframes (Q4 2024, FY2024, etc.)
- Source context

Format: [Metric]: [Value] ([Timeframe])

Example: Revenue: $119.6B (Q4 2024)"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You find quantitative business data. Always include numbers and timeframes."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=500
            )
            
            content = response.choices[0].message.content.strip()
            facts = [line.strip() for line in content.split('\n') if line.strip() and len(line.strip()) > 20]
            
            return facts[:5]
            
        except Exception as e:
            logger.error(f"Search failed: {e}")
            return [f"{query}: See latest financial reports"]
