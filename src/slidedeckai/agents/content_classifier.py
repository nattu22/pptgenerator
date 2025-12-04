import logging
from typing import List
from openai import OpenAI

logger = logging.getLogger(__name__)

class ContentTypeClassifier:
    """Intelligently selects content type based on data and placeholder"""
    
    def __init__(self, api_key: str):
        self.client = OpenAI(api_key=api_key)
    
    def select_content_type(
        self,
        content_description: str,
        search_queries: List[str],
        placeholder_type: str,
        optimal_types: List[str]
    ) -> str:
        """Select best content type for this placeholder"""
        
        # Rule-based selection first
        if 'comparison' in content_description.lower() or 'vs' in content_description.lower():
            if 'table' in optimal_types:
                return 'comparison_table'
            elif any('chart' in t for t in optimal_types):
                return 'bar_chart'
        
        if 'trend' in content_description.lower() or 'over time' in content_description.lower():
            if any('chart' in t for t in optimal_types):
                return 'line_chart'
        
        if 'breakdown' in content_description.lower() or 'distribution' in content_description.lower():
            if 'pie_chart' in optimal_types:
                return 'pie_chart'
            elif 'column_chart' in optimal_types:
                return 'column_chart'
        
        if any(word in content_description.lower() for word in ['metric', 'kpi', 'number', 'stat']):
            if 'kpi' in optimal_types:
                return 'kpi'
        
        # Default to first optimal type
        return optimal_types[0] if optimal_types else 'text'