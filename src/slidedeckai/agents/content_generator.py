# slidedeckai/agents/content_generator.py - SMART CONTENT GENERATION
"""
Generate actual slide content using GPT with quantitative data
"""
import logging
import json
from typing import List, Dict
from openai import OpenAI

logger = logging.getLogger(__name__)


class ContentGenerator:
    """
    Generate slide content using GPT
    Each method generates specific content type
    """
    
    def __init__(self, api_key: str):
        self.client = OpenAI(api_key=api_key)
        # Use GPT-4 family for content generation (best available GPT-4 model by default)
        self.model = "gpt-4.1-mini"
    
    def generate_subtitle(self, slide_title: str, purpose: str, 
                         search_facts: List[str]) -> str:
        """
        Generate contextual subtitle (2-5 words)
        """
        
        facts_text = "\n".join(search_facts[:3]) if search_facts else "No data"
        
        prompt = f"""Generate a SHORT subtitle (2-5 words) for this slide:

Title: {slide_title}
Purpose: {purpose}
Key Facts: {facts_text}

The subtitle should be:
- 2-5 words MAXIMUM
- Contextual to the data
- Professional tone

Return ONLY the subtitle text, nothing else."""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Generate concise subtitles."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.4,
                max_tokens=20
            )
            
            subtitle = response.choices[0].message.content.strip().strip('"\'')
            return subtitle if subtitle else "Key Insights"
            
        except Exception as e:
            logger.error(f"Subtitle generation failed: {e}")
            return "Analysis"
    
    def generate_bullets(self, slide_title: str, purpose: str,
                        search_facts: List[str], max_bullets: int = 5,
                        max_words_per_bullet: int = 15) -> List[str]:
        """
        Generate bullet points from search facts with strict length control
        """
        
        facts_text = "\n".join(search_facts) if search_facts else "No data available"
        
        prompt = f"""Generate {max_bullets} bullet points for this slide:

Title: {slide_title}
Purpose: {purpose}

Available Data:
{facts_text}

Requirements:
- Generate EXACTLY {max_bullets} bullet points
- Each bullet MUST be under {max_words_per_bullet} words to fit layout
- Include QUANTITATIVE data (numbers, percentages)
- Professional, executive-level tone
- NO preamble, ONLY bullet points

Return as plain text, one bullet per line."""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Generate concise, data-driven bullet points."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=300
            )
            
            content = response.choices[0].message.content.strip()
            bullets = [line.strip('- ').strip() for line in content.split('\n') 
                      if line.strip() and not line.startswith('```')]
            
            logger.info(f"        ✓ {len(bullets)} bullets")
            return bullets[:max_bullets]
            
        except Exception as e:
            logger.error(f"Bullet generation failed: {e}")
            return [f"Analysis of {slide_title}", "Key findings pending", "Data review in progress"]
    
    def generate_chart(self, slide_title: str, purpose: str,
                      search_facts: List[str], chart_type: str = 'column') -> Dict:
        """
        Generate chart data from search facts
        """
        
        facts_text = "\n".join(search_facts) if search_facts else "No data"
        
        prompt = f"""Generate chart data for: {slide_title}

Purpose: {purpose}
Chart Type: {chart_type}

Available Data:
{facts_text}

Create a chart with:
- 3-5 categories (short labels)
- 1-2 data series
- Real numbers from the facts above
- Meaningful title

Return ONLY valid JSON:
{{
  "title": "Chart Title",
  "type": "{chart_type}",
  "categories": ["Cat1", "Cat2", "Cat3"],
  "series": [
    {{"name": "Series 1", "values": [10, 20, 30]}}
  ]
}}"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Generate chart data in JSON format. Return ONLY valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                max_tokens=400,
                response_format={"type": "json_object"}
            )
            
            chart_data = json.loads(response.choices[0].message.content)
            logger.info(f"        ✓ Chart: {len(chart_data.get('categories', []))} cats")
            return chart_data
            
        except Exception as e:
            logger.error(f"Chart generation failed: {e}")
            return {
                "title": slide_title,
                "type": chart_type,
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [{"name": "Data", "values": [100, 120, 140, 160]}]
            }
    
    def generate_table(self, slide_title: str, purpose: str,
                      search_facts: List[str]) -> Dict:
        """
        Generate table data from search facts
        """
        
        facts_text = "\n".join(search_facts) if search_facts else "No data"
        
        prompt = f"""Generate table data for: {slide_title}

Purpose: {purpose}

Available Data:
{facts_text}

Create a comparison table with:
- 3-4 column headers
- 4-6 data rows
- Real numbers from facts
- Clear labels

Return ONLY valid JSON:
{{
  "headers": ["Metric", "Q3 2024", "Q4 2024"],
  "rows": [
    ["Revenue", "$X.XB", "$X.XB"],
    ["Profit", "$X.XB", "$X.XB"]
  ]
}}"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Generate table data in JSON. Return ONLY valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                max_tokens=500,
                response_format={"type": "json_object"}
            )
            
            table_data = json.loads(response.choices[0].message.content)
            logger.info(f"        ✓ Table: {len(table_data.get('headers', []))} cols")
            return table_data
            
        except Exception as e:
            logger.error(f"Table generation failed: {e}")
            return {
                "headers": ["Metric", "Value", "Change"],
                "rows": [
                    ["Revenue", "$XXB", "+X%"],
                    ["Profit", "$XXB", "+X%"]
                ]
            }
    
    def generate_kpi(self, slide_title: str, fact: str) -> Dict:
        """
        Generate KPI from a fact
        Extract: Big Number + Label
        """
        
        prompt = f"""Extract KPI from this fact:

Fact: {fact}
Context: {slide_title}

Extract:
- value: The BIG NUMBER (e.g., "$119.6B", "25%", "5M")
- label: Short description (3-5 words)

Return ONLY valid JSON:
{{
  "value": "$119.6B",
  "label": "Q4 Revenue"
}}"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Extract KPI data. Return ONLY valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=100,
                response_format={"type": "json_object"}
            )
            
            kpi_data = json.loads(response.choices[0].message.content)
            logger.info(f"        ✓ KPI: {kpi_data.get('label', 'N/A')}")
            return kpi_data
            
        except Exception as e:
            logger.error(f"KPI generation failed: {e}")
            return {"value": "N/A", "label": slide_title[:20]}

    def generate_speaker_notes(self, slide_title: str, bullets: List[str], key_facts: List[str]) -> str:
        """
        Generate conversational speaker notes
        """

        bullet_text = "\n- ".join(bullets) if bullets else "N/A"
        fact_text = "\n- ".join(key_facts[:3]) if key_facts else "N/A"

        prompt = f"""Generate speaker notes for this slide:

Title: {slide_title}

Visual Content:
- {bullet_text}

Supporting Data:
- {fact_text}

Requirements:
- Conversational tone ("Welcome to this slide...", "Here we see...")
- Explain the key points, don't just read them
- Add a transition sentence to the next topic if applicable
- Keep it under 150 words
- Professional and engaging

Return ONLY the speaker notes text."""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "Generate professional speaker notes."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=250
            )

            notes = response.choices[0].message.content.strip()
            logger.info(f"        ✓ Speaker notes generated")
            return notes

        except Exception as e:
            logger.error(f"Speaker notes generation failed: {e}")
            return f"Speaker notes for {slide_title}: Please cover the key points listed on the slide."