"""
Utilities for processing various file formats uploaded by the user.

This module handles extracting text from documents (TXT, CSV, Excel) and
extracting structured chart data from images or data files using LLMs.
"""
import pandas as pd
from PIL import Image
import io
import logging
from typing import Union, List, Dict, Optional, Any

logger = logging.getLogger(__name__)

class FileProcessor:
    """
    A class containing static methods for file processing.
    """

    @staticmethod
    def extract_text(file_storage: Any) -> str:
        """
        Extract text content from uploaded files (txt, csv, xlsx).

        Args:
            file_storage (Any): The Flask FileStorage object.

        Returns:
            str: The extracted text content.
        """
        try:
            filename = file_storage.filename.lower()
            if filename.endswith('.txt'):
                return file_storage.read().decode('utf-8')
            elif filename.endswith('.csv'):
                # Reset pointer just in case
                if hasattr(file_storage, 'stream'):
                    file_storage.stream.seek(0)
                else:
                    file_storage.seek(0)
                df = pd.read_csv(file_storage)
                return df.to_string()
            elif filename.endswith('.xlsx') or filename.endswith('.xls'):
                if hasattr(file_storage, 'stream'):
                    file_storage.stream.seek(0)
                else:
                    file_storage.seek(0)
                df = pd.read_excel(file_storage)
                return df.to_string()
            else:
                logger.warning(f"Unsupported file type for text extraction: {filename}")
                return ""
        except Exception as e:
            logger.error(f"Failed to extract text from {file_storage.filename}: {e}")
            return ""

    @staticmethod
    def extract_chart_data(file_storage: Any, client: Any, model: Optional[str] = None) -> Optional[Dict]:
        """
        Extract chart data from uploaded file (Image, Excel, CSV) using LLM/Vision.

        Args:
            file_storage (Any): The file storage object.
            client (Any): The OpenAI client.
            model (Optional[str]): The model to use.

        Returns:
            Optional[Dict]: A JSON object suitable for chart generation, or None if extraction failed.
        """
        from slidedeckai.global_config import GlobalConfig
        if not model:
            model = GlobalConfig.LLM_MODEL_FAST

        filename = file_storage.filename.lower()
        content = ""

        try:
            if filename.endswith(('.png', '.jpg', '.jpeg', '.webp')):
                # Process image with GPT Vision
                # We need to base64 encode the image or pass the URL if it were hosted,
                # but here we have the file stream.
                import base64
                file_storage.stream.seek(0)
                image_data = base64.b64encode(file_storage.read()).decode('utf-8')

                response = client.chat.completions.create(
                    model=GlobalConfig.LLM_MODEL_VISION, # Use vision capable model
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": "Analyze this chart image and extract the data points. Return a JSON with 'title', 'type' (bar, column, line, pie), 'categories' (list of strings), and 'series' (list of objects with 'name' and 'values')."},
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_data}"}}
                            ]
                        }
                    ],
                    max_tokens=500,
                    response_format={"type": "json_object"}
                )
                import json
                return json.loads(response.choices[0].message.content)

            elif filename.endswith('.csv'):
                file_storage.stream.seek(0)
                df = pd.read_csv(file_storage)
                content = df.to_string()
            elif filename.endswith('.xlsx') or filename.endswith('.xls'):
                file_storage.stream.seek(0)
                df = pd.read_excel(file_storage)
                content = df.to_string()
            elif filename.endswith('.txt'):
                file_storage.stream.seek(0)
                content = file_storage.read().decode('utf-8')

            if content:
                # Use LLM to structure data
                prompt = f"""Extract chart data from this content:

{content[:5000]} # Limit content length

Return ONLY valid JSON:
{{
  "title": "Chart Title",
  "type": "column", # or bar, line, pie
  "categories": ["Cat1", "Cat2"],
  "series": [
    {{"name": "Series 1", "values": [10, 20]}}
  ]
}}"""
                response = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": "Extract chart data to JSON."},
                        {"role": "user", "content": prompt}
                    ],
                    response_format={"type": "json_object"}
                )
                import json
                return json.loads(response.choices[0].message.content)

        except Exception as e:
            logger.error(f"Failed to extract chart data from {filename}: {e}")
            return None
