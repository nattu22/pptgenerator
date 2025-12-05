"""
Core functionality of SlideDeck AI.

This module provides the main `SlideDeckAI` class which orchestrates the generation
of slide decks, including interacting with LLMs, planning story structures,
and generating content for slides.
"""
import logging
import os
import pathlib
import tempfile
from typing import Union, Any, Dict, List, Optional
import time
import json5, json
from dotenv import load_dotenv
load_dotenv()

from . import global_config as gcfg
from .global_config import GlobalConfig
from .helpers import file_manager as filem
from .helpers import llm_helper, pptx_helper, text_helper
from .helpers.chat_helper import ChatMessageHistory
from .layout_analyzer import TemplateAnalyzer, LayoutCapability, PlaceholderInfo
from .content_matcher import ContentLayoutMatcher
from pptx import Presentation


RUN_IN_OFFLINE_MODE = os.getenv('RUN_IN_OFFLINE_MODE', 'False').lower() == 'true'
VALID_MODEL_NAMES = list(GlobalConfig.VALID_MODELS.keys())
VALID_TEMPLATE_NAMES = list(GlobalConfig.PPTX_TEMPLATE_FILES.keys())

logger = logging.getLogger(__name__)


def _process_llm_chunk(chunk: Any) -> str:
    """
    Helper function to process LLM response chunks consistently.

    Args:
        chunk: The chunk received from the LLM stream. It can be a string or an object with a 'content' attribute.

    Returns:
        str: The processed text content from the chunk.
    """
    if isinstance(chunk, str):
        return chunk

    content = getattr(chunk, 'content', None)
    return content if content is not None else str(chunk)


def _stream_llm_response(llm: Any, prompt: str, progress_callback=None) -> str:
    """
    Helper function to stream LLM responses with consistent handling.

    Args:
        llm: The LLM instance to use for generating responses.
        prompt: The prompt string to send to the LLM.
        progress_callback: Optional callback function that receives the current length of the response.

    Returns:
        str: The complete accumulated response from the LLM.

    Raises:
        RuntimeError: If there is an error during the streaming process.
    """
    response = ''
    try:
        for chunk in llm.stream(prompt):
            chunk_text = _process_llm_chunk(chunk)
            response += chunk_text
            if progress_callback:
                progress_callback(len(response))
        return response
    except Exception as e:
        logger.error('Error streaming LLM response: %s', str(e))
        raise RuntimeError(f'Failed to get response from LLM: {str(e)}') from e


class SlideDeckAI:
    """
    The main class for generating slide decks using AI.

    This class handles the end-to-end process of creating a presentation,
    from story planning and content generation to the final PPTX file creation.
    """

    def __init__(
            self,
            model: str,
            topic: str,
            api_key: str = None,
            pdf_path_or_stream=None,
            pdf_page_range=None,
            template_idx: int = 0
    ):
        """
        Initialize the SlideDeckAI object.

        Args:
            model (str): The name of the LLM model to use. Must be one of `GlobalConfig.VALID_MODELS`.
            topic (str): The main topic or title of the slide deck.
            api_key (str, optional): The API key for the LLM provider. Defaults to None.
            pdf_path_or_stream (Union[str, IO], optional): The path to a PDF file or a file-like object to use as source material. Defaults to None.
            pdf_page_range (tuple, optional): A tuple representing the page range to use from the PDF file. Defaults to None.
            template_idx (int, optional): The index of the PowerPoint template to use from `GlobalConfig.PPTX_TEMPLATE_FILES`. Defaults to 0.

        Raises:
            ValueError: If the model name is not in `VALID_MODELS`.
        """
        if model not in GlobalConfig.VALID_MODELS:
            raise ValueError(
                f'Invalid model name: {model}.'
                f' Must be one of: {", ".join(VALID_MODEL_NAMES)}.'
            )

        self.model: str = model
        self.topic: str = topic
        self.api_key: str = api_key
        self.pdf_path_or_stream = pdf_path_or_stream
        self.pdf_page_range = pdf_page_range
        # Validate template_idx is within valid range
        num_templates = len(GlobalConfig.PPTX_TEMPLATE_FILES)
        self.template_idx: int = template_idx if 0 <= template_idx < num_templates else 0
        self.chat_history = ChatMessageHistory()
        self.last_response = None
        logger.info('Using model: %s', model)

    def _initialize_llm(self) -> Any:
        """
        Initialize and return an LLM instance with the current configuration.

        Returns:
            Any: A configured LLM instance (e.g., from LiteLLM).
        """
        provider, llm_name = llm_helper.get_provider_model(
            self.model,
            use_ollama=RUN_IN_OFFLINE_MODE
        )

        return llm_helper.get_litellm_llm(
            provider=provider,
            model=llm_name,
            max_new_tokens=gcfg.get_max_output_tokens(self.model),
            api_key=self.api_key,
        )

    def _get_prompt_template(self, is_refinement: bool) -> str:
        """
        Retrieve the appropriate prompt template.

        Args:
            is_refinement (bool): Whether to get the refinement prompt template (True) or the initial prompt template (False).

        Returns:
            str: The content of the prompt template file.
        """
        if is_refinement:
            with open(GlobalConfig.REFINEMENT_PROMPT_TEMPLATE, 'r', encoding='utf-8') as in_file:
                template = in_file.read()
        else:
            with open(GlobalConfig.INITIAL_PROMPT_TEMPLATE, 'r', encoding='utf-8') as in_file:
                template = in_file.read()
        return template
    
    def _build_executive_story_plan(self, topic: str, template_name: str) -> Dict[str, Any]:
        """
        Plan the story structure before generating content.

        This method defines a sequence of sections for an executive presentation,
        tailoring the layout requirements for each section.

        Args:
            topic (str): The topic of the presentation.
            template_name (str): The name of the template being used.

        Returns:
            Dict[str, Any]: A dictionary containing the story plan, including sections and the analyzer instance.
        """
        
        # Get template analyzer
        template_file = GlobalConfig.PPTX_TEMPLATE_FILES[template_name]['file']
        presentation = Presentation(template_file)
        analyzer = TemplateAnalyzer(presentation)
        
        # Get available layouts sorted by executive suitability
        # exec_layouts = sorted(
        #     analyzer.layouts.items(),
        #     key=lambda x: x[1].executive_suitability,
        #     reverse=True
        # )
        
        # Build story sections (10-12 slides typical)
        # num_slides = 10
        
        sections = []
        
        # 1. OPENING: Strong visual opener
        sections.append({
            "type": "opening",
            "purpose": "Hook attention with key insight",
            "preferred_story": "focused_message",
            "content_type": "bullets",
            "layout_requirements": {
                "min_executive_score": 70,
                "preferred_types": ["focused_message", "data_visualization"]
            }
        })
        
        # 2. OVERVIEW: Set context
        sections.append({
            "type": "overview",
            "purpose": "Establish scope and framework",
            "preferred_story": "balanced_comparison",
            "content_type": "bullets",
            "layout_requirements": {
                "min_executive_score": 60,
                "preferred_types": ["balanced_comparison", "hierarchical_story"]
            }
        })
        
        # 3-4. DATA SECTIONS: Charts/tables
        sections.extend([
            {
                "type": "data_analysis",
                "purpose": "Present quantitative evidence",
                "preferred_story": "data_visualization",
                "content_type": "chart",
                "layout_requirements": {
                    "min_executive_score": 60,
                    "must_have": "chart_suitable",
                    "preferred_types": ["data_visualization"]
                }
            },
            {
                "type": "data_breakdown",
                "purpose": "Detailed data comparison",
                "preferred_story": "data_visualization",
                "content_type": "table",
                "layout_requirements": {
                    "min_executive_score": 50,
                    "must_have": "table_suitable",
                    "preferred_types": ["data_visualization", "hierarchical_story"]
                }
            }
        ])
        
        # 5-6. COMPARISON/ANALYSIS
        sections.extend([
            {
                "type": "comparison",
                "purpose": "Contrast key dimensions",
                "preferred_story": "balanced_comparison",
                "content_type": "double_column",
                "layout_requirements": {
                    "min_executive_score": 65,
                    "preferred_types": ["balanced_comparison"]
                }
            },
            {
                "type": "deep_dive",
                "purpose": "Detailed examination",
                "preferred_story": "detailed_analysis",
                "content_type": "bullets",
                "layout_requirements": {
                    "min_executive_score": 55,
                    "preferred_types": ["detailed_analysis", "hierarchical_story"]
                }
            }
        ])
        
        # 7. METRICS: KPI dashboard
        sections.append({
            "type": "metrics",
            "purpose": "Key performance indicators",
            "preferred_story": "metrics_dashboard",
            "content_type": "kpi_dashboard",
            "layout_requirements": {
                "min_executive_score": 70,
                "must_have": "kpi_grid",
                "preferred_types": ["metrics_dashboard"]
            }
        })
        
        # 8. VISUAL: Icons/pictograms
        sections.append({
            "type": "concept_visual",
            "purpose": "Illustrate key concepts",
            "preferred_story": "feature_grid",
            "content_type": "pictogram",
            "layout_requirements": {
                "min_executive_score": 60,
                "preferred_types": ["feature_grid", "hierarchical_story"]
            }
        })
        
        # 9. IMPLICATIONS
        sections.append({
            "type": "implications",
            "purpose": "Strategic implications",
            "preferred_story": "three_stage_narrative",
            "content_type": "bullets",
            "layout_requirements": {
                "min_executive_score": 65,
                "preferred_types": ["three_stage_narrative", "hierarchical_story"]
            }
        })
        
        # 10. CLOSING: Call to action
        sections.append({
            "type": "closing",
            "purpose": "Clear next steps",
            "preferred_story": "focused_message",
            "content_type": "bullets",
            "layout_requirements": {
                "min_executive_score": 75,
                "preferred_types": ["focused_message"]
            }
        })
        
        return {
            "topic": topic,
            "template": template_name,
            "total_slides": len(sections),
            "sections": sections,
            "analyzer": analyzer
        }


    def generate(self) -> Union[pathlib.Path, None]:
        """
        Generate the slide deck based on the initialized topic and settings.

        This method orchestrates the story planning, prompt construction, LLM interaction,
        and finally the slide deck creation.

        Returns:
            Union[pathlib.Path, None]: The path to the generated PPTX file, or None if generation failed.

        Raises:
            RuntimeError: If there is a failure in getting a response from the LLM.
        """
        
        start_time = time.time()
        logger.info(f'🚀 Generating executive deck on: {self.topic}')
        
        # GET TEMPLATE NAME
        template_name = list(GlobalConfig.PPTX_TEMPLATE_FILES.keys())[self.template_idx]
        
        # STEP 1: BUILD STORY PLAN (NEW)
        logger.info('📋 Building executive story plan...')
        story_plan = self._build_executive_story_plan(self.topic, template_name)
        
        logger.info(f"✓ Story plan: {story_plan['total_slides']} sections")
        for idx, section in enumerate(story_plan['sections'], 1):
            logger.info(f"  {idx}. {section['type']} - {section['purpose']}")
        
        # STEP 2: ENHANCE PROMPT WITH STORY PLAN
        prompt_template = self._get_prompt_template(is_refinement=False)
        
        additional_info = ''
        if self.pdf_path_or_stream:
            additional_info = filem.get_pdf_contents(
                self.pdf_path_or_stream, 
                self.pdf_page_range
            )
        
        # ADD STORY GUIDANCE TO PROMPT
        story_guidance = "\n\n### EXECUTIVE STORY STRUCTURE:\n"
        story_guidance += f"Create exactly {story_plan['total_slides']} slides following this structure:\n\n"
        
        for idx, section in enumerate(story_plan['sections'], 1):
            story_guidance += f"{idx}. **{section['type'].upper()}**: {section['purpose']}\n"
            story_guidance += f"   - Content type: {section['content_type']}\n"
            story_guidance += f"   - Style: {section['preferred_story']}\n\n"
        
        story_guidance += "\nIMPORTANT RULES:\n"
        story_guidance += "- NO duplicate section types\n"
        story_guidance += "- Each section must have UNIQUE purpose\n"
        story_guidance += "- Use varied content types (charts, tables, bullets, icons)\n"
        story_guidance += "- Executive verbosity: concise yet complete (level 7)\n"
        story_guidance += "- Every slide must tell ONE clear story\n"
        
        # FORMAT PROMPT
        try:
            formatted_prompt = prompt_template.format(
                topic=self.topic,
                question=self.topic,
                additional_info=additional_info
            )
            # INJECT STORY GUIDANCE
            formatted_prompt = formatted_prompt.replace(
                "### Topic:",
                story_guidance + "\n### Topic:"
            )
        except KeyError as e:
            logger.warning(f"Template format error: {e}")
            formatted_prompt = prompt_template.replace('{topic}', self.topic)
            formatted_prompt = formatted_prompt.replace('{question}', self.topic)
            formatted_prompt = formatted_prompt.replace('{additional_info}', additional_info)
            formatted_prompt = story_guidance + "\n" + formatted_prompt
        
        # STEP 3: GET LLM RESPONSE (existing code)
        llm = self._initialize_llm()
        response = ''
        
        try:
            logger.info('🤖 Streaming LLM response with story guidance...')
            for chunk in llm.stream(formatted_prompt):
                chunk_text = _process_llm_chunk(chunk)
                response += chunk_text
            logger.info(f'✓ Received {len(response)} characters')
        except Exception as e:
            logger.error(f'LLM streaming failed: {e}')
            raise RuntimeError(f'Failed to get response from LLM: {e}') from e

        # STEP 4: GENERATE PPTX
        self.last_response = text_helper.get_clean_json(response)
        return self._generate_slide_deck(self.last_response)
    
    def _generate_section_plan(self, layouts_info: dict) -> list:
        """
        Generate a high-level section plan based on available layouts.

        Args:
            layouts_info (dict): Information about available layouts.

        Returns:
            list: A list of dictionaries, where each dictionary represents a section plan
                  containing keys like "section_title", "layout_idx", "purpose", etc.
        """
        llm = self._initialize_llm()
        
        # Create planning prompt
        planning_prompt = f"""You are planning an executive presentation on: {self.topic}
    
    Available layouts:
    {self._format_layouts_for_planning(layouts_info)}
    
    Create a section plan with 8-12 sections. Each section should:
    1. Have a clear purpose
    2. Use an appropriate layout (specify layout index)
    3. Not repeat layout types consecutively
    4. Follow a logical flow
    
    Include these section types:
    - Introduction (bullets)
    - Key data (table or chart)
    - Comparison (2-3 column layout)
    - Highlights (KPI cards or icons)
    - Analysis (bullets or chart)
    - Conclusion (bullets)
    
    Return ONLY a JSON array:
    [
      {{
        "section_title": "Introduction",
        "layout_idx": 1,
        "layout_type": "single_column",
        "purpose": "Set context",
        "content_type": "bullets"
      }},
      ...
    ]
    """
        
        response = ''
        for chunk in llm.stream(planning_prompt):
            response += _process_llm_chunk(chunk)
        
        # Parse plan
        try:
            cleaned = text_helper.get_clean_json(response)
            plan = json5.loads(cleaned)
            
            # Validate and ensure diversity
            plan = self._enforce_layout_diversity(plan, layouts_info)
            
            logger.info(f'✅ Section plan created: {len(plan)} sections')
            return plan
        except Exception as e:
            logger.error(f'Planning failed: {e}')
            raise
    
    def _format_layouts_for_planning(self, layouts_info: dict) -> str:
        """
        Format layout information into a string for the LLM prompt.

        Args:
            layouts_info (dict): A dictionary containing layout information.

        Returns:
            str: A formatted string describing the layouts.
        """
        formatted = []
        for idx, layout in layouts_info['layouts'].items():
            formatted.append(
                f"Layout {idx}: {layout['name']}\n"
                f"  Type: {layout['layout_type']}\n"
                f"  Best for: {', '.join(layout['best_for'][:3])}\n"
                f"  Sections: {layout['semantic_sections']}\n"
                f"  Executive score: {layout.get('executive_score', 50)}/100"
            )
        return '\n\n'.join(formatted)
    
    def _enforce_layout_diversity(self, plan: list, layouts_info: dict) -> list:
        """
        Ensure no three consecutive sections use the same layout.

        Args:
            plan (list): The initial section plan.
            layouts_info (dict): Information about available layouts.

        Returns:
            list: The modified section plan with better layout diversity.
        """
        for i in range(2, len(plan)):
            if plan[i-2]['layout_idx'] == plan[i-1]['layout_idx'] == plan[i]['layout_idx']:
                # Find alternative layout
                current_type = plan[i]['content_type']
                alternatives = [
                    idx for idx, layout in layouts_info['layouts'].items()
                    if current_type in layout['best_for'] and idx != plan[i]['layout_idx']
                ]
                
                if alternatives:
                    plan[i]['layout_idx'] = alternatives[0]
                    logger.info(f"🔄 Diversified section {i}: layout {plan[i]['layout_idx']}")
        
        return plan
    
    def _generate_content_for_sections(self, section_plan: list) -> dict:
        """
        Generate actual content for each planned section.

        Args:
            section_plan (list): The list of planned sections.

        Returns:
            dict: A dictionary containing the presentation title and a list of generated slides.
        """
        llm = self._initialize_llm()
        
        all_slides = []
        
        for idx, section in enumerate(section_plan):
            logger.info(f"  Generating section {idx+1}/{len(section_plan)}: {section['section_title']}")
            
            # Create section-specific prompt
            section_prompt = f"""Generate content for this presentation section:
    
    Topic: {self.topic}
    Section: {section['section_title']}
    Purpose: {section['purpose']}
    Content Type: {section['content_type']}
    Layout: {section['layout_type']}
    
    Generate appropriate content (bullets, table, chart, or comparison format).
    Be concise and executive-focused.
    
    Return ONLY a JSON object for ONE slide:
    {{
      "heading": "Section Title",
      "layout_idx": {section['layout_idx']},
      "bullet_points": [...] or "table": {{...}} or "chart": {{...}}
    }}
    """
            
            response = ''
            for chunk in llm.stream(section_prompt):
                response += _process_llm_chunk(chunk)
            
            try:
                cleaned = text_helper.get_clean_json(response)
                slide_data = json5.loads(cleaned)
                all_slides.append(slide_data)
            except Exception as e:
                logger.error(f'Section {idx} generation failed: {e}')
                # Add placeholder
                all_slides.append({
                    'heading': section['section_title'],
                    'layout_idx': section['layout_idx'],
                    'bullet_points': ['Content generation failed']
                })
        
        return {
            'title': self.topic,
            'slides': all_slides
        }
    
    def revise(self, instructions: str, progress_callback=None) -> Union[pathlib.Path, None]:
        """
        Revise the slide deck with new instructions.

        Args:
            instructions (str): The instructions for revising the slide deck.
            progress_callback (callable, optional): Optional callback function to report progress.

        Returns:
            Union[pathlib.Path, None]: The path to the revised .pptx file, or None if failed.

        Raises:
            ValueError: If no slide deck exists or chat history is full.
        """
        if not self.last_response:
            raise ValueError('You must generate a slide deck before you can revise it.')

        if len(self.chat_history.messages) >= 16:
            raise ValueError('Chat history is full. Please reset to continue.')

        self.chat_history.add_user_message(instructions)

        prompt_template = self._get_prompt_template(is_refinement=True)

        list_of_msgs = [
            f'{idx + 1}. {msg.content}'
            for idx, msg in enumerate(self.chat_history.messages) if msg.role == 'user'
        ]

        additional_info = ''
        if self.pdf_path_or_stream:
            additional_info = filem.get_pdf_contents(self.pdf_path_or_stream, self.pdf_page_range)

        formatted_template = prompt_template.format(
            instructions='\n'.join(list_of_msgs),
            previous_content=self.last_response,
            additional_info=additional_info,
        )

        llm = self._initialize_llm()
        response = _stream_llm_response(llm, formatted_template, progress_callback)

        self.last_response = text_helper.get_clean_json(response)
        self.chat_history.add_ai_message(self.last_response)

        return self._generate_slide_deck(self.last_response)

    def _generate_slide_deck(self, json_str: str) -> Union[pathlib.Path, None]:
        """
        Create a slide deck and return the file path.

        Args:
            json_str (str): The content in valid JSON format.

        Returns:
            Union[pathlib.Path, None]: The path to the .pptx file or None in case of error.
        """
        try:
            parsed_data = json5.loads(json_str)
            with open("/home/loft_user_3531/slide-deck-ai/output.json", "w", encoding="utf-8") as f:
                json.dump(parsed_data, f, indent=4, ensure_ascii=False)
        except (ValueError, RecursionError) as e:
            logger.error('Error parsing JSON: %s', e)
            try:
                parsed_data = json5.loads(text_helper.fix_malformed_json(json_str))
            except (ValueError, RecursionError) as e2:
                logger.error('Error parsing fixed JSON: %s', e2)
                return None

        temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        path = pathlib.Path(temp.name)
        temp.close()

        try:
            pptx_helper.generate_powerpoint_presentation(
                parsed_data,
                slides_template=VALID_TEMPLATE_NAMES[self.template_idx],
                output_file_path=path
            )
        except Exception as ex:
            logger.exception('Caught a generic exception: %s', str(ex))
            return None

        return path

    def set_model(self, model_name: str, api_key: str | None = None):
        """
        Set the LLM model (and API key) to use.

        Args:
            model_name (str): The name of the model to use.
            api_key (str, optional): The API key for the LLM provider.

        Raises:
            ValueError: If the model name is not in `VALID_MODELS`.
        """
        if model_name not in GlobalConfig.VALID_MODELS:
            raise ValueError(
                f'Invalid model name: {model_name}.'
                f' Must be one of: {", ".join(VALID_MODEL_NAMES)}.'
            )
        self.model = model_name
        if api_key:
            self.api_key = api_key
        logger.debug('Model set to: %s', model_name)

    def set_template(self, idx: int):
        """
        Set the PowerPoint template to use.

        Args:
            idx (int): The index of the template to use from `GlobalConfig.PPTX_TEMPLATE_FILES`.
        """
        num_templates = len(GlobalConfig.PPTX_TEMPLATE_FILES)
        self.template_idx = idx if 0 <= idx < num_templates else 0

    def reset(self):
        """
        Reset the chat history and internal state.
        """
        self.chat_history = ChatMessageHistory()
        self.last_response = None
        self.template_idx = 0
        self.topic = ''
        
    def generate_from_plan(self, plan: Any, progress_callback=None) -> Union[pathlib.Path, None]:
        """
        Generate slides from an approved research plan.
        
        Args:
            plan (ResearchPlan): ResearchPlan object with sections and queries.
            progress_callback (callable, optional): Optional callback for progress updates.
        
        Returns:
            Union[pathlib.Path, None]: Path to generated PPTX file.
        """
        # Note: Importing inside method to avoid circular imports if core_agents imports this
        from slidedeckai.agents.core_agents import ResearchPlan
        
        # Convert plan sections to SlideDeck format
        sections_text = []
        
        for section in plan.sections:
            section_text = f"\n## {section.section_title}\n"
            section_text += f"{section.section_purpose}\n\n"
            
            # Add visualization hint
            section_text += f"*Visualization: {section.visualization_hint}*\n\n"
            
            # Add search queries as bullet points
            for query in section.search_queries:
                section_text += f"- {query.query}\n"
            
            sections_text.append(section_text)
        
        # Combine into single prompt
        enhanced_topic = f"{plan.query}\n\n" + "\n".join(sections_text)
        
        # Update the topic
        self.topic = enhanced_topic
        
        # Generate slides using existing SlideDeck AI logic
        return self.generate()
