"""
Core functionality of SlideDeck AI.
"""
import logging
import os
import pathlib
import tempfile
from typing import Union, Any
import time
import json5, json
from dotenv import load_dotenv
load_dotenv()

from . import global_config as gcfg
from .global_config import GlobalConfig
from .helpers import file_manager as filem
from .helpers import llm_helper, pptx_helper, text_helper
from .helpers.chat_helper import ChatMessageHistory
from .layout_analyzer import TemplateAnalyzer
from .content_matcher import ContentLayoutMatcher


RUN_IN_OFFLINE_MODE = os.getenv('RUN_IN_OFFLINE_MODE', 'False').lower() == 'true'
VALID_MODEL_NAMES = list(GlobalConfig.VALID_MODELS.keys())
VALID_TEMPLATE_NAMES = list(GlobalConfig.PPTX_TEMPLATE_FILES.keys())

logger = logging.getLogger(__name__)


def _process_llm_chunk(chunk: Any) -> str:
    """
    Helper function to process LLM response chunks consistently.

    Args:
        chunk: The chunk received from the LLM stream.

    Returns:
        The processed text from the chunk.
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
        prompt: The prompt to send to the LLM.
        progress_callback: A callback function to report progress.

    Returns:
        The complete response from the LLM.

    Raises:
        RuntimeError: If there's an error getting response from LLM.
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
    The main class for generating slide decks.
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
            model: The name of the LLM model to use.
            topic: The topic of the slide deck.
            api_key: The API key for the LLM provider.
            pdf_path_or_stream: The path to a PDF file or a file-like object.
            pdf_page_range: A tuple representing the page range to use from the PDF file.
            template_idx: The index of the PowerPoint template to use.

        Raises:
            ValueError: If the model name is not in VALID_MODELS.
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

    def _initialize_llm(self):
        """
        Initialize and return an LLM instance with the current configuration.

        Returns:
            Configured LLM instance.
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
        Return a prompt template.

        Args:
            is_refinement: Whether this is the initial or refinement prompt.

        Returns:
            The prompt template as f-string.
        """
        if is_refinement:
            with open(GlobalConfig.REFINEMENT_PROMPT_TEMPLATE, 'r', encoding='utf-8') as in_file:
                template = in_file.read()
        else:
            with open(GlobalConfig.INITIAL_PROMPT_TEMPLATE, 'r', encoding='utf-8') as in_file:
                template = in_file.read()
        return template
    
    def _build_executive_story_plan(self, topic: str, template_name: str) -> Dict:
        """
        CRITICAL: Plan the story BEFORE generating content
        Returns section structure with layout assignments
        """
        
        # Get template analyzer
        template_file = GlobalConfig.PPTX_TEMPLATE_FILES[template_name]['file']
        presentation = Presentation(template_file)
        analyzer = TemplateAnalyzer(presentation)
        
        # Get available layouts sorted by executive suitability
        exec_layouts = sorted(
            analyzer.layouts.items(),
            key=lambda x: x[1].executive_suitability,
            reverse=True
        )
        
        # Build story sections (10-12 slides typical)
        num_slides = 10
        
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


    def generate(self) -> pathlib.Path:
        """ENHANCED with executive story planning"""
        
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
    
    def _generate_section_plan(self, layouts_info: dict) -> list:
        """
        Generate high-level section plan based on available layouts
        Returns: [{"section_title": ..., "layout_idx": ..., "purpose": ...}, ...]
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
        """Format layout info for LLM"""
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
        """Ensure no 3 consecutive same layout types"""
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
        """Generate actual content for each planned section"""
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
    
    def revise(self, instructions, progress_callback=None):
        """
        Revise the slide deck with new instructions.

        Args:
            instructions: The instructions for revising the slide deck.
            progress_callback: Optional callback function to report progress.

        Returns:
            The path to the revised .pptx file.

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
            json_str: The content in valid JSON format.

        Returns:
            The path to the .pptx file or None in case of error.
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
            model_name: The name of the model to use.
            api_key: The API key for the LLM provider.

        Raises:
            ValueError: If the model name is not in VALID_MODELS.
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

    def set_template(self, idx):
        """
        Set the PowerPoint template to use.

        Args:
            idx: The index of the template to use.
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
        
    def generate_from_plan(self, plan: 'ResearchPlan', progress_callback=None):
        """
        Generate slides from an approved research plan.
        
        Args:
            plan: ResearchPlan object with sections and queries
            progress_callback: Optional callback for progress updates
        
        Returns:
            Path to generated PPTX file
        """
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
        return self.generate(progress_callback=progress_callback)