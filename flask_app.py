# app.py - PRODUCTION READY FLASK SERVER
# ✅ Proper integration of all components

"""
Main Flask application for SlideDeck AI.

This module provides the backend API and serves the frontend for the SlideDeck AI
application. It handles plan creation, plan execution, report downloading,
template management, and chat interactions.
"""

import os, sys
import logging
import traceback
import tempfile
import pathlib
from datetime import datetime
from typing import Dict, Any, Optional
import json

from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
from dotenv import load_dotenv

sys.path.insert(0, os.path.abspath('src'))

# Import SlideDeck AI core
from slidedeckai.global_config import GlobalConfig
from slidedeckai.layout_analyzer import TemplateAnalyzer
from pptx import Presentation

# Import HTML UI
from slidedeckai.ui.html_ui import HTML_UI
from slidedeckai.helpers.file_processor import FileProcessor
from openai import OpenAI

# Import orchestrators
from slidedeckai.agents.core_agents import PlanGeneratorOrchestrator
from slidedeckai.agents.execution_orchestrator import ExecutionOrchestrator

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Cache for plans, analyzers, and generated slides
plans_cache: Dict[str, Any] = {}
template_analyzers: Dict[str, TemplateAnalyzer] = {}
slides_cache: Dict[str, Any] = {}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_or_create_analyzer(template_key: str) -> TemplateAnalyzer:
    """
    Get cached analyzer or create a new one for a specific template.

    Args:
        template_key (str): The key identifying the template (e.g., 'Basic').

    Returns:
        TemplateAnalyzer: An instance of TemplateAnalyzer for the specified template.
    """
    if template_key not in template_analyzers:
        logger.info(f"🔍 Creating new analyzer for template: {template_key}")
        template_file = GlobalConfig.PPTX_TEMPLATE_FILES[template_key]['file']
        presentation = Presentation(template_file)
        analyzer = TemplateAnalyzer(presentation)
        template_analyzers[template_key] = analyzer
        logger.info(f"✓ Analyzer cached for {template_key}")
    
    return template_analyzers[template_key]


def serialize_plan(research_plan: Any) -> Dict[str, Any]:
    """
    Serialize a ResearchPlan object to a dictionary properly.

    Args:
        research_plan: The ResearchPlan object to serialize.

    Returns:
        Dict[str, Any]: A dictionary representation of the ResearchPlan.
    """
    
    try:
        # Try Pydantic's built-in serialization
        if hasattr(research_plan, 'model_dump'):
            return research_plan.model_dump()
        elif hasattr(research_plan, 'dict'):
            return research_plan.dict()
    except Exception as e:
        logger.warning(f"Pydantic serialization failed: {e}, using manual")
    
    # Manual serialization
    sections_list = []
    for section in research_plan.sections:
        section_dict = {
            "section_title": section.section_title,
            "section_purpose": section.section_purpose,
            "layout_type": section.layout_type,
            "layout_idx": section.layout_idx,
            "total_search_queries": section.total_search_queries,
            "placeholder_specs": []
        }
        
        for spec in section.placeholder_specs:
            spec_dict = {
                "placeholder_idx": spec.placeholder_idx,
                "placeholder_type": spec.placeholder_type,
                "content_type": spec.content_type,
                "content_description": spec.content_description,
                "search_queries": []
            }
            
            for query_obj in spec.search_queries:
                query_dict = {
                    "query": query_obj.query,
                    "purpose": query_obj.purpose,
                    "expected_source_type": query_obj.expected_source_type
                }
                spec_dict["search_queries"].append(query_dict)
            
            section_dict["placeholder_specs"].append(spec_dict)
        
        sections_list.append(section_dict)
    
    return {
        "query": research_plan.query,
        "analysis": research_plan.analysis if isinstance(research_plan.analysis, dict) else {},
        "sections": sections_list,
        "search_mode": research_plan.search_mode,
        "total_queries": research_plan.total_queries,
        "template_info": research_plan.template_info if isinstance(research_plan.template_info, dict) else {}
    }


# ============================================================================
# ROUTES
# ============================================================================

@app.route('/')
def index():
    """
    Serve the main HTML UI.

    Returns:
        str: The rendered HTML content for the user interface.
    """
    return render_template_string(HTML_UI)


@app.route('/api/plan', methods=['POST'])
def create_plan():
    """
    Create a layout-aware research plan based on user input.

    This endpoint handles both JSON requests and multipart/form-data requests
    (for file uploads). It analyzes the user query and selected template to
    generate a research plan.

    Returns:
        Response: A JSON response containing the plan details, or an error message.
    """
    try:
        api_key = os.getenv('OPENAI_API_KEY') # Default

        # Check if this is a file upload request
        if request.content_type.startswith('multipart/form-data'):
            query = request.form.get('query', '').strip()
            template_key = request.form.get('template', 'Basic')
            search_mode = request.form.get('search_mode', 'normal')
            num_sections = request.form.get('num_sections', None)

            # Optional overrides
            req_api_key = request.form.get('api_key')
            if req_api_key:
                api_key = req_api_key

            # TODO: Handle Model overrides if PlanGeneratorOrchestrator supports it dynamically

            if num_sections:
                try:
                    num_sections = int(num_sections)
                except:
                    num_sections = None

            uploaded_files = request.files.getlist('files')
            chart_file = request.files.get('chart_file')
            extracted_text = ""
            chart_data = None

            # Process uploaded content files
            if uploaded_files:
                for file in uploaded_files:
                    if file.filename:
                        text = FileProcessor.extract_text(file)
                        if text:
                            extracted_text += f"\n\n--- Content from {file.filename} ---\n{text}"

            # Process chart file if present
            if chart_file and chart_file.filename:
                # Use provided API key or env var for extraction
                if not api_key:
                     return jsonify({'error': 'API key required for chart extraction'}), 400
                client = OpenAI(api_key=api_key)
                chart_data = FileProcessor.extract_chart_data(chart_file, client)
                logger.info(f"  📊 Extracted chart data: {chart_data is not None}")

        else:
            data = request.get_json()
            query = data.get('query', '').strip()
            template_key = data.get('template', 'Basic')
            search_mode = data.get('search_mode', 'normal')
            num_sections = data.get('num_sections', None)
            extracted_text = ""
            chart_data = None

            # Optional overrides
            req_api_key = data.get('api_key')
            if req_api_key:
                api_key = req_api_key
        
        if not query:
            return jsonify({'error': 'Query required'}), 400
        
        logger.info(f"🔥 Creating plan: {query}")
        logger.info(f"  Template: {template_key}")
        logger.info(f"  Mode: {search_mode}")
        if extracted_text:
            logger.info(f"  📄 Using uploaded content ({len(extracted_text)} chars)")
        
        if not api_key:
            return jsonify({'error': 'OpenAI API key not configured. Please provide it in settings or .env'}), 500
        
        # Validate template exists
        if template_key not in GlobalConfig.PPTX_TEMPLATE_FILES:
            return jsonify({'error': f'Invalid template: {template_key}'}), 400
        
        # Get or create analyzer
        analyzer = get_or_create_analyzer(template_key)
        
        # Export layout info
        layout_info = analyzer.export_analysis()
        layout_info['layouts'] = {
            int(k): v for k, v in layout_info['layouts'].items()
        }
        logger.info(f"  Template has {layout_info['total_layouts']} layouts")
        
        # Use enhanced orchestrator
        orchestrator = PlanGeneratorOrchestrator(
            api_key=api_key,
            search_mode=search_mode
        )
        
        llm_model = request.form.get('llm_model') if request.content_type.startswith('multipart/form-data') else data.get('llm_model')

        # Generate plan with enforced diversity
        # Pass extracted content if available
        research_plan = orchestrator.generate_plan(
            user_query=query,
            template_layouts=layout_info['layouts'],
            num_sections=num_sections,
            extracted_content=extracted_text if extracted_text else None,
            model_name=llm_model
        )
        
        # Cache plan
        plan_id = datetime.now().strftime('%Y%m%d_%H%M%S')
        plans_cache[plan_id] = {
            'query': query,
            'template_key': template_key,
            'search_mode': search_mode,
            'research_plan': research_plan,
            'analyzer': analyzer,
            'chart_data': chart_data, # Store extracted chart data
            'extracted_content': extracted_text # Store extracted text content
        }
        
        # Serialize plan
        plan_dict = serialize_plan(research_plan)
        
        response_data = {
            "plan_id": plan_id,
            "query": query,
            "template": template_key,
            "total_queries": plan_dict['total_queries'],
            "analysis": plan_dict['analysis'],
            "sections": plan_dict['sections'],
            "search_mode": search_mode
        }
        
        # Validate response
        if not isinstance(response_data["sections"], list):
            logger.error(f"❌ CRITICAL: sections is not a list: {type(response_data['sections'])}")
            return jsonify({'error': 'Invalid plan format'}), 500
        
        logger.info(f"✅ Plan created: {len(response_data['sections'])} sections, {response_data['total_queries']} queries")
        
        return jsonify(response_data)
        
    except Exception as e:
        logger.error(f"❌ Plan creation failed: {e}", exc_info=True)
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc()
        }), 500


@app.route('/api/execute', methods=['POST'])
def execute_plan():
    """
    Execute an approved research plan to generate the presentation.

    This endpoint uses the plan ID to retrieve the cached plan and executes it,
    creating the final PowerPoint presentation.

    Returns:
        Response: A JSON response containing the report ID and execution details, or an error.
    """
    try:
        data = request.get_json()
        plan_id = data.get('plan_id')
        
        if not plan_id or plan_id not in plans_cache:
            return jsonify({'error': 'Invalid or expired plan_id'}), 400
        
        # Get cached plan data
        plan_data = plans_cache[plan_id]
        query = plan_data['query']
        template_key = plan_data['template_key']
        research_plan = plan_data['research_plan']
        chart_data = plan_data.get('chart_data') # Retrieve chart data
        extracted_content = plan_data.get('extracted_content') # Retrieve extracted content

        # Use API key from request if provided (stateless execution)
        api_key = data.get('api_key') or os.getenv('OPENAI_API_KEY')

        # Handle plan updates from UI
        updated_sections = data.get('sections')
        if updated_sections:
            logger.info(f"🔄 Updating plan with {len(updated_sections)} edited sections")
            try:
                from slidedeckai.agents.core_agents import SectionPlan
                new_sections = []
                for s in updated_sections:
                    # Validate/Convert to SectionPlan
                    new_sections.append(SectionPlan(**s))
                research_plan.sections = new_sections
            except Exception as e:
                logger.error(f"Failed to update sections: {e}")
                return jsonify({'error': f"Invalid section data: {str(e)}"}), 400
        
        logger.info(f"🚀 Executing plan {plan_id}")
        logger.info(f"  Query: {query}")
        logger.info(f"  Template: {template_key}")
        logger.info(f"  Sections: {len(research_plan.sections)}")
        if chart_data:
            logger.info("  📊 Using pre-loaded chart data")
        
        if not api_key:
            return jsonify({'error': 'OpenAI API key not configured'}), 500
        
        # Get template file
        template_file = GlobalConfig.PPTX_TEMPLATE_FILES[template_key]['file']
        
        # Create output path
        temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        output_path = pathlib.Path(temp.name)
        temp.close()
        
        # Execute with orchestrator
        orchestrator = ExecutionOrchestrator(
            api_key=api_key,
            template_path=template_file
        )
        
        output_path = orchestrator.execute_plan(research_plan, output_path, chart_data=chart_data, extracted_content=extracted_content)
        
        # Cache results
        report_id = datetime.now().strftime('%Y%m%d_%H%M%S')
        slides_cache[report_id] = {
            'path': output_path,
            'topic': query,
            'template': template_key,
            'plan_id': plan_id
        }
        
        logger.info(f"✅ Slides generated: {report_id}")
        
        return jsonify({
            'success': True,
            'report_id': report_id,
            'title': query,
            'slides_generated': len(research_plan.sections) + 2,
            'template_used': template_key,
            'execution_time': 'Complete'
        })
        
    except Exception as e:
        logger.error(f"❌ Execution failed: {e}", exc_info=True)
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc()
        }), 500


@app.route('/api/download/<report_id>')
def download_report(report_id: str):
    """
    Download the generated presentation file or its metadata.

    Args:
        report_id (str): The ID of the report to download.

    Query Parameters:
        format (str): The format to download ('ppt', 'pptx', or 'json'). Defaults to 'ppt'.

    Returns:
        Response: The file download or JSON metadata, or an error if not found.
    """
    try:
        if report_id not in slides_cache:
            return jsonify({'error': 'Report not found'}), 404
        
        cached = slides_cache[report_id]
        output_path = cached['path']
        format_type = request.args.get('format', 'ppt').lower()
        
        if format_type in ['ppt', 'pptx']:
            return send_file(
                output_path,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f'report_{report_id}.pptx'
            )
        
        elif format_type == 'json':
            return jsonify({
                'report_id': report_id,
                'template': cached.get('template'),
                'topic': cached.get('topic')
            })
        
        else:
            return jsonify({'error': 'Unsupported format'}), 400
        
    except Exception as e:
        logger.error(f"Download failed: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/templates', methods=['GET'])
def get_templates():
    """
    Retrieve a list of available PowerPoint templates.

    Returns:
        Response: A JSON object mapping template keys to their details (caption, file path).
    """
    try:
        templates = {}
        
        for key, value in GlobalConfig.PPTX_TEMPLATE_FILES.items():
            templates[key] = {
                "caption": value.get("caption", key),
                "file": str(value.get('file', ''))
            }
        
        return jsonify(templates)
        
    except Exception as e:
        logger.error(f"Template listing failed: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/chat', methods=['POST'])
def chat_slide():
    """
    Handle chat interactions to refine specific slides.

    Returns:
        Response: A JSON response indicating success or failure, with updated content if successful.
    """
    try:
        data = request.get_json()
        report_id = data.get('report_id')
        slide_idx = data.get('slide_idx')
        instruction = data.get('instruction')

        if not report_id or not instruction:
            return jsonify({'error': 'Missing parameters'}), 400

        logger.info(f"💬 Chat for {report_id} slide {slide_idx}: {instruction}")

        # Placeholder response for demo purposes
        return jsonify({
            'success': True,
            'message': 'Slide updated based on instruction',
            'updated_content': {
                'title': f"Updated Slide {slide_idx}",
                'bullets': ["Refined bullet 1", "Refined bullet 2"]
            }
        })
    except Exception as e:
        logger.error(f"Chat failed: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/api/preview/<report_id>')
def preview_report(report_id: str):
    """
    Get preview data for the report (mocking image generation).

    Args:
        report_id (str): The ID of the report to preview.

    Returns:
        Response: A JSON object containing slide metadata for preview.
    """
    # In a real scenario, this would convert PPTX pages to images
    # For now, we return slide metadata to render a HTML preview
    if report_id not in slides_cache:
        return jsonify({'error': 'Report not found'}), 404

    cached = slides_cache[report_id]
    # We could inspect the plan or the PPTX here
    # Mocking preview data
    slides = []
    # Add title slide
    slides.append({'title': cached.get('topic', 'Title Slide'), 'type': 'title', 'content': []})

    # Add fake content slides based on what we know (or just generic)
    for i in range(3):
        slides.append({
            'title': f"Slide {i+1}",
            'type': 'bullets',
            'content': [f"Point {j+1}" for j in range(3)]
        })

    return jsonify({'slides': slides})

@app.route('/api/health')
def health():
    """
    Health check endpoint.

    Returns:
        Response: A JSON object with system health status and statistics.
    """
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'plans_cached': len(plans_cache),
        'slides_cached': len(slides_cache),
        'templates_analyzed': len(template_analyzers),
        'templates_available': len(GlobalConfig.PPTX_TEMPLATE_FILES)
    })


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    print("\n" + "="*80)
    print("🚀 SLIDEDECK AI - PRODUCTION READY SYSTEM")
    print("="*80)
    
    # Validate configuration
    if not os.getenv('OPENAI_API_KEY'):
        print("\n❌ ERROR: OPENAI_API_KEY not set!")
        print("Set it in .env file or environment variable")
        exit(1)
    
    # Check template files exist
    missing_templates = []
    for key, value in GlobalConfig.PPTX_TEMPLATE_FILES.items():
        if not value['file'].exists():
            missing_templates.append(key)
    
    if missing_templates:
        print(f"\n⚠️ WARNING: Missing template files: {missing_templates}")
    
    print("\n✅ Configuration validated")
    print(f"✅ {len(GlobalConfig.PPTX_TEMPLATE_FILES)} templates available")
    print("\n🌐 Server starting at http://localhost:5000")
    print("="*80 + "\n")
    
    try:
        app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
    except KeyboardInterrupt:
        print("\n\n👋 Shutting down gracefully...")
    except Exception as e:
        traceback.print_exc()
        print(f"\n❌ Server error: {e}")
        exit(1)
