# app.py - PRODUCTION READY FLASK SERVER
# ✅ Proper integration of all components

import os, sys
import logging
import traceback
import tempfile
import pathlib
from datetime import datetime
from typing import Dict, Any
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
    """Get cached analyzer or create new one for template"""
    if template_key not in template_analyzers:
        logger.info(f"🔍 Creating new analyzer for template: {template_key}")
        template_file = GlobalConfig.PPTX_TEMPLATE_FILES[template_key]['file']
        presentation = Presentation(template_file)
        analyzer = TemplateAnalyzer(presentation)
        template_analyzers[template_key] = analyzer
        logger.info(f"✓ Analyzer cached for {template_key}")
    
    return template_analyzers[template_key]


def serialize_plan(research_plan) -> Dict:
    """Serialize ResearchPlan to dict properly"""
    
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
    """Serve the HTML UI"""
    return render_template_string(HTML_UI)


@app.route('/api/plan', methods=['POST'])
def create_plan():
    """Phase 1: Create layout-aware research plan with enforced diversity"""
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
    """Phase 2: Execute approved plan with proper mapping"""
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
        # However, for consistency, if the user provided an API key during plan generation, we should probably stick to it or ask for it again.
        # Ideally, we should receive it again here or store it in cache (not recommended for secrets).
        # Let's assume the user has to provide it if not in env, or it's passed in data.
        # But `html_ui` currently only sends `plan_id`.
        # I'll stick to env var for now unless I update `execute` frontend call too.
        # Wait, I should update frontend `approvePlan` to send API key if it was set in settings.
        # But `approvePlan` logic is separate.
        # Let's rely on `orchestrator`'s API key.
        # Actually, `plans_cache` is in-memory. I can store the API key there TEMPORARILY for the session?
        # A better practice is to pass it from frontend.

        # Retrieve potential API key from plans_cache if I decided to store it there (I didn't).
        # So I will check if data has api_key (I need to update frontend to send it).

        api_key = data.get('api_key') or os.getenv('OPENAI_API_KEY')
        
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
def download_report(report_id):
    """Download generated presentation"""
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
    """Get all available templates"""
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


@app.route('/api/health')
def health():
    """Health check endpoint"""
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