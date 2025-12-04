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
from werkzeug.utils import secure_filename
from flask_cors import CORS
from dotenv import load_dotenv

sys.path.insert(0, os.path.abspath('src'))

# Import SlideDeck AI core
from slidedeckai.global_config import GlobalConfig
from slidedeckai.layout_analyzer import TemplateAnalyzer
from pptx import Presentation

# Import HTML UI
from slidedeckai.ui.html_ui import HTML_UI

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
        data = request.get_json()
        query = data.get('query', '').strip()
        template_key = data.get('template', 'Basic')
        search_mode = data.get('search_mode', 'normal')
        num_sections = data.get('num_sections', None)
        
        if not query:
            return jsonify({'error': 'Query required'}), 400
        
        logger.info(f"🔥 Creating plan: {query}")
        logger.info(f"  Template: {template_key}")
        logger.info(f"  Mode: {search_mode}")
        
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            return jsonify({'error': 'OpenAI API key not configured'}), 500
        
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
        
        # Generate plan with enforced diversity
        research_plan = orchestrator.generate_plan(
            user_query=query,
            template_layouts=layout_info['layouts'],
            num_sections=num_sections
        )
        
        # Cache plan
        plan_id = datetime.now().strftime('%Y%m%d_%H%M%S')
        plans_cache[plan_id] = {
            'query': query,
            'template_key': template_key,
            'search_mode': search_mode,
            'research_plan': research_plan,
            'analyzer': analyzer
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
        
        logger.info(f"🚀 Executing plan {plan_id}")
        logger.info(f"  Query: {query}")
        logger.info(f"  Template: {template_key}")
        logger.info(f"  Sections: {len(research_plan.sections)}")
        
        api_key = os.getenv('OPENAI_API_KEY')
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
        
        output_path = orchestrator.execute_plan(research_plan, output_path)
        
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


@app.route('/api/upload_template', methods=['POST'])
def upload_template():
    """Upload a new PPTX template"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        if file and file.filename.endswith('.pptx'):
            filename = secure_filename(file.filename)
            template_dir = pathlib.Path('src/slidedeckai/pptx_templates')
            template_dir.mkdir(parents=True, exist_ok=True)

            save_path = template_dir / filename
            file.save(save_path)

            # Register in GlobalConfig
            template_name = filename.replace('.pptx', '').replace('_', ' ').title()
            GlobalConfig.PPTX_TEMPLATE_FILES[template_name] = {
                'file': save_path,
                'caption': 'Uploaded Template'
            }

            logger.info(f"✅ Template uploaded: {template_name}")
            return jsonify({'success': True, 'template': template_name})

        return jsonify({'error': 'Invalid file type'}), 400

    except Exception as e:
        logger.error(f"Upload failed: {e}", exc_info=True)
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