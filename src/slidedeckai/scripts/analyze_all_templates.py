"""
Analyze all templates and save their capabilities.
Run this once to understand your templates.
"""
import sys
import os
import json

sys.path.insert(0, os.path.abspath('src'))

from slidedeckai.global_config import GlobalConfig
from slidedeckai.layout_analyzer import TemplateAnalyzer
from pptx import Presentation


def analyze_all_templates():
    """Analyze all configured templates."""
    print("\n" + "="*80)
    print("ANALYZING ALL TEMPLATES")
    print("="*80 + "\n")
    
    results = {}
    
    for key, value in GlobalConfig.PPTX_TEMPLATE_FILES.items():
        print(f"\n📄 Analyzing template: {key}")
        print(f"   File: {value['file']}")
        
        try:
            presentation = Presentation(value['file'])
            analyzer = TemplateAnalyzer(presentation)
            
            # Print summary
            analyzer.print_summary()
            
            # Save analysis
            analysis = analyzer.export_analysis()
            results[key] = analysis
            
            # Save to file
            output_file = f'template_analysis_{key}.json'
            with open(output_file, 'w') as f:
                json.dump(analysis, f, indent=2)
            
            print(f"   ✓ Analysis saved to {output_file}")
            
        except Exception as e:
            print(f"   ✗ Failed: {e}")
            results[key] = {'error': str(e)}
    
    # Save combined results
    with open('all_templates_analysis.json', 'w') as f:
        json.dump(results, f, indent=2)
    
    print("\n" + "="*80)
    print("✅ Analysis complete! Check all_templates_analysis.json")
    print("="*80 + "\n")


if __name__ == '__main__':
    analyze_all_templates()