"""
FIXED content_matcher.py - ONLY enhancing existing functions
"""
import logging
from typing import Dict, Optional, List
from .layout_analyzer import TemplateAnalyzer

logger = logging.getLogger(__name__)


class ContentLayoutMatcher:
    """ENHANCED - Same class, better intelligence"""
    
    def __init__(self, analyzer: TemplateAnalyzer):
        self.analyzer = analyzer
        self.used_layouts = []
        self.used_story_types = []  # NEW: Track story types used
        self.section_sequence = []  # NEW: Planned story arc

    def _is_compatible_story_type(self, layout_story: str, preferred_story: str) -> bool:
        """Check if layout story type is compatible with preferred"""
        
        compatible_groups = [
            {"data_visualization", "metrics_dashboard"},
            {"balanced_comparison", "hierarchical_story"},
            {"three_stage_narrative", "feature_grid"},
            {"focused_message", "main_supporting"}
        ]
        
        for group in compatible_groups:
            if layout_story in group and preferred_story in group:
                return True
        
        return False
        
    def select_layout_with_story_awareness(self, 
                                       slide_json: dict, 
                                       slide_index: int,
                                       total_slides: int) -> int:
        """
        CRITICAL: Select layout based on:
        1. Content type (chart/table/bullets)
        2. Story position (opening/body/closing)
        3. Diversity (no 3 consecutive same type)
        4. Executive suitability
        """
        
        # Build sequence if not done
        if not self.section_sequence:
            self.section_sequence = self._build_section_sequence(total_slides)
        
        # Get preferred story type for this position
        preferred_story = self.section_sequence[min(slide_index, len(self.section_sequence)-1)]
        
        # Get content type
        content_type = self._infer_content_type_from_json(slide_json)
        
        # Score all layouts
        scored_layouts = []
        
        for idx, layout in self.analyzer.layouts.items():
            score = 0.0
            
            # Base content match (40 points)
            score += self._score_layout_for_content(layout, content_type, slide_json)
            
            # Story alignment (30 points)
            if layout.semantic_story_type == preferred_story:
                score += 30
            elif self._is_compatible_story_type(layout.semantic_story_type, preferred_story):
                score += 15
            
            # Executive suitability (20 points)
            score += (layout.executive_suitability / 100) * 20
            
            # Diversity bonus (10 points)
            if len(self.used_layouts) >= 2:
                if self.used_layouts[-1] != idx and self.used_layouts[-2] != idx:
                    score += 10
            
            # Penalize if used 3 times recently
            recent_uses = self.used_layouts[-5:]  # Last 5 slides
            use_count = recent_uses.count(idx)
            if use_count >= 2:
                score -= 20  # Heavy penalty
            
            scored_layouts.append((score, idx, layout))
        
        scored_layouts.sort(reverse=True, key=lambda x: x[0])
        
        if scored_layouts:
            best_score, best_idx, best_layout = scored_layouts[0]
            
            logger.info(
                f"Slide {slide_index+1}/{total_slides}: "
                f"Layout {best_idx} ({best_layout.name}) - "
                f"Score: {best_score:.1f}/100, "
                f"Story: {best_layout.semantic_story_type}, "
                f"ExecScore: {best_layout.executive_suitability:.0f}"
            )
            
            self.used_layouts.append(best_idx)
            self.used_story_types.append(best_layout.semantic_story_type)
            
            return best_idx
        
        logger.warning(f"No suitable layout found for slide {slide_index+1}")
        return 1
    
    def _build_section_sequence(self, total_sections: int) -> List[str]:
        """
        CRITICAL: Build executive story arc
        Opening -> Body -> Closing structure
        
        Returns list of preferred story types in order
        """
        
        sequence = []
        
        # OPENING (10% of slides)
        opening_count = max(1, total_sections // 10)
        for _ in range(opening_count):
            sequence.append("focused_message")  # Clear opener
        
        # BODY - Varied content (70%)
        body_count = int(total_sections * 0.7)
        
        # Alternate between analytical and visual
        body_types = [
            "data_visualization",      # Charts/graphs
            "balanced_comparison",     # Contrast
            "three_stage_narrative",   # Process
            "metrics_dashboard",       # KPIs
            "detailed_analysis",       # Deep dive
            "hierarchical_story",      # Structured
            "feature_grid"             # Overview
        ]
        
        # Cycle through body types
        for i in range(body_count):
            sequence.append(body_types[i % len(body_types)])
        
        # CLOSING (20%)
        closing_count = total_sections - opening_count - body_count
        for i in range(closing_count):
            if i == closing_count - 1:
                sequence.append("focused_message")  # Clear conclusion
            else:
                sequence.append("metrics_dashboard")  # Summary stats
        
        return sequence
        
    def select_layout_for_slide(self, slide_json: dict, slide_index: int = 0, total_slides: int = 10) -> int:
        """ENHANCED with story awareness"""
        
        # Use new story-aware selection to get initial pick
        layout_idx = self.select_layout_with_story_awareness(
            slide_json, slide_index, total_slides
        )

        # Stronger diversity enforcement and smarter tie-breaking:
        # - Avoid selecting a layout that would cause 3 or more consecutive slides
        #   with the same story type.
        # - If best layout repeats story type too often, prefer alternative with
        #   similar score but different story type.

        # If we don't have history yet, accept the pick and record it
        if not self.used_layouts:
            self.used_layouts.append(layout_idx)
            try:
                self.used_story_types.append(self.analyzer.layouts[layout_idx].semantic_story_type)
            except Exception:
                self.used_story_types.append(None)
            return layout_idx

        # Determine the candidate's story type
        candidate_story = None
        try:
            candidate_story = self.analyzer.layouts[layout_idx].semantic_story_type
        except Exception:
            candidate_story = None

        # If the last two chosen story types match candidate_story, avoid it.
        if len(self.used_story_types) >= 2 and self.used_story_types[-1] == self.used_story_types[-2] == candidate_story:
            logger.info(f"⚠️ Avoiding 3rd consecutive story type '{candidate_story}' for layout {layout_idx}")

            # Look for best alternative layout with a different story type and similar score
            content_type = self._infer_content_type_from_json(slide_json)
            # Compute baseline score for chosen layout
            baseline_score = self._score_layout_for_content(self.analyzer.layouts[layout_idx], content_type, slide_json)

            best_alt = (None, -999.0)
            for idx, layout in self.analyzer.layouts.items():
                if idx == layout_idx:
                    continue
                alt_story = getattr(layout, 'semantic_story_type', None)
                if alt_story == candidate_story:
                    continue

                alt_score = self._score_layout_for_content(layout, content_type, slide_json)

                # Accept alternatives that are within a small margin of the baseline
                if alt_score >= baseline_score - 12:  # allow slight quality drop for diversity
                    # Prefer layouts that introduce a new story type not used recently
                    recent_story_penalty = 0
                    if alt_story in self.used_story_types[-3:]:
                        recent_story_penalty = -5

                    adj_score = alt_score + recent_story_penalty
                    if adj_score > best_alt[1]:
                        best_alt = (idx, adj_score)

            if best_alt[0] is not None:
                logger.info(f"→ Switching to alternative layout {best_alt[0]} to improve diversity")
                layout_idx = best_alt[0]

        # Finally, ensure the used_layouts and used_story_types lists are updated
        # We keep these lists bounded to the last N entries for efficiency
        self.used_layouts.append(layout_idx)
        if len(self.used_layouts) > 50:
            self.used_layouts = self.used_layouts[-50:]

        try:
            self.used_story_types.append(self.analyzer.layouts[layout_idx].semantic_story_type)
        except Exception:
            self.used_story_types.append(None)

        if len(self.used_story_types) > 50:
            self.used_story_types = self.used_story_types[-50:]

        return layout_idx

    def map_content_to_placeholders(self, slide_json: dict, layout_capability) -> dict:
        """
        CRITICAL: Map slide content to specific placeholders intelligently
        Returns: {placeholder_idx: content_spec}
        """
        mapping = {}
        
        # Extract content from JSON
        heading = slide_json.get('heading', '')
        bullets = slide_json.get('bullet_points', [])
        table = slide_json.get('table')
        chart = slide_json.get('chart')
        
        # Get layout structure
        sections = layout_capability.semantic_sections
        content_phs = layout_capability.content_placeholders
        subtitle_phs = layout_capability.subtitle_placeholders
        
        # CASE 1: Chart/Table - use largest placeholder
        if chart:
            largest = max(content_phs, key=lambda x: x.area)
            mapping[largest.idx] = {'type': 'chart', 'data': chart}
            return mapping
        
        if table:
            largest = max(content_phs, key=lambda x: x.area)
            mapping[largest.idx] = {'type': 'table', 'data': table}
            return mapping
        
        # CASE 2: Multi-section layout (subtitles + content)
        if sections and len(sections) >= 2:
            # This is comparison/multi-topic layout
            if isinstance(bullets, list) and bullets and isinstance(bullets[0], dict):
                # bullets = [{heading: ..., bullet_points: ...}, ...]
                for idx, section_data in enumerate(bullets[:len(sections)]):
                    section = sections[idx]
                    
                    # Map subtitle
                    if section['subtitle']:
                        mapping[section['subtitle'].idx] = {
                            'type': 'subtitle',
                            'text': section_data.get('heading', f'Section {idx+1}')
                        }
                    
                    # Map content to first content area under this subtitle
                    if section['content_areas']:
                        content_ph = section['content_areas'][0]
                        mapping[content_ph.idx] = {
                            'type': 'bullets',
                            'items': section_data.get('bullet_points', [])
                        }
            return mapping
        
        # CASE 3: Icon/Pictogram layout (small boxes or medium-wide)
        if self._is_icon_slide(slide_json):
            icon_items = [item for item in bullets if isinstance(item, str) and '[[' in item]
            
            # Use KPI grid if available
            if layout_capability.kpi_grid:
                boxes = layout_capability.kpi_grid['boxes']
                for idx, item in enumerate(icon_items[:len(boxes)]):
                    mapping[boxes[idx].idx] = {'type': 'icon', 'spec': item}
            else:
                # Use content placeholders ordered left-to-right
                sorted_phs = sorted(content_phs, key=lambda x: x.left)
                for idx, item in enumerate(icon_items[:len(sorted_phs)]):
                    mapping[sorted_phs[idx].idx] = {'type': 'icon', 'spec': item}
            
            return mapping
        
        # CASE 4: Simple bullets - use largest content placeholder
        if bullets:
            largest = max(content_phs, key=lambda x: x.area)
            mapping[largest.idx] = {'type': 'bullets', 'items': bullets}
        
        return mapping
    
    def select_layout_with_scoring(self, slide_json: dict) -> int:
        """ENHANCED scoring with space awareness"""
        
        content_type = self._infer_content_type_from_json(slide_json)
        
        scored_layouts = []
        
        for idx, layout_capability in self.analyzer.layouts.items():
            score = self._score_layout_for_content(
                layout_capability, 
                content_type, 
                slide_json
            )
            scored_layouts.append((score, idx, layout_capability))
        
        scored_layouts.sort(reverse=True, key=lambda x: x[0])
        
        if scored_layouts:
            best_score, best_idx, best_layout = scored_layouts[0]
            logger.info(f"✅ Layout {best_idx} ({best_layout.name}) - Score: {best_score:.1f}/100")
            logger.info(f"   Sections: {len(best_layout.semantic_sections)}, Type: {best_layout.layout_type}")
            return best_idx
        
        logger.warning("No suitable layout found, using layout 1")
        return 1
    
    def _find_alternative_layout(self, current_idx: int, slide_json: dict) -> int:
        """ADDED: Find alternative to avoid repetition"""
        content_type = self._infer_content_type_from_json(slide_json)
        
        # Get all suitable layouts excluding current
        alternatives = []
        for idx, layout in self.analyzer.layouts.items():
            if idx == current_idx:
                continue
            
            score = self._score_layout_for_content(layout, content_type, slide_json)
            if score > 50:  # Decent fit
                alternatives.append((score, idx))
        
        if alternatives:
            alternatives.sort(reverse=True)
            logger.info(f"✓ Found alternative: layout {alternatives[0][1]}")
            return alternatives[0][1]
        
        return current_idx

    def _score_layout_for_content(self, layout_capability, content_type: str, 
                              slide_json: dict) -> float:
        """ENHANCED scoring with space awareness"""
        score = 0.0
        
        # Base match
        if content_type in layout_capability.best_for:
            score += 40  # Increased from 50 to allow space for other factors
        
        # ENHANCED: Content-specific scoring
        if content_type == 'chart':
            score += self._score_for_chart(layout_capability, slide_json)
        
        elif content_type == 'table':
            score += self._score_for_table(layout_capability, slide_json)
        
        elif content_type == 'kpi_dashboard':
            score += self._score_for_kpi(layout_capability, slide_json)
        
        elif content_type == 'pictogram':
            score += self._score_for_pictogram(layout_capability, slide_json)
        
        elif content_type == 'comparison':
            score += self._score_for_comparison(layout_capability, slide_json)
        
        elif content_type == 'bullets':
            score += self._score_for_bullets(layout_capability, slide_json)
        
        # ADDED: Executive quality bonuses
        if layout_capability.visual_balance > 70:
            score += 5
        
        if layout_capability.fill_difficulty == "easy":
            score += 3
        
        return min(100.0, score)
    
    def _score_for_chart(self, layout, slide_json: dict) -> float:
        """ADDED: Smart chart scoring"""
        score = 0.0
        
        # Check capacity
        if layout.content_capacity['chart']['suitable']:
            score += 30
            
            # Bonus for very large area
            if layout.content_capacity['chart'].get('available_area', 0) > 50:
                score += 10
        
        # Prefer single large section
        if len(layout.semantic_sections) == 1:
            section = layout.semantic_sections[0]
            if len(section['content_areas']) == 1:
                if section['content_areas'][0].is_large_box:
                    score += 20
        
        return score
    
    def _score_for_table(self, layout, slide_json: dict) -> float:
        """ADDED: Smart table scoring"""
        score = 0.0
        
        table_data = slide_json.get('table', {})
        needed_cols = len(table_data.get('headers', []))
        needed_rows = len(table_data.get('rows', []))
        
        capacity = layout.content_capacity['table']
        
        # Can it fit?
        if capacity['max_cols'] >= needed_cols and capacity['max_rows'] >= needed_rows:
            score += 40
            
            # Bonus for good fit
            if capacity['max_cols'] <= needed_cols + 2:
                score += 10
        else:
            score += 10  # Partial
        
        # Prefer single large area
        if len(layout.semantic_sections) == 1:
            score += 10
        
        return score
    
    def _score_for_kpi(self, layout, slide_json: dict) -> float:
        """ADDED: Smart KPI scoring"""
        score = 0.0
        
        bullets = slide_json.get('bullet_points', [])
        needed_kpis = len(bullets) if isinstance(bullets, list) else 0
        
        if layout.kpi_grid:
            available = layout.content_capacity['kpis']['count']
            
            if available >= needed_kpis:
                score += 50
                
                if available == needed_kpis:
                    score += 10  # Perfect match
        else:
            # Check small boxes
            small_boxes = [ph for ph in layout.content_placeholders if ph.is_small_box]
            if len(small_boxes) >= needed_kpis:
                score += 30
        
        return score
    
    def _score_for_pictogram(self, layout, slide_json: dict) -> float:
        """ADDED: Smart pictogram scoring"""
        score = 0.0
        
        bullets = slide_json.get('bullet_points', [])
        needed_icons = len(bullets) if isinstance(bullets, list) else 0
        
        if layout.content_capacity['pictograms']['suitable']:
            estimated = layout.content_capacity['pictograms']['estimated_count']
            
            if estimated >= needed_icons:
                score += 40
                
                if abs(estimated - needed_icons) <= 1:
                    score += 10
        
        # Prefer medium-wide areas
        medium_wide = [ph for ph in layout.content_placeholders if ph.is_medium_box and ph.is_wide]
        if medium_wide:
            score += 10
        
        return score
    
    def _score_for_comparison(self, layout, slide_json: dict) -> float:
        """ENHANCED comparison scoring"""
        score = 0.0
        
        bullets = slide_json.get('bullet_points', [])
        
        if isinstance(bullets, list) and bullets and isinstance(bullets[0], dict):
            needed_cols = len(bullets)
        else:
            needed_cols = 2
        
        # Check semantic sections
        if len(layout.semantic_sections) == needed_cols:
            score += 50
        elif abs(len(layout.semantic_sections) - needed_cols) == 1:
            score += 30
        
        # Check spatial groups
        if needed_cols == 2 and 'left_column' in layout.spatial_groups:
            score += 10
        
        return score
    
    def _score_for_bullets(self, layout, slide_json: dict) -> float:
        """ENHANCED with density awareness"""
        score = 0.0
        
        bullets = slide_json.get('bullet_points', [])
        estimated_lines = self._estimate_bullet_lines(bullets)
        
        # Get recommended density
        density_rec = layout.content_density_recommendation
        target_bullets = density_rec.get('bullets_recommended', 10) if density_rec else 10
        
        capacity_lines = layout.content_capacity['bullets']['max_lines']
        
        # Perfect fit bonus
        if abs(estimated_lines - target_bullets) <= 2:
            score += 50  # Ideal fit
        # Can fit with good spacing
        elif capacity_lines >= estimated_lines:
            score += 40
            # Not too much empty space
            if capacity_lines <= estimated_lines + 5:
                score += 10
        else:
            score += 20  # Will be tight
        
        # Prefer layouts with good executive suitability
        if layout.executive_suitability >= 70:
            score += 10
        
        return score
    
    def _estimate_bullet_lines(self, bullets) -> int:
        """ADDED: Estimate lines needed"""
        if not isinstance(bullets, list):
            return 5
        
        lines = 0
        for item in bullets:
            if isinstance(item, str):
                lines += max(1, len(item) // 50)
            elif isinstance(item, list):
                lines += self._estimate_bullet_lines(item)
            elif isinstance(item, dict):
                lines += 2
                if 'bullet_points' in item:
                    lines += self._estimate_bullet_lines(item['bullet_points'])
        
        return lines

    def _infer_content_type_from_json(self, slide_json: dict) -> str:
        """Existing inference - unchanged"""
        
        if 'chart' in slide_json and slide_json['chart']:
            return 'chart'
        
        if 'table' in slide_json and slide_json['table']:
            return 'table'
        
        if 'bullet_points' in slide_json:
            bullets = slide_json['bullet_points']
            
            if not isinstance(bullets, list):
                return 'bullets'
            
            if not bullets:
                return 'bullets'
            
            if all(isinstance(item, str) and '[[' in item for item in bullets):
                return 'pictogram'
            
            if all(isinstance(item, dict) and 'heading' in item for item in bullets):
                if len(bullets) >= 4 and all(len(str(item.get('heading', ''))) < 20 for item in bullets):
                    return 'kpi_dashboard'
                else:
                    return 'comparison'
        
        return 'bullets'
    