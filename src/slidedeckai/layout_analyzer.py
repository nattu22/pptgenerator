"""
ENHANCED EXISTING layout_analyzer.py - NO NEW CLASSES
Only improvements to existing TemplateAnalyzer class
"""
import logging
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
from pptx import Presentation

logger = logging.getLogger(__name__)


@dataclass
class PlaceholderInfo:
    """Info about a single placeholder - ENHANCED"""
    idx: int
    type: str
    type_id: int
    left: float
    top: float
    width: float
    height: float
    area: float
    
    role: str = "content"
    position_group: str = ""
    
    # ADDED: Spatial characteristics for smart decisions
    aspect_ratio: float = 0.0
    is_small_box: bool = False
    is_medium_box: bool = False
    is_large_box: bool = False
    is_wide: bool = False
    is_tall: bool = False
    
    def __post_init__(self):
        """Calculate spatial characteristics"""
        self.aspect_ratio = self.width / self.height if self.height > 0 else 1.0
        self.is_small_box = self.area < 3.0
        self.is_medium_box = 3.0 <= self.area < 15.0
        self.is_large_box = self.area >= 15.0
        self.is_wide = self.aspect_ratio > 2.0
        self.is_tall = self.aspect_ratio < 0.5
    
    def to_dict(self):
        return {
            'idx': self.idx,
            'type': self.type,
            'type_id': self.type_id,
            'left': self.left,
            'top': self.top,
            'width': self.width,
            'height': self.height,
            'area': self.area,
            'role': self.role,
            'position_group': self.position_group,
            'aspect_ratio': self.aspect_ratio,
            'is_small': self.is_small_box,
            'is_medium': self.is_medium_box,
            'is_large': self.is_large_box
        }


@dataclass
class LayoutCapability:
    """Layout with proper role classification - ENHANCED"""
    idx: int
    name: str
    has_title: bool
    has_subtitle: bool
    has_chart: bool
    has_table: bool
    has_picture: bool
    
    subtitle_placeholders: List[PlaceholderInfo]
    content_placeholders: List[PlaceholderInfo]
    text_placeholders: List[PlaceholderInfo]
    all_placeholders: List[PlaceholderInfo]
    
    layout_type: str
    best_for: List[str]
    spatial_groups: Dict[str, List[PlaceholderInfo]]
    layout_story: str
    
    semantic_sections: List[Dict] = None
    kpi_grid: Optional[Dict] = None
    usable_content_area: float = 0.0
    content_capacity: Dict = None
    
    # ADDED: Executive metrics
    complexity_score: float = 0.0
    visual_balance: float = 0.0
    fill_difficulty: str = "medium"
    recommended_verbosity: int = 7
    executive_score: float = 0.0 

    semantic_story_type: str = "general_content"
    executive_suitability: float = 0.0
    content_density_recommendation: Dict = None
    layout_category: str = "large_content"  # NEW: blank, cover, section_divider, kpicards, small_content, large_content
    
    def __post_init__(self):
        if self.semantic_sections is None:
            self.semantic_sections = []
        if self.content_capacity is None:
            self.content_capacity = {}
        if self.content_density_recommendation is None:  # NEW
            self.content_density_recommendation = {}
    
    def to_dict(self):
        return {
            'idx': self.idx,
            'name': self.name,
            'has_title': self.has_title,
            'has_subtitle': self.has_subtitle,
            'has_chart': self.has_chart,
            'has_table': self.has_table,
            'has_picture': self.has_picture,
            'content_count': len(self.content_placeholders),
            'subtitle_count': len(self.subtitle_placeholders),
            'text_count': len(self.text_placeholders),
            'layout_type': self.layout_type,
            'best_for': self.best_for,
            'spatial_groups': {k: [p.to_dict() for p in v] for k, v in self.spatial_groups.items()},
            'layout_story': self.layout_story,
            'semantic_sections': self.semantic_sections,
            'kpi_grid': self.kpi_grid,
            'usable_content_area': self.usable_content_area,
            'content_capacity': self.content_capacity,
            'complexity_score': self.complexity_score,
            'visual_balance': self.visual_balance,
            'fill_difficulty': self.fill_difficulty,
            'recommended_verbosity': self.recommended_verbosity,
            'placeholders': {
                'subtitles': [p.to_dict() for p in self.subtitle_placeholders],
                'content': [p.to_dict() for p in self.content_placeholders],
                'text': [p.to_dict() for p in self.text_placeholders],
                'all': [p.to_dict() for p in self.all_placeholders]
            }
        }


class TemplateAnalyzer:
    """ENHANCED - Same class name, better intelligence"""
    
    PLACEHOLDER_TYPE_NAMES = {
        1: 'TITLE', 2: 'BODY', 3: 'CENTER_TITLE', 4: 'SUBTITLE',
        5: 'DATE', 6: 'SLIDE_NUMBER', 7: 'FOOTER', 8: 'HEADER',
        9: 'OBJECT', 10: 'CHART', 11: 'TABLE', 12: 'CLIP_ART',
        13: 'ORG_CHART', 14: 'MEDIA', 15: 'PICTURE',
        16: 'VERTICAL_BODY', 17: 'VERTICAL_OBJECT', 18: 'VERTICAL_TITLE',
    }
    
    def __init__(self, presentation: Presentation):
        self.presentation = presentation
        self.layouts: Dict[int, LayoutCapability] = {}
        self.slide_width = 10.0
        self.slide_height = 7.5
        self._analyze_all_layouts()
    
    def _analyze_all_layouts(self):
        """ENHANCED analysis with executive intelligence"""
        logger.info("🔍 Starting ENHANCED template analysis...")
        
        for idx, layout in enumerate(self.presentation.slide_layouts):
            try:
                capability = self._analyze_single_layout(idx, layout)
                self.layouts[idx] = capability
                logger.info(f"  ✓ Layout {idx}: {layout.name}")
                logger.info(f"    Sections: {len(capability.semantic_sections)}, Content: {len(capability.content_placeholders)}")
                logger.info(f"    Best for: {', '.join(capability.best_for[:3])}")
            except Exception as e:
                logger.error(f"  ✗ Failed layout {idx}: {e}")
                # ✅ ADD: Create minimal fallback instead of skipping
                self.layouts[idx] = self._create_fallback_layout(idx, layout)
        
        # ✅ ADD: Verify no gaps in indices
        expected_indices = set(range(len(self.presentation.slide_layouts)))
        actual_indices = set(self.layouts.keys())
        missing = expected_indices - actual_indices
        
        if missing:
            logger.warning(f"  ⚠️ Missing layouts: {missing}")
        
        logger.info(f"  ✓ Analyzed {len(self.layouts)} layouts successfully")
        
    def _analyze_single_layout(self, idx: int, layout) -> LayoutCapability:
        """ENHANCED with smart grouping and metrics"""
        
        has_title = False
        has_subtitle = False
        has_chart = False
        has_table = False
        has_picture = False
        
        all_placeholders = []
        subtitle_placeholders = []
        content_placeholders = []
        text_placeholders = []

        
        for shape in layout.placeholders:
            try:
                ph_idx = shape.placeholder_format.idx
                ph_type_id = shape.placeholder_format.type
                ph_type_name = self.PLACEHOLDER_TYPE_NAMES.get(ph_type_id, f'UNKNOWN_{ph_type_id}')
                
                left = shape.left / 914400.0
                top = shape.top / 914400.0
                width = shape.width / 914400.0
                height = shape.height / 914400.0
                area = width * height
                
                # ENHANCED role classification
                role = self._classify_placeholder_role(ph_type_id, ph_type_name, width, height, area)
                
                placeholder_info = PlaceholderInfo(
                    idx=ph_idx,
                    type=ph_type_name,
                    type_id=ph_type_id,
                    left=left,
                    top=top,
                    width=width,
                    height=height,
                    area=area,
                    role=role
                )
                
                all_placeholders.append(placeholder_info)
                
                if ph_type_id == 1:
                    has_title = True
                    placeholder_info.role = 'title'
                
                elif ph_type_id == 4:
                    has_subtitle = True
                    placeholder_info.role = 'subtitle'
                    subtitle_placeholders.append(placeholder_info)
                
                elif ph_type_id == 10:
                    has_chart = True
                    content_placeholders.append(placeholder_info)
                
                elif ph_type_id == 11:
                    has_table = True
                    content_placeholders.append(placeholder_info)
                
                elif ph_type_id == 15:
                    has_picture = True
                    content_placeholders.append(placeholder_info)
                
                elif ph_type_id in [2, 9, 16, 17]:
                    if role == 'subtitle':
                        subtitle_placeholders.append(placeholder_info)
                    else:
                        content_placeholders.append(placeholder_info)
                        if ph_type_id in [2, 16]:
                            text_placeholders.append(placeholder_info)
                
            except Exception as e:
                logger.warning(f"Could not process placeholder: {e}")
                continue
        
        # 1. Detect KPI grid first
        kpi_grid = self._detect_kpi_grid(content_placeholders)
        
        # 2. Then spatial grouping
        spatial_groups = self._group_by_spatial_position(content_placeholders)
        
        # 3. Then semantic sections
        semantic_sections = self._group_placeholders_semantically(subtitle_placeholders, content_placeholders)
        
        # 4. Then story type (needs kpi_grid)
        story_type = self._infer_semantic_story_type(semantic_sections, content_placeholders, kpi_grid)
        
        # 5. Then content capacity (needs kpi_grid)
        content_capacity = self._calculate_content_capacity(content_placeholders, semantic_sections, kpi_grid)
    
        executive_score = self._calculate_executive_score(
            semantic_sections, content_placeholders, subtitle_placeholders
        )
        self._match_subtitles_to_groups(subtitle_placeholders, spatial_groups)
        
        layout_story = self._infer_layout_story(spatial_groups, has_chart, has_table, kpi_grid, semantic_sections)
        
        layout_type = self._infer_layout_type(
            has_chart, has_table, has_picture,
            len(content_placeholders), len(text_placeholders), len(semantic_sections), kpi_grid
        )
        
        best_for = self._determine_best_use(
            has_chart, has_table, has_picture,
            content_placeholders, text_placeholders, spatial_groups, semantic_sections, kpi_grid
        )
        
        # ADDED: Executive metrics
        complexity_score = self._calculate_complexity(semantic_sections, content_placeholders)
        visual_balance = self._calculate_balance(content_placeholders)
        fill_difficulty, recommended_verbosity = self._assess_fill_difficulty(semantic_sections, content_placeholders)
        executive_suitability = self._calculate_executive_suitability(
        visual_balance, complexity_score, semantic_sections, story_type
        )
        
        content_density_rec = self._recommend_content_density(
            sum(ph.area for ph in content_placeholders),
            semantic_sections,
            story_type
        )

        # ADDED: Determine specific layout category
        layout_category = self._determine_layout_category(
            has_title, has_subtitle, kpi_grid, content_placeholders,
            text_placeholders, semantic_sections, layout.name
        )

        return LayoutCapability(
            idx=idx,
            name=layout.name,
            has_title=has_title,
            has_subtitle=has_subtitle,
            has_chart=has_chart,
            has_table=has_table,
            has_picture=has_picture,
            subtitle_placeholders=subtitle_placeholders,
            content_placeholders=content_placeholders,
            text_placeholders=text_placeholders,
            all_placeholders=all_placeholders,
            layout_type=layout_type,
            best_for=best_for,
            spatial_groups=spatial_groups,
            layout_story=layout_story,
            semantic_sections=semantic_sections,
            kpi_grid=kpi_grid,
            usable_content_area=sum(ph.area for ph in content_placeholders),
            content_capacity=content_capacity,
            complexity_score=complexity_score,
            visual_balance=visual_balance,
            fill_difficulty=fill_difficulty,
            recommended_verbosity=recommended_verbosity,
            executive_score=executive_score,
            semantic_story_type=story_type,  # NEW
            executive_suitability=executive_suitability,  # NEW
            content_density_recommendation=content_density_rec,  # NEW
            layout_category=layout_category  # NEW
        )

    def _group_placeholders_semantically(self, 
                                      subtitles: List[PlaceholderInfo],
                                      content_areas: List[PlaceholderInfo]) -> List[Dict]:
        """ENHANCED: Group subtitle with content areas directly below"""
        sections = []
        used_content = set()
        
        for subtitle in subtitles:
            related_content = []
            
            for content in content_areas:
                if content.idx in used_content:
                    continue
                
                # Check if content is 0-1.0" below subtitle (INCREASED tolerance)
                vertical_distance = content.top - subtitle.top
                if not (0 < vertical_distance < 1.0):
                    continue
                
                # Check horizontal alignment (±1.5" tolerance)
                if abs(content.left - subtitle.left) > 1.5:
                    continue
                
                related_content.append(content)
                used_content.add(content.idx)
            
            if related_content:
                # ADDED: Detect pattern
                pattern = self._detect_section_pattern(related_content)
                section_best_for = self._infer_section_best_for(related_content, pattern)
                
                sections.append({
                    'subtitle': subtitle,
                    'content_areas': related_content,
                    'total_capacity': sum(c.area for c in related_content),
                    'section_id': f"section_{subtitle.idx}",
                    'layout_pattern': pattern,
                    'best_for': section_best_for
                })
        
        logger.info(f"    Found {len(sections)} semantic sections")
        return sections
    
    def _detect_section_pattern(self, content_areas: List[PlaceholderInfo]) -> str:
        """ADDED: Detect pattern within a section"""
        if len(content_areas) == 1:
            return "single"
        
        small_count = sum(1 for c in content_areas if c.is_small_box)
        if small_count >= 3:
            return "grid"
        
        if len(content_areas) >= 2:
            sorted_by_left = sorted(content_areas, key=lambda x: x.left)
            if abs(sorted_by_left[0].top - sorted_by_left[1].top) < 0.5:
                return "columns"
        
        return "mixed"
    
    def _infer_section_best_for(self, content_areas: List[PlaceholderInfo], pattern: str) -> List[str]:
        """ADDED: Infer what content fits this section"""
        best_for = []
        
        if pattern == "single":
            if content_areas[0].is_large_box:
                best_for.extend(['chart', 'table', 'bullets'])
            elif content_areas[0].is_medium_box:
                best_for.extend(['bullets', 'pictogram'])
        elif pattern == "grid":
            best_for.extend(['kpi_dashboard', 'icon_grid'])
        elif pattern == "columns":
            best_for.extend(['comparison', 'bullets'])
        
        return best_for
    
    def _detect_kpi_grid(self, placeholders: List[PlaceholderInfo]) -> Optional[Dict]:
        """ENHANCED KPI grid detection"""
        small_boxes = [ph for ph in placeholders if ph.is_small_box]
        
        if len(small_boxes) < 4:
            return None
        
        # Group by row (0.3" tolerance)
        rows = {}
        for box in small_boxes:
            row_key = round(box.top * 3) / 3
            if row_key not in rows:
                rows[row_key] = []
            rows[row_key].append(box)
        
        if len(rows) < 2 or any(len(row) < 2 for row in rows.values()):
            return None
        
        areas = [box.area for box in small_boxes]
        avg_area = sum(areas) / len(areas)
        max_deviation = max(abs(a - avg_area) for a in areas)
        
        if max_deviation > avg_area * 0.3:
            return None
        
        grid_rows = len(rows)
        grid_cols = len(rows[list(rows.keys())[0]])
        
        logger.info(f"    ✅ KPI Grid: {grid_rows}x{grid_cols}")
        
        return {
            'boxes': small_boxes,
            'rows': grid_rows,
            'cols': grid_cols,
            'total_area': sum(areas),
            'avg_box_size': avg_area
        }
    
    def _calculate_content_capacity(self, 
                                    content_placeholders: List[PlaceholderInfo],
                                    semantic_sections: List[Dict],
                                    kpi_grid: Optional[Dict]) -> Dict:
        """ENHANCED capacity calculation"""
        capacity = {
            'bullets': {'max_lines': 0, 'chars_per_line': 0, 'estimated_words': 0},
            'table': {'max_cols': 0, 'max_rows': 0},
            'chart': {'suitable': False, 'min_area': 30, 'available_area': 0},
            'kpis': {'count': 0},
            'pictograms': {'suitable': False, 'estimated_count': 0},
            'sections': len(semantic_sections)
        }
        
        if kpi_grid:
            capacity['kpis']['count'] = len(kpi_grid['boxes'])
        
        # ENHANCED bullet capacity
        text_areas = [ph for ph in content_placeholders if ph.height > 1.0]
        if text_areas:
            largest = max(text_areas, key=lambda x: x.area)
            capacity['bullets']['max_lines'] = int(largest.height / 0.3)
            capacity['bullets']['chars_per_line'] = int(largest.width * 8)
            capacity['bullets']['estimated_words'] = int(largest.area * 20)
        
        # Table capacity
        if content_placeholders:
            largest = max(content_placeholders, key=lambda x: x.area)
            capacity['table']['max_cols'] = max(2, int(largest.width / 1.5))
            capacity['table']['max_rows'] = max(3, int(largest.height / 0.4))
        
        # ENHANCED chart suitability
        large_areas = [ph for ph in content_placeholders if ph.is_large_box]
        if large_areas:
            capacity['chart']['suitable'] = True
            capacity['chart']['available_area'] = max(ph.area for ph in large_areas)
        
        # Pictogram suitability
        medium_areas = [ph for ph in content_placeholders if ph.is_medium_box and ph.is_wide]
        if medium_areas:
            capacity['pictograms']['suitable'] = True
            capacity['pictograms']['estimated_count'] = int(medium_areas[0].width / 1.5)
        
        return capacity
    
    def _calculate_complexity(self, semantic_sections: List[Dict], 
                             content_placeholders: List[PlaceholderInfo]) -> float:
        """ADDED: Calculate complexity score"""
        score = 0.0
        score += min(len(semantic_sections) * 15, 40)
        score += min(len(content_placeholders) * 8, 40)
        small_count = sum(1 for ph in content_placeholders if ph.is_small_box)
        score += min(small_count * 5, 20)
        return min(score, 100.0)
    
    def _calculate_balance(self, content_placeholders: List[PlaceholderInfo]) -> float:
        """ADDED: Calculate visual balance"""
        if not content_placeholders:
            return 0.0
        
        areas = [ph.area for ph in content_placeholders]
        avg_area = sum(areas) / len(areas)
        max_deviation = max(abs(a - avg_area) for a in areas) if areas else 0
        
        balance = 100 - min((max_deviation / avg_area * 100) if avg_area > 0 else 100, 100)
        return balance
    
    def _assess_fill_difficulty(self, semantic_sections: List[Dict],
                                content_placeholders: List[PlaceholderInfo]) -> Tuple[str, int]:
        """ADDED: Assess fill difficulty"""
        section_count = len(semantic_sections)
        ph_count = len(content_placeholders)
        
        if section_count <= 2 and ph_count <= 3:
            return "easy", 7
        if section_count <= 4 and ph_count <= 6:
            return "medium", 8
        return "hard", 9

    def _determine_layout_category(self, has_title: bool, has_subtitle: bool,
                                   kpi_grid: Optional[Dict],
                                   content_placeholders: List[PlaceholderInfo],
                                   text_placeholders: List[PlaceholderInfo],
                                   semantic_sections: List[Dict],
                                   layout_name: str) -> str:
        """NEW: Classify into 6 strict categories"""

        # 1. Blank
        if not content_placeholders and not has_title:
            return "blank"

        # 2. Cover / Title Only
        if has_title and not content_placeholders:
            # If it looks like a title slide
            if "title" in layout_name.lower() and "only" not in layout_name.lower():
                return "cover"
            return "section_divider"

        # 3. Section Divider
        if has_title and not content_placeholders and "section" in layout_name.lower():
            return "section_divider"

        if len(content_placeholders) == 0 and len(text_placeholders) == 0:
            if has_title:
                return "section_divider"
            return "blank"

        # 4. KPI Cards
        if kpi_grid:
            return "kpicards"

        # Check if it has many small boxes
        small_boxes = sum(1 for ph in content_placeholders if ph.is_small_box)
        if small_boxes >= 4:
            return "kpicards"

        # 5. Large Content (Charts, Tables, Big Text)
        large_areas = sum(1 for ph in content_placeholders if ph.is_large_box)
        if large_areas >= 1:
            return "large_content"

        if len(content_placeholders) == 1 and content_placeholders[0].area > 10:
            return "large_content"

        # 6. Small Content (Bullets, comparisons, multi-column)
        return "small_content"

    def _calculate_executive_suitability(self, 
                                     visual_balance: float,
                                     complexity_score: float,
                                     semantic_sections: List[Dict],
                                     story_type: str) -> float:
        """
        CRITICAL: Rate this layout for EXECUTIVE presentations
        High score = Clear, impactful, professional
        """
        
        score = 0.0
        
        # Balance is king (execs hate clutter)
        score += (visual_balance / 100) * 40  # Max 40 points
        
        # Moderate complexity is good (not too simple, not overwhelming)
        if 30 <= complexity_score <= 60:
            score += 30
        elif complexity_score < 30:
            score += 20  # Too simple
        else:
            score += 10  # Too complex
        
        # Story clarity bonus
        executive_story_types = [
            "metrics_dashboard",
            "data_visualization", 
            "balanced_comparison",
            "three_stage_narrative"
        ]
        if story_type in executive_story_types:
            score += 20
        elif story_type in ["focused_message", "main_supporting"]:
            score += 15
        else:
            score += 5
        
        # Section count sweet spot (execs prefer 1-3 clear sections)
        if 1 <= len(semantic_sections) <= 3:
            score += 10
        else:
            score += 3
        
        return min(score, 100.0)
    
    
    def _recommend_content_density(self, 
                                    usable_area: float,
                                    semantic_sections: List[Dict],
                                    story_type: str) -> Dict:
        """
        CRITICAL: Tell content generator HOW MUCH to write
        Based on ACTUAL available space
        """
        
        # Calculate words per square inch (executive style = sparse)
        exec_density = 15  # words per sq inch for professional look
        student_density = 30  # cramped, avoid this
        
        target_density = exec_density
        
        # Adjust by story type
        if story_type in ["metrics_dashboard", "feature_grid"]:
            target_density = 10  # Very sparse, number-focused
        elif story_type in ["detailed_analysis"]:
            target_density = 20  # Can be denser
        
        total_words = int(usable_area * target_density)
        
        # Distribute across sections
        if semantic_sections:
            words_per_section = total_words // len(semantic_sections)
        else:
            words_per_section = total_words
        
        # Bullet density
        if story_type == "metrics_dashboard":
            bullets_recommended = 4 + (len(semantic_sections) * 2)  # Sparse
        elif story_type in ["balanced_comparison", "three_stage_narrative"]:
            bullets_recommended = 6 + (len(semantic_sections) * 3)  # Moderate
        else:
            bullets_recommended = 8 + (len(semantic_sections) * 4)  # Detailed
        
        return {
            "total_words_target": total_words,
            "words_per_section": words_per_section,
            "density_style": "executive" if target_density <= 15 else "detailed",
            "bullets_recommended": bullets_recommended,
            "verbosity_level": 6 if target_density <= 15 else 8,
            "avoid_overflow": True
        }
        
    def _infer_semantic_story_type(self, semantic_sections: List[Dict], 
                                content_placeholders: List[PlaceholderInfo],
                                kpi_grid: Optional[Dict]) -> str:
        """
        CRITICAL: Understand WHAT STORY this layout tells
        Not just "2 column" but "comparison narrative"
        """
        
        if kpi_grid:
            return "metrics_dashboard"  # Executive summary style
        
        section_count = len(semantic_sections)
        total_area = sum(ph.area for ph in content_placeholders)
        
        # SINGLE LARGE AREA = Deep dive content
        if section_count == 1:
            largest = max(content_placeholders, key=lambda x: x.area)
            if largest.area > 40:  # Very large
                if largest.aspect_ratio > 1.5:
                    return "data_visualization"  # Chart/table focus
                else:
                    return "detailed_analysis"  # Text-heavy deep dive
            return "focused_message"  # Single key point
        
        # DUAL SECTIONS = Comparison/contrast
        if section_count == 2:
            areas = [sum(ph.area for ph in s['content_areas']) for s in semantic_sections]
            if abs(areas[0] - areas[1]) < 5:  # Balanced
                return "balanced_comparison"  # Before/after, pros/cons
            else:
                return "main_supporting"  # Primary + evidence
        
        # TRIPLE = Process/stages
        if section_count == 3:
            return "three_stage_narrative"  # Problem-solution-outcome
        
        # MANY SMALL = Grid showcase
        if len(content_placeholders) >= 6 and all(ph.is_small_box for ph in content_placeholders):
            return "feature_grid"  # Multiple parallel points
        
        # MIXED SIZES = Hierarchy
        large_count = sum(1 for ph in content_placeholders if ph.is_large_box)
        small_count = sum(1 for ph in content_placeholders if ph.is_small_box)
        if large_count >= 1 and small_count >= 2:
            return "hierarchical_story"  # Main point + supporting facts
        
        return "general_content"
    def _calculate_executive_score(self, sections, content_phs, subtitle_phs) -> float:
        """
        Score layout for executive presentations
        Executives want: clear hierarchy, visual focus, minimal text density
        """
        score = 50.0  # baseline
        
        # Prefer 1-3 clear sections (not overwhelming)
        if 1 <= len(sections) <= 3:
            score += 20
        elif len(sections) > 5:
            score -= 15
        
        # Prefer layouts with subtitles (clear structure)
        if subtitle_phs:
            score += 15
        
        # Penalize text-heavy layouts
        text_heavy = sum(1 for ph in content_phs if ph.height > 3.0)
        if text_heavy > 2:
            score -= 10
        
        # Reward visual balance
        if self._has_visual_balance(content_phs):
            score += 15
        
        return min(100.0, max(0.0, score))
    
    def _has_visual_balance(self, placeholders) -> bool:
        """Check if placeholders are visually balanced"""
        if len(placeholders) < 2:
            return True
        
        areas = [ph.area for ph in placeholders]
        avg_area = sum(areas) / len(areas)
        max_deviation = max(abs(a - avg_area) for a in areas)
        
        # Balanced if no placeholder is 2x larger than average
        return max_deviation / avg_area < 1.0
        
    def _classify_placeholder_role(self, type_id: int, type_name: str, 
                                    width: float, height: float, area: float) -> str:
        """ENHANCED role classification"""
        if type_id == 4:
            return 'subtitle'
        if type_id == 1:
            return 'title'
        if type_id in [5, 6, 7, 8]:
            return 'footer'
        if type_id in [10, 11, 15]:
            return 'content'
        
        # ENHANCED heuristics
        if type_id in [2, 9, 16, 17]:
            if height < 0.5:
                return 'subtitle'
            elif area < 1.0:
                return 'subtitle'
            else:
                aspect = width / height if height > 0 else 1.0
                if aspect > 3.0 and height < 0.8:
                    return 'subtitle'
                return 'content'
        
        return 'content'
    
    def _group_by_spatial_position(self, placeholders: List[PlaceholderInfo]) -> Dict[str, List[PlaceholderInfo]]:
        """Existing spatial grouping - unchanged"""
        if not placeholders:
            return {}
        
        left_positions = sorted(set(round(p.left, 1) for p in placeholders))
        top_positions = sorted(set(round(p.top, 1) for p in placeholders))
        
        groups = {}
        
        if len(left_positions) == 1:
            if len(top_positions) == 1:
                groups['center'] = placeholders
            else:
                for i, top in enumerate(top_positions):
                    groups[f'row_{i+1}'] = [p for p in placeholders if round(p.top, 1) == top]
        
        elif len(left_positions) == 2:
            mid_x = sum(left_positions) / 2
            groups['left_column'] = [p for p in placeholders if p.left < mid_x]
            groups['right_column'] = [p for p in placeholders if p.left >= mid_x]
        
        elif len(left_positions) == 3:
            sorted_x = sorted(left_positions)
            groups['left_column'] = [p for p in placeholders if round(p.left, 1) == sorted_x[0]]
            groups['center_column'] = [p for p in placeholders if round(p.left, 1) == sorted_x[1]]
            groups['right_column'] = [p for p in placeholders if round(p.left, 1) == sorted_x[2]]
        
        else:
            for i, p in enumerate(placeholders):
                groups[f'cell_{i+1}'] = [p]
        
        for group_name, group_phs in groups.items():
            for ph in group_phs:
                ph.position_group = group_name
        
        return groups
    
    def _match_subtitles_to_groups(self, subtitles: List[PlaceholderInfo], 
                                     spatial_groups: Dict[str, List[PlaceholderInfo]]):
        """Existing subtitle matching - unchanged"""
        for subtitle in subtitles:
            min_dist = float('inf')
            closest_group = None
            
            for group_name, content_phs in spatial_groups.items():
                if not content_phs:
                    continue
                
                content_top = content_phs[0].top
                dist = abs(subtitle.top - content_top)
                
                if dist < min_dist:
                    min_dist = dist
                    closest_group = group_name
            
            if closest_group:
                subtitle.position_group = f"{closest_group}_subtitle"
    
    def _infer_layout_story(self, spatial_groups: Dict, has_chart: bool, 
                       has_table: bool, kpi_grid=None, semantic_sections=None) -> str:
        """ENHANCED layout story"""
        if kpi_grid:
            return f"KPI Dashboard ({kpi_grid['rows']}x{kpi_grid['cols']} metrics)"
        
        if semantic_sections and len(semantic_sections) >= 3:
            return f"{len(semantic_sections)} topic sections"
        
        num_groups = len(spatial_groups)
        group_names = list(spatial_groups.keys())
        
        if has_chart:
            return "Chart with supporting text"
        if has_table:
            return "Data table presentation"
        
        if 'left_column' in group_names and 'right_column' in group_names:
            return "Two column comparison"
        
        if num_groups == 3 and all('column' in g for g in group_names):
            return "Three column layout"
        
        if num_groups >= 1 and all('row' in g for g in group_names):
            return f"Vertical stack ({num_groups} sections)"
        
        if num_groups == 1:
            return "Single content area"
        
        return f"Multi-area layout ({num_groups} areas)"
    
    def _infer_layout_type(self, has_chart: bool, has_table: bool, has_picture: bool,
                          content_count: int, text_count: int, section_count: int = 0, kpi_grid=None) -> str:
        """ENHANCED layout type"""
        if kpi_grid:
            return 'kpi_dashboard'
        if has_chart:
            return 'chart_layout'
        if has_table:
            return 'table_layout'
        if has_picture:
            return 'image_layout'
        
        if section_count >= 3:
            return 'multi_section'
        elif section_count == 2:
            return 'double_section'
        elif section_count == 1:
            return 'single_section'
        
        if text_count == 0:
            return 'title_only'
        elif text_count == 1:
            return 'single_column'
        elif text_count == 2:
            return 'double_column'
        elif text_count == 3:
            return 'triple_column'
        else:
            return 'multi_column'
    
    def _determine_best_use(self, has_chart: bool, has_table: bool, has_picture: bool,
                       content_placeholders: List[PlaceholderInfo],
                       text_placeholders: List[PlaceholderInfo],
                       spatial_groups: Dict,
                       semantic_sections: List[Dict] = None,
                       kpi_grid=None) -> List[str]:
        """ENHANCED best use determination"""
        best_for = []
        
        if kpi_grid:
            best_for.extend(['kpi_dashboard', 'metrics', 'scorecard'])
        
        if has_chart:
            best_for.append('chart')
        if has_table:
            best_for.append('table')
        
        # ADDED: Section-based inference
        if semantic_sections:
            for section in semantic_sections:
                if 'best_for' in section:
                    best_for.extend(section['best_for'])
        
        num_groups = len(spatial_groups)
        
        if 'left_column' in spatial_groups and 'right_column' in spatial_groups:
            best_for.extend(['comparison', 'before_after'])
        
        if num_groups == 3:
            best_for.extend(['three_points', 'process_steps'])
        
        if num_groups >= 4 and not kpi_grid:
            best_for.append('icon_grid')
        
        medium_areas = [ph for ph in content_placeholders if ph.is_medium_box]
        if medium_areas:
            best_for.append('pictogram')
        
        if not best_for:
            best_for.append('bullets')
        
        return list(set(best_for))  # Remove duplicates
    
    def export_analysis(self) -> dict:
        """Existing export - unchanged"""
        return {
            'template_name': 'Analyzed Template',
            'total_layouts': len(self.layouts),
            'layouts': {idx: layout.to_dict() for idx, layout in self.layouts.items()}
        }

    def print_summary(self):
        """ENHANCED summary"""
        logger.info("\n" + "="*80)
        logger.info(f"TEMPLATE ANALYSIS SUMMARY")
        logger.info("="*80)
        logger.info(f"Total layouts: {len(self.layouts)}")
        
        for idx, layout in self.layouts.items():
            logger.info(f"\n  Layout {idx}: {layout.name}")
            logger.info(f"    Type: {layout.layout_type}")
            logger.info(f"    Best for: {', '.join(layout.best_for[:3])}")
            logger.info(f"    Placeholders: {len(layout.content_placeholders)} content, {len(layout.subtitle_placeholders)} subtitle")
            logger.info(f"    Sections: {len(layout.semantic_sections)}")
            logger.info(f"    Complexity: {layout.complexity_score:.0f}/100, Balance: {layout.visual_balance:.0f}/100")
            
            if hasattr(layout, 'kpi_grid') and layout.kpi_grid:
                logger.info(f"    KPI Grid: {layout.kpi_grid['rows']}x{layout.kpi_grid['cols']}")
        
        logger.info("="*80 + "\n")