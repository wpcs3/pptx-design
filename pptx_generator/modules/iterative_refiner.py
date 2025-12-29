"""
Iterative Refinement Module

PPTAgent-inspired iterative refinement that evaluates presentations
and refines them based on quality feedback.

Key Features:
- Multi-pass refinement loop
- LLM-powered outline improvement based on evaluation issues
- Configurable quality thresholds
- Refinement history tracking

Phase 4 Enhancement (2025-12-29):
Based on PPTAgent's iterative generation approach (github.com/icip-cas/PPTAgent)
"""

import copy
import json
import logging
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


# =============================================================================
# Refinement Types
# =============================================================================

class RefinementType(Enum):
    """Types of refinements that can be applied."""
    ADD_CONTENT = "add_content"           # Slide needs more content
    REDUCE_CONTENT = "reduce_content"     # Slide has too much content
    SPLIT_SLIDE = "split_slide"           # Slide should be split into multiple
    MERGE_SLIDES = "merge_slides"         # Multiple slides should be combined
    ADD_VISUAL = "add_visual"             # Slide needs chart/table/image
    FIX_STRUCTURE = "fix_structure"       # Section structure issues
    IMPROVE_FLOW = "improve_flow"         # Narrative flow issues
    CLARIFY_TITLE = "clarify_title"       # Title needs improvement


@dataclass
class RefinementAction:
    """A specific refinement action to apply."""
    refinement_type: RefinementType
    target_slide_index: int
    description: str
    suggested_change: Dict[str, Any] = field(default_factory=dict)
    priority: int = 1  # 1 = high, 2 = medium, 3 = low

    def __repr__(self):
        return f"Refinement({self.refinement_type.value}, slide {self.target_slide_index})"


@dataclass
class RefinementResult:
    """Result of a refinement iteration."""
    iteration: int
    original_score: float
    new_score: float
    grade: str
    actions_applied: List[RefinementAction]
    issues_remaining: List[str]
    improved: bool

    @property
    def improvement(self) -> float:
        return self.new_score - self.original_score


@dataclass
class RefinementHistory:
    """Complete history of refinement iterations."""
    initial_score: float
    initial_grade: str
    iterations: List[RefinementResult] = field(default_factory=list)
    final_score: float = 0.0
    final_grade: str = "F"
    total_iterations: int = 0
    converged: bool = False
    reason: str = ""

    def add_iteration(self, result: RefinementResult):
        self.iterations.append(result)
        self.total_iterations = len(self.iterations)
        self.final_score = result.new_score
        self.final_grade = result.grade

    def summary(self) -> str:
        lines = [
            "Refinement Summary",
            "=" * 40,
            f"Initial: {self.initial_score:.2f} ({self.initial_grade})",
            f"Final:   {self.final_score:.2f} ({self.final_grade})",
            f"Iterations: {self.total_iterations}",
            f"Improvement: {self.final_score - self.initial_score:+.2f}",
            f"Status: {'Converged' if self.converged else 'Stopped'} - {self.reason}",
        ]
        return "\n".join(lines)


# =============================================================================
# Issue Analyzer
# =============================================================================

class IssueAnalyzer:
    """Analyzes evaluation issues and generates refinement actions."""

    # Patterns to identify issue types
    ISSUE_PATTERNS = {
        "too many": RefinementType.REDUCE_CONTENT,
        "too few": RefinementType.ADD_CONTENT,
        "exceeds": RefinementType.REDUCE_CONTENT,
        "missing": RefinementType.ADD_CONTENT,
        "no visual": RefinementType.ADD_VISUAL,
        "needs chart": RefinementType.ADD_VISUAL,
        "needs image": RefinementType.ADD_VISUAL,
        "unclear title": RefinementType.CLARIFY_TITLE,
        "vague title": RefinementType.CLARIFY_TITLE,
        "flow": RefinementType.IMPROVE_FLOW,
        "transition": RefinementType.IMPROVE_FLOW,
        "structure": RefinementType.FIX_STRUCTURE,
        "section": RefinementType.FIX_STRUCTURE,
        "split": RefinementType.SPLIT_SLIDE,
        "dense": RefinementType.SPLIT_SLIDE,
        "empty": RefinementType.ADD_CONTENT,
        "sparse": RefinementType.ADD_CONTENT,
    }

    def analyze_issues(
        self,
        issues: List[str],
        slide_count: int
    ) -> List[RefinementAction]:
        """
        Analyze evaluation issues and generate refinement actions.

        Args:
            issues: List of issue strings from evaluation
            slide_count: Total number of slides

        Returns:
            List of RefinementActions sorted by priority
        """
        actions = []

        for issue in issues:
            action = self._parse_issue(issue, slide_count)
            if action:
                actions.append(action)

        # Sort by priority
        actions.sort(key=lambda a: a.priority)

        return actions

    def _parse_issue(self, issue: str, slide_count: int) -> Optional[RefinementAction]:
        """Parse a single issue into a refinement action."""
        issue_lower = issue.lower()

        # Extract slide index if mentioned
        slide_idx = self._extract_slide_index(issue)
        if slide_idx is None:
            slide_idx = -1  # Apply to presentation level

        # Determine refinement type
        ref_type = None
        for pattern, rtype in self.ISSUE_PATTERNS.items():
            if pattern in issue_lower:
                ref_type = rtype
                break

        if ref_type is None:
            ref_type = RefinementType.ADD_CONTENT  # Default

        # Determine priority based on issue severity
        priority = 2  # Default medium
        if any(w in issue_lower for w in ["critical", "major", "too many", "exceeds"]):
            priority = 1
        elif any(w in issue_lower for w in ["minor", "could", "consider"]):
            priority = 3

        return RefinementAction(
            refinement_type=ref_type,
            target_slide_index=slide_idx,
            description=issue,
            priority=priority
        )

    def _extract_slide_index(self, issue: str) -> Optional[int]:
        """Extract slide index from issue string."""
        import re

        # Look for patterns like "Slide 3", "slide 3:", "Slide #3"
        patterns = [
            r"slide\s*#?\s*(\d+)",
            r"slide\s*(\d+)",
            r"\bslide:\s*(\d+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, issue.lower())
            if match:
                return int(match.group(1)) - 1  # Convert to 0-indexed

        return None


# =============================================================================
# Outline Refiner
# =============================================================================

class OutlineRefiner:
    """Refines presentation outlines based on evaluation feedback."""

    def __init__(self, llm_manager=None):
        """
        Initialize refiner.

        Args:
            llm_manager: Optional LLMManager for AI-powered refinement
        """
        self.llm_manager = llm_manager
        self.issue_analyzer = IssueAnalyzer()

    def refine_outline(
        self,
        outline: Dict[str, Any],
        evaluation_result: Any,
        max_actions: int = 5
    ) -> Tuple[Dict[str, Any], List[RefinementAction]]:
        """
        Refine an outline based on evaluation results.

        Args:
            outline: Original presentation outline
            evaluation_result: EvaluationResult from PresentationEvaluator
            max_actions: Maximum number of refinement actions to apply

        Returns:
            Tuple of (refined_outline, actions_applied)
        """
        # Create a deep copy to avoid mutating original
        refined = copy.deepcopy(outline)

        # Gather all issues
        all_issues = []
        if hasattr(evaluation_result, 'content') and evaluation_result.content.issues:
            all_issues.extend(evaluation_result.content.issues)
        if hasattr(evaluation_result, 'design') and evaluation_result.design.issues:
            all_issues.extend(evaluation_result.design.issues)
        if hasattr(evaluation_result, 'coherence') and evaluation_result.coherence.issues:
            all_issues.extend(evaluation_result.coherence.issues)

        # Analyze issues
        slide_count = self._count_slides(outline)
        actions = self.issue_analyzer.analyze_issues(all_issues, slide_count)

        # Apply top actions
        applied_actions = []
        for action in actions[:max_actions]:
            success = self._apply_action(refined, action)
            if success:
                applied_actions.append(action)

        # If we have an LLM, use it for smarter refinement
        if self.llm_manager and applied_actions:
            refined = self._llm_enhance_refinement(refined, applied_actions, all_issues)

        return refined, applied_actions

    def _count_slides(self, outline: Dict[str, Any]) -> int:
        """Count total slides in outline."""
        count = 0
        for section in outline.get("sections", []):
            count += len(section.get("slides", []))
        return count

    def _apply_action(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Apply a refinement action to the outline."""
        try:
            if action.refinement_type == RefinementType.ADD_CONTENT:
                return self._add_content(outline, action)

            elif action.refinement_type == RefinementType.REDUCE_CONTENT:
                return self._reduce_content(outline, action)

            elif action.refinement_type == RefinementType.ADD_VISUAL:
                return self._add_visual(outline, action)

            elif action.refinement_type == RefinementType.SPLIT_SLIDE:
                return self._split_slide(outline, action)

            elif action.refinement_type == RefinementType.FIX_STRUCTURE:
                return self._fix_structure(outline, action)

            elif action.refinement_type == RefinementType.CLARIFY_TITLE:
                return self._clarify_title(outline, action)

            return False

        except Exception as e:
            logger.warning(f"Failed to apply action {action}: {e}")
            return False

    def _get_slide(self, outline: Dict[str, Any], slide_idx: int) -> Optional[Dict[str, Any]]:
        """Get slide by global index."""
        current_idx = 0
        for section in outline.get("sections", []):
            for slide in section.get("slides", []):
                if current_idx == slide_idx:
                    return slide
                current_idx += 1
        return None

    def _add_content(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Add content to a sparse slide."""
        if action.target_slide_index < 0:
            return False

        slide = self._get_slide(outline, action.target_slide_index)
        if not slide:
            return False

        content = slide.get("content", {})
        bullets = content.get("bullets", [])

        # Add placeholder bullets if too few
        if len(bullets) < 3:
            while len(bullets) < 3:
                bullets.append(f"[Additional point {len(bullets) + 1}]")
            content["bullets"] = bullets
            slide["content"] = content
            return True

        return False

    def _reduce_content(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Reduce content on an overcrowded slide."""
        if action.target_slide_index < 0:
            return False

        slide = self._get_slide(outline, action.target_slide_index)
        if not slide:
            return False

        content = slide.get("content", {})
        bullets = content.get("bullets", [])

        # Trim to max 6 bullets
        if len(bullets) > 6:
            content["bullets"] = bullets[:6]
            slide["content"] = content
            return True

        return False

    def _add_visual(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Suggest adding a visual element."""
        if action.target_slide_index < 0:
            return False

        slide = self._get_slide(outline, action.target_slide_index)
        if not slide:
            return False

        content = slide.get("content", {})

        # Add a note suggesting visual
        if "visual_suggestion" not in content:
            content["visual_suggestion"] = "Consider adding a chart, diagram, or image"
            slide["content"] = content
            return True

        return False

    def _split_slide(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Mark a slide for splitting."""
        # Note: Actual splitting would require more complex logic
        # For now, just reduce content
        return self._reduce_content(outline, action)

    def _fix_structure(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Fix structural issues."""
        sections = outline.get("sections", [])

        # Ensure each section has a name
        for i, section in enumerate(sections):
            if not section.get("name"):
                section["name"] = f"Section {i + 1}"

        return True

    def _clarify_title(self, outline: Dict[str, Any], action: RefinementAction) -> bool:
        """Mark title for improvement."""
        if action.target_slide_index < 0:
            return False

        slide = self._get_slide(outline, action.target_slide_index)
        if not slide:
            return False

        content = slide.get("content", {})
        title = content.get("title", "")

        # Add improvement note
        if title and "[CLARIFY]" not in title:
            content["title_note"] = "Title may need clarification"
            slide["content"] = content
            return True

        return False

    def _llm_enhance_refinement(
        self,
        outline: Dict[str, Any],
        actions: List[RefinementAction],
        issues: List[str]
    ) -> Dict[str, Any]:
        """Use LLM to enhance refinement (if available)."""
        if not self.llm_manager or not self.llm_manager.is_available():
            return outline

        try:
            # Build prompt for LLM
            prompt = f"""You are refining a presentation outline based on quality feedback.

Issues identified:
{chr(10).join(f'- {issue}' for issue in issues[:10])}

Actions being taken:
{chr(10).join(f'- {action.description}' for action in actions)}

Current outline structure:
{json.dumps(outline.get('sections', [])[:3], indent=2)}

Provide specific content improvements in JSON format:
{{
  "improvements": [
    {{"slide_index": 0, "field": "title", "value": "improved title"}},
    {{"slide_index": 0, "field": "bullets", "value": ["bullet 1", "bullet 2"]}}
  ]
}}

Only output valid JSON, no other text."""

            response = self.llm_manager.generate(
                prompt,
                system_prompt="You are a presentation quality improvement assistant. Output only valid JSON.",
                max_tokens=1000,
                temperature=0.3
            )

            # Parse LLM response
            import re
            json_match = re.search(r'\{.*\}', response.content, re.DOTALL)
            if json_match:
                improvements = json.loads(json_match.group())
                outline = self._apply_llm_improvements(outline, improvements)

        except Exception as e:
            logger.debug(f"LLM enhancement failed: {e}")

        return outline

    def _apply_llm_improvements(
        self,
        outline: Dict[str, Any],
        improvements: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Apply LLM-suggested improvements to outline."""
        for imp in improvements.get("improvements", []):
            slide_idx = imp.get("slide_index", -1)
            field = imp.get("field")
            value = imp.get("value")

            if slide_idx >= 0 and field and value:
                slide = self._get_slide(outline, slide_idx)
                if slide:
                    content = slide.get("content", {})
                    content[field] = value
                    slide["content"] = content

        return outline


# =============================================================================
# Iterative Refiner (Main Class)
# =============================================================================

class IterativeRefiner:
    """
    Main iterative refinement controller.

    Runs a refinement loop that evaluates presentations and refines
    them until quality thresholds are met or max iterations reached.

    Usage:
        refiner = IterativeRefiner(evaluator, llm_manager)
        history = refiner.refine(
            outline,
            generator_fn,  # Function that generates PPTX from outline
            target_grade="B",
            max_iterations=3
        )
    """

    def __init__(
        self,
        evaluator=None,
        llm_manager=None,
        outline_refiner: OutlineRefiner = None
    ):
        """
        Initialize iterative refiner.

        Args:
            evaluator: PresentationEvaluator instance
            llm_manager: Optional LLMManager for AI-powered refinement
            outline_refiner: Optional custom OutlineRefiner
        """
        self.evaluator = evaluator
        self.llm_manager = llm_manager
        self.outline_refiner = outline_refiner or OutlineRefiner(llm_manager)

    def refine(
        self,
        initial_outline: Dict[str, Any],
        generator_fn: Callable[[Dict[str, Any]], str],
        target_grade: str = "B",
        target_score: float = None,
        max_iterations: int = 3,
        min_improvement: float = 0.05
    ) -> Tuple[Dict[str, Any], RefinementHistory]:
        """
        Run iterative refinement loop.

        Args:
            initial_outline: Starting presentation outline
            generator_fn: Function that takes outline and returns PPTX path
            target_grade: Target letter grade (A, B, C, D)
            target_score: Optional target score (0-1), overrides target_grade
            max_iterations: Maximum refinement iterations
            min_improvement: Minimum score improvement to continue

        Returns:
            Tuple of (final_outline, RefinementHistory)
        """
        # Convert grade to score threshold
        grade_thresholds = {"A": 0.9, "B": 0.8, "C": 0.7, "D": 0.6}
        threshold = target_score or grade_thresholds.get(target_grade, 0.7)

        # Generate initial presentation
        current_outline = copy.deepcopy(initial_outline)
        pptx_path = generator_fn(current_outline)

        # Initial evaluation
        if not self.evaluator:
            logger.warning("No evaluator provided, skipping refinement")
            return current_outline, RefinementHistory(
                initial_score=0.0,
                initial_grade="?",
                converged=False,
                reason="No evaluator"
            )

        eval_result = self.evaluator.evaluate(pptx_path)
        current_score = eval_result.overall_score
        current_grade = eval_result.grade

        history = RefinementHistory(
            initial_score=current_score,
            initial_grade=current_grade
        )

        logger.info(f"Starting refinement: {current_score:.2f} ({current_grade}) -> target {threshold:.2f}")

        # Check if already meets threshold
        if current_score >= threshold:
            history.converged = True
            history.reason = "Already meets target"
            history.final_score = current_score
            history.final_grade = current_grade
            return current_outline, history

        # Refinement loop
        for iteration in range(1, max_iterations + 1):
            logger.info(f"Refinement iteration {iteration}/{max_iterations}")

            # Refine outline based on evaluation
            refined_outline, actions = self.outline_refiner.refine_outline(
                current_outline,
                eval_result
            )

            if not actions:
                history.converged = False
                history.reason = "No refinement actions available"
                break

            # Generate new presentation
            try:
                pptx_path = generator_fn(refined_outline)
            except Exception as e:
                logger.error(f"Generation failed: {e}")
                history.reason = f"Generation error: {e}"
                break

            # Evaluate new presentation
            new_eval = self.evaluator.evaluate(pptx_path)
            new_score = new_eval.overall_score
            new_grade = new_eval.grade

            # Record iteration
            result = RefinementResult(
                iteration=iteration,
                original_score=current_score,
                new_score=new_score,
                grade=new_grade,
                actions_applied=actions,
                issues_remaining=new_eval.content.issues + new_eval.design.issues,
                improved=new_score > current_score
            )
            history.add_iteration(result)

            logger.info(f"  Score: {current_score:.2f} -> {new_score:.2f} ({new_grade})")

            # Check if target reached
            if new_score >= threshold:
                history.converged = True
                history.reason = f"Reached target grade {target_grade}"
                current_outline = refined_outline
                break

            # Check if improvement is sufficient
            improvement = new_score - current_score
            if improvement < min_improvement and iteration > 1:
                history.reason = f"Insufficient improvement ({improvement:.3f} < {min_improvement})"
                # Keep better version
                if new_score > current_score:
                    current_outline = refined_outline
                break

            # Update for next iteration
            if new_score > current_score:
                current_outline = refined_outline
                current_score = new_score
                eval_result = new_eval
            else:
                # Score decreased, stop refinement
                history.reason = "Score decreased, reverting"
                break

        if not history.converged and not history.reason:
            history.reason = "Max iterations reached"

        return current_outline, history


# =============================================================================
# Convenience Functions
# =============================================================================

def quick_refine(
    outline: Dict[str, Any],
    generator_fn: Callable[[Dict[str, Any]], str],
    evaluator=None,
    target_grade: str = "B"
) -> Tuple[Dict[str, Any], float]:
    """
    Quick refinement with defaults.

    Returns:
        Tuple of (refined_outline, final_score)
    """
    refiner = IterativeRefiner(evaluator=evaluator)
    refined, history = refiner.refine(
        outline,
        generator_fn,
        target_grade=target_grade,
        max_iterations=3
    )
    return refined, history.final_score
