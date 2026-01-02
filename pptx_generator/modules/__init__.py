"""
Core modules for the PowerPoint generator.
"""

# Import individual modules - these can be imported independently
# to avoid circular imports

__all__ = [
    "TemplateAnalyzer",
    "SlideLibrary",
    "SlideRenderer",
    "OutlineGenerator",
    "ResearchAgent",
    "PresentationOrchestrator",
    "ComponentLibrary",
    "PresentationReviewer",
    "review_presentation",
    "OutputOrganizer",
    "organize_output_directory",
]


def __getattr__(name):
    """Lazy import of modules to avoid circular dependencies."""
    if name == "TemplateAnalyzer":
        from .template_analyzer import TemplateAnalyzer
        return TemplateAnalyzer
    elif name == "SlideLibrary":
        from .slide_library import SlideLibrary
        return SlideLibrary
    elif name == "SlideRenderer":
        from .slide_renderer import SlideRenderer
        return SlideRenderer
    elif name == "OutlineGenerator":
        from .outline_generator import OutlineGenerator
        return OutlineGenerator
    elif name == "ResearchAgent":
        from .research_agent import ResearchAgent
        return ResearchAgent
    elif name == "PresentationOrchestrator":
        from .orchestrator import PresentationOrchestrator
        return PresentationOrchestrator
    elif name == "ComponentLibrary":
        from .component_library import ComponentLibrary
        return ComponentLibrary
    elif name == "PresentationReviewer":
        from .presentation_review import PresentationReviewer
        return PresentationReviewer
    elif name == "review_presentation":
        from .presentation_review import review_presentation
        return review_presentation
    elif name == "OutputOrganizer":
        from .output_organizer import OutputOrganizer
        return OutputOrganizer
    elif name == "organize_output_directory":
        from .output_organizer import organize_output_directory
        return organize_output_directory
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
