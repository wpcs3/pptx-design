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
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
