"""Generate Light Industrial Investment Thesis Presentation."""

import json
from pathlib import Path

# Define the presentation outline based on the markdown thesis
outline = {
    "presentation_type": "investment_pitch",
    "title": "Light Industrial Investment Thesis",
    "subtitle": "US Portfolio Acquisition Strategy | Institutional Joint Venture",
    "template": "consulting_toolkit",
    "sections": [
        {
            "name": "Executive Summary",
            "slides": [
                {
                    "slide_type": "title_slide",
                    "content": {
                        "title": "Light Industrial Investment Thesis",
                        "subtitle": "US Portfolio Acquisition Strategy | Institutional Joint Venture"
                    }
                },
                {
                    "slide_type": "key_metrics",
                    "content": {
                        "title": "Investment Opportunity Overview",
                        "metrics": [
                            {"label": "JV Structure", "value": "49/49/2"},
                            {"label": "Target IRR", "value": "10-15%"},
                            {"label": "Small-Bay Vacancy", "value": "3.4%"},
                            {"label": "10-Yr Returns", "value": "12.4%"}
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Why Light Industrial, Why Now",
                        "bullets": [
                            "12.4% ten-year annualized returns—highest among all property types",
                            "Small-bay vacancy: 3.4% nationally vs 7.1% overall industrial",
                            "Construction pipeline at decade lows (0.3% of stock)",
                            "Cap rates repriced 130 bps from 2022 trough—attractive entry",
                            "Structural demand from e-commerce, nearshoring, obsolescence"
                        ]
                    }
                }
            ]
        },
        {
            "name": "Market Fundamentals",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "Market Fundamentals"}
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "US Industrial Market Overview",
                        "headers": ["Metric", "Value"],
                        "data": [
                            ["Total Inventory", "20 billion SF"],
                            ["National Vacancy", "7.1% (Q3 2025)"],
                            ["Small-Bay Vacancy", "3.4%"],
                            ["Avg. Asking Rent", "$10.10/SF"],
                            ["Rent Growth (5-yr)", "+60%"],
                            ["Under Construction", "382.7M SF"]
                        ]
                    }
                },
                {
                    "slide_type": "two_column",
                    "content": {
                        "title": "Light Industrial vs. Bulk Logistics",
                        "left_column": {
                            "header": "Light Industrial",
                            "bullets": [
                                "Size: <50,000 SF",
                                "Clear Heights: 20-28 ft",
                                "Vacancy: 3.4%",
                                "Pipeline: 0.3% of stock",
                                "Tenants: Diversified SMBs"
                            ]
                        },
                        "right_column": {
                            "header": "Bulk Logistics",
                            "bullets": [
                                "Size: 200,000+ SF",
                                "Clear Heights: 36-40+ ft",
                                "Vacancy: 9-10%",
                                "Pipeline: 2.0%+ of stock",
                                "Tenants: National corps"
                            ]
                        }
                    }
                },
                {
                    "slide_type": "data_chart",
                    "content": {
                        "title": "Supply Pipeline Contraction",
                        "chart_data": {
                            "type": "column",
                            "categories": ["2021", "2022", "2023", "2024", "2025"],
                            "series": [{"name": "Under Construction (M SF)", "values": [550, 1000, 850, 550, 383]}]
                        },
                        "narrative": "Construction down 62% from peak"
                    }
                },
                {
                    "slide_type": "data_chart",
                    "content": {
                        "title": "Cap Rate Evolution",
                        "chart_data": {
                            "type": "line",
                            "categories": ["Q2 2022", "Q4 2022", "Q2 2023", "Q4 2023", "Q2 2024", "Q4 2024"],
                            "series": [{"name": "Cap Rate (%)", "values": [5.22, 5.75, 6.10, 6.40, 6.51, 6.29]}]
                        },
                        "narrative": "Cap rates expanded 130 bps from trough"
                    }
                }
            ]
        },
        {
            "name": "Structural Demand Drivers",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "Structural Demand Drivers"}
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Three Converging Megatrends",
                        "bullets": [
                            "E-Commerce: 23.2% of retail, targeting 32% by 2035",
                            "Nearshoring: 300+ mfg announcements, $400B+ investment",
                            "Obsolescence: Pre-2000 buildings posting negative absorption"
                        ]
                    }
                },
                {
                    "slide_type": "data_chart",
                    "content": {
                        "title": "Manufacturing Renaissance",
                        "chart_data": {
                            "type": "column",
                            "categories": ["2020", "2021", "2022", "2023", "2024"],
                            "series": [{"name": "Annual Spending ($B)", "values": [80, 95, 128, 195, 237]}]
                        },
                        "narrative": "86% increase in manufacturing construction spending"
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Interest Rate Sensitivity",
                        "headers": ["Property Type", "Cap Rate Δ per 100 bps"],
                        "data": [
                            ["Industrial", "41 bps"],
                            ["Office", "70 bps"],
                            ["Multifamily", "75 bps"],
                            ["Retail", "78 bps"]
                        ]
                    }
                }
            ]
        },
        {
            "name": "Target Market Analysis",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "Target Market Analysis"}
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Market Selection Framework",
                        "bullets": [
                            "Tier 1 (Core Growth): Nashville, Tampa, Raleigh-Durham",
                            "Tier 2 (Scale Markets): Dallas-Fort Worth, Atlanta, Phoenix",
                            "Tier 3 (Yield): San Antonio, Jacksonville, Salt Lake City"
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Tier 1: Core Growth Markets",
                        "headers": ["Market", "Vacancy", "Rent Growth", "Cap Rate"],
                        "data": [
                            ["Nashville", "4.1%", "+8% YoY", "5.5-6.5%"],
                            ["Tampa", "3.0-3.2%", "+69% (5-yr)", "5.5-6.5%"],
                            ["Raleigh-Durham", "6.3%", "+39%", "5.5-6.5%"]
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Nashville Deep Dive",
                        "bullets": [
                            "Overall vacancy: 4.1-4.4% (lowest among targets)",
                            "2024 investment volume: $1.4B (+37% YoY)",
                            "East submarket: 1.5M SF absorption",
                            "79% of relocations involve industrial users"
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Tampa Deep Dive",
                        "bullets": [
                            "Small-bay vacancy: 3.0-3.2% (constrained)",
                            "Pre-leasing rate: 75% (highest in Florida)",
                            "Construction: Down 70% from 2022 peak",
                            "5-year rent growth: 69.1%"
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Tier 2: Scale Markets",
                        "headers": ["Market", "Vacancy", "Key Metric"],
                        "data": [
                            ["Dallas-Fort Worth", "9.7%", "55 qtrs positive absorption"],
                            ["Atlanta", "8.2-8.6%", "$11B Amazon investment"],
                            ["Phoenix", "10.6%", "#1 mfg job growth"]
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Tier 3: Yield Opportunities",
                        "headers": ["Market", "Cap Rate", "Vacancy"],
                        "data": [
                            ["San Antonio", "7.0-8.5%", "2.7% (mfg)"],
                            ["Jacksonville", "6.5-7.5%", "Moderate"],
                            ["Salt Lake City", "6.0-7.0%", "2.5% (small-bay)"]
                        ]
                    }
                }
            ]
        },
        {
            "name": "Investment Strategy",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "Investment Strategy"}
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Portfolio Construction",
                        "bullets": [
                            "Target: 2-4M SF across 15-25 properties",
                            "Focus: Multi-tenant small-bay (<50,000 SF)",
                            "Geographic: 7+ markets, max 20% concentration",
                            "Vintage: 2000+ construction"
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Return Expectations by Strategy",
                        "headers": ["Strategy", "Target IRR", "Multiple", "Leverage"],
                        "data": [
                            ["Core", "7-10%", "1.3-1.6x", "40-45%"],
                            ["Core-Plus", "8-12%", "1.4-1.7x", "45-60%"],
                            ["Value-Add", "13-20%", "1.7-2.0x", "60-80%"]
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Value Creation Levers",
                        "bullets": [
                            "Mark-to-Market: 20-40% rent upside on rollover",
                            "Occupancy: Target 95%+ stabilized",
                            "OpEx Efficiency: Scale and technology",
                            "ESG: Solar/EV for tenant attraction"
                        ]
                    }
                },
                {
                    "slide_type": "data_chart",
                    "content": {
                        "title": "Historical Performance",
                        "chart_data": {
                            "type": "bar",
                            "categories": ["Industrial", "Multifamily", "Retail", "Office"],
                            "series": [{"name": "10-Yr Return (%)", "values": [12.4, 9.8, 8.2, 6.5]}]
                        },
                        "narrative": "Industrial outperformed all property types"
                    }
                }
            ]
        },
        {
            "name": "Risk Factors",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "Risk Factors & Mitigants"}
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Key Risks and Mitigation",
                        "headers": ["Risk", "Mitigation"],
                        "data": [
                            ["Supply overhang", "Avoid >5% pipeline markets"],
                            ["Tenant credit", "15+ tenant diversification"],
                            ["Rate volatility", "Lowest sensitivity (41 bps)"],
                            ["Recession", "E-commerce counter-cyclical"],
                            ["Geographic risk", "7+ market diversification"]
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Defensive Characteristics",
                        "bullets": [
                            "Triple-net leases pass OpEx to tenants",
                            "3-5 year terms enable mark-to-market",
                            "Diversified SMB tenant mix",
                            "Essential distribution infrastructure"
                        ]
                    }
                }
            ]
        },
        {
            "name": "ESG & Sustainability",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "ESG & Sustainability"}
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "ESG Competitive Advantages",
                        "bullets": [
                            "Lowest GHG intensity: 21.1 kg/m²",
                            "Large roof areas ideal for solar",
                            "GRESB: 2,200+ companies, $7T AUM",
                            "170 investors with $51T+ use GRESB data"
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Green Building Economics",
                        "headers": ["Certification", "Rent Premium", "Sales Premium"],
                        "data": [
                            ["LEED Certified", "6-31%", "7.6-21%"],
                            ["ENERGY STAR", "3-7%", "5-10%"],
                            ["Solar-Equipped", "2-5%", "3-8%"]
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "ESG Integration Strategy",
                        "bullets": [
                            "Year 1: GRESB baseline assessment",
                            "Year 2-3: LED and water efficiency",
                            "Year 3-5: Solar on 50%+ roof area",
                            "Target: Top-quartile GRESB by Year 5"
                        ]
                    }
                }
            ]
        },
        {
            "name": "JV Structure",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "JV Structure & Governance"}
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Capital Structure",
                        "headers": ["Partner", "Commitment", "Role"],
                        "data": [
                            ["US Public Pension", "49%", "Limited Partner"],
                            ["International SWF", "49%", "Limited Partner"],
                            ["GP Sponsor", "2%", "General Partner"]
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "GP Responsibilities",
                        "bullets": [
                            "Sourcing and underwriting acquisitions",
                            "Asset management and leasing oversight",
                            "Capital improvement execution",
                            "Investor reporting and GRESB compliance"
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Fee Structure",
                        "headers": ["Fee Type", "Rate", "Basis"],
                        "data": [
                            ["Acquisition Fee", "1.0%", "Gross purchase price"],
                            ["Asset Management", "0.75%", "Gross asset value"],
                            ["Disposition Fee", "0.5%", "Gross sales price"],
                            ["Promote", "20%", "Above 8% pref return"]
                        ]
                    }
                }
            ]
        },
        {
            "name": "Conclusion",
            "slides": [
                {
                    "slide_type": "section_divider",
                    "content": {"title": "Conclusion"}
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Investment Highlights",
                        "bullets": [
                            "Timing: Attractive entry after 130 bps expansion",
                            "Fundamentals: 3.4% vacancy, decade-low construction",
                            "Demand: E-commerce, nearshoring, obsolescence",
                            "Returns: 10-15% target exceeds 6.91% pension hurdle"
                        ]
                    }
                },
                {
                    "slide_type": "table_slide",
                    "content": {
                        "title": "Target Portfolio Summary",
                        "headers": ["Attribute", "Target"],
                        "data": [
                            ["Portfolio Size", "2-4 million SF"],
                            ["Property Count", "15-25 assets"],
                            ["Markets", "7+ (Sunbelt emphasis)"],
                            ["Net IRR", "10-15%"],
                            ["Hold Period", "5-7 years"]
                        ]
                    }
                },
                {
                    "slide_type": "title_content",
                    "content": {
                        "title": "Next Steps",
                        "bullets": [
                            "Phase 1: Finalize JV documentation",
                            "Phase 2: Activate acquisition pipeline",
                            "Phase 3: Execute 3-5 seed investments (12 mo)",
                            "Phase 4: Scale through programmatic acquisitions"
                        ]
                    }
                }
            ]
        }
    ]
}


def main():
    """Generate the presentation."""
    # Save the outline as JSON
    outline_path = Path("pptx_generator/output/light_industrial_outline.json")
    outline_path.parent.mkdir(parents=True, exist_ok=True)

    with open(outline_path, "w") as f:
        json.dump(outline, f, indent=2)
    print(f"Saved outline to: {outline_path}")

    # Import and run the orchestrator
    from pptx_generator.modules.orchestrator import PresentationOrchestrator, GenerationOptions

    # Set up paths
    config_dir = Path("pptx_generator/config")
    templates_dir = Path("pptx_templates")
    output_dir = Path("pptx_generator/output")

    # Create orchestrator with options
    options = GenerationOptions(
        auto_layout=True,
        auto_section_headers=False,  # We already have section dividers
        evaluate_after=True,
        use_slide_pool=False  # Disable to use simpler rendering
    )

    orchestrator = PresentationOrchestrator(
        config_dir=str(config_dir),
        templates_dir=str(templates_dir),
        output_dir=str(output_dir),
        options=options
    )

    # Generate the presentation
    output_path = output_dir / "Light_Industrial_Thesis_20251229.pptx"

    result = orchestrator.generate_pptx_with_evaluation(
        outline=outline,
        context={"request": "Light Industrial Investment Thesis"}
    )

    # Save the presentation
    result.presentation.save(str(output_path))

    print(f"\nGenerated presentation: {output_path}")
    print(f"Slide count: {len(result.presentation.slides)}")

    if result.evaluation:
        print(f"Quality Grade: {result.evaluation.grade}")
        print(f"Overall Score: {result.evaluation.overall_score:.1f}")

    return output_path


if __name__ == "__main__":
    main()
