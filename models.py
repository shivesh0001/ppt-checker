#!/usr/bin/env python3

from dataclasses import dataclass
from typing import List

@dataclass
class Inconsistency:
    type: str
    confidence: float
    slides: List[int]
    issue: str
    evidence: List[str]

@dataclass
class SlideContent:
    slide_number: int
    text: str
    ocr_text: str = ""

def generate_report(inconsistencies: List[Inconsistency], total_slides: int) -> str:
    # Create a nice  report for the console
    lines = [
        "=" * 50,
        "=== POWERPOINT INCONSISTENCY REPORT ===", 
        "=" * 50,
        f"Slides Analyzed: {total_slides}",
        f"Issues Found: {len(inconsistencies)}",
        ""
    ]
    
    if not inconsistencies:
        lines.append(" No significant inconsistencies detected!")
        lines.append("")
        lines.append("Note: Only inconsistencies with 70%+ confidence are shown.")
    else:
        for i, issue in enumerate(inconsistencies, 1):
            lines.extend([
                f"{i}. {issue.type} (Confidence: {issue.confidence:.1f})",
                f"   Slides: {', '.join(map(str, issue.slides))}",
                f"   Issue: {issue.issue}",
                "   Evidence:"
            ])
            
            for evidence in issue.evidence:
                lines.append(f"   - {evidence}")
            
            lines.append("")
    
    lines.extend([
        "=" * 50,
        "Analysis complete. Review findings carefully for business impact.",
        "=" * 50
    ])
    
    return "\n".join(lines)
