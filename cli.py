#!/usr/bin/env python3

import argparse
import sys
from pathlib import Path

try:
    import pytesseract
except ImportError as e:
    print(f"Error: Missing required dependency - {e}")
    print("Please install dependencies: pip install python-pptx google-generativeai pillow pytesseract")
    sys.exit(1)

from ppt_analyzer import PPTAnalyzer
from models import generate_report

def main():
    parser = argparse.ArgumentParser(
        description="Detect inconsistencies in PowerPoint presentations using AI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python cli.py presentation.pptx --api-key YOUR_GEMINI_KEY
  python cli.py slides.pptx --api-key YOUR_KEY --ocr
  python cli.py deck.pptx --api-key YOUR_KEY --batch-size 4

Inconsistency Types Detected:
  • Numerical conflicts (revenue figures, percentages that don't add up)
  • Timeline mismatches (date conflicts, impossible sequences)  
  • Logical contradictions (mutually exclusive statements)
  • Data relationship errors (parts don't sum to whole)
        """
    )
    
    parser.add_argument("pptx_file", help="Path to PowerPoint (.pptx) file")
    parser.add_argument("--api-key", required=True, help="Google Gemini API key")
    parser.add_argument("--ocr", action="store_true", help="Enable OCR for image text extraction")
    parser.add_argument("--batch-size", type=int, default=6, help="Number of slides per API batch (default: 6)")
    parser.add_argument("--output", help="Save report to file (optional)")
    
    args = parser.parse_args()
    
    # Check if file exists
    pptx_path = Path(args.pptx_file)
    if not pptx_path.exists():
        print(f"Error: File '{args.pptx_file}' not found.")
        sys.exit(1)
    
    if not pptx_path.suffix.lower() == '.pptx':
        print("Error: Please provide a .pptx file.")
        sys.exit(1)
    
    # Check if OCR is available
    if args.ocr:
        try:
            pytesseract.get_tesseract_version()
        except:
            print("Warning: Tesseract OCR not found. Install tesseract-ocr to enable OCR functionality.")
            print("Continuing without OCR...")
            args.ocr = False
    
    try:
        print(" Starting PowerPoint inconsistency analysis...")
        print(f" File: {pptx_path.name}")
        
        # Setup the analyzer
        analyzer = PPTAnalyzer(args.api_key, args.batch_size)
        
        # Get slide content
        print(" Extracting slide content...")
        slides = analyzer.extract_slide_content(str(pptx_path), args.ocr)
        print(f" Extracted content from {len(slides)} slides")
        
        # Run the analysis
        issues = analyzer.analyze_inconsistencies(slides)
        
        # Show the results
        report = generate_report(issues, len(slides))
        print("\n" + report)
        
        # Save to file if they want it
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(report)
            print(f"\n Report saved to: {args.output}")
        
    except KeyboardInterrupt:
        print("\n Analysis interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f" Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
