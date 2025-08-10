# PowerPoint Inconsistency Detector

An AI-powered tool that automatically detects business logic inconsistencies, numerical conflicts, timeline mismatches, and data relationship errors in PowerPoint presentations using Google's Gemini AI.

**What it does**

The tool finds these types of problems in your presentations:
- **Numerical conflicts**: Revenue figures that don't match, percentages that don't add up
- **Timeline issues**: Dates that conflict, impossible project sequences  
- **Logic problems**: Contradictory statements about markets or competition
- **Data errors**: Parts that don't sum to the whole, inconsistent breakdowns

Other features:
- Analyzes slides in batches plus does a full cross-presentation check
- Can extract text from images using OCR
- Scores each finding with confidence level (0.0-1.0)
- Only shows high-confidence issues (70%+ by default)
- Removes duplicate findings automatically
- Shows exact quotes and slide numbers for each problem

**How it works**

The system processes PowerPoint files in two passes:
1. **Batch analysis** - checks slides in groups of 6 for local inconsistencies
2. **Cross-slide analysis** - looks across the entire presentation for global conflicts

```
PowerPoint File (.pptx)
    ↓
Text Extraction (python-pptx)
    ↓
OCR Processing (pytesseract) [Optional]
    ↓
Two-Pass Analysis:
├── Batch Analysis (6 slides per batch)
└── Cross-Slide Analysis (entire presentation)
    ↓
AI Processing (Google Gemini 2.0 Flash)
    ↓
Result Filtering & Deduplication
    ↓
Report Generation (Console + File)
```

**Project structure**

```
ppt-checker/
├── cli.py              # Command-line interface
├── ppt_analyzer.py     # Core analysis engine
├── models.py           # Data structures and report generation
├── presentation.pptx   # Sample presentation file
├── ppt-env/           # Virtual environment (not in repo)
└── README.md          # This file
``` 


**Installation and Setup**

You'll need:
- Python 3.10+ (I used Python 3.10.11)
- Windows PowerShell 
- Google Gemini API Key (get one at https://makersuite.google.com/app/apikey)
- Tesseract OCR (optional, for reading text from images)

*Step 1: Create virtual environment*
```powershell
# Check your Python version
python --version

# Create the virtual environment
python -m venv ppt-env

# Activate it
ppt-env\Scripts\activate
```

*Step 2: Install the packages*
```powershell
pip install python-pptx google-generativeai pillow pytesseract
```

*Step 3: Setup Tesseract OCR (optional)*
1. Download from UB-Mannheim GitHub releases
2. Install to `C:\Program Files\Tesseract-OCR`
3. Add to your PATH:
```powershell
$env:PATH += ";C:\Program Files\Tesseract-OCR"
```

**How to use it**

*Basic usage*
```powershell
python cli.py presentation.pptx --api-key YOUR_GEMINI_API_KEY
```

*More options*
```powershell
# Turn on OCR for images
python cli.py presentation.pptx --api-key YOUR_KEY --ocr

# Change batch size for big presentations
python cli.py presentation.pptx --api-key YOUR_KEY --batch-size 4

# Save the report to a file
python cli.py presentation.pptx --api-key YOUR_KEY --output report.txt

# All options together(recommended)
python cli.py presentation.pptx --api-key YOUR_KEY --ocr --batch-size 8 --output detailed_report.txt
```



##  Performance & Scalability

### Current Performance
| Presentation Size | Processing Time | Bottleneck |
|-------------------|-----------------|------------|
| 10-20 slides | ~30-45 seconds | None |
| 50-100 slides | ~2-3 minutes | API rate limits |
| 200+ slides | ~8-12 minutes | Sequential processing |
| 500+ slides | ~25-35 minutes | Memory + API limits |

### Batch Processing
- **Default Batch Size**: 6 slides per API call
- **Rate Limiting**: 1-second delay between batches
- **Memory Efficient**: Processes slides in chunks
- **Error Isolation**: Failed batches don't stop analysis

### API Usage
- **Model**: Google Gemini 2.0 Flash Experimental
- **Average API Calls**: 1 calls per batch + 1 global analysis
- **Cost Estimation**: ~$0.01-0.05 per presentation (varies by size)
- **Rate Limits**: Respects Google's default limits with built-in delays



**Current limitations**

*Processing issues*
- Batches are processed one at a time (no parallel processing)
- Single-threaded execution
- Memory usage grows with presentation size
- Basic OCR implementation

*AI model issues*
- Results can vary slightly between runs (that's normal for AI)
- Limited by Gemini's token limits
- Works best with English presentations
- Designed for business presentations

*Technical constraints*
- Currently set up for Windows
- Needs internet connection for API calls
- Depends on Google Gemini being available

**Scaling for large presentations**

*Problems with big decks (200+ slides)*
- Processing time grows linearly (~1 minute per 20 slides)
- Memory usage can hit 500MB for 500-slide presentations
- Gemini has daily/hourly API quotas
- Long processes can fail if network connection drops
- Single API failure can stop the whole analysis


### Performance targets

#### What we could achieve with improvements
| Presentation Size | Current Time | Target Time | How to get there |
|-------------------|--------------|-------------|------------------|
| 100 slides | 3-4 minutes | 45-60 seconds | Parallel processing |
| 500 slides | 25-30 minutes | 4-6 minutes | Async + checkpoints |
| 1000 slides | 60+ minutes | 8-12 minutes | Distributed processing |
| 5000 slides | Doesn't work | 30-45 minutes | Microservices |



*Implementation plan*

**Phase 1: Quick wins**

1. Add adaptive batching

2. Basic checkpointing

3. Retry logic with backoff


**Phase 2: Async processing (2-4 weeks)**
- Implement async-based parallel batch processing
- Add connection pooling for API calls
- Create adaptive rate limiting system



*What I've tested this on*
- Python 3.10.11 on Windows 11
- Tesseract v5.5.0 with English language pack
- Google Gemini 2.0 Flash Experimental model

Built for business people who need their presentations to be accurate.

Last updated: August 10, 2025
