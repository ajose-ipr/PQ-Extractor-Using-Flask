# Harmonic Analysis Report Generator

A Flask web application that converts machine-generated Power Quality (PQ) reports into standardized Word documents compliant with IEEE 519-2022 standards for regulatory submission.

## Purpose

Transform raw harmonic analysis data into professional reports that meet regulatory authority requirements, featuring:
- Standardized table formatting per IEEE 519-2022
- Automatic violation detection and highlighting
- Professional Word document export

## Key Features

### Data Processing
- Processes voltage and current harmonic measurements
- Maps raw data (V1N, V2N, V3N, I1, I2, I3) from PQ reports to standard R, Y, B phase notation
- Displays limit exceeded values in a separate section

### Export Options
- Professional Word document (.doc) export
- Excel for current file or all uploaded files' downloads
- CSV for harmonic violation reports

## Installation

### Prerequisites
- Python 3.7+
- Flask
- Pandas
- Others (see requirements.txt)

### Setup
Run `exe.bat` file to open application

## Usage

### 1. Upload Data
- Upload PQ report files through the web interface

### 2. Review Analysis
- View system metadata (component, block, bay/feeder info)
- Check violation summary and statistics
- Review formatted harmonic tables

## File Structure

```
harmonic-analysis-app/
├── app.py                 # Flask application main file
├── templates/
│   ├── base.html         # Base template
│   ├── select_file.html  # File selection page
│   └── results.html      # Results display (integrated template)
├── static/
│   ├── css/             # Stylesheets
│   └── js/              # JavaScript files
├── uploads/             # Uploaded files directory
├── utils/           # Processed data directory
└── requirements.txt     # Python dependencies
```

## Technical Details

### Table Configurations
| Table | Time Interval | Percentile | Measurement Type |
|-------|---------------|------------|------------------|
| Voltage Daily | 3-second | 99th | Voltage |
| Voltage Full | 10-minute | 95th | Voltage |
| Current Daily | 3-second | 99th | Current |
| Current Full 99 | 10-minute | 99th | Current |
| Current Full 95 | 10-minute | 95th | Current |


## Browser Compatibility

- Modern browsers (prefer Chrome)
- JavaScript-enabled functionality required

## Troubleshooting

### Common Issues
1. **File Upload Errors**: Check file format and size limits
2. **Missing Data**: Ensure input files contain required columns
3. **Violation Detection**: Verify regulation limit data is present
4. **Export Problems**: Check browser download permissions

