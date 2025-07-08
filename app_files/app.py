from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.utils import secure_filename
import os
import logging
from io import BytesIO
import pandas as pd
from config import Config
from utils.processing import (
    extract_metadata, extract_tables_from_pdf, 
    process_table_data, split_table, analyze_failures,
    create_excel_download, create_bulk_excel_download,
    extract_thd_tdd_summary_tables_from_pdf,
    process_thd_tdd_summary_data,
    create_enhanced_excel_download
)

# Initialize Flask app
app = Flask(__name__)
app.config.from_object(Config)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
handler = logging.FileHandler('app.log')
handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
logger.info(f"Ensured upload directory exists at: {app.config['UPLOAD_FOLDER']}")

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET', 'POST'])
def index():
    logger.info(f"Accessed index route with method: {request.method}")
    if request.method == 'POST':
        if 'files' not in request.files:
            logger.warning("No files part in the request")
            flash('No files selected', 'warning')
            return redirect(request.url)
        
        files = request.files.getlist('files')
        if not files or all(file.filename == '' for file in files):
            logger.warning("No files selected for upload")
            flash('No files selected', 'warning')
            return redirect(request.url)
        
        filenames = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                try:
                    file.save(filepath)
                    filenames.append(filename)
                    logger.info(f"Successfully saved file: {filename} at {filepath}")
                    logger.debug(f"File details - Size: {os.path.getsize(filepath)} bytes, "
                               f"Content-Type: {file.content_type}")
                except Exception as e:
                    logger.error(f"Failed to save file {filename}: {str(e)}", exc_info=True)
                    flash(f'Error saving file {filename}', 'danger')
                    continue
            else:
                logger.warning(f"Invalid file type: {file.filename}")
                flash(f'Invalid file type: {file.filename}. Only PDF files are allowed.', 'warning')
        
        if not filenames:
            logger.error("No valid PDF files were uploaded")
            flash('No valid PDF files uploaded', 'danger')
            return redirect(request.url)
        
        session['uploaded_files'] = filenames
        logger.info(f"Stored {len(filenames)} files in session: {filenames}")
        return redirect(url_for('select_file'))
    
    return render_template('index.html')

@app.route('/select', methods=['GET', 'POST'])
def select_file():
    logger.info("Accessed select file route")
    if 'uploaded_files' not in session or not session['uploaded_files']:
        logger.warning("No uploaded files in session, redirecting to index")
        flash('No files uploaded. Please upload files first.', 'warning')
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        selected_file = request.form.get('selected_file')
        if not selected_file or selected_file not in session['uploaded_files']:
            logger.error(f"Invalid file selection: {selected_file}")
            flash('Invalid file selection', 'danger')
            return redirect(url_for('select_file'))
        
        session['selected_file'] = selected_file
        logger.info(f"Selected file: {selected_file}")
        return redirect(url_for('process_file'))
    
    logger.debug(f"Showing selection page with files: {session['uploaded_files']}")
    return render_template('select.html', files=session['uploaded_files'])

@app.route('/process')
def process_file():
    logger.info("Accessed process file route")
    if 'selected_file' not in session:
        logger.warning("No file selected, redirecting to select")
        flash('No file selected. Please select a file first.', 'warning')
        return redirect(url_for('select_file'))
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], session['selected_file'])
    if not os.path.exists(filepath):
        logger.error(f"File not found at path: {filepath}")
        flash('Selected file not found', 'danger')
        return redirect(url_for('select_file'))
    
    logger.info(f"Processing file: {session['selected_file']}")
    
    try:
        # Extract metadata
        component_text, block, feeder, company, report_info = extract_metadata(
            filepath, session['selected_file'])
        logger.info(f"Extracted metadata - Component: {component_text}, "
                   f"Block: {block}, Feeder: {feeder}, Company: {company}")
        
        # Extract individual harmonic tables with page tracking
        tables = extract_tables_from_pdf(filepath)
        logger.info(f"Extracted {len(tables)} table types from PDF")
        
        # Log table extraction summary
        for table_name, table_data in tables.items():
            if table_data:
                # Check for page numbers in the data
                page_numbers = []
                for row in table_data:
                    if len(row) >= 10:  # Should have page number as last column
                        page_numbers.append(row[-1])
                
                unique_pages = set(page_numbers) if page_numbers else set()
                logger.info(f"Table '{table_name}': {len(table_data)} rows from pages {sorted(unique_pages)}")
            else:
                logger.warning(f"Table '{table_name}': No data extracted")
        
        # Extract THD/TDD summary tables
        thd_tdd_raw_data = extract_thd_tdd_summary_tables_from_pdf(filepath)
        thd_tdd_tables = process_thd_tdd_summary_data(thd_tdd_raw_data)
        logger.info(f"Extracted {len(thd_tdd_tables)} THD/TDD summary tables")
        
        # Process individual harmonic tables and split by odd/even harmonics
        processed_tables = {}
        for table_name, table_data in tables.items():
            if table_data:
                df = process_table_data(table_data, table_name)
                if not df.empty:
                    processed_tables[table_name] = split_table(df)
                    logger.debug(f"Processed {table_name} with {len(df)} rows")
                    
                    # Log page distribution for this table
                    if 'Page_Number' in df.columns:
                        page_counts = df['Page_Number'].value_counts().to_dict()
                        logger.info(f"Page distribution for {table_name}: {page_counts}")
                else:
                    logger.warning(f"No valid data found in {table_name}")
        
        # Analyze violations with comprehensive page tracking
        violations = []
        total_violations_found = 0
        
        for table_name, table_data in tables.items():
            if table_data:
                df = process_table_data(table_data, table_name)
                if not df.empty:
                    # Pass table_name to analyze_failures for proper page tracking
                    table_violations = analyze_failures(df, table_name)
                    if not table_violations.empty:
                        violations.append(table_violations)
                        table_violation_count = len(table_violations)
                        total_violations_found += table_violation_count
                        
                        # Log detailed violation information with page tracking
                        if 'Page' in table_violations.columns:
                            page_counts = table_violations['Page'].value_counts().to_dict()
                            logger.info(f"Found {table_violation_count} violations in {table_name} "
                                      f"across pages: {sorted(page_counts.keys())}")
                            logger.info(f"Page distribution for violations in {table_name}: {page_counts}")
                        else:
                            logger.warning(f"No page information in violations for {table_name}")
        
        combined_violations = pd.concat(violations) if violations else pd.DataFrame()
        
        # Log comprehensive violation summary
        if not combined_violations.empty:
            total_violations = len(combined_violations)
            unique_harmonics = combined_violations['Harmonic'].nunique()
            
            if 'Page' in combined_violations.columns:
                unique_pages = combined_violations['Page'].nunique()
                page_distribution = combined_violations['Page'].value_counts().to_dict()
                max_exceedance = combined_violations['Exceedance (%)'].max()
                
                logger.info(f"=== VIOLATION SUMMARY ===")
                logger.info(f"Total violations: {total_violations}")
                logger.info(f"Unique harmonics affected: {unique_harmonics}")
                logger.info(f"Pages with violations: {unique_pages}")
                logger.info(f"Max exceedance: {max_exceedance}%")
                logger.info(f"Violations by page: {page_distribution}")
                logger.info(f"========================")
            else:
                logger.warning("Combined violations DataFrame missing 'Page' column")
        else:
            logger.info("No violations found in any table")
        
        return render_template('results.html',
                            filename=session['selected_file'],
                            metadata={
                                'component': component_text,
                                'block': block,
                                'feeder': feeder,
                                'company': company,
                                'report_info': report_info
                            },
                            tables=processed_tables,
                            thd_tdd_tables=thd_tdd_tables,
                            violations=combined_violations,
                            violations_exist=not combined_violations.empty)
    
    except Exception as e:
        logger.error(f"Error processing file {session['selected_file']}: {str(e)}", exc_info=True)
        flash(f'Error processing file: {str(e)}', 'danger')
        return redirect(url_for('select_file'))

@app.route('/download/<filename>')
def download_file(filename):
    logger.info(f"Download request for file: {filename}")
    if 'uploaded_files' not in session or filename not in session['uploaded_files']:
        logger.error(f"Unauthorized download attempt: {filename}")
        flash('File not available for download', 'danger')
        return redirect(url_for('index'))
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        logger.error(f"File not found for download: {filepath}")
        flash('File not found', 'danger')
        return redirect(url_for('index'))
    
    try:
        # Extract individual harmonic tables
        tables = extract_tables_from_pdf(filepath)
        
        # Extract THD/TDD summary tables
        thd_tdd_raw_data = extract_thd_tdd_summary_tables_from_pdf(filepath)
        thd_tdd_tables = process_thd_tdd_summary_data(thd_tdd_raw_data)
        
        if not any(tables.values()) and not any(df.empty == False for df in thd_tdd_tables.values()):
            logger.warning(f"No tables extracted from {filename}")
            flash('No data available for download', 'warning')
            return redirect(url_for('index'))
        
        # Create enhanced Excel with both harmonic and THD/TDD tables
        excel_data = create_enhanced_excel_download(tables, thd_tdd_tables, filename)
        logger.info(f"Successfully created enhanced Excel download for {filename}")
        
        return send_file(
            BytesIO(excel_data),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{filename.replace('.pdf', '')}_complete_analysis_v3.xlsx"
        )
    
    except Exception as e:
        logger.error(f"Error generating download for {filename}: {str(e)}", exc_info=True)
        flash(f'Error generating download: {str(e)}', 'danger')
        return redirect(url_for('index'))

@app.route('/download_violations/<filename>')
def download_violations(filename):
    logger.info(f"Violations download request for file: {filename}")
    if 'uploaded_files' not in session or filename not in session['uploaded_files']:
        logger.error(f"Unauthorized violations download attempt: {filename}")
        flash('File not available for download', 'danger')
        return redirect(url_for('index'))
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        logger.error(f"File not found for violations download: {filepath}")
        flash('File not found', 'danger')
        return redirect(url_for('index'))
    
    try:
        tables = extract_tables_from_pdf(filepath)
        violations = []
        
        for table_name, table_data in tables.items():
            if table_data:
                df = process_table_data(table_data, table_name)
                if not df.empty:
                    # Ensure table_name is passed for proper page tracking
                    table_violations = analyze_failures(df, table_name)
                    if not table_violations.empty:
                        violations.append(table_violations)
                        logger.info(f"Including {len(table_violations)} violations from {table_name} in CSV download")
        
        if not violations:
            flash('No violations found in this file', 'info')
            return redirect(url_for('process_file'))
        
        combined_violations = pd.concat(violations)
        
        # Reorder columns to put Page before Table for better readability
        if 'Page' in combined_violations.columns:
            cols = combined_violations.columns.tolist()
            if 'Table' in cols:
                cols.remove('Page')
                cols.remove('Table')
                cols.extend(['Page', 'Table'])
                combined_violations = combined_violations[cols]
            
            # Log final CSV content summary
            page_distribution = combined_violations['Page'].value_counts().to_dict()
            logger.info(f"CSV download contains {len(combined_violations)} violations from pages: {page_distribution}")
        else:
            logger.warning("No page information available in violations CSV")
        
        csv_data = combined_violations.to_csv(index=False)
        
        return send_file(
            BytesIO(csv_data.encode('utf-8')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f"{filename.replace('.pdf', '')}_violations_with_pages_v3.csv"
        )
    
    except Exception as e:
        logger.error(f"Error generating violations download for {filename}: {str(e)}", exc_info=True)
        flash(f'Error generating violations download: {str(e)}', 'danger')
        return redirect(url_for('process_file'))

@app.route('/bulk_download')
def bulk_download():
    logger.info("Bulk download requested")
    if 'uploaded_files' not in session or not session['uploaded_files']:
        logger.warning("No files available for bulk download")
        flash('No files available for bulk download', 'warning')
        return redirect(url_for('index'))
    
    all_files_data = {}
    for filename in session['uploaded_files']:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            tables = extract_tables_from_pdf(filepath)
            if any(tables.values()):
                all_files_data[filename] = tables
                logger.debug(f"Processed {filename} for bulk download")
                
                # Log page information for bulk download
                for table_name, table_data in tables.items():
                    if table_data:
                        page_numbers = [row[-1] for row in table_data if len(row) >= 10]
                        unique_pages = set(page_numbers) if page_numbers else set()
                        logger.debug(f"Bulk: {filename} - {table_name}: {len(table_data)} rows from pages {sorted(unique_pages)}")
        except Exception as e:
            logger.error(f"Error processing {filename} for bulk download: {str(e)}")
            continue
    
    if not all_files_data:
        logger.error("No valid data found for bulk download")
        flash('No valid data available for bulk download', 'warning')
        return redirect(url_for('index'))
    
    try:
        excel_data = create_bulk_excel_download(all_files_data)
        logger.info(f"Created bulk download with {len(all_files_data)} files")
        
        return send_file(
            BytesIO(excel_data),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name="bulk_harmonic_reports_v3.xlsx"
        )
    
    except Exception as e:
        logger.error(f"Error generating bulk download: {str(e)}", exc_info=True)
        flash(f'Error generating bulk download: {str(e)}', 'danger')
        return redirect(url_for('index'))

@app.errorhandler(413)
def too_large(e):
    flash("File is too large. Maximum file size is 50MB.", 'danger')
    return redirect(url_for('index'))

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

@app.errorhandler(500)
def server_error(e):
    logger.error(f"Server error: {str(e)}", exc_info=True)
    flash('An internal server error occurred', 'danger')
    return redirect(url_for('index'))

if __name__ == '__main__':
    logger.info("Starting Flask application v3 with enhanced page tracking")
    app.run(debug=True, host='0.0.0.0', port=5000)