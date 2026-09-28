import os
import shutil
import openpyxl
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, url_for
import pypdfium2 as pdfium
import generator

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'static', 'uploads')
app.config['GENERATED_FOLDER'] = os.path.join(os.getcwd(), 'generated_pdfs')
app.config['SIGNATURE_FOLDER'] = os.path.join(os.getcwd(), 'static', 'uploads', 'signatures')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)
os.makedirs(app.config['SIGNATURE_FOLDER'], exist_ok=True)

# Default template & excel paths if available in workspace
DEFAULT_TEMPLATE = os.path.join(os.getcwd(), '250371509001.pdf')
DEFAULT_EXCEL = os.path.join(os.getcwd(), 'I SEM DIGIVAL PURPOSE-25 AB F+RV.xlsx')

generation_status = {
    "status": "idle",
    "total": 0,
    "current": 0,
    "generated": 0,
    "failed": 0,
    "message": "Ready to generate",
    "zip_available": False,
    "preview_img": None
}

@app.route('/')
def index():
    template_exists = os.path.exists(DEFAULT_TEMPLATE)
    excel_exists = os.path.exists(DEFAULT_EXCEL)
    sheets = []
    if excel_exists:
        try:
            wb = openpyxl.load_workbook(DEFAULT_EXCEL, read_only=True)
            sheets = wb.sheetnames
        except Exception:
            pass
            
    return render_template(
        'index.html',
        template_exists=template_exists,
        excel_exists=excel_exists,
        sheets=sheets
    )

@app.route('/api/get_sheets', methods=['POST'])
def get_sheets():
    excel_file = request.files.get('excel')
    if excel_file:
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], 'uploaded_excel.xlsx')
        excel_file.save(save_path)
    else:
        save_path = DEFAULT_EXCEL
        
    if not os.path.exists(save_path):
        return jsonify({'error': 'Excel file not found'}), 400
        
    try:
        wb = openpyxl.load_workbook(save_path, read_only=True)
        return jsonify({'sheets': wb.sheetnames})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate', methods=['POST'])
def generate():
    global generation_status
    generation_status = {
        "status": "running",
        "total": 0,
        "current": 0,
        "generated": 0,
        "failed": 0,
        "message": "Initializing...",
        "zip_available": False,
        "preview_img": None
    }

    # Handle Uploads
    template_file = request.files.get('template')
    if template_file and template_file.filename:
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template.pdf')
        template_file.save(template_path)
    else:
        template_path = DEFAULT_TEMPLATE

    excel_file = request.files.get('excel')
    if excel_file and excel_file.filename:
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        excel_file.save(excel_path)
    else:
        excel_path = DEFAULT_EXCEL

    signature_files = request.files.getlist('signatures')
    sig_folder = app.config['SIGNATURE_FOLDER']
    if signature_files and any(f.filename for f in signature_files):
        # Clear previous signatures
        shutil.rmtree(sig_folder, ignore_errors=True)
        os.makedirs(sig_folder, exist_ok=True)
        for f in signature_files:
            if f.filename:
                f.save(os.path.join(sig_folder, f.filename))
    else:
        generator.ensure_default_signatures(sig_folder)

    selected_sheet = request.form.get('sheet', 'ALL')
    sheets_param = None if selected_sheet == 'ALL' else [selected_sheet]

    date_mode = request.form.get('date_mode', 'fixed')
    fixed_date = request.form.get('fixed_date', '06.04.2026')
    start_date = request.form.get('start_date', '2026-04-01')
    end_date = request.form.get('end_date', '2026-04-15')

    def update_progress(current, total, gen, fail, msg):
        global generation_status
        generation_status["current"] = current
        generation_status["total"] = total
        generation_status["generated"] = gen
        generation_status["failed"] = fail
        generation_status["message"] = msg

    try:
        res = generator.process_excel_and_generate_all(
            template_path=template_path,
            excel_path=excel_path,
            signature_folder=sig_folder,
            output_base_dir=app.config['GENERATED_FOLDER'],
            selected_sheets=sheets_param,
            date_mode=date_mode,
            fixed_date=fixed_date,
            start_date=start_date,
            end_date=end_date,
            progress_callback=update_progress
        )

        # Generate sample preview image
        preview_rel_url = None
        if res["sample_pdf"] and os.path.exists(res["sample_pdf"]):
            try:
                pdf = pdfium.PdfDocument(res["sample_pdf"])
                preview_img_path = os.path.join(os.getcwd(), 'static', 'preview_sample.png')
                pdf[0].render(scale=2.0).to_pil().save(preview_img_path)
                preview_rel_url = '/static/preview_sample.png'
            except Exception as pe:
                print("Preview render error:", pe)

        generation_status = {
            "status": "completed",
            "total": res["total"],
            "current": res["total"],
            "generated": res["generated"],
            "failed": res["failed"],
            "message": "Generation completed successfully!",
            "zip_available": True,
            "preview_img": preview_rel_url
        }

        return jsonify({
            "status": "completed",
            "results": res,
            "preview_img": preview_rel_url
        })

    except Exception as e:
        generation_status["status"] = "error"
        generation_status["message"] = f"Error: {str(e)}"
        return jsonify({"error": str(e)}), 500

@app.route('/api/status')
def get_status():
    return jsonify(generation_status)

@app.route('/download/zip')
def download_zip():
    zip_path = os.path.join(app.config['GENERATED_FOLDER'], 'output.zip')
    if os.path.exists(zip_path):
        return send_file(zip_path, as_attachment=True, download_name='output.zip')
    return "ZIP file not generated yet.", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
