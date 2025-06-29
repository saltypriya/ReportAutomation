from flask import Flask, render_template, request, send_file, url_for
import os
import zipfile
import io
from report_generator import generate_perfect_reports

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
REPORT_FOLDER = 'reports'
PHOTOS_FOLDER = 'photos'

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

# Clear old reports on startup
for f in os.listdir(REPORT_FOLDER):
    os.remove(os.path.join(REPORT_FOLDER, f))

@app.route('/', methods=['GET', 'POST'])
def index():
    generated_files = []
    show_results = False  # Only show files if Excel was uploaded

    if request.method == 'POST':
        file = request.files.get('excel')
        if file and file.filename.endswith('.xlsx'):
            # Save Excel file
            excel_path = os.path.join(UPLOAD_FOLDER, 'claim_data.xlsx')
            file.save(excel_path)

            # Clear old reports
            for f in os.listdir(REPORT_FOLDER):
                os.remove(os.path.join(REPORT_FOLDER, f))

            # Generate new reports
            generate_perfect_reports(
                excel_path=excel_path,
                output_folder=REPORT_FOLDER,
                photos_path=PHOTOS_FOLDER
            )

            generated_files = os.listdir(REPORT_FOLDER)
            show_results = True

    return render_template('index.html', files=generated_files, show_results=show_results)

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(REPORT_FOLDER, filename)
    return send_file(file_path, as_attachment=True)

@app.route('/download-zip')
def download_zip():
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for filename in os.listdir(REPORT_FOLDER):
            file_path = os.path.join(REPORT_FOLDER, filename)
            zipf.write(file_path, arcname=filename)
    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name="Generated_Reports.zip",
        mimetype='application/zip'
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
