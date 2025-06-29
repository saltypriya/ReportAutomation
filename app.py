from flask import Flask, render_template, request, send_file, url_for
import os
import zipfile
import io
from report_generator import generate_perfect_reports

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
REPORT_FOLDER = 'reports'
PHOTOS_FOLDER = 'photos'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    generated_files = []
    if request.method == 'POST':
        file = request.files.get('excel')
        if file and file.filename.endswith('.xlsx'):
            excel_path = os.path.join(UPLOAD_FOLDER, 'claim_data.xlsx')
            file.save(excel_path)

            # Clear previous reports
            for f in os.listdir(REPORT_FOLDER):
                os.remove(os.path.join(REPORT_FOLDER, f))

            # Generate reports
            generate_perfect_reports(
                excel_path=excel_path,
                output_folder=REPORT_FOLDER,
                photos_path=PHOTOS_FOLDER
            )

            # List files (DON'T trigger download)
            generated_files = os.listdir(REPORT_FOLDER)

    else:
        # If just GET request, show existing ones if needed
        generated_files = os.listdir(REPORT_FOLDER)

    return render_template('index.html', files=generated_files)


@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(REPORT_FOLDER, filename)
    return send_file(path, as_attachment=True)

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
    app.run(debug=True)
