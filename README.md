# 📝 Report Automation Web App

A smart, browser-based tool that lets your team **upload an Excel file** and instantly **generate multiple Word reports** with photos, formatted text, and downloadable links — all in seconds.

> Built with Python 🐍, Flask ⚙️, and Microsoft Word automation 💼

---

## 🚀 Live Demo

🔗 [Visit the app on Render →](https://reportautomation.onrender.com)

---

## 📦 Features

- ✅ Upload `.xlsx` Excel files
- ✅ Auto-generate multiple `.docx` reports
- ✅ Add header/footer images (branding)
- ✅ Embed room-wise photos (like Kitchen, Bedroom, etc.)
- ✅ Download each report individually or as a ZIP
- ✅ Responsive & easy-to-use frontend (Bootstrap)
- ✅ Deployable to Render with one click

---

## 🧠 How It Works

1. **Prepare your Excel file** (`claim_data.xlsx`)  
   Each row should contain: name, address, insurer, claim info, indemnity, listing cost, etc.

2. **Upload the file** using the browser form  
   The app reads all entries and uses `python-docx` to generate formatted reports.

3. **Download your results**  
   Download individual reports or all of them as a ZIP.

---

## 💻 Technologies Used

| Backend        | Frontend     | File Handling     |
|----------------|--------------|-------------------|
| Flask (Python) | HTML/CSS     | openpyxl (Excel)  |
| python-docx    | Bootstrap 5  | docx image inserts |
| Render (Hosting) |             | zipfile (for bundling) |

---

## 🛠 Folder Structure

ReportAutomation/
├── app.py # Flask app logic
├── report_generator.py # Core logic for generating DOCX reports
├── templates/
│ └── index.html # Upload + result UI
├── uploads/ # Temporary Excel storage
├── reports/ # Generated DOCX files
├── photos/ # Header/footer and room photos
├── requirements.txt # Flask, openpyxl, python-docx
└── README.md


---

## ⚙️ Setup Instructions

> 💡 Or just use the hosted version at [https://reportautomation.onrender.com](https://reportautomation.onrender.com)

### Run Locally

```bash
git clone https://github.com/saltypriya/ReportAutomation.git
cd ReportAutomation

# (Optional) create virtual environment
python -m venv venv
venv\Scripts\activate  # For Windows

# Install dependencies
pip install -r requirements.txt

# Run the app
python app.py

Then visit: http://localhost:5000 in your browser.
