import io
import zipfile
from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files.get('excel_file')
        if not excel_file:
            return "No file uploaded", 400

        # Read the uploaded Excel file (works with .xlsx and .xls)
        df = pd.read_excel(excel_file)

        # Get the first detected column to be used as identifier for file names
        first_col = df.columns[0]

        # Create an in-memory ZIP file
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            for index, row in df.iterrows():
                identifier = str(row[first_col])
                pdf_buffer = io.BytesIO()
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt="Faculty CV", ln=True, align="C")
                pdf.ln(10)
                # Add each column's data to the PDF
                for col in df.columns:
                    pdf.cell(0, 10, txt=f"{col}: {row[col]}", ln=True)
                pdf.output(pdf_buffer)
                pdf_buffer.seek(0)
                zf.writestr(f"{identifier}.pdf", pdf_buffer.getvalue())
        memory_file.seek(0)
        return send_file(memory_file, download_name="faculty_cvs.zip", as_attachment=True)
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)
