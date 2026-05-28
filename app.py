import os
import tempfile
from datetime import date
from flask import Flask, request, render_template, send_file, jsonify
from po_parser import parse_po
from excel_generator import generate_excel

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload


@app.route('/')
def index():
    today = date.today().strftime('%Y-%m-%d')
    return render_template('index.html', today=today)


@app.route('/generate', methods=['POST'])
def generate():
    # ── validate inputs ────────────────────────────────────────────
    if 'po_pdf' not in request.files:
        return jsonify({'error': 'No PDF file uploaded'}), 400

    pdf_file = request.files['po_pdf']
    if not pdf_file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    invoice_number = request.form.get('invoice_number', '').strip()
    if not invoice_number:
        return jsonify({'error': 'Invoice number is required'}), 400

    invoice_date_raw = request.form.get('invoice_date', '')
    try:
        # HTML date input gives YYYY-MM-DD; convert to DD.MM.YYYY
        y, m, d = invoice_date_raw.split('-')
        invoice_date = f'{d}.{m}.{y}'
    except Exception:
        invoice_date = date.today().strftime('%d.%m.%Y')

    try:
        carton_capacity = int(request.form.get('carton_capacity', 100))
        if carton_capacity <= 0:
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({'error': 'Carton capacity must be a positive integer'}), 400

    # ── save PDF to temp file, parse, generate ─────────────────────
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        pdf_file.save(tmp.name)
        tmp_path = tmp.name

    try:
        po_data = parse_po(tmp_path)
    except Exception as e:
        os.unlink(tmp_path)
        return jsonify({'error': f'Could not parse PDF: {str(e)}'}), 422
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    if not po_data.get('items'):
        return jsonify({'error': 'No line items found in the PDF. '
                                 'Make sure it is a valid purchase order.'}), 422

    try:
        excel_bytes = generate_excel(po_data, invoice_number, carton_capacity, invoice_date)
    except Exception as e:
        return jsonify({'error': f'Excel generation failed: {str(e)}'}), 500

    po_no = po_data.get('po_number', 'invoice').replace('/', '-')
    filename = f'{invoice_number.replace("/", "-")}-{po_no}.xlsx'

    return send_file(
        excel_bytes,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
