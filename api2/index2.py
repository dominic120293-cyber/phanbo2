import os
import tempfile
from flask import Flask, request, render_template, send_file
from allo import run_allocation   # import hàm từ allo.py

app = Flask(__name__, template_folder='../templates')
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Không có file nào được chọn', 400
    file = request.files['file']
    if file.filename == '':
        return 'Tên file trống', 400
    if not file.filename.endswith('.xlsx'):
        return 'Chỉ hỗ trợ file .xlsx', 400

    input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input.xlsx')
    file.save(input_path)

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Allocation_Result.xlsx')
    try:
        run_allocation(input_path, output_path)   # gọi hàm từ allo
    except Exception as e:
        return f'Lỗi xử lý: {str(e)}', 500

    return send_file(output_path, as_attachment=True, download_name='Allocation_Result.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
