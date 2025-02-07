from flask import Flask, render_template, request, send_file
import pandas as pd
# import openpyxl
import os

# create instance of Flask application
app = Flask(__name__)

# 업로드된 파일을 저장할 폴더
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 기본 페이지
@app.route('/')
def index():
  return render_template('index.html')

# 파일 처리 및 업로드 처리
@app.route('/upload', methods=['POST'])
def upload_file():
  if 'file' not in request.files:
    return "No file part"
  
  file = request.files['file']
  if file.filename == '':
    return "No selected file"
  
  if file and file.filename.endswith('.xlsx'):
    filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filename)
    
    # 엑셀 파일 처리
    try:
      # 파일 읽기
      df = pd.read_excel(filename)

      # 여기에 파일 처리하는 로직 추가 (예: 데이터를 가공)
      # 예시로 'Year' 컬럼을 1년 더하는 작업을 해본다
      if 'Year' in df.columns:
          df['Year'] = df['Year'] + 1  # Year 컬럼에 1 더하기

      # 결과를 새로운 엑셀 파일로 저장
      output_filename = 'processed_file.xlsx'
      output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
      df.to_excel(output_path, index=False)

      # 처리된 파일 다운로드 링크 제공
      return send_file(output_path, as_attachment=True)

    except Exception as e:
      return f"Error processing the file: {str(e)}"
  
  return "Invalid file format. Please upload a .xlsx file."

if __name__ == '__main__':
    app.run(debug=True)
