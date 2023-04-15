from flask import Flask, render_template, request, send_file, url_for, redirect
from datetime import datetime, timedelta
import read_excel
from io import BytesIO
from zipfile import ZipFile
import os
import webbrowser

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('input_form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        date = request.form.get("date")
        next_hw_str = request.form.get("next_hw")
        grade = request.form.get("grade")
        if request.files['excel_test'].filename == '':
            excel_test = ''
        else:
            excel_test = request.files.get("excel_test")
        if request.files['excel_hw'].filename == '':
            excel_hw = ''
        else:
            excel_hw = request.files.get("excel_hw")
        if request.files['excel_daily'].filename == '':
            excel_daily = ''
        else:
            excel_daily = request.files.get("excel_daily")
        
        saturday = datetime.strptime(date, '%Y-%m-%d')   
        sunday = saturday + timedelta(days=1)
        date_list = [[str(saturday.month), str(saturday.day)], [str(sunday.month), str(sunday.day)]]
        print('date_list', date_list)
        for i in range(2):
            for j in range(2):
                if len(date_list[i][j]) <= 1:
                    date_list[i][j] = '0'+date_list[i][j]
                
        
        read_e = read_excel.Create_Message(date_list)
        df_list = []
        if excel_hw:
            df_hw = read_e.find_hw(excel_hw, date_list)
            df_list.append(df_hw)
        else:
            df_hw =''
        if excel_daily:
            df_daily = read_e.find_daily(excel_daily, date_list)
            df_list.append(df_daily)
        else:
            df_daily=''
        if excel_test:
            df_test, mean_score = read_e.find_test(excel_test, date_list)
            df_list.append(df_test)
        else:
            df_test=''
            mean_score=''
        if len(df_list):
            df_students = read_e.merge_df2(df_list, '이름')
            
            df_students['text'] = df_students.apply(read_e.make_text, args=(date_list, mean_score, next_hw_str), axis=1)
            # folder_name = f'고{grade}_{date_list[0][0]}월_{date_list[0][1]}일_{date_list[1][0]}월_{date_list[1][1]}일_문자'
            folder_name = 'folder_text'
            read_e.save_text_files(folder_name, df_students)
            
            return redirect(url_for('download_files', folder_name=folder_name))
        else:
            return render_template('fail.html',error=error)
    except Exception as error:
            return render_template('fail.html',error=error)


@app.route('/download_zip/<folder_name>')
def download_files(folder_name):
    # 폴더에서 텍스트 파일들을 불러옴
    files = [f for f in os.listdir("folder_text") if f.endswith(".txt")]
    # 텍스트 파일들과 함께 HTML 페이지 렌더링
    return render_template("text_files.html", files=files)

@app.route("/download/<filename>")
def download_file(filename):
    # 다운로드 링크를 눌렀을 때 해당 파일 다운로드
    return send_file(os.path.join("folder_text", filename), as_attachment=True)

@app.route('/download_all')
def download_all_files():
    folder_path = 'folder_text'
    zip_filename = 'all_files.zip'
    # create a zip file with all the files in the folder
    with ZipFile(zip_filename, 'w') as zip_file:
        for file in os.listdir(folder_path):
            zip_file.write(os.path.join(folder_path, file), file)
    # send the zip file to the client
    return send_file(zip_filename, as_attachment=True)

if __name__ == '__main__':
    webbrowser.open_new('http://127.0.0.1:5000')
    app.run(debug=True)
