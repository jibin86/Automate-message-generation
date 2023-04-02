from flask import Flask, render_template, request
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('input_form.html')

@app.route('/submit', methods=['POST'])
def submit():
    date = request.form.get("date")
    next_hw = request.form.get("next_hw")
    grade = request.form.get("grade")
    excel_test = request.files.get("excel_test")
    excel_hw = request.files.get("excel_hw")
    excel_daily = request.files.get("excel_daily")
    
    

    return "Form submitted successfully"

if __name__ == '__main__':
    app.run(debug=True)
