from flask import Flask, render_template, request, send_from_directory
from xyj import *
import os

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


@app.route('/res', methods=['POST'])
def result():
    if request.method == 'POST':
        tm985 = request.form['985推免']
        tm211 = request.form['211推免']
        tk985 = request.form['985统考']
        tk211 = request.form['211统考']
        normal = request.form['双非推免']
        m2020 = request.form['2020级']
        m2021 = request.form['2021级']
        m2022 = request.form['2022级']
        xyj = XueYeJiang(tm985, tm211, tk985, tk211, normal, m2020, m2021, m2022)
        return send_from_directory(os.path.join(os.path.dirname(__file__), 'static/'), 'res.xlsx')


if __name__ == '__main__':
    app.run(debug=True)
