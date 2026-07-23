from flask import Flask, render_template, request, jsonify
import os

app = Flask(__name__)

@app.route('/tg/app/')
def index():
    return render_template('index.html')

@app.route('/tg/app/submit', methods=['POST'])
def submit():
    data = request.get_json()
    print(f"Получены данные: {data}")
    return jsonify({"status": "ok", "message": "Данные получены!"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)