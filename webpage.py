from flask import Flask
from waitress import serve
app = Flask('app')


@app.route('/')
def run():
    with open('setupmonitor.html', 'r', encoding='utf-8') as f:
        page_content = f.read()
    return page_content


def start():
    print('Web server started\nHost: http://192.168.255.158:8080/')
    serve(app, host='0.0.0.0', port=8080)


if __name__ == '__main__':
    start()
