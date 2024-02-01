# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import datetime
import json
import os
import urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib import parse

host = ('', 12007)


class MyHandler(BaseHTTPRequestHandler):

    def do_POST(self):
        length = int(self.headers['Content-Length'])
        print('in handle ' + str(length))
        query_string = self.rfile.read(length)
        content = json.loads(str(query_string, 'latin-1'))
        # print(content)

        os.makedirs('data', exist_ok=True)
        cur_time = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S_%f")
        with open(os.path.join('data', cur_time + '.json'), 'w') as f:
            json.dump(content, f)

        print('in handle over ')
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps("response").encode())
        pass

    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(response_view(self.path)).encode())

    def reponse_json(self, content):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(content).encode())


def response_view(path):
    print('request path = ' + path)
    if path == '/dataSize':
        return 'dataSize = ' + str(len(os.listdir('data')))
    return "response"


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print(datetime.datetime.now().strftime("%Y_%m_%d_%H:%M:%S_%f"))
    server = HTTPServer(host, MyHandler)
    print("Starting MTS_ShoujoKageki_Server, listen at: %s:%s" % host)
    server.serve_forever()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
