import socket

from flask import Flask
from flask import request, jsonify

app = Flask(__name__)


@app.route('/success', methods=['POST', 'GET'])
def hello_world():
    if request.method == 'POST':
        return '', 200
    elif request.method == 'GET':
        return jsonify(
            {
                "modules": [
                    {
                        "file_name": "permission.py",
                        "sha256": "f63b45df73a6567d8144421364e01a843d203c1f0bf300ddc363d767705d2b56",
                        "machine_names": [
                            "*"
                        ]
                    },
                    {
                        "file_name": "permission2.py",
                        "sha256": "355200bb9ae00fcec1d7b660e7dd95fb3dbf246a9db397a6daa2471458a8e6cb",
                        "machine_names": [
                            socket.gethostname()
                        ]
                    }
                ]
            }

        )


if __name__ == '__main__':
    app.run(host='127.0.0.1', port='5000', debug=True)
