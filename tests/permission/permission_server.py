import socket

from flask import Flask
from flask import request, jsonify

app = Flask(__name__)


@app.route('/success', methods=['POST', 'GET'])
def success():
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


@app.route('/success-embedded', methods=['POST', 'GET'])
def success_embedded():
    # Different hashes because Git manipulated the line endings
    if request.method == 'POST':
        return '', 200
    elif request.method == 'GET':
        return jsonify(
            {
                "modules": [
                    {
                        "file_name": "permission.py",
                        "sha256": "492f87359ec97642d6bbe5c2dfc519f76f78d45ee997bb426c3d3b39255b6a9f",
                        "machine_names": [
                            "*"
                        ]
                    },
                    {
                        "file_name": "permission2.py",
                        "sha256": "2857cda216b0b54232e0e3c813436d7025d08c8b6c3e9a2d7e8cc125e22d9d57",
                        "machine_names": [
                            socket.gethostname()
                        ]
                    }
                ]
            }

        )


@app.route('/fail-machinename', methods=['POST', 'GET'])
def fail_hostname():
    if request.method == 'POST':
        return '', 403
    elif request.method == 'GET':
        return jsonify(
            {
                "modules": [
                    {
                        "file_name": "permission.py",
                        "sha256": "f63b45df73a6567d8144421364e01a843d203c1f0bf300ddc363d767705d2b56",
                        "machine_names": [
                        ]
                    },
                    {
                        "file_name": "permission2.py",
                        "sha256": "355200bb9ae00fcec1d7b660e7dd95fb3dbf246a9db397a6daa2471458a8e6cb",
                        "machine_names": [
                            "abc", "def"
                        ]
                    }
                ]
            }

        )


@app.route('/fail-hash', methods=['POST', 'GET'])
def fail_hash():
    if request.method == 'POST':
        return '', 403
    elif request.method == 'GET':
        return jsonify(
            {
                "modules": [
                    {
                        "file_name": "permission.py",
                        "sha256": "x63b45df73a6567d8144421364e01a843d203c1f0bf300ddc363d767705d2b56",
                        "machine_names": ["*"
                        ]
                    },
                    {
                        "file_name": "permission2.py",
                        "sha256": "x55200bb9ae00fcec1d7b660e7dd95fb3dbf246a9db397a6daa2471458a8e6cb",
                        "machine_names": [
                            socket.gethostname()
                        ]
                    }
                ]
            }
        )


@app.route('/fail-filename', methods=['POST', 'GET'])
def fail_filename():
    if request.method == 'POST':
        return '', 403
    elif request.method == 'GET':
        return jsonify(
            {
                "modules": [
                    {
                        "file_name": "notexisting1.py",
                        "sha256": "f63b45df73a6567d8144421364e01a843d203c1f0bf300ddc363d767705d2b56",
                        "machine_names": ["*"
                        ]
                    },
                    {
                        "file_name": "notexisting2.py",
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
