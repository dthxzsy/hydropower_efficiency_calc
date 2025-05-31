from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO
import time

app = Flask(__name__)
app.config['SECRET_KEY'] = 'secret!'
socketio = SocketIO(app)

def long_task(sid):
    """Simulates a long running task and sends progress updates via WebSocket."""
    total = 100
    for i in range(total):
        time.sleep(0.1)  # 模拟处理时间
        progress = int((i + 1) / total * 100)
        socketio.emit('progress', {'progress': progress}, room=sid)  # 发送进度
    socketio.emit('complete', {'message': 'Processing complete!'}, room=sid)  # 发送完成消息

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    #使用request.sid 获取本次会话ID
    socketio.start_background_task(long_task, request.sid)  # 启动后台任务
    return jsonify({'status': 'processing'})

@socketio.on('connect')
def connect_handler():
    print('Client connected')

if __name__ == "__main__":
    socketio.run(app, debug=True)