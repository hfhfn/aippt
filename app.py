import logging

from flask import Flask, request, send_file, jsonify
import os
import json
from aippt import generate_ppt_content, generate_ppt_file, cache_dir, ppt_dir
from dotenv import load_dotenv

load_dotenv()

last_topic = ""
last_pages = 0

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate_ppt():
    global last_topic, last_pages

    data = request.json
    topic = data.get('topic')
    pages = data.get('pages')
    design_number = data.get('design_number')
    layout_index = data.get('layout_index')
    design_number = design_number if design_number else 0
    layout_index = int(layout_index) if layout_index else 0

    if not all([topic, pages]):
        return jsonify({"error": "Missing required parameters topic or pages"}), 400

    # 生成PPT内容
    if os.path.exists(f"{cache_dir}/{topic}.txt") and last_topic == topic and last_pages == pages:
        ppt_content = json.load(open(f"{cache_dir}/{topic}.txt", "r", encoding="utf-8"))
        logging.info(f"从缓存中读取PPT内容...\n\n{ppt_content}")
    else:
        ppt_content = generate_ppt_content(topic, pages)
    last_topic = topic
    last_pages = pages

    # 生成PPT文件
    # ppt_filename = f"../output/ppt/{topic}.pptx"
    generate_ppt_file(topic, ppt_content, design_number, layout_index)

    # 获取 host_url
    host_url = request.host_url
    host_url = host_url if host_url != "http://host.docker.internal/" else "http://localhost:8000/"

    # 返回生成的PPT文件
    # return send_file(ppt_filename, as_attachment=True, download_name=ppt_filename)
    download_url = f"{host_url}ppt/download/{topic}.pptx"
    return f"[点击下载 PPT 文件]({download_url})"


@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    file_path = os.path.join("../output/ppt", filename)
    if not os.path.exists(file_path):
        app.logger.error(f'File not found: {file_path}')
        return jsonify({'error': 'File not found'}), 404

    try:
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        app.logger.error(f'Error sending file: {str(e)}')
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
