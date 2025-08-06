#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Railway部署版 PDF轉Excel Web應用
使用完整的Python PDF抽取器邏輯
"""

from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import os
import tempfile
from datetime import datetime
from final.pdf_extractor import FinalPDFExtractor
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# 設定上傳檔案大小限制 (50MB for Railway)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# HTML模板
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF 轉 Excel 工具 - 完整版</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft JhengHei', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            padding: 40px;
            max-width: 700px;
            width: 90%;
            text-align: center;
        }

        .title {
            color: #333;
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 700;
        }

        .subtitle {
            color: #666;
            font-size: 1.1em;
            margin-bottom: 40px;
        }

        .drop-zone {
            border: 3px dashed #ddd;
            border-radius: 15px;
            padding: 60px 20px;
            margin: 30px 0;
            background: #fafafa;
            transition: all 0.3s ease;
            cursor: pointer;
        }

        .drop-zone:hover, .drop-zone.dragover {
            border-color: #667eea;
            background: #f0f4ff;
            transform: scale(1.02);
        }

        .drop-icon {
            font-size: 4em;
            color: #ddd;
            margin-bottom: 20px;
        }

        .drop-zone.dragover .drop-icon {
            color: #667eea;
        }

        .drop-text {
            font-size: 1.3em;
            color: #666;
            margin-bottom: 15px;
        }

        .drop-hint {
            font-size: 0.9em;
            color: #999;
        }

        .file-input { display: none; }

        .upload-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 25px;
            font-size: 1.1em;
            cursor: pointer;
            margin: 20px 10px;
            transition: transform 0.2s ease;
        }

        .upload-btn:hover { transform: translateY(-2px); }

        .status {
            margin: 20px 0;
            padding: 15px;
            border-radius: 10px;
            display: none;
        }

        .status.success {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }

        .status.error {
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }

        .status.processing {
            background: #d1ecf1;
            border: 1px solid #bee5eb;
            color: #0c5460;
        }

        .progress {
            width: 100%;
            height: 8px;
            background: #e0e0e0;
            border-radius: 4px;
            overflow: hidden;
            margin: 10px 0;
            display: none;
        }

        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #667eea, #764ba2);
            width: 0%;
            transition: width 0.3s ease;
        }

        .download-btn {
            background: #28a745;
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 25px;
            font-size: 1.1em;
            cursor: pointer;
            margin: 20px 0;
            display: none;
            transition: transform 0.2s ease;
        }

        .download-btn:hover {
            transform: translateY(-2px);
            background: #218838;
        }

        .file-info {
            background: #f8f9fa;
            border-radius: 10px;
            padding: 15px;
            margin: 20px 0;
            display: none;
        }

        .features {
            text-align: left;
            margin-top: 30px;
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
        }

        .features h3 {
            color: #333;
            margin-bottom: 15px;
            text-align: center;
        }

        .features ul {
            list-style: none;
            padding-left: 0;
        }

        .features li {
            padding: 8px 0;
            color: #555;
            position: relative;
            padding-left: 20px;
        }

        .features li:before {
            content: "✅";
            position: absolute;
            left: 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="title">📄→📊</div>
        <h1 class="title">PDF 轉 Excel</h1>
        <p class="subtitle">完整版 - 使用Python後端精確解析</p>

        <div class="drop-zone" id="dropZone">
            <div class="drop-icon">📁</div>
            <div class="drop-text">拖拽PDF檔案到這裡</div>
            <div class="drop-hint">或點擊選擇檔案</div>
        </div>

        <input type="file" id="fileInput" class="file-input" accept=".pdf">
        
        <button class="upload-btn" onclick="document.getElementById('fileInput').click()">
            選擇檔案
        </button>

        <div class="file-info" id="fileInfo">
            <div class="file-name" id="fileName"></div>
            <div class="file-size" id="fileSize"></div>
        </div>

        <div class="progress" id="progress">
            <div class="progress-bar" id="progressBar"></div>
        </div>

        <div class="status" id="status"></div>

        <button class="download-btn" id="downloadBtn">
            📥 下載 Excel 檔案
        </button>

        <div class="features">
            <h3>🚀 完整版功能特色</h3>
            <ul>
                <li>使用Python後端，100%精確解析</li>
                <li>支援所有PDF格式和複雜結構</li>
                <li>完整的H/I系列材料分類</li>
                <li>可變區塊解析（PD到下一個PD）</li>
                <li>多材料追加模式</li>
                <li>純數字格式便於Excel運算</li>
                <li>支援特殊I系列代碼識別</li>
                <li>容錯處理缺漏欄位</li>
            </ul>
        </div>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const status = document.getElementById('status');
        const progress = document.getElementById('progress');
        const progressBar = document.getElementById('progressBar');
        const downloadBtn = document.getElementById('downloadBtn');

        let currentFile = null;

        // 拖拽功能
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFile(files[0]);
            }
        });

        dropZone.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        });

        function handleFile(file) {
            if (file.type !== 'application/pdf') {
                showStatus('請選擇PDF檔案', 'error');
                return;
            }

            currentFile = file;
            
            fileName.textContent = file.name;
            fileSize.textContent = formatFileSize(file.size);
            fileInfo.style.display = 'block';

            convertToExcel(file);
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        function showStatus(message, type) {
            status.textContent = message;
            status.className = `status ${type}`;
            status.style.display = 'block';
        }

        function updateProgress(percent) {
            progress.style.display = 'block';
            progressBar.style.width = percent + '%';
        }

        function convertToExcel(file) {
            showStatus('正在上傳並處理PDF檔案...', 'processing');
            updateProgress(10);

            const formData = new FormData();
            formData.append('pdf_file', file);

            let progressValue = 10;
            const progressInterval = setInterval(() => {
                progressValue += Math.random() * 15;
                if (progressValue < 85) {
                    updateProgress(progressValue);
                }
            }, 800);

            fetch('/api/convert-pdf', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                clearInterval(progressInterval);
                updateProgress(100);
                
                if (!response.ok) {
                    return response.json().then(err => {
                        throw new Error(err.error || '轉換失敗');
                    });
                }
                return response.blob();
            })
            .then(blob => {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = currentFile.name.replace('.pdf', '_extracted.xlsx');
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                showStatus('✅ 轉換完成！Excel檔案已下載', 'success');
                progress.style.display = 'none';
            })
            .catch(error => {
                clearInterval(progressInterval);
                console.error('Error:', error);
                showStatus(`❌ ${error.message}`, 'error');
                progress.style.display = 'none';
            });
        }

        window.addEventListener('load', () => {
            showStatus('歡迎使用完整版PDF轉Excel工具！', 'success');
        });
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    """提供HTML界面"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/convert-pdf', methods=['POST'])
def convert_pdf():
    """處理PDF轉Excel的API端點"""
    try:
        logger.info("收到PDF轉換請求")
        
        if 'pdf_file' not in request.files:
            return jsonify({'error': '未上傳檔案'}), 400
        
        file = request.files['pdf_file']
        
        if file.filename == '':
            return jsonify({'error': '未選擇檔案'}), 400
        
        if not file.filename.lower().endswith('.pdf'):
            return jsonify({'error': '請上傳PDF檔案'}), 400
        
        logger.info(f"處理檔案: {file.filename}")
        
        # 創建臨時檔案來保存上傳的PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            file.save(temp_pdf.name)
            temp_pdf_path = temp_pdf.name
        
        try:
            # 使用完整版PDF抽取器
            logger.info("開始PDF解析")
            extractor = FinalPDFExtractor(temp_pdf_path)
            orders = extractor.extract_orders()
            
            if not orders:
                return jsonify({'error': '未能從PDF中抽取到訂單資料，請檢查PDF格式'}), 400
            
            logger.info(f"成功解析 {len(orders)} 筆訂單")
            
            # 創建臨時Excel檔案
            temp_dir = tempfile.mkdtemp()
            excel_filename = f"{file.filename.replace('.pdf', '')}_extracted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            excel_path = os.path.join(temp_dir, excel_filename)
            
            # 使用完整版Excel輸出功能
            extractor._save_to_excel(excel_path)
            
            logger.info("Excel檔案生成完成")
            
            return send_file(
                excel_path,
                as_attachment=True,
                download_name=excel_filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        finally:
            # 清理臨時PDF檔案
            try:
                os.unlink(temp_pdf_path)
            except OSError:
                pass
                
    except Exception as e:
        logger.error(f"轉換過程中發生錯誤: {str(e)}")
        return jsonify({'error': f'處理失敗: {str(e)}'}), 500

@app.route('/health', methods=['GET'])
def health_check():
    """健康檢查端點"""
    return jsonify({
        'status': 'ok',
        'message': 'PDF轉Excel服務運行正常',
        'timestamp': datetime.now().isoformat()
    })

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': '檔案過大，請上傳小於50MB的PDF檔案'}), 413

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)