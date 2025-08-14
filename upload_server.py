#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import shutil
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import cgi
import tempfile
import subprocess
import sys
from pathlib import Path

class UploadHandler(BaseHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        self.upload_dir = Path(__file__).parent
        self.excel_file_path = self.upload_dir / "需求工单统计表.xlsx"
        super().__init__(*args, **kwargs)
    
    def do_OPTIONS(self):
        """处理预检请求"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def do_POST(self):
        """处理POST请求"""
        try:
            if self.path == '/upload':
                self.handle_file_upload()
            else:
                self.send_error(404, "Not Found")
        except Exception as e:
            print(f"处理POST请求时出错: {e}")
            self.send_json_response({'success': False, 'message': str(e)}, 500)
    
    def do_GET(self):
        """处理GET请求 - 提供静态文件服务"""
        try:
            # 解析URL路径
            parsed_path = urlparse(self.path)
            file_path = parsed_path.path.lstrip('/')
            
            # 如果是根路径，返回index.html
            if not file_path or file_path == '/':
                file_path = 'index.html'
            
            # 构建完整文件路径
            full_path = self.upload_dir / file_path
            
            # 检查文件是否存在
            if not full_path.exists() or not full_path.is_file():
                self.send_error(404, "File not found")
                return
            
            # 确定MIME类型
            content_type = self.get_content_type(file_path)
            
            # 读取并发送文件
            with open(full_path, 'rb') as f:
                content = f.read()
            
            self.send_response(200)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(content)
            
        except Exception as e:
            print(f"处理GET请求时出错: {e}")
            self.send_error(500, "Internal Server Error")
    
    def get_content_type(self, file_path):
        """根据文件扩展名确定MIME类型"""
        ext = Path(file_path).suffix.lower()
        content_types = {
            '.html': 'text/html; charset=utf-8',
            '.css': 'text/css; charset=utf-8',
            '.js': 'application/javascript; charset=utf-8',
            '.json': 'application/json; charset=utf-8',
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.svg': 'image/svg+xml',
            '.ico': 'image/x-icon'
        }
        return content_types.get(ext, 'application/octet-stream')
    
    def handle_file_upload(self):
        """处理文件上传"""
        try:
            # 解析multipart/form-data
            content_type = self.headers.get('Content-Type', '')
            if not content_type.startswith('multipart/form-data'):
                self.send_json_response({'success': False, 'message': '无效的内容类型'}, 400)
                return
            
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # 解析表单数据
                form = cgi.FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ={
                        'REQUEST_METHOD': 'POST',
                        'CONTENT_TYPE': content_type
                    }
                )
                
                # 获取上传的文件
                if 'file' not in form:
                    self.send_json_response({'success': False, 'message': '未找到上传文件'}, 400)
                    return
                
                file_item = form['file']
                if not file_item.filename:
                    self.send_json_response({'success': False, 'message': '文件名为空'}, 400)
                    return
                
                # 验证文件类型
                filename = file_item.filename.lower()
                if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
                    self.send_json_response({'success': False, 'message': '只支持Excel文件(.xlsx/.xls)'}, 400)
                    return
                
                # 保存临时文件
                temp_file = temp_path / 'uploaded_file.xlsx'
                with open(temp_file, 'wb') as f:
                    f.write(file_item.file.read())
                
                # 备份原文件
                backup_file = self.upload_dir / "需求工单统计表_backup.xlsx"
                if self.excel_file_path.exists():
                    shutil.copy2(self.excel_file_path, backup_file)
                
                # 替换原文件
                shutil.copy2(temp_file, self.excel_file_path)
                
                # 重新生成JSON数据
                self.regenerate_data()
                
                self.send_json_response({'success': True, 'message': '文件上传成功'})
                
        except Exception as e:
            print(f"文件上传处理错误: {e}")
            # 如果有备份文件，恢复原文件
            backup_file = self.upload_dir / "需求工单统计表_backup.xlsx"
            if backup_file.exists():
                try:
                    shutil.copy2(backup_file, self.excel_file_path)
                    backup_file.unlink()  # 删除备份文件
                except:
                    pass
            
            self.send_json_response({'success': False, 'message': f'处理失败: {str(e)}'}, 500)
    
    def regenerate_data(self):
        """重新生成JSON数据"""
        try:
            # 运行数据处理脚本
            data_processor_path = self.upload_dir / "data_processor.py"
            if not data_processor_path.exists():
                raise Exception("数据处理脚本不存在")
            
            # 执行数据处理脚本
            result = subprocess.run(
                [sys.executable, str(data_processor_path)],
                cwd=str(self.upload_dir),
                capture_output=True,
                text=True,
                timeout=60  # 60秒超时
            )
            
            if result.returncode != 0:
                raise Exception(f"数据处理失败: {result.stderr}")
            
            print("数据重新生成成功")
            
        except subprocess.TimeoutExpired:
            raise Exception("数据处理超时")
        except Exception as e:
            raise Exception(f"数据重新生成失败: {str(e)}")
    
    def send_json_response(self, data, status_code=200):
        """发送JSON响应"""
        response = json.dumps(data, ensure_ascii=False)
        response_bytes = response.encode('utf-8')
        
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(response_bytes)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(response_bytes)
    
    def log_message(self, format, *args):
        """自定义日志格式"""
        print(f"[{self.date_time_string()}] {format % args}")

def run_server(port=8001):
    """启动服务器"""
    server_address = ('', port)
    httpd = HTTPServer(server_address, UploadHandler)
    print(f"上传服务器启动在端口 {port}")
    print(f"访问地址: http://localhost:{port}")
    
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
        httpd.server_close()

if __name__ == '__main__':
    # 检查端口参数
    port = 8001
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            print("无效的端口号，使用默认端口8001")
    
    run_server(port)