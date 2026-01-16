from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from openpyxl import load_workbook
from datetime import datetime
import os
import uuid
from pathlib import Path

app = Flask(__name__)
CORS(app)

# Configuration
UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'generated_documents'
TEMPLATE_PATH = 'visa_booking_template.xlsx'

# Create directories
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

# Store for generated documents
documents_store = []

# 每日计数器文件路径
COUNTER_FILE = 'daily_counters.json'

def cleanup_storage():
    """云平台环境下清理临时存储（防止重启后文件堆积）"""
    try:
        print("Performing storage cleanup for cloud environment...")
        
        # 清理生成的文件目录
        if os.path.exists(GENERATED_FOLDER):
            for filename in os.listdir(GENERATED_FOLDER):
                filepath = os.path.join(GENERATED_FOLDER, filename)
                try:
                    if os.path.isfile(filepath):
                        os.remove(filepath)
                        print(f"Cleaned: {filename}")
                except Exception as e:
                    print(f"Error cleaning {filename}: {e}")
        
        # 清理上传目录
        if os.path.exists(UPLOAD_FOLDER):
            for filename in os.listdir(UPLOAD_FOLDER):
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                try:
                    if os.path.isfile(filepath):
                        os.remove(filepath)
                except:
                    pass
        
        # 清空内存存储（可选，根据你的需求）
        # documents_store.clear()
        
        print("Storage cleanup completed")
        
    except Exception as e:
        print(f"Error during storage cleanup: {e}")

def cleanup_old_documents():
    """自动清除超过48小时的文档"""
    import time
    from datetime import datetime, timedelta
    
    current_time = datetime.now()
    cutoff_time = current_time - timedelta(hours=48)
    
    # 清理内存中的文档记录
    original_count = len(documents_store)
    documents_store[:] = [doc for doc in documents_store 
                          if datetime.strptime(doc['generated_date'], '%Y-%m-%d %H:%M:%S') > cutoff_time]
    
    # 清理文件系统中的Excel文件
    if os.path.exists(GENERATED_FOLDER):
        for filename in os.listdir(GENERATED_FOLDER):
            filepath = os.path.join(GENERATED_FOLDER, filename)
            if os.path.isfile(filepath):
                file_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                if file_time < cutoff_time:
                    try:
                        os.remove(filepath)
                        print(f"Cleaned up old file: {filename}")
                    except Exception as e:
                        print(f"Error removing file {filename}: {e}")
    
    cleaned_count = original_count - len(documents_store)
    if cleaned_count > 0:
        print(f"Cleaned up {cleaned_count} old documents (older than 48 hours)")

def load_daily_counters():
    """加载每日计数器"""
    import json
    try:
        if os.path.exists(COUNTER_FILE):
            with open(COUNTER_FILE, 'r', encoding='utf-8') as f:
                counters = json.load(f)
                # 确保计数器值是数字
                for date in counters:
                    if isinstance(counters[date], str):
                        counters[date] = int(counters[date])
                return counters
    except:
        return {}

def save_daily_counters(counters):
    """保存每日计数器"""
    import json
    try:
        with open(COUNTER_FILE, 'w', encoding='utf-8') as f:
            json.dump(counters, f, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving counters: {e}")

def generate_confirmation_number():
    """Generate a unique confirmation number: YYMMDDXXXX"""
    today = datetime.now().strftime('%Y%m%d')
    counters = load_daily_counters()
    
    # 每天都从0001重新开始
    counters[today] = 1
    
    save_daily_counters(counters)
    return f"{today}{str(counters[today]).zfill(4)}"

@app.route('/admin')
def admin_panel():
    """后端管理页面"""
    return render_template('admin.html')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate-document', methods=['POST'])
def generate_document():
    try:
        data = request.json
        
        # Validate required fields
        required_fields = ['guestName', 'email', 'company', 'arrivalDate', 'departureDate']
        for field in required_fields:
            if not data.get(field):
                return jsonify({
                    'success': False,
                    'message': f'Missing required field: {field}'
                }), 400
        
        # Generate unique confirmation number
        confirmation_number = generate_confirmation_number()
        
        # Calculate nights
        arrival_date = datetime.strptime(data['arrivalDate'], '%Y-%m-%d')
        departure_date = datetime.strptime(data['departureDate'], '%Y-%m-%d')
        nights = (departure_date - arrival_date).days
        if nights < 1:
            nights = 1
        
        # Calculate total amount
        room_rate = 98000
        quantity = data.get('quantity', 1)
        total_amount = nights * room_rate * quantity
        
        # Check if template exists, create if not
        if not os.path.exists(TEMPLATE_PATH):
            create_template_file()
        
        # Load the template - 创建副本避免修改原文件
        import shutil
        temp_template = TEMPLATE_PATH.replace('.xlsx', '_temp.xlsx')
        shutil.copy2(TEMPLATE_PATH, temp_template)
        
        wb = load_workbook(temp_template)
        ws = wb.active
        
        # Fill in the data - 智能处理合并单元格
        # 保存原始合并区域
        original_merges = list(ws.merged_cells.ranges)
        
        # 定义需要写入数据的单元格
        data_cells = ['J5', 'J19', 'D22', 'B7', 'H22', 'K22', 'J8', 'J17', 'J9', 'J10', 'L22', 'Q22', 'T22', 'V22', 'L23', 'Q23', 'T23', 'V23']
        
        # 只取消包含我们需要写入数据的合并区域
        merges_to_remove = []
        for merge_range in original_merges:
            should_remove = False
            for cell_addr in data_cells:
                cell = ws[cell_addr]
                if merge_range.min_row <= cell.row <= merge_range.max_row and \
                   merge_range.min_col <= cell.column <= merge_range.max_col:
                    should_remove = True
                    break
            if should_remove:
                merges_to_remove.append(merge_range)
        
        # 取消特定的合并区域
        for merge_range in merges_to_remove:
            ws.unmerge_cells(str(merge_range))
        
        # Guest Information
        ws['J5'] = data['guestName']    # Guest Name in contact
        ws['J19'] = data['guestName']   # Guest Name in reservation
        ws['D22'] = data['guestName']   # Guest Name in table
        
        # Company Information - 不写入Excel，只保留在后台
        # ws['B7'] = data['company']  # 注释掉，不在Excel中显示
        
        # Dates
        ws['H22'] = arrival_date.strftime('%Y-%m-%d')    # Arrival Date
        ws['K22'] = departure_date.strftime('%Y-%m-%d')  # Departure Date
        ws['J8'] = datetime.now().strftime('%Y-%m-%d')   # Booking Date
        
        # Confirmation Number
        ws['J17'] = confirmation_number
        
        # Email and Remarks - 不写入Excel，只保留在后台
        # ws['J9'] = data['email']  # 注释掉，不在Excel中显示
        # remark = data.get('remark', '')
        # if data.get('purpose') == 'VISA_APPLICATION_ONLY':
        #     remark = "FOR VISA APPLICATION PURPOSES ONLY - NOT AN ACTUAL BOOKING. " + remark
        # ws['J10'] = remark  # 注释掉，不在Excel中显示
        
        # Room Information - 写入合并区域的左上角单元格
        room_info_mapping = {
            'L22': ('M22', data.get('roomType', 'Classic Queen')),  # M22:Q22合并
            'Q22': ('M22', quantity),  # M22:Q22合并，左上角是M22
            'T22': ('T22', nights),  # T22:U22合并，左上角是T22
            'V22': ('V22', room_rate)   # V22:Y22合并，左上角是V22
        }
        
        for cell_addr, (target_cell, value) in room_info_mapping.items():
            try:
                ws[target_cell] = value
            except Exception as e:
                print(f"Warning: Could not write to {target_cell}: {e}")
        
        # 重新合并我们取消的区域（保持原有格式）
        for merge_range in merges_to_remove:
            try:
                ws.merge_cells(str(merge_range))
            except Exception as e:
                print(f"Warning: Could not re-merge {merge_range}: {e}")
        
        # Add metadata
        ws['AA1'] = f"Company: {data['company']}"
        ws['AA2'] = f"Email: {data['email']}"
        ws['AA3'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws['AA4'] = f"Document ID: {confirmation_number}"
        
        # Generate filename
        safe_company = "".join(c for c in data['company'] if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_company = safe_company.replace(' ', '_')[:30]
        filename = f"Visa_Booking_{confirmation_number}_{safe_company}.xlsx"
        filepath = os.path.join(GENERATED_FOLDER, filename)
        
        # Save the workbook
        wb.save(filepath)
        
        # Clean up temp file
        try:
            os.remove(temp_template)
        except:
            pass
        
        # Store document information
        document_info = {
            'id': confirmation_number,
            'filename': filename,
            'company': data['company'],
            'email': data['email'],
            'guest_name': data['guestName'],
            'arrival_date': data['arrivalDate'],
            'departure_date': data['departureDate'],
            'nights': nights,
            'total_amount': total_amount,
            'generated_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'filepath': filepath,
            'purpose': 'VISA_APPLICATION_ONLY',
            'download_url': f'/download/{confirmation_number}',
            'print_url': f'/print/{confirmation_number}'
        }
        
        documents_store.append(document_info)
        
        # Print to console
        print("\n" + "="*60)
        print("NEW VISA BOOKING DOCUMENT GENERATED")
        print("="*60)
        print(f"Company: {data['company']}")
        print(f"Email: {data['email']}")
        print(f"Guest: {data['guestName']}")
        print(f"Dates: {data['arrivalDate']} to {data['departureDate']}")
        print(f"Nights: {nights}")
        print(f"Total: {total_amount:,} CFA")
        print(f"Document ID: {confirmation_number}")
        print(f"File: {filename}")
        print(f"Location: {filepath}")
        print("="*60 + "\n")
        
        return jsonify({
            'success': True,
            'message': 'Visa booking document generated successfully!',
            'document': {
                'id': confirmation_number,
                'filename': filename,
                'company': data['company'],
                'email': data['email'],
                'guest_name': data['guestName'],
                'nights': nights,
                'total_amount': total_amount,
                'download_url': f'/download/{confirmation_number}',
                'view_url': f'/documents/{confirmation_number}'
            }
        })
        
    except Exception as e:
        print(f"Error generating document: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return jsonify({
            'success': False,
            'message': f'Error generating document: {str(e)}'
        }), 500

@app.route('/documents', methods=['GET'])
def list_documents():
    """View all generated documents"""
    return jsonify({
        'success': True,
        'count': len(documents_store),
        'documents': [
            {
                'id': doc['id'],
                'filename': doc['filename'],
                'company': doc['company'],
                'email': doc['email'],
                'guest_name': doc['guest_name'],
                'dates': f"{doc['arrival_date']} to {doc['departure_date']}",
                'nights': doc['nights'],
                'total_amount': doc['total_amount'],
                'generated_date': doc['generated_date'],
                'download_url': doc['download_url'],
                'print_url': doc['print_url']
            }
            for doc in documents_store
        ]
    })

@app.route('/documents/<document_id>', methods=['GET'])
def get_document(document_id):
    """Get specific document information"""
    for doc in documents_store:
        if doc['id'] == document_id:
            return jsonify({
                'success': True,
                'document': doc
            })
    
    return jsonify({
        'success': False,
        'message': 'Document not found'
    }), 404

@app.route('/download/<document_id>', methods=['GET'])
def download_document(document_id):
    """Download the Excel file"""
    for doc in documents_store:
        if doc['id'] == document_id:
            if os.path.exists(doc['filepath']):
                return send_file(
                    doc['filepath'],
                    as_attachment=True,
                    download_name=doc['filename'],
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
    
    return jsonify({
        'success': False,
        'message': 'File not found'
    }), 404

@app.route('/print/<document_id>', methods=['GET'])
def print_document(document_id):
    """打印文档信息到控制台并尝试实际打印"""
    for doc in documents_store:
        if doc['id'] == document_id:
            print("\n" + "="*60)
            print("DOCUMENT PRINT REQUEST")
            print("="*60)
            print(f"Company: {doc['company']}")
            print(f"Email: {doc['email']}")
            print(f"Guest: {doc['guest_name']}")
            print(f"Dates: {doc['arrival_date']} to {doc['departure_date']}")
            print(f"Nights: {doc['nights']}")
            print(f"Total: {doc['total_amount']:,} CFA")
            print(f"Document ID: {doc['id']}")
            print(f"Generated: {doc['generated_date']}")
            print(f"File: {doc['filename']}")
            print(f"Path: {doc['filepath']}")
            print("="*60 + "\n")
            
            # 尝试实际打印Excel文件
            try:
                import subprocess
                import platform
                
                if os.path.exists(doc['filepath']):
                    system = platform.system()
                    
                    if system == 'Windows':
                        # Windows系统使用默认程序打印
                        subprocess.run(['start', '/min', doc['filepath']], shell=True, check=False)
                        print(f"已发送打印命令到系统: {doc['filename']}")
                    elif system == 'Darwin':  # macOS
                        subprocess.run(['lpr', doc['filepath']], check=False)
                        print(f"已发送打印命令到系统: {doc['filename']}")
                    elif system == 'Linux':
                        subprocess.run(['lp', doc['filepath']], check=False)
                        print(f"已发送打印命令到系统: {doc['filename']}")
                    else:
                        print(f"不支持的操作系统，无法自动打印")
                else:
                    print(f"文件不存在: {doc['filepath']}")
                    
            except Exception as e:
                print(f"打印失败: {e}")
            
            return jsonify({
                'success': True,
                'message': 'Document information printed to console and sent to printer',
                'document': {
                    'id': doc['id'],
                    'company': doc['company'],
                    'email': doc['email'],
                    'guest_name': doc['guest_name'],
                    'dates': f"{doc['arrival_date']} to {doc['departure_date']}",
                    'nights': doc['nights'],
                    'total_amount': doc['total_amount'],
                    'filename': doc['filename']
                }
            })
    
    return jsonify({
        'success': False,
        'message': 'Document not found'
    }), 404

@app.route('/cleanup', methods=['POST'])
def cleanup_documents():
    """手动清理超过48小时的文档"""
    try:
        cleanup_old_documents()
        return jsonify({
            'success': True,
            'message': 'Cleanup completed successfully',
            'remaining_documents': len(documents_store)
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Cleanup failed: {str(e)}'
        }), 500

def create_template_file():
    """Create a basic template matching your structure"""
    print("Creating template file...")
    
    wb = load_workbook()
    ws = wb.active
    ws.title = "ipms_master_bill"
    
    # Add headers and structure based on your Excel
    ws['C3'] = "Reservation Confirmation"
    
    # Left column labels
    ws['B5'] = "Booking Name"
    ws['C6'] = "Phone No."
    ws['B7'] = "Company Name"
    ws['B8'] = "Booking Date"
    ws['C9'] = "Email"
    ws['D10'] = "Remark"
    
    # Right column labels
    ws['O5'] = "Hotel"
    ws['O6'] = "Page"
    ws['O7'] = "Address"
    ws['O8'] = "Deposit(CFA)"
    
    # Separators
    ws['F5'] = ":"
    ws['F6'] = ":"
    ws['F7'] = ":"
    ws['F8'] = ":"
    ws['F9'] = ":"
    ws['F10'] = ":"
    ws['S5'] = ":"
    ws['S6'] = ":"
    ws['S7'] = ":"
    ws['S8'] = ":"
    
    # Fixed values
    ws['J6'] = "+240 333091088"
    ws['W5'] = "Hotel Anda Malabo"
    ws['W6'] = "1/ of 1"
    ws['W7'] = "Malabo II, Malabo, G.E"
    ws['W8'] = "0"
    
    # Thank you message
    ws['C13'] = "Thank you for choosing to stay at Hotel Anda Malabo. We are pleased to confirm the following reservation for you"
    
    # Confirmation and Guest Name
    ws['C16'] = "Confirmation No"
    ws['F16'] = ":"
    ws['C18'] = "Guest Name"
    ws['F18'] = ":"
    
    # Table headers
    ws['D21'] = "Name"
    ws['H21'] = "Arrival Date"
    ws['K21'] = "Departure Date"
    ws['L21'] = "Room Type"
    ws['Q21'] = "Quantity"
    ws['T21'] = "Nights"
    ws['V21'] = "Room Rate"
    ws['Z21'] = "Total (CFA)"
    
    # Table data row (row 22)
    ws['L22'] = "Classic Queen"
    ws['V22'] = "98000"
    ws['Z22'] = "=T22*V22"  # Formula for total
    
    # Save template
    wb.save(TEMPLATE_PATH)
    print(f"Template created: {TEMPLATE_PATH}")
    return True

@app.route('/check-template', methods=['GET'])
def check_template():
    """Check if template exists and its structure"""
    if os.path.exists(TEMPLATE_PATH):
        try:
            wb = load_workbook(TEMPLATE_PATH)
            ws = wb.active
            sheet_name = ws.title
            
            # Check some key cells
            key_cells = {
                'C3': ws['C3'].value,
                'B5': ws['B5'].value,
                'J6': ws['J6'].value,
                'W5': ws['W5'].value,
                'Z22': ws['Z22'].value
            }
            
            return jsonify({
                'success': True,
                'message': 'Template found and loaded successfully',
                'sheet_name': sheet_name,
                'key_cells': key_cells
            })
        except Exception as e:
            return jsonify({
                'success': False,
                'message': f'Error loading template: {str(e)}'
            }), 500
    else:
        return jsonify({
            'success': False,
            'message': f'Template file not found at: {TEMPLATE_PATH}'
        }), 404

if __name__ == '__main__':
    print("Starting Visa Booking Document Generator")
    print("="*50)
    
    # 云环境端口配置
    port = int(os.environ.get('PORT', 5000))  # 本地默认5000，云平台会自动设置
    
    # 云环境需要绑定到 0.0.0.0
    host = '0.0.0.0'
    
    print(f"Server: http://{host}:{port}")
    print(f"Template: {TEMPLATE_PATH}")
    print(f"Output folder: {GENERATED_FOLDER}")
    print(f"Environment: {'PRODUCTION' if os.environ.get('PORT') else 'DEVELOPMENT'}")
    print("\nAvailable endpoints:")
    print("  GET  /                    - Frontend form")
    print("  POST /generate-document   - Submit booking data")
    print("  GET  /documents           - List all documents")
    print("  GET  /documents/{id}      - View document info")
    print("  GET  /download/{id}       - Download Excel file")
    print("  GET  /print/{id}          - Print to console")
    print("  GET  /check-template      - Check template status")
    print("  GET  /admin               - Admin panel")
    print("\nWaiting for submissions...")
    print("="*50 + "\n")
    
    # Check template
    if not os.path.exists(TEMPLATE_PATH):
        print("Template not found, creating basic template...")
        create_template_file()
    else:
        print("Template found and ready")
    
    # 云平台环境下自动清理存储（防止文件堆积）
    if os.environ.get('PORT'):
        print("Cloud environment detected, initializing cleanup...")
        cleanup_storage()
    
    # 启动应用
    # 云平台：debug=False，本地：debug=True
    debug_mode = not bool(os.environ.get('PORT'))
    app.run(host=host, port=port, debug=debug_mode)