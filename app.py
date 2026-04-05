import os
import json
import io
import urllib.request
import pandas as pd
import openpyxl
from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

LOCAL_EXCEL_PATH = os.environ.get('EXCEL_PATH', 'BaoCaoNgay_TTVT_XLC OK.xlsx')
USERS_FILE = os.environ.get('USERS_PATH', 'users.json')
GSHEET_URL = 'https://docs.google.com/spreadsheets/d/1NqIySkzy9XXSprd6SyBuJjj7qPm4rcaCGVJZQkWEVa8/export?format=csv&gid=230040167'

def get_users():
    if not os.path.exists(USERS_FILE):
        return []
    with open(USERS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=4)

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    cccd = data.get('cccd')
    password = data.get('password')
    
    users = get_users()
    for user in users:
        # Check against hashed password, or handle the initial plain 'admin' case gracefully
        is_valid = False
        stored_pass = str(user.get('password'))
        if stored_pass.startswith('scrypt:') or stored_pass.startswith('pbkdf2:'):
            is_valid = check_password_hash(stored_pass, str(password))
        else:
            is_valid = (stored_pass == str(password))
            
        if str(user.get('cccd')) == str(cccd) and is_valid:
            return jsonify({
                "success": True, 
                "user": {"cccd": user['cccd'], "ten": user.get('ten', ''), "role": user['role']}
            })
            
    return jsonify({"success": False, "message": "Sai thông tin đăng nhập"}), 401

@app.route('/api/upload-users', methods=['POST'])
def upload_users():
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "Không tìm thấy file"}), 400
        
    file = request.files['file']
    try:
        # Expected columns: STT, Tên đăng nhập (là số CCCD của người dùng), Password , Phân quyền (Admin/Nhân viên)
        # Maybe 'Họ và tên' too
        df = pd.read_excel(file)
        
        # We try to find the columns flexibly
        cols = {c.lower().strip() for c in df.columns}
        cccd_col = next((c for c in df.columns if 'cccd' in c.lower() or 'tên đăng nhập' in c.lower()), None)
        pass_col = next((c for c in df.columns if 'mật khẩu' in c.lower() or 'password' in c.lower()), None)
        role_col = next((c for c in df.columns if 'quyền' in c.lower() or 'vai trò' in c.lower()), None)
        name_col = next((c for c in df.columns if 'họ và tên' in c.lower() or 'tên' in c.lower()), None)
        
        if not (cccd_col and pass_col and role_col and name_col):
            return jsonify({"success": False, "message": "File không đúng định dạng. Cần các cột: Họ và tên, CCCD, Mật khẩu, Phân quyền"}), 400
            
        users = []
        # Retain the existing super admin in memory to avoid locking out during bad uploads
        existing_admin = {"cccd": "admin", "ten": "Administrator", "password": generate_password_hash("admin"), "role": "Admin"}
        users.append(existing_admin)
        
        for _, row in df.iterrows():
            if pd.isna(row[cccd_col]) or pd.isna(row[role_col]): continue
            
            raw_cccd = str(row[cccd_col]).split('.')[0].strip()
            if raw_cccd.lower() == 'admin': continue # Managed manually above
            
            raw_pass = str(row[pass_col]).split('.')[0].strip() if not pd.isna(row[pass_col]) else "123456"
            hashed_pass = generate_password_hash(raw_pass)
            users.append({
                "cccd": raw_cccd,
                "ten": str(row[name_col]).strip() if not pd.isna(row[name_col]) else "",
                "password": hashed_pass,
                "role": "Admin" if "admin" in str(row[role_col]).lower() else "Nhân viên"
            })
        
        save_users(users)
        return jsonify({"success": True, "message": f"Đã cập nhật {len(users)-1} người dùng"})
        
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/api/change-password', methods=['POST'])
def change_password():
    data = request.json
    cccd = str(data.get('cccd'))
    old_password = str(data.get('old_password'))
    new_password = str(data.get('new_password'))
    
    users = get_users()
    user_found = False
    
    for i, user in enumerate(users):
        if str(user.get('cccd')) == cccd:
            user_found = True
            stored_pass = str(user.get('password'))
            # Check old password
            is_valid = False
            if stored_pass.startswith('scrypt:') or stored_pass.startswith('pbkdf2:'):
                is_valid = check_password_hash(stored_pass, old_password)
            else:
                is_valid = (stored_pass == old_password)
                
            if not is_valid:
                return jsonify({"success": False, "message": "Mật khẩu cũ không chính xác."})
                
            # Update password
            users[i]['password'] = generate_password_hash(new_password)
            break
            
    if not user_found:
        return jsonify({"success": False, "message": "Không tìm thấy người dùng."})
        
    save_users(users)
    return jsonify({"success": True, "message": "Đổi mật khẩu thành công!"})

@app.route('/api/sync', methods=['POST'])
def sync_data():
    try:
        # 1. Download Google Sheet Data
        req = urllib.request.Request(GSHEET_URL)
        with urllib.request.urlopen(req) as response:
            html = response.read().decode('utf-8')
            df_gsheet = pd.read_csv(io.StringIO(html))
            
        # 2. Read local excel (just the DuLieu sheet) to find existing Dấu thời gian
        # Use pandas just to get existing keys
        df_local = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='DuLieu')
        existing_timestamps = set(df_local['Dấu thời gian'].astype(str).tolist())
        
        # 3. Find newly added rows
        new_rows = []
        for _, row in df_gsheet.iterrows():
            ts = str(row.get('Dấu thời gian'))
            if ts not in existing_timestamps and ts != 'nan' and ts != 'Tổng':
                new_rows.append(row)
                
        if len(new_rows) == 0:
            return jsonify({"success": True, "message": "Không có dữ liệu mới để đồng bộ."})
            
        # 4. Append to local Excel using openpyxl to not corrupt formulas
        wb = openpyxl.load_workbook(LOCAL_EXCEL_PATH, data_only=False)
        ws = wb['DuLieu']
        
        # Determine the GS columns mapping. Assuming identical order.
        # GS CSV has the string values. Let's write array directly.
        for r in new_rows:
            # list of values in the row
            row_vals = r.tolist()
            ws.append(row_vals)
            
        wb.save(LOCAL_EXCEL_PATH)
        
        return jsonify({"success": True, "message": f"Đã chèn thêm {len(new_rows)} dòng mới thành công."})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/api/stats', methods=['GET'])
def get_stats():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    try:
        df = pd.read_excel(LOCAL_EXCEL_PATH, sheet_name='DuLieu')
        # Loại bỏ dòng Tổng
        df = df[df['Họ và tên Nhân Viên'].astype(str) != 'Tổng']
        
        # Tiền xử lý ngày
        if 'Ngày báo Cáo' in df.columns:
            df['Ngày_parsed'] = pd.to_datetime(df['Ngày báo Cáo'], errors='coerce')
            
            if start_date:
                start_dt = pd.to_datetime(start_date, errors='coerce')
                if not pd.isnull(start_dt):
                    df = df[df['Ngày_parsed'] >= start_dt]
                    
            if end_date:
                end_dt = pd.to_datetime(end_date, errors='coerce')
                if not pd.isnull(end_dt):
                    # To include entire end day
                    df = df[df['Ngày_parsed'] <= end_dt]
                    
        # Gom nhóm theo nhân viên, tính tổng các dịch vụ
        cols_to_sum = [
            'Bán Fiber', 'Bán MyTV', 'Bán Camera_Mesh', 'Ngưng Fiber', 
            'Ngưng Mytv', 'Chuyển ONT 2B', 'Chuyển XGSPON', 'GHTT tháng T',
            'GHTT tháng T+1', 'Lắp đặt/Dịch chuyển', 'Thu hồi ONT 2B', 'Thu hồi Mesh',
            'Xử Lý Suy Hao', 'Sửa chữa', 'B2A', 'Tiền Suy Hao'
        ]
        
        # Filter matching columns
        cols = [c for c in cols_to_sum if c in df.columns]
        
        for c in cols:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        grouped = df.groupby('Họ và tên Nhân Viên')[cols].sum().reset_index()
        grouped['Tổng điểm'] = grouped[cols].sum(axis=1) # Pseudo total
        
        # Sort by total
        grouped = grouped.sort_values(by='Tổng điểm', ascending=False)
        
        records = grouped.to_dict(orient='records')
        return jsonify({"success": True, "data": records, "columns": cols})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
