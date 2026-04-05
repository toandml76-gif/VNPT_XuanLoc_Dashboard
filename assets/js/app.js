const API_BASE = window.location.hostname === 'localhost' ? 'http://localhost:5000/api' : '/api';

let currentUser = null;
const charts = {}; // Store canvas instances

// DOM Init
const loginView = document.getElementById('loginView');
const dashboardView = document.getElementById('dashboardView');
const loginForm = document.getElementById('loginForm');
const btnLogout = document.getElementById('btnLogout');
const userNameDisplay = document.getElementById('userNameDisplay');
const userRoleBadge = document.getElementById('userRoleBadge');
const menuAdmin = document.getElementById('menuAdmin');
const navLinks = document.querySelectorAll('.nav-links a');
const panels = document.querySelectorAll('.panel');
const toastBox = document.getElementById('toast');

// Utils
function showToast(message, type = 'success') {
    toastBox.textContent = message;
    toastBox.className = `toast show ${type}`;
    setTimeout(() => toastBox.className = 'toast hidden', 3000);
}

function switchView(isLoggedIn) {
    if (isLoggedIn) {
        loginView.classList.remove('active-view');
        dashboardView.classList.add('active-view');
        
        // Setup initial default dates
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('filterDay').value = today;
        document.getElementById('filterCustomStart').value = today;
        document.getElementById('filterCustomEnd').value = today;
        
        const currentMonth = new Date().getMonth() + 1;
        const currentYear = new Date().getFullYear();
        document.getElementById('filterWeekMonth').value = currentMonth;
        document.getElementById('filterWeekYear').value = currentYear;
        document.getElementById('filterMonthMonth').value = currentMonth;
        document.getElementById('filterMonthYear').value = currentYear;
        
    } else {
        dashboardView.classList.remove('active-view');
        loginView.classList.add('active-view');
    }
}

function switchPanel(targetId) {
    navLinks.forEach(l => l.classList.remove('active'));
    panels.forEach(p => p.classList.remove('active-panel'));
    document.querySelector(`[data-target="${targetId}"]`).classList.add('active');
    document.getElementById(targetId).classList.add('active-panel');
    
    // Auto-fetch data on lazy load optionally, but user should click buttons based on filters.
}

navLinks.forEach(link => {
    link.addEventListener('click', (e) => {
        e.preventDefault();
        switchPanel(e.target.getAttribute('data-target'));
    });
});

// Auth
loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const cccd = document.getElementById('cccd').value;
    const password = document.getElementById('password').value;
    try {
        const res = await fetch(`${API_BASE}/login`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ cccd, password })
        });
        const data = await res.json();
        if (data.success) {
            currentUser = data.user;
            showToast('Đăng nhập thành công!');
            userNameDisplay.textContent = `Chào, ${currentUser.ten || currentUser.cccd}`;
            userRoleBadge.textContent = currentUser.role;
            if (currentUser.role === 'Admin') menuAdmin.classList.remove('hidden');
            else menuAdmin.classList.add('hidden');
            
            // Fill account page
            document.getElementById('accName').textContent = currentUser.ten || "Chưa cập nhật";
            document.getElementById('accCCCD').textContent = currentUser.cccd;
            document.getElementById('accRole').textContent = currentUser.role;
            
            switchView(true);
            switchPanel('panel-day');
            window.fetchDayStats(); // trigger default load
        } else {
            showToast(data.message, 'error');
        }
    } catch (err) { showToast('Lỗi kết nối máy chủ', 'error'); }
});

btnLogout.addEventListener('click', () => {
    currentUser = null;
    document.getElementById('cccd').value = '';
    document.getElementById('password').value = '';
    switchView(false);
});

// Render logic
function renderTable(tableId, dataArray) {
    const table = document.getElementById(tableId);
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (dataArray && dataArray.length > 0) {
        // Define tracking columns manually to respect ordering
        const columnsToTrack = [
            'Bán Fiber', 'Bán MyTV', 'Bán Camera_Mesh', 
            'Ngưng Fiber', 'Ngưng Mytv', 'Chuyển ONT 2B', 
            'Chuyển XGSPON', 'GHTT tháng T', 'GHTT tháng T+1',
            'Lắp đặt/Dịch chuyển', 'Thu hồi ONT 2B', 'Thu hồi Mesh', 
            'Xử Lý Suy Hao', 'Sửa chữa', 'B2A', 'Tiền Suy Hao'
        ];
        
        // Calculate Totals
        const totals = {};
        columnsToTrack.forEach(c => totals[c] = 0);
        totals['Tổng điểm'] = 0;
        
        dataArray.forEach(row => {
            columnsToTrack.forEach(c => totals[c] += Number(row[c] || 0));
            totals['Tổng điểm'] += Number(row['Tổng điểm'] || 0);
        });

        // Determine active columns (total > 0)
        const activeColumns = columnsToTrack.filter(col => totals[col] > 0);

        // Render THEAD
        let thHtml = '<tr><th>STT</th><th>Họ và Tên</th>';
        activeColumns.forEach(c => { thHtml += `<th>${c}</th>`; });
        thHtml += '<th>Tổng Điểm</th></tr>';
        thead.innerHTML = thHtml;

        // Render Total Row under Header
        let trTotal = document.createElement('tr');
        trTotal.style.backgroundColor = '#e1f5fe';
        trTotal.style.fontWeight = 'bold';
        let tdHtmlTotal = `<td>-</td><td style="white-space: nowrap; color: var(--primary);">TỔNG CỘNG</td>`;
        activeColumns.forEach(c => { tdHtmlTotal += `<td>${totals[c]}</td>`; });
        tdHtmlTotal += `<td style="color: var(--danger);">${totals['Tổng điểm']}</td>`;
        trTotal.innerHTML = tdHtmlTotal;
        tbody.appendChild(trTotal);

        // Render Data Rows
        dataArray.forEach((row, idx) => {
            let tr = document.createElement('tr');
            let tdHtml = `<td>${idx + 1}</td><td style="font-weight: 600; white-space: nowrap;">${row['Họ và tên Nhân Viên']}</td>`;
            activeColumns.forEach(c => { tdHtml += `<td>${row[c] || 0}</td>`; });
            tdHtml += `<td style="font-weight: bold; color: var(--primary);">${row['Tổng điểm'] || 0}</td>`;
            tr.innerHTML = tdHtml;
            tbody.appendChild(tr);
        });
    }
}

function renderChart(chartId, dataArray, titleStr) {
    const ctx = document.getElementById(chartId).getContext('2d');
    const topData = dataArray.slice(0, 10);
    const labels = topData.map(d => d['Họ và tên Nhân Viên']);
    
    if (charts[chartId]) charts[chartId].destroy();

    charts[chartId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                { label: 'Bán Fiber', data: topData.map(d => d['Bán Fiber'] || 0), backgroundColor: '#1a5f7a', borderRadius: 4 },
                { label: 'Bán MyTV', data: topData.map(d => d['Bán MyTV'] || 0), backgroundColor: '#e74c3c', borderRadius: 4 },
                { label: 'Bán Camera_Mesh', data: topData.map(d => d['Bán Camera_Mesh'] || 0), backgroundColor: '#f39c12', borderRadius: 4 },
                { label: 'Ngưng Fiber', data: topData.map(d => d['Ngưng Fiber'] || 0), backgroundColor: '#8b0000', borderRadius: 4 },
                { label: 'Ngưng Mytv', data: topData.map(d => d['Ngưng Mytv'] || 0), backgroundColor: '#e4c400', borderRadius: 4 },
                { label: 'Chuyển ONT 2B', data: topData.map(d => d['Chuyển ONT 2B'] || 0), backgroundColor: '#9b59b6', borderRadius: 4 },
                { label: 'Chuyển XGSPON', data: topData.map(d => d['Chuyển XGSPON'] || 0), backgroundColor: '#8e44ad', borderRadius: 4 },
                { label: 'GHTT tháng T', data: topData.map(d => d['GHTT tháng T'] || 0), backgroundColor: '#2ecc71', borderRadius: 4 },
                { label: 'GHTT tháng T+1', data: topData.map(d => d['GHTT tháng T+1'] || 0), backgroundColor: '#27ae60', borderRadius: 4 },
                { label: 'Lắp đặt/Dịch chuyển', data: topData.map(d => d['Lắp đặt/Dịch chuyển'] || 0), backgroundColor: '#3498db', borderRadius: 4 },
                { label: 'Thu hồi ONT 2B', data: topData.map(d => d['Thu hồi ONT 2B'] || 0), backgroundColor: '#e67e22', borderRadius: 4 },
                { label: 'Thu hồi Mesh', data: topData.map(d => d['Thu hồi Mesh'] || 0), backgroundColor: '#d35400', borderRadius: 4 },
                { label: 'Xử lý Suy Hao', data: topData.map(d => d['Xử Lý Suy Hao'] || 0), backgroundColor: '#1abc9c', borderRadius: 4 },
                { label: 'Sửa chữa', data: topData.map(d => d['Sửa chữa'] || 0), backgroundColor: '#16a085', borderRadius: 4 },
                { label: 'B2A', data: topData.map(d => d['B2A'] || 0), backgroundColor: '#34495e', borderRadius: 4 },
                { label: 'Tiền Suy Hao', data: topData.map(d => d['Tiền Suy Hao'] || 0), backgroundColor: '#f1c40f', borderRadius: 4 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { title: { display: true, text: titleStr, font: { size: 16 } } }
        }
    });
}

async function doFetchStats(params, titleElId, titleStr, tableId, chartId) {
    try {
        const qs = new URLSearchParams(params).toString();
        const res = await fetch(`${API_BASE}/stats?${qs}`);
        const result = await res.json();
        
        if (result.success) {
            document.getElementById(titleElId).innerText = titleStr;
            renderTable(tableId, result.data);
            renderChart(chartId, result.data, `Top 10 - ${titleStr}`);
        } else {
            showToast(result.message, 'error');
        }
    } catch (err) { showToast("Lỗi lấy dữ liệu", "error"); }
}

// Fetch Handlers
window.exportTableToExcel = async (tableId, type) => {
    const table = document.getElementById(tableId);
    if (!table || table.querySelector('tbody').rows.length === 0) {
        return showToast('Không có dữ liệu để xuất!', 'error');
    }

    const pad = n => n.toString().padStart(2, '0');
    const yy = year => year.toString().slice(-2);

    let reportTitle = "";
    let fileName = "";

    if (type === 'day') {
        const d = document.getElementById('filterDay').value;
        if (!d) return;
        const [y, m, day] = d.split('-');
        reportTitle = `Kết quả thống kê ngày ${pad(day)}/${pad(m)}/${y}`;
        fileName = `BC_${pad(day)}_${pad(m)}_${yy(y)}.xlsx`;
    } else if (type === 'week') {
        const month = parseInt(document.getElementById('filterWeekMonth').value);
        const year = document.getElementById('filterWeekYear').value;
        const week = document.getElementById('filterWeekNum').value;
        
        const daysInMonth = new Date(year, month, 0).getDate();
        let startDay, endDay;
        const w = parseInt(week);
        if (w === 1) { startDay = 1; endDay = 7; }
        else if (w === 2) { startDay = 8; endDay = 14; }
        else if (w === 3) { startDay = 15; endDay = 21; }
        else if (w === 4) { startDay = 22; endDay = daysInMonth; }
        
        reportTitle = `Kết quả thống kê tuần ${week} từ ngày ${pad(startDay)}/${pad(month)}/${year} đến ngày ${pad(endDay)}/${pad(month)}/${year}`;
        fileName = `BC_Tuan${week}_${pad(month)}_${yy(year)}.xlsx`;
    } else if (type === 'month') {
        const month = document.getElementById('filterMonthMonth').value;
        const year = document.getElementById('filterMonthYear').value;
        reportTitle = `Kết quả thống kê tháng ${pad(month)} năm ${year}`;
        fileName = `BC_${pad(month)}_${yy(year)}.xlsx`;
    } else if (type === 'custom') {
        const sd = document.getElementById('filterCustomStart').value;
        const ed = document.getElementById('filterCustomEnd').value;
        if (!sd || !ed) return;
        const [sy, sm, sday] = sd.split('-');
        const [ey, em, eday] = ed.split('-');
        // Example: Kết quả thống kê từ 12/03/2006 đến 02/04/2006
        reportTitle = `Kết quả thống kê từ ${pad(sday)}/${pad(sm)}/${sy} đến ${pad(eday)}/${pad(em)}/${ey}`;
        // BC_dd1_mm1_dd2_mm2_yy. Using ey and em for the yy and base if they differ
        fileName = `BC_${pad(sday)}_${pad(sm)}_${pad(eday)}_${pad(em)}_${yy(ey)}.xlsx`;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('BaoCao');

    // Extract Headers
    const thead = table.querySelector('thead');
    const headers = [];
    thead.querySelectorAll('th').forEach(th => headers.push(th.innerText.trim()));

    // Extract Rows
    const tbody = table.querySelector('tbody');
    const rows = [];
    tbody.querySelectorAll('tr').forEach(tr => {
        const rowData = [];
        tr.querySelectorAll('td').forEach(td => rowData.push(td.innerText.trim()));
        rows.push(rowData);
    });

    // Write and style Title Row (1)
    const titleRow = worksheet.addRow([reportTitle]);
    worksheet.mergeCells(1, 1, 1, headers.length);
    titleRow.height = 30; // giving space
    const titleCell = titleRow.getCell(1);
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Write and style Header Row (2)
    const headerRow = worksheet.addRow(headers);
    headerRow.height = 35; 
    headerRow.eachCell((cell, colNumber) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 12 };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };
    });

    // Write and style Data Rows
    rows.forEach((rowData, index) => {
        const row = worksheet.addRow(rowData);
        // Hàng số 3 trong Excel tương ứng với index === 0 của mảng rows dữ liệu (bởi vì hàng 1 là Title, hàng 2 là Header)
        let isRow3 = (index === 0);

        row.eachCell((cell, colNumber) => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
            
            if (isRow3) {
                cell.font = { bold: true, color: { argb: 'FF0070C0' } };
            }
            
            // Format date correctly if date pattern is matched yyyy-mm-dd
            if (typeof cell.value === 'string' && cell.value.match(/^\d{4}-\d{2}-\d{2}$/)) {
                let [y, m, d] = cell.value.split('-');
                cell.value = `${d}/${m}/${y}`;
            }

            // Parse numbers slightly so Excel understands it
            let valStr = String(cell.value);
            if (!isNaN(parseFloat(valStr)) && isFinite(valStr) && valStr.trim() !== "") {
                cell.value = parseFloat(valStr);
            }
        });
    });

    // Handle columns width
    worksheet.columns.forEach((col, i) => {
        if (i === 0) col.width = 8; // STT
        else if (i === 1) col.width = 25; // Họ tên
        else col.width = 15; // Numeric properties
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), fileName);
};

window.fetchDayStats = () => {
    const d = document.getElementById('filterDay').value;
    if (!d) return showToast('Vui lòng chọn ngày', 'error');
    
    // Convert YYYY-MM-DD to DD/MM/YYYY for text
    const [y, m, day] = d.split('-');
    doFetchStats({ start_date: d, end_date: d }, 'dayTitle', `Thống kê ngày ${day}/${m}/${y}`, 'tableDay', 'chartDay');
};

window.fetchWeekStats = () => {
    const month = parseInt(document.getElementById('filterWeekMonth').value);
    const year = parseInt(document.getElementById('filterWeekYear').value);
    const week = parseInt(document.getElementById('filterWeekNum').value);
    
    if(!year) return showToast('Nhập năm hợp lệ', 'error');
    
    // Logic week
    // Lấy số ngày trong tháng đó
    const daysInMonth = new Date(year, month, 0).getDate();
    
    let startDay, endDay;
    if (week === 1) { startDay = 1; endDay = 7; }
    else if (week === 2) { startDay = 8; endDay = 14; }
    else if (week === 3) { startDay = 15; endDay = 21; }
    else if (week === 4) { startDay = 22; endDay = daysInMonth; }
    
    const pad = n => n.toString().padStart(2, '0');
    const sm = pad(month);
    
    const startObjStr = `${year}-${sm}-${pad(startDay)}`;
    const endObjStr = `${year}-${sm}-${pad(endDay)}`;
    
    const titleStr = `Tuần thứ ${week} từ ngày ${pad(startDay)}/${sm}/${year} đến ngày ${pad(endDay)}/${sm}/${year}`;
    
    doFetchStats({ start_date: startObjStr, end_date: endObjStr }, 'weekTitle', titleStr, 'tableWeek', 'chartWeek');
};

window.fetchMonthStats = () => {
    const month = parseInt(document.getElementById('filterMonthMonth').value);
    const year = parseInt(document.getElementById('filterMonthYear').value);
    if(!year) return showToast('Nhập năm hợp lệ', 'error');
    
    const daysInMonth = new Date(year, month, 0).getDate();
    const sm = month.toString().padStart(2, '0');
    
    const startStr = `${year}-${sm}-01`;
    const endStr = `${year}-${sm}-${daysInMonth}`;
    
    doFetchStats({ start_date: startStr, end_date: endStr }, 'monthTitle', `Kết quả của tháng ${month} năm ${year}`, 'tableMonth', 'chartMonth');
};

window.fetchCustomStats = () => {
    const sd = document.getElementById('filterCustomStart').value;
    const ed = document.getElementById('filterCustomEnd').value;
    if(!sd || !ed) return showToast('Chọn đầy đủ ngày bắt đầu và kết thúc', 'error');
    
    const fmt = d => { let [y,m,day]=d.split('-'); return `${day}/${m}/${y}`; };
    
    doFetchStats({ start_date: sd, end_date: ed }, 'customTitle', `Kết quả từ ngày ${fmt(sd)} đến ngày ${fmt(ed)}`, 'tableCustom', 'chartCustom');
};

// Admin
document.getElementById('uploadUsersForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const file = document.getElementById('fileExcel').files[0];
    if (!file) return;
    const fd = new FormData(); fd.append('file', file);
    try {
        const res = await fetch(`${API_BASE}/upload-users`, { method: 'POST', body: fd });
        const data = await res.json();
        showToast(data.message, data.success ? 'success' : 'error');
    } catch(err) { showToast('Lỗi tải file', 'error'); }
});

document.getElementById('btnSyncGoogle').addEventListener('click', async (e) => {
    const btn = e.target;
    btn.textContent = "Đang đồng bộ..."; btn.disabled = true;
    try {
        const res = await fetch(`${API_BASE}/sync`, { method: 'POST' });
        const data = await res.json();
        showToast(data.message, data.success ? 'success' : 'error');
    } catch(err) { showToast('Lỗi đồng bộ', 'error'); }
    finally { btn.textContent = "Tiến Hành Đồng Bộ"; btn.disabled = false; }
});

// Account / Change Password
document.getElementById('changePasswordForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const old_password = document.getElementById('cpOld').value;
    const new_password = document.getElementById('cpNew').value;
    const confirm_password = document.getElementById('cpConfirm').value;
    
    if (new_password !== confirm_password) {
        return showToast('Mật khẩu mới và Xác nhận mật khẩu mới không khớp!', 'error');
    }
    
    try {
        const res = await fetch(`${API_BASE}/change-password`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                cccd: currentUser.cccd,
                old_password,
                new_password
            })
        });
        const data = await res.json();
        showToast(data.message, data.success ? 'success' : 'error');
        if (data.success) {
            e.target.reset(); // if canceled this runs on the other button, but here it submits successfully
        }
    } catch (err) {
        showToast('Lỗi máy chủ khi đổi mật khẩu', 'error');
    }
});
