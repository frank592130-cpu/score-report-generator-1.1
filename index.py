from flask import Flask, request, send_file, render_template_string, jsonify
import math
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ════════════════════════════════
# 核心 Excel 繪製邏輯
# ════════════════════════════════
FONT_NAME = "新細明體"
FILLS = {
    "A++": None,
    "A+":  PatternFill("solid", fgColor="E6E6E6"),
    "A":   PatternFill("solid", fgColor="BFBFBF"),
    "B++": PatternFill("solid", fgColor="808080"),
    "":    PatternFill("solid", fgColor="808080"),
}

def _med(): return Side(style="medium", color="000000")
def _thn(): return Side(style="thin",   color="000000")
def all_thin():
    s = _thn()
    return Border(left=s, right=s, top=s, bottom=s)

def outer_med(r, c, r1, c1, r2, c2):
    return Border(
        left   = _med() if c == c1 else _thn(),
        right  = _med() if c == c2 else _thn(),
        top    = _med() if r == r1 else _thn(),
        bottom = _med() if r == r2 else _thn(),
    )

def sc(ws, row, col, value, bold=False, size=10, fill=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name=FONT_NAME, bold=bold, size=size)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if border: cell.border = border
    if fill: cell.fill = fill
    return cell

def read_students_initial(file_stream):
    try:
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        students = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 10: continue
            name = str(row[5]).strip() if row[5] is not None else "" 
            student_id = str(row[4]).strip() if row[4] is not None else "" 
            if name in ["", "預設標準答案", "None"]: continue
            try:
                x_val = float(row[9]) if row[9] is not None else 0.0
                students.append({"id": student_id, "name": name, "x": x_val, "y": 0, "is_leave": False})
            except (ValueError, TypeError): continue
        return students
    except Exception as e:
        print(f"Error reading file: {e}")
        return []

def build_excel(students_data, exam_lines, ths):
    normal_list = []
    leave_list = []
    for s in students_data:
        if s.get('is_leave', False):
            leave_list.append({'name': s.get('name', ''), 'is_leave': True})
        else:
            x, y = float(s.get('x', 0)), float(s.get('y', 0))
            total = (x / 25.0) * 85.0 + (y / 6.0) * 15.0
            normal_list.append({'name': s.get('name', ''), 'x': x, 'y': y, 'total': round(total, 2), 'is_leave': False})
    
    sorted_normal = sorted(normal_list, key=lambda x: -x['total'])
    final_students = sorted_normal + leave_list
    n_total, n_normal = len(final_students), len(sorted_normal)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for s in sorted_normal:
        t = s['total']
        if t >= ths['th_app']: counts["A++"] += 1
        elif t >= ths['th_ap']: counts["A+"] += 1
        elif t >= ths['th_a']: counts["A"] += 1
        elif t >= ths['th_bpp']: counts["B++"] += 1

    rows_per_block = math.ceil((n_total + 2) / 3) if n_total > 0 else 1
    HEADER_ROW, DATA_START = 1, 2
    FINAL_ROW = DATA_START + rows_per_block - 1
    
    wb = openpyxl.Workbook()
    ws = wb.active
    
    for b in range(3):
        base = b * 4 + 1
        for offset, w in enumerate([9, 7.5, 7.5, 7.5]):
            ws.column_dimensions[get_column_letter(base + offset)].width = w
    ws.column_dimensions["M"].width = 0.4
    for col_let in ["N", "O", "P"]: ws.column_dimensions[col_let].width = 7

    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, HEADER_ROW, base + i, h, border=all_thin())

    for idx, s in enumerate(final_students):
        b, r = idx // rows_per_block, DATA_START + (idx % rows_per_block)
        col = b * 4 + 1
        if s['is_leave']:
            sc(ws, r, col, s['name'], border=all_thin())
            ws.merge_cells(start_row=r, start_column=col+1, end_row=r, end_column=col+3)
            sc(ws, r, col+1, "假", border=all_thin())
            for i in range(1, 4): ws.cell(row=r, column=col+i).border = all_thin()
        else:
            g = ""
            total = s['total']
            if total >= ths['th_app']: g = "A++"
            elif total >= ths['th_ap']: g = "A+"
            elif total >= ths['th_a']: g = "A"
            elif total >= ths['th_bpp']: g = "B++"
            f = FILLS.get(g)
            vals = [s['name'], s['x'], s['y'], total]
            for i, val in enumerate(vals): sc(ws, r, col + i, val, fill=f, border=all_thin())

    if n_total > 0:
        avg_pos = n_total
        b_avg, r_avg_start = avg_pos // rows_per_block, DATA_START + (avg_pos % rows_per_block)
        col_avg_base = b_avg * 4 + 1
        if n_normal > 0:
            avg_vals = ["平均", round(sum(s['x'] for s in sorted_normal)/n_normal, 2), round(sum(s['y'] for s in sorted_normal)/n_normal, 2), round(sum(s['total'] for s in sorted_normal)/n_normal, 2)]
        else:
            avg_vals = ["平均", 0, 0, 0]
        for i, val in enumerate(avg_vals):
            curr_col = col_avg_base + i
            for fill_r in range(r_avg_start, FINAL_ROW + 1): sc(ws, fill_r, curr_col, "", border=all_thin())
            if FINAL_ROW > r_avg_start: ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=FINAL_ROW, end_column=curr_col)
            sc(ws, r_avg_start, curr_col, val, bold=True, border=all_thin())

    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font, tc.alignment = Font(name=FONT_NAME, bold=True, size=18), Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1): ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)
    
    visible_grades = [("A++", counts["A++"]), ("A+", counts["A+"]), ("A", counts["A"]), ("B++", counts["B++"])]
    GRADE_R1 = TITLE_R2 + 4 
    for i, (g, cnt) in enumerate(visible_grades):
        row = GRADE_R1 + i
        f = FILLS.get(g)
        for col, val in [(14, g), (15, cnt)]:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font, cell.alignment = Font(name=FONT_NAME, bold=True, size=11), Alignment(horizontal="center", vertical="center")
            cell.border = outer_med(row, col, GRADE_R1, 14, GRADE_R1 + 3, 15)
            if f: cell.fill = f
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ════════════════════════════════
# UI 模板 
# ════════════════════════════════
HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TIME SAVER</title>
    
    <link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>⚡</text></svg>">
    
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            background-color: #020617;
            background-image: radial-gradient(at 0% 0%, #1e1b4b 0, transparent 50%), radial-gradient(at 100% 0%, #312e81 0, transparent 50%), radial-gradient(at 50% 100%, #0f172a 0, transparent 50%);
            background-attachment: fixed; min-height: 100vh; color: #f8fafc; font-family: system-ui, -apple-system, sans-serif;
        }
        .glass-card { background: rgba(15, 23, 42, 0.6); backdrop-filter: blur(12px); border: 1px solid rgba(255, 255, 255, 0.1); }
        .btn-main { background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%); transition: all 0.3s; }
        .btn-main:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 0 20px rgba(168, 85, 247, 0.4); }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #334155; border-radius: 10px; }
    </style>
</head>
<body class="p-4 md:p-12 flex justify-center">
    <div class="max-w-4xl w-full space-y-8">
        <header class="text-center"><h1 class="text-4xl font-extrabold tracking-tighter bg-gradient-to-r from-indigo-400 via-purple-400 to-pink-400 bg-clip-text text-transparent">TIME SAVER</h1></header>

        <div class="glass-card rounded-2xl p-8 space-y-6 shadow-2xl">
            <form id="main-form" action="/generate" method="post" class="space-y-6">
                <input type="hidden" id="students-json" name="students_json" value="">
                
                <div>
                    <label class="text-[10px] uppercase tracking-[0.2em] text-slate-500 font-bold mb-2 block">考試名稱</label>
                    <input type="text" name="exam_name" placeholder="例如：國三 金安模擬考 第一回" class="w-full bg-slate-950/50 border border-slate-800 rounded-lg px-4 py-3 focus:ring-2 focus:ring-indigo-500 outline-none">
                </div>

                <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-center">
                    <div class="space-y-1"><label class="text-[10px] text-slate-500 font-bold">A++ 門檻</label><input type="number" step="0.1" inputmode="decimal" id="th_app" name="th_app" value="93.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center text-indigo-400"></div>
                    <div class="space-y-1"><label class="text-[10px] text-slate-500 font-bold">A+ 門檻</label><input type="number" step="0.1" inputmode="decimal" id="th_ap" name="th_ap" value="85.7" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center text-emerald-400"></div>
                    <div class="space-y-1"><label class="text-[10px] text-slate-500 font-bold">A 門檻</label><input type="number" step="0.1" inputmode="decimal" id="th_a" name="th_a" value="76.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center text-amber-400"></div>
                    <div class="space-y-1"><label class="text-[10px] text-slate-500 font-bold">B++ 門檻</label><input type="number" step="0.1" inputmode="decimal" id="th_bpp" name="th_bpp" value="67.1" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center text-pink-400"></div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div id="dropzone" class="border-2 border-dashed border-slate-800 rounded-xl p-8 text-center hover:border-indigo-500 cursor-pointer">
                        <input type="file" id="file-input" accept=".xlsx, .xlsm" class="hidden">
                        <div class="text-2xl mb-1">📁</div><p class="text-xs text-slate-500">讀卡機 XLSX 或 XLSM 檔案</p>
                    </div>
                    <div class="border border-slate-800 rounded-xl p-4 bg-slate-900/30 space-y-3">
                        <p class="text-[10px] text-indigo-400 font-bold uppercase">手動新增</p>
                        <div class="flex flex-wrap gap-2 items-center">
                            <input type="text" id="manual-name" placeholder="姓名" class="flex-1 min-w-[80px] bg-slate-950 border border-slate-700 rounded px-2 py-1 text-sm outline-none">
                            <input type="number" id="manual-x" inputmode="decimal" placeholder="選擇" class="w-16 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-sm text-center">
                            <label class="flex items-center gap-1 text-xs text-slate-400 cursor-pointer">
                                <input type="checkbox" id="manual-leave" class="accent-indigo-500"> 請假
                            </label>
                            <button type="button" onclick="addManualStudent()" class="bg-indigo-600 hover:bg-indigo-500 px-3 py-1 rounded text-xs font-bold transition-colors">新增</button>
                        </div>
                    </div>
                </div>

                <div id="success-bar" class="hidden bg-indigo-500/10 border border-indigo-500/20 py-3 px-4 rounded-lg flex items-center justify-between">
                    <span class="text-xs text-indigo-300 font-medium">✨ 載入成功：<span id="st-count">0</span> 位學生</span>
                    <span class="text-[10px] px-2 py-0.5 bg-indigo-500/20 rounded-full text-indigo-400 font-bold uppercase">Ready</span>
                </div>

                <div id="students-scores-container" class="hidden space-y-4">
                    <div class="flex justify-between items-center border-b border-slate-800 pb-2">
                        <h3 class="text-sm font-bold text-indigo-400">登錄非選分數 (滿分 6)</h3>
                        <span class="text-[10px] text-slate-500 italic">總分 = (X/25)*85 + (Y/6)*15</span>
                    </div>
                    <div id="students-list" class="max-h-64 overflow-y-auto pr-1 space-y-2 custom-scrollbar"></div>
                </div>

                <div class="space-y-2">
                    <label class="text-[10px] uppercase tracking-[0.2em] text-indigo-400 font-bold block">名單順序 (每人一行)</label>
                    <textarea id="ordered_names" name="ordered_names" rows="5" placeholder="請直接貼入補習班系統的名單順序..." class="w-full bg-slate-950/50 border border-slate-800 rounded-lg px-4 py-3 text-sm focus:ring-2 focus:ring-indigo-500 outline-none custom-scrollbar"></textarea>
                </div>

                <div class="flex flex-col md:flex-row gap-4 pt-4">
                    <button type="submit" id="submit-btn" class="btn-main flex-[2] py-4 rounded-xl font-bold text-white shadow-xl opacity-50 cursor-not-allowed" disabled>下載正式成績報表 (.xlsx)</button>
                    <button type="button" id="copy-btn" onclick="downloadCopyList()" class="flex-1 bg-slate-800 border border-slate-700 hover:bg-slate-700 text-slate-200 py-4 rounded-xl font-bold opacity-50 cursor-not-allowed" disabled>生成 APP 貼上表</button>
                </div>
            </form>
        </div>

        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-indigo-400 font-bold">A++</p><h2 id="sum-app" class="text-3xl font-black">--</h2></div>
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-emerald-400 font-bold">A+</p><h2 id="sum-ap" class="text-3xl font-black">--</h2></div>
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-amber-400 font-bold">A</p><h2 id="sum-a" class="text-3xl font-black">--</h2></div>
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-pink-400 font-bold">B++</p><h2 id="sum-bpp" class="text-3xl font-black">--</h2></div>
        </div>
    </div>

    <script>
        const fileInput = document.getElementById('file-input');
        const dropzone = document.getElementById('dropzone');
        let studentsData = []; 

        dropzone.onclick = () => fileInput.click();
        fileInput.onchange = async () => {
            const file = fileInput.files[0];
            if (!file) return;
            const formData = new FormData();
            formData.append('file', file);
            try {
                const res = await fetch('/upload_read', { method: 'POST', body: formData });
                const d = await res.json();
                if (d.students && d.students.length > 0) {
                    d.students.forEach(newSt => {
                        if(!studentsData.some(s => s.name === newSt.name)) {
                            studentsData.push({ name: newSt.name, x: newSt.x, y: 0, is_leave: false });
                        }
                    });
                    refreshUI();
                } else { alert("檔案讀取失敗（確保學生姓名在第 F 欄，選擇分數在第 J 欄）"); }
            } catch (e) { console.error(e); }
        };

        function addManualStudent() {
            const nameInp = document.getElementById('manual-name'), xInp = document.getElementById('manual-x'), leaveInp = document.getElementById('manual-leave');
            const name = nameInp.value.trim(), x = parseFloat(xInp.value) || 0, is_leave = leaveInp.checked;
            if (!name) return;
            studentsData.push({ name, x, y: 0, is_leave });
            nameInp.value = ''; xInp.value = ''; leaveInp.checked = false;
            refreshUI();
        }

        function removeStudent(index) { studentsData.splice(index, 1); refreshUI(); }

        function refreshUI() {
            const hasData = studentsData.length > 0;
            const submitBtn = document.getElementById('submit-btn'), copyBtn = document.getElementById('copy-btn');
            if (hasData) {
                document.getElementById('st-count').innerText = studentsData.length;
                document.getElementById('success-bar').classList.remove('hidden');
                [submitBtn, copyBtn].forEach(btn => { btn.disabled = false; btn.classList.remove('opacity-50', 'cursor-not-allowed'); });
                renderStudentsInputs(); updateDashboard();
            } else {
                document.getElementById('success-bar').classList.add('hidden');
                [submitBtn, copyBtn].forEach(btn => { btn.disabled = true; btn.classList.add('opacity-50', 'cursor-not-allowed'); });
            }
        }

        function renderStudentsInputs() {
            const listDiv = document.getElementById('students-list');
            listDiv.innerHTML = '';
            studentsData.forEach((s, idx) => {
                const row = document.createElement('div');
                row.className = "flex items-center justify-between bg-slate-950/40 border border-slate-800/80 px-4 py-2 rounded-lg gap-4";
                let scoreSection = s.is_leave ? `<span class="text-xs font-bold text-slate-500 bg-slate-800 px-2 py-0.5 rounded">請假</span>` : `<span class="text-[10px] text-slate-500">選擇: <b class="text-indigo-300">${s.x}</b></span>`;
                let inputSection = !s.is_leave ? `<input type="number" inputmode="decimal" min="0" max="6" step="0.5" value="${s.y}" data-idx="${idx}" class="student-y-input w-12 bg-slate-950 border border-slate-700 rounded p-1 text-center text-xs font-bold text-emerald-400 outline-none">` : `<div class="w-12"></div>`;
                row.innerHTML = `<div class="flex-1 flex items-center justify-between"><span class="text-sm font-semibold text-slate-200">${s.name}</span>${scoreSection}</div><div class="flex items-center gap-3">${inputSection}<button type="button" onclick="removeStudent(${idx})" class="text-slate-600 hover:text-red-500 text-lg transition-colors">×</button></div>`;
                listDiv.appendChild(row);
            });
            document.getElementById('students-scores-container').classList.remove('hidden');
            document.querySelectorAll('.student-y-input').forEach(input => {
                input.addEventListener('input', (e) => {
                    const idx = e.target.getAttribute('data-idx');
                    let val = parseFloat(e.target.value) || 0;
                    if (val < 0) val = 0; if (val > 6) val = 6;
                    studentsData[idx].y = val; updateDashboard();
                });
            });
        }

        async function updateDashboard() {
            if (studentsData.length === 0) return;
            const payload = {
                students: studentsData,
                th_app: parseFloat(document.getElementById('th_app').value) || 0,
                th_ap: parseFloat(document.getElementById('th_ap').value) || 0,
                th_a: parseFloat(document.getElementById('th_a').value) || 0,
                th_bpp: parseFloat(document.getElementById('th_bpp').value) || 0
            };
            try {
                const res = await fetch('/analyze_full', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
                const d = await res.json();
                document.getElementById('sum-app').innerText = d.counts["A++"];
                document.getElementById('sum-ap').innerText = d.counts["A+"];
                document.getElementById('sum-a').innerText = d.counts["A"];
                document.getElementById('sum-bpp').innerText = d.counts["B++"];
                document.getElementById('students-json').value = JSON.stringify(studentsData);
            } catch (e) { console.error(e); }
        }

        function downloadCopyList() {
            const names = document.getElementById('ordered_names').value.trim();
            if (!names) { alert("請先貼入 APP 名單順序！"); return; }
            const form = document.getElementById('main-form');
            const originalAction = form.action;
            form.action = '/generate_copy_list'; form.submit();
            setTimeout(() => { form.action = originalAction; }, 500);
        }
        document.querySelectorAll('.th-input').forEach(i => i.onchange = updateDashboard);
    </script>
</body>
</html>'''

@app.route('/')
def index(): return render_template_string(HTML_TEMPLATE)

@app.route('/upload_read', methods=['POST'])
def upload_read():
    file = request.files.get('file')
    if not file: return jsonify({"students": []})
    return jsonify({"students": read_students_initial(io.BytesIO(file.read()))})

@app.route('/analyze_full', methods=['POST'])
def analyze_full():
    data = request.get_json() or {}
    students = data.get('students', [])
    ths = {k: float(data.get(k, 0)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for s in students:
        if s.get('is_leave', False): continue
        total = (float(s.get('x', 0)) / 25.0) * 85.0 + (float(s.get('y', 0)) / 6.0) * 15.0
        if total >= ths['th_app']: counts["A++"] += 1
        elif total >= ths['th_ap']: counts["A+"] += 1
        elif total >= ths['th_a']: counts["A"] += 1
        elif total >= ths['th_bpp']: counts["B++"] += 1
    return jsonify({"counts": counts})

@app.route('/generate', methods=['POST'])
def generate():
    exam_name = request.form.get('exam_name', '成績報表')
    ths = {k: float(request.form.get(k, 0)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    students = json.loads(request.form.get('students_json', '[]'))
    parts = exam_name.strip().split()
    lines = [parts[0], " ".join(parts[1:-1]), parts[-1]] if len(parts) >= 3 else ["", exam_name, ""]
    return send_file(build_excel(students, lines, ths), as_attachment=True, download_name=f"{exam_name}.xlsx")

@app.route('/generate_copy_list', methods=['POST'])
def generate_copy_list():
    students_json = request.form.get('students_json', '[]')
    ordered_names_raw = request.form.get('ordered_names', '')
    students = json.loads(students_json)
    student_map = {s['name']: s for s in students}
    ordered_names = [n.strip() for n in ordered_names_raw.split('\n') if n.strip()]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "補習班貼上專用"
    ws.cell(row=1, column=1, value="APP名單順序")
    ws.cell(row=1, column=2, value="總分 (表現)")
    for i, target_name in enumerate(ordered_names, start=2):
        ws.cell(row=i, column=1, value=target_name)
        if target_name in student_map:
            s = student_map[target_name]
            if s.get('is_leave', False): ws.cell(row=i, column=2, value="假")
            else:
                x, y = float(s.get('x', 0)), float(s.get('y', 0))
                total = (x / 25.0) * 85.0 + (y / 6.0) * 15.0
                ws.cell(row=i, column=2, value=round(total, 2) if total > 0 else "")
        else: ws.cell(row=i, column=2, value="")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="App_Score_Import.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
