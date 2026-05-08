from flask import Flask, request, send_file, render_template_string, jsonify
import math
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ════════════════════════════════
# 核心 Excel 樣式設定
# ════════════════════════════════
FONT_NAME = "新細明體"
FILLS = {
    "A++": None,
    "A+":  PatternFill("solid", fgColor="E6E6E6"),
    "A":   PatternFill("solid", fgColor="BFBFBF"),
    "B++": PatternFill("solid", fgColor="808080"),
    "":    PatternFill("solid", fgColor="808080"),
}

def _med(): return Side(style="medium", color="000000")
def _thn(): return Side(style="thin",   color="000000")
def all_thin():
    s = _thn()
    return Border(left=s, right=s, top=s, bottom=s)

def outer_med(r, c, r1, c1, r2, c2):
    return Border(
        left   = _med() if c == c1 else _thn(),
        right  = _med() if c == c2 else _thn(),
        top    = _med() if r == r1 else _thn(),
        bottom = _med() if r == r2 else _thn(),
    )

def sc(ws, row, col, value, bold=False, size=10, fill=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name=FONT_NAME, bold=bold, size=size)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if border: cell.border = border
    if fill: cell.fill = fill
    return cell

# ---------------------------------------------------------
# 1. 讀取邏輯：抓取 P 欄 (Index 15)
# ---------------------------------------------------------
def read_students_initial(file_stream):
    try:
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        students = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 16: continue
            
            # O 欄姓名 (Index 14)，P 欄答對數 (Index 15)
            name = str(row[14]).strip() if row[14] is not None else "" 
            student_id = str(row[4]).strip() if row[4] is not None else "" 
            
            if name in ["", "預設標準答案", "None"]: continue
                
            try:
                # 修改處：改抓 P 欄
                x_val = float(row[15]) if row[15] is not None else 0.0
                students.append({
                    "id": student_id, 
                    "name": name, 
                    "x": x_val,
                    "y": 0 
                })
            except (ValueError, TypeError): continue
        return students
    except Exception as e:
        return []

# ---------------------------------------------------------
# 2. 生成 APP 貼上專用表 (補習班系統用)
# ---------------------------------------------------------
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
            x = float(s.get('x', 0))
            y = float(s.get('y', 0))
            total = (x / 25.0) * 85.0 + (y / 6.0) * 15.0
            ws.cell(row=i, column=2, value=round(total, 2) if total > 0 else "")
        else:
            ws.cell(row=i, column=2, value="")
            
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="App_Score_Import.xlsx")

# ---------------------------------------------------------
# 3. 正式成績報表繪製 (優化欄位寬度)
# ---------------------------------------------------------
def build_excel(students_data, exam_lines, ths):
    students_for_render = []
    for s in students_data:
        x = float(s.get('x', 0))
        y = float(s.get('y', 0))
        total = (x / 25.0) * 85.0 + (y / 6.0) * 15.0
        students_for_render.append((s.get('name', ''), x, y, round(total, 2)))

    sorted_s = sorted(students_for_render, key=lambda x: -x[3])
    n = len(sorted_s)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for s in sorted_s:
        t = s[3]
        if t >= ths['th_app']: counts["A++"] += 1
        elif t >= ths['th_ap']: counts["A+"] += 1
        elif t >= ths['th_a']: counts["A"] += 1
        elif t >= ths['th_bpp']: counts["B++"] += 1

    rows_per_block = math.ceil((n + 2) / 3) if n > 0 else 1
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

    # --- 寬度優化：縮短數據欄位 ---
    for b in range(3):
        base = b * 4 + 1
        ws.column_dimensions[get_column_letter(base)].width = 9       # 姓名 (A, E, I)
        ws.column_dimensions[get_column_letter(base + 1)].width = 5   # 選擇 (B, F, J)
        ws.column_dimensions[get_column_letter(base + 2)].width = 5   # 非選 (C, G, K)
        ws.column_dimensions[get_column_letter(base + 3)].width = 5.5 # 總分 (D, H, L)
    
    ws.column_dimensions["M"].width = 0.4
    for col_let in ["N", "O", "P"]: ws.column_dimensions[col_let].width = 7

    # 標題列
    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, 1, base + i, h, border=all_thin())

    # 資料填入
    for idx, (name, sel, nonsel, total) in enumerate(sorted_s):
        b, r = idx // rows_per_block, 2 + (idx % rows_per_block)
        col = b * 4 + 1
        g = ""
        if total >= ths['th_app']: g = "A++"
        elif total >= ths['th_ap']: g = "A+"
        elif total >= ths['th_a']: g = "A"
        elif total >= ths['th_bpp']: g = "B++"
        f = FILLS.get(g)
        for i, val in enumerate([name, sel, nonsel, total]):
            sc(ws, r, col + i, val, fill=f, border=all_thin())

    # 平均區塊
    if n > 0:
        avg_pos = n
        b_avg, r_avg_start = avg_pos // rows_per_block, 2 + (avg_pos % rows_per_block)
        col_avg_base = b_avg * 4 + 1
        avg_vals = ["平均", 
                    round(sum(s[1] for s in students_for_render)/n, 2), 
                    round(sum(s[2] for s in students_for_render)/n, 2), 
                    round(sum(s[3] for s in students_for_render)/n, 2)]
        
        final_row = 2 + rows_per_block - 1
        for i, val in enumerate(avg_vals):
            curr_col = col_avg_base + i
            for fill_r in range(r_avg_start, final_row + 1):
                sc(ws, fill_r, curr_col, "", border=all_thin())
            if final_row > r_avg_start:
                ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=final_row, end_column=curr_col)
            sc(ws, r_avg_start, curr_col, val, bold=True, border=all_thin())

    # 右側大型標題框
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font, tc.alignment = Font(name=FONT_NAME, bold=True, size=18), Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    # 右側等級統計
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
    <title>S R G - Grade Report</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            background-color: #020617;
            background-image: 
                radial-gradient(at 0% 0%, #1e1b4b 0, transparent 50%), 
                radial-gradient(at 100% 0%, #312e81 0, transparent 50%),
                radial-gradient(at 50% 100%, #0f172a 0, transparent 50%);
            background-attachment: fixed;
            min-height: 100vh;
            color: #f8fafc;
            font-family: 'Inter', sans-serif;
        }
        .glass-card {
            background: rgba(15, 23, 42, 0.6);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        .btn-main {
            background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
            transition: all 0.3s ease;
        }
        .btn-main:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 0 20px rgba(168, 85, 247, 0.4); }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #334155; border-radius: 10px; }
    </style>
</head>
<body class="p-6 md:p-12 flex justify-center">
    <div class="max-w-4xl w-full space-y-8">
        <header class="text-center">
            <h1 class="text-4xl font-black tracking-tighter bg-gradient-to-r from-indigo-400 to-pink-400 bg-clip-text text-transparent">
                SCORE REPORT GENERATOR
            </h1>
        </header>

        <div class="glass-card rounded-2xl p-8 space-y-6 shadow-2xl">
            <form id="main-form" action="/generate" method="post" class="space-y-6">
                <input type="hidden" id="students-json" name="students_json" value="">
                
                <div>
                    <label class="text-[10px] uppercase tracking-widest text-slate-500 font-bold mb-2 block">考試名稱</label>
                    <input type="text" name="exam_name" placeholder="例如：113學年度 第一次 模擬考" class="w-full bg-slate-950/50 border border-slate-800 rounded-lg px-4 py-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all">
                </div>

                <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div class="space-y-1 text-center">
                        <label class="text-[10px] text-slate-500 font-bold">A++</label>
                        <input type="number" step="0.1" id="th_app" name="th_app" value="93.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1 text-center">
                        <label class="text-[10px] text-slate-500 font-bold">A+</label>
                        <input type="number" step="0.1" id="th_ap" name="th_ap" value="85.7" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1 text-center">
                        <label class="text-[10px] text-slate-500 font-bold">A</label>
                        <input type="number" step="0.1" id="th_a" name="th_a" value="76.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1 text-center">
                        <label class="text-[10px] text-slate-500 font-bold">B++</label>
                        <input type="number" step="0.1" id="th_bpp" name="th_bpp" value="67.1" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div id="dropzone" class="border-2 border-dashed border-slate-800 rounded-xl p-8 text-center hover:border-indigo-500 transition-colors cursor-pointer group">
                        <input type="file" id="file-input" accept=".xlsx, .xlsm" class="hidden">
                        <div class="text-2xl mb-1 group-hover:scale-110 transition-transform">📁</div>
                        <p class="text-xs text-slate-500" id="file-status">讀卡機檔案 (.xlsx/.xlsm)</p>
                    </div>
                    <div class="border border-slate-800 rounded-xl p-4 bg-slate-900/30 space-y-3">
                        <p class="text-[10px] text-indigo-400 font-bold uppercase">手動新增</p>
                        <div class="flex gap-2">
                            <input type="text" id="manual-name" placeholder="姓名" class="flex-1 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-sm outline-none">
                            <input type="number" id="manual-x" placeholder="選擇" class="w-16 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-sm text-center outline-none">
                            <button type="button" onclick="addManualStudent()" class="bg-indigo-600 hover:bg-indigo-500 px-3 py-1 rounded text-xs font-bold transition-colors">新增</button>
                        </div>
                    </div>
                </div>

                <div class="space-y-2">
                    <label class="text-[10px] uppercase tracking-widest text-indigo-400 font-bold block">1. 貼入補習班 APP 名單順序</label>
                    <textarea id="ordered_names" name="ordered_names" rows="4" placeholder="每行一個姓名..." class="w-full bg-slate-950/50 border border-slate-800 rounded-lg px-4 py-3 text-sm focus:ring-2 focus:ring-indigo-500 outline-none custom-scrollbar"></textarea>
                </div>

                <div id="success-bar" class="hidden bg-indigo-500/10 border border-indigo-500/20 py-3 px-4 rounded-lg flex items-center justify-between">
                    <span class="text-xs text-indigo-300 font-medium">✨ 已載入：<span id="st-count">0</span> 位成員</span>
                </div>

                <div id="students-scores-container" class="hidden space-y-4">
                    <h3 class="text-sm font-bold text-indigo-400 border-b border-slate-800 pb-2">2. 登錄非選分數 (滿分 6)</h3>
                    <div id="students-list" class="max-h-64 overflow-y-auto pr-1 space-y-2 custom-scrollbar"></div>
                </div>

                <div class="flex flex-col md:flex-row gap-4 pt-4">
                    <button type="submit" id="submit-btn" class="btn-main flex-[2] py-4 rounded-xl font-bold text-white shadow-xl opacity-50 cursor-not-allowed" disabled>下載正式報表</button>
                    <button type="button" id="copy-btn" onclick="downloadCopyList()" class="flex-1 bg-slate-800 border border-slate-700 hover:bg-slate-700 text-slate-200 py-4 rounded-xl font-bold opacity-50 cursor-not-allowed" disabled>生成 APP 貼上表</button>
                </div>
            </form>
        </div>

        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-indigo-400 font-bold uppercase">A++</p><h2 id="sum-app" class="text-3xl font-black mt-1">--</h2></div>
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-emerald-400 font-bold uppercase">A+</p><h2 id="sum-ap" class="text-3xl font-black mt-1">--</h2></div>
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-amber-400 font-bold uppercase">A</p><h2 id="sum-a" class="text-3xl font-black mt-1">--</h2></div>
            <div class="glass-card rounded-2xl p-6 text-center"><p class="text-[10px] text-pink-400 font-bold uppercase">B++</p><h2 id="sum-bpp" class="text-3xl font-black mt-1">--</h2></div>
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
            const res = await fetch('/upload_read', { method: 'POST', body: formData });
            const d = await res.json();
            if (d.students) {
                d.students.forEach(newSt => {
                    if(!studentsData.some(s => s.name === newSt.name)) studentsData.push(newSt);
                });
                refreshUI();
            }
        };

        function addManualStudent() {
            const n = document.getElementById('manual-name'), x = document.getElementById('manual-x');
            if (!n.value.trim()) return;
            studentsData.push({ name: n.value.trim(), x: parseFloat(x.value) || 0, y: 0 });
            n.value = ''; x.value = ''; refreshUI();
        }

        function refreshUI() {
            const hasData = studentsData.length > 0;
            document.getElementById('submit-btn').disabled = !hasData;
            document.getElementById('copy-btn').disabled = !hasData;
            document.getElementById('submit-btn').classList.toggle('opacity-50', !hasData);
            document.getElementById('copy-btn').classList.toggle('opacity-50', !hasData);
            document.getElementById('st-count').innerText = studentsData.length;
            document.getElementById('success-bar').classList.toggle('hidden', !hasData);
            renderList(); updateDashboard();
        }

        function renderList() {
            const list = document.getElementById('students-list');
            list.innerHTML = '';
            studentsData.forEach((s, i) => {
                const row = document.createElement('div');
                row.className = "flex items-center justify-between bg-slate-900/50 p-3 rounded-lg border border-slate-800";
                row.innerHTML = `<span class="text-sm font-medium">${s.name} (選: ${s.x})</span>
                    <input type="number" step="0.5" value="${s.y}" onchange="updateY(${i}, this.value)" class="w-16 bg-slate-950 border border-slate-700 rounded p-1 text-center text-xs text-emerald-400 outline-none">`;
                list.appendChild(row);
            });
            document.getElementById('students-scores-container').classList.remove('hidden');
        }

        function updateY(i, v) { studentsData[i].y = parseFloat(v) || 0; updateDashboard(); }

        async function updateDashboard() {
            const payload = {
                students: studentsData,
                th_app: parseFloat(document.getElementById('th_app').value),
                th_ap: parseFloat(document.getElementById('th_ap').value),
                th_a: parseFloat(document.getElementById('th_a').value),
                th_bpp: parseFloat(document.getElementById('th_bpp').value)
            };
            const res = await fetch('/analyze_full', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
            const d = await res.json();
            ['app', 'ap', 'a', 'bpp'].forEach(k => document.getElementById(`sum-${k}`).innerText = d.counts[k.toUpperCase().replace('PP', '++').replace('P', '+')]);
            document.getElementById('students-json').value = JSON.stringify(studentsData);
        }

        function downloadCopyList() {
            const f = document.getElementById('main-form');
            const old = f.action; f.action = '/generate_copy_list'; f.submit();
            setTimeout(() => f.action = old, 500);
        }
    </script>
</body>
</html>'''

# ════════════════════════════════
# Flask 路由
# ════════════════════════════════
@app.route('/')
def index(): return render_template_string(HTML_TEMPLATE)

@app.route('/upload_read', methods=['POST'])
def upload_read():
    file = request.files.get('file')
    return jsonify({"students": read_students_initial(io.BytesIO(file.read()))}) if file else jsonify({"students": []})

@app.route('/analyze_full', methods=['POST'])
def analyze_full():
    data = request.get_json() or {}
    students = data.get('students', [])
    ths = {k: float(data.get(k, 0)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for s in students:
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
    buf = build_excel(students, lines, ths)
    return send_file(buf, as_attachment=True, download_name=f"{exam_name}.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
