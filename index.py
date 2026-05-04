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

def read_students_initial(file_stream):
    # data_only=True 會讀取公式計算後的數值，而非公式本身
    # keep_vba=False 讀取時不處理巨集，這能增加穩定性
    try:
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        students = []
        
        # 從第 2 列開始讀取 (跳過標題)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 15: continue
            
            # 取得姓名 (第 O 欄，索引 14)
            name = str(row[14]).strip() if row[14] is not None else ""
            
            # 過濾無效資料與標題行
            if name in ["", "預設標準答案", "None", "None"]:
                continue
                
            try:
                # 取得選擇題分數 X (第 H 欄「客觀題」，索引 7)
                x_val = float(row[7]) if row[7] is not None else 0.0
                students.append({"name": name, "x": x_val})
            except (ValueError, TypeError):
                continue
                
        return students
    except Exception as e:
        print(f"讀取 Excel 失敗: {e}")
        return []

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
    HEADER_ROW, DATA_START = 1, 2
    FINAL_ROW = DATA_START + rows_per_block - 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

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

    for idx, (name, sel, nonsel, total) in enumerate(sorted_s):
        b, r = idx // rows_per_block, DATA_START + (idx % rows_per_block)
        col = b * 4 + 1
        g = ""
        if total >= ths['th_app']: g = "A++"
        elif total >= ths['th_ap']: g = "A+"
        elif total >= ths['th_a']: g = "A"
        elif total >= ths['th_bpp']: g = "B++"
        f = FILLS.get(g)
        for i, val in enumerate([name, sel, nonsel, total]):
            sc(ws, r, col + i, val, fill=f, border=all_thin())

    if n > 0:
        avg_pos = n
        b_avg, r_avg_start = avg_pos // rows_per_block, DATA_START + (avg_pos % rows_per_block)
        col_avg_base = b_avg * 4 + 1
        avg_vals = ["平均", 
                    round(sum(s[1] for s in students_for_render)/n, 2), 
                    round(sum(s[2] for s in students_for_render)/n, 2), 
                    round(sum(s[3] for s in students_for_render)/n, 2)]
        
        for i, val in enumerate(avg_vals):
            curr_col = col_avg_base + i
            for fill_r in range(r_avg_start, FINAL_ROW + 1):
                sc(ws, fill_r, curr_col, "", border=all_thin())
            if FINAL_ROW > r_avg_start:
                ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=FINAL_ROW, end_column=curr_col)
            sc(ws, r_avg_start, curr_col, val, bold=True, border=all_thin())

    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font, tc.alignment = Font(name=FONT_NAME, bold=True, size=18), Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

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
# UI 模板 (新增手動輸入功能)
# ════════════════════════════════

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>S R G</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        :root {
            --mesh-color-1: #1e1b4b;
            --mesh-color-2: #312e81;
            --mesh-color-3: #1e293b;
            --mesh-color-4: #020617;
        }
        body {
            background-color: #020617;
            background-image: 
                radial-gradient(at 0% 0%, #1e1b4b 0, transparent 50%), 
                radial-gradient(at 100% 0%, #312e81 0, transparent 50%),
                radial-gradient(at 50% 100%, #0f172a 0, transparent 50%);
            background-attachment: fixed;
            min-height: 100vh;
            margin: 0;
            font-family: 'Inter', sans-serif;
            color: #f8fafc;
        }
        .glass-card {
            background: rgba(15, 23, 42, 0.6);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        }
        .btn-main {
            background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .btn-main:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 0 20px rgba(168, 85, 247, 0.4); }
        .stat-card {
            background: linear-gradient(180deg, rgba(255,255,255,0.05) 0%, rgba(255,255,255,0) 100%);
            border: 1px solid rgba(255,255,255,0.1);
        }
        input::-webkit-outer-spin-button,
        input::-webkit-inner-spin-button {
            -webkit-appearance: none;
            margin: 0;
        }
        input[type=number] {
            -moz-appearance: textfield;
        }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #334155; border-radius: 10px; }
    </style>
</head>
<body class="p-4 md:p-12 flex justify-center">
    <div class="max-w-4xl w-full space-y-8">
        <header class="text-center space-y-2">
            <h1 class="text-4xl font-extrabold tracking-tighter bg-gradient-to-r from-indigo-400 via-purple-400 to-pink-400 bg-clip-text text-transparent">
                SCORE REPORT GENERATOR
            </h1>
        </header>

        <div class="glass-card rounded-2xl p-8 space-y-6">
            <form id="main-form" action="/generate" method="post" class="space-y-6">
                <input type="hidden" id="students-json" name="students_json" value="">
                
                <div>
                    <label class="text-[10px] uppercase tracking-[0.2em] text-slate-500 font-bold mb-2 block">考試名稱</label>
                    <input type="text" name="exam_name" placeholder="" class="w-full bg-slate-950/50 border border-slate-800 rounded-lg px-4 py-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all">
                </div>

                <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">A++ 門檻</label>
                        <input type="number" step="0.1" id="th_app" name="th_app" value="93.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">A+ 門檻</label>
                        <input type="number" step="0.1" id="th_ap" name="th_ap" value="85.7" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">A 門檻</label>
                        <input type="number" step="0.1" id="th_a" name="th_a" value="76.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">B++ 門檻</label>
                        <input type="number" step="0.1" id="th_bpp" name="th_bpp" value="67.1" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <!-- 讀卡機上傳 -->
                    <!-- 找到這一段並替換 -->
                <div id="dropzone" class="...">
                    <input type="file" id="file-input" accept=".xlsx, .xlsm" class="hidden">
                    <div id="upload-content">
                        <div class="text-2xl mb-1 group-hover:scale-110 transition-transform">📁</div>
                        <p class="text-xs text-slate-500" id="file-status">讀卡機 XLSX 或 XLSM 檔案</p>
                    </div>
                </div>

                    <!-- 手動新增學生 -->
                    <div class="border border-slate-800 rounded-xl p-4 bg-slate-900/30 space-y-3">
                        <p class="text-[10px] text-indigo-400 font-bold uppercase tracking-wider">手動新增名單</p>
                        <div class="flex gap-2">
                            <input type="text" id="manual-name" placeholder="姓名" class="flex-1 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-sm outline-none focus:border-indigo-500">
                            <input type="number" id="manual-x" placeholder="X (選擇)" class="w-20 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-sm text-center outline-none focus:border-indigo-500">
                            <button type="button" onclick="addManualStudent()" class="bg-indigo-600 hover:bg-indigo-500 px-3 py-1 rounded text-xs font-bold transition-colors">新增</button>
                        </div>
                    </div>
                </div>

                <div id="success-bar" class="hidden bg-indigo-500/10 border border-indigo-500/20 py-3 px-4 rounded-lg flex items-center justify-between">
                    <span class="text-xs text-indigo-300 font-medium">✨ 目前清單：<span id="st-count">0</span> 位成員</span>
                    <span class="text-[10px] px-2 py-0.5 bg-indigo-500/20 rounded-full text-indigo-400">READY</span>
                </div>

                <div id="students-scores-container" class="hidden space-y-4">
                    <div class="flex justify-between items-center border-b border-slate-800 pb-2">
                        <h3 class="text-sm font-bold text-indigo-400">分數登錄區 (手寫滿分 6 分)</h3>
                        <span class="text-[10px] text-slate-500">總分 = (X/25)*85 + (Y/6)*15</span>
                    </div>
                    <div id="students-list" class="max-h-64 overflow-y-auto pr-1 space-y-2 custom-scrollbar">
                        <!-- 學生項目將在此產生 -->
                    </div>
                </div>

                <button type="submit" id="submit-btn" class="btn-main w-full py-4 rounded-xl font-bold text-white shadow-xl opacity-50 cursor-not-allowed" disabled>
                    下載成績單
                </button>
            </form>
        </div>

        <div class="space-y-4">
            <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                <div class="stat-card rounded-2xl p-6 text-center">
                    <p class="text-[10px] text-indigo-400 font-bold uppercase">A++</p>
                    <h2 id="sum-app" class="text-3xl font-black mt-1">--</h2>
                </div>
                <div class="stat-card rounded-2xl p-6 text-center">
                    <p class="text-[10px] text-emerald-400 font-bold uppercase">A+</p>
                    <h2 id="sum-ap" class="text-3xl font-black mt-1">--</h2>
                </div>
                <div class="stat-card rounded-2xl p-6 text-center">
                    <p class="text-[10px] text-amber-400 font-bold uppercase">A</p>
                    <h2 id="sum-a" class="text-3xl font-black mt-1">--</h2>
                </div>
                <div class="stat-card rounded-2xl p-6 text-center">
                    <p class="text-[10px] text-pink-400 font-bold uppercase">B++</p>
                    <h2 id="sum-bpp" class="text-3xl font-black mt-1">--</h2>
                </div>
            </div>

            <div class="glass-card rounded-xl p-5 flex flex-wrap gap-6 justify-between items-center text-xs font-bold text-slate-400">
                <div class="flex gap-8">
                    <div class="flex flex-col">
                        <span class="text-[8px] text-slate-500">選擇平均 (題數)</span>
                        <span id="sum-sel" class="text-slate-200">--</span>
                    </div>
                    <div class="flex flex-col">
                        <span class="text-[8px] text-slate-500">非選平均 (分數)</span>
                        <span id="sum-nonsel" class="text-slate-200">--</span>
                    </div>
                    <div class="flex flex-col">
                        <span class="text-[8px] text-slate-500">總分平均</span>
                        <span id="sum-total" class="text-slate-200">--</span>
                    </div>
                </div>
                <div class="bg-slate-950/50 px-4 py-2 rounded-lg border border-slate-800">
                    共 <span id="sum-count" class="text-indigo-400">--</span> 位學生
                </div>
            </div>
        </div>
    </div>

    <script>
        const fileInput = document.getElementById('file-input');
        const dropzone = document.getElementById('dropzone');
        let studentsData = []; 

        dropzone.onclick = () => fileInput.click();

        // 處理讀卡機上傳
        fileInput.onchange = async () => {
            const file = fileInput.files[0];
            if (!file) return;

            const formData = new FormData();
            formData.append('file', file);

            try {
                const res = await fetch('/upload_read', { method: 'POST', body: formData });
                const d = await res.json();
                if (d.students && d.students.length > 0) {
                    // 合併新舊資料，以姓名為基準避免重複
                    d.students.forEach(newSt => {
                        if(!studentsData.some(s => s.name === newSt.name)) {
                            studentsData.push({ name: newSt.name, x: newSt.x, y: 0 });
                        }
                    });
                    refreshUI();
                }
            } catch (e) { console.error(e); }
        };

        // 手動新增學生邏輯
        function addManualStudent() {
            const nameInp = document.getElementById('manual-name');
            const xInp = document.getElementById('manual-x');
            const name = nameInp.value.trim();
            const x = parseFloat(xInp.value) || 0;

            if (!name) { alert("請輸入學生姓名"); return; }
            
            if (studentsData.some(s => s.name === name)) {
                if(!confirm("名單中已有此學生，是否要重複新增？")) return;
            }

            studentsData.push({ name: name, x: x, y: 0 });
            
            // 清空輸入
            nameInp.value = '';
            xInp.value = '';
            
            refreshUI();
        }

        // 移除學生
        function removeStudent(index) {
            studentsData.splice(index, 1);
            refreshUI();
        }

        function refreshUI() {
            if (studentsData.length > 0) {
                document.getElementById('st-count').innerText = studentsData.length;
                document.getElementById('success-bar').classList.remove('hidden');
                
                const submitBtn = document.getElementById('submit-btn');
                submitBtn.disabled = false;
                submitBtn.classList.remove('opacity-50', 'cursor-not-allowed');
                
                renderStudentsInputs();
                updateDashboard();
            } else {
                document.getElementById('success-bar').classList.add('hidden');
                document.getElementById('students-scores-container').classList.add('hidden');
                document.getElementById('submit-btn').disabled = true;
            }
        }

        function renderStudentsInputs() {
            const listDiv = document.getElementById('students-list');
            listDiv.innerHTML = '';
            
            studentsData.forEach((s, idx) => {
                const row = document.createElement('div');
                row.className = "flex items-center justify-between bg-slate-950/40 border border-slate-800/80 px-4 py-2 rounded-lg gap-4";
                row.innerHTML = `
                    <div class="flex-1 flex items-center justify-between">
                        <span class="text-sm font-semibold text-slate-200">${s.name}</span>
                        <span class="text-[10px] text-slate-500">選擇: <b class="text-indigo-300">${s.x}</b> 題</span>
                    </div>
                    <div class="flex items-center gap-3">
                        <div class="flex items-center gap-1">
                            <label class="text-[9px] font-bold text-slate-500 uppercase">手寫 Y</label>
                            <input type="number" min="0" max="6" step="0.5" value="${s.y}" data-idx="${idx}" class="student-y-input w-12 bg-slate-950 border border-slate-700 rounded p-1 text-center text-xs font-bold text-emerald-400 outline-none focus:border-indigo-500">
                        </div>
                        <button type="button" onclick="removeStudent(${idx})" class="text-slate-600 hover:text-red-500 transition-colors text-lg">×</button>
                    </div>
                `;
                listDiv.appendChild(row);
            });

            document.getElementById('students-scores-container').classList.remove('hidden');

            document.querySelectorAll('.student-y-input').forEach(input => {
                input.addEventListener('input', (e) => {
                    const idx = e.target.getAttribute('data-idx');
                    let val = parseFloat(e.target.value) || 0;
                    if (val < 0) val = 0;
                    if (val > 6) val = 6;
                    studentsData[idx].y = val;
                    updateDashboard();
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
                const res = await fetch('/analyze_full', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const d = await res.json();

                document.getElementById('sum-app').innerText = d.counts["A++"];
                document.getElementById('sum-ap').innerText = d.counts["A+"];
                document.getElementById('sum-a').innerText = d.counts["A"];
                document.getElementById('sum-bpp').innerText = d.counts["B++"];
                document.getElementById('sum-sel').innerText = d.avg_sel;
                document.getElementById('sum-nonsel').innerText = d.avg_nonsel;
                document.getElementById('sum-total').innerText = d.avg_total;
                document.getElementById('sum-count').innerText = d.count;

                document.getElementById('students-json').value = JSON.stringify(studentsData);
            } catch (e) { console.error(e); }
        }

        document.querySelectorAll('.th-input').forEach(i => i.onchange = updateDashboard);
    </script>
</body>
</html>'''

@app.route('/')
def index(): 
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload_read', methods=['POST'])
def upload_read():
    file = request.files.get('file')
    if not file:
        return jsonify({"students": []})
    students = read_students_initial(io.BytesIO(file.read()))
    return jsonify({"students": students})

@app.route('/analyze_full', methods=['POST'])
def analyze_full():
    data = request.get_json() or {}
    students = data.get('students', [])
    ths = {
        'th_app': float(data.get('th_app', 93.2)),
        'th_ap': float(data.get('th_ap', 85.7)),
        'th_a': float(data.get('th_a', 76.2)),
        'th_bpp': float(data.get('th_bpp', 67.1))
    }
    n = len(students)
    if n == 0: 
        return jsonify({"count":0, "avg_sel":0, "avg_nonsel":0, "avg_total":0, "counts":{"A++":0,"A+":0,"A":0,"B++":0}})

    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    sum_sel = 0
    sum_nonsel = 0
    sum_total = 0

    for s in students:
        x = float(s.get('x', 0))
        y = float(s.get('y', 0))
        total = (x / 25.0) * 85.0 + (y / 6.0) * 15.0

        sum_sel += x
        sum_nonsel += y
        sum_total += total

        if total >= ths['th_app']: counts["A++"] += 1
        elif total >= ths['th_ap']: counts["A+"] += 1
        elif total >= ths['th_a']: counts["A"] += 1
        elif total >= ths['th_bpp']: counts["B++"] += 1

    return jsonify({
        "count": n,
        "avg_sel": round(sum_sel / n, 2),
        "avg_nonsel": round(sum_nonsel / n, 2),
        "avg_total": round(sum_total / n, 2),
        "counts": counts
    })

@app.route('/generate', methods=['POST'])
def generate():
    exam_name = request.form.get('exam_name', '成績報表')
    ths = {k: float(request.form.get(k, 0)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    
    students_json = request.form.get('students_json', '[]')
    students = json.loads(students_json)
    
    # 標題處理邏輯
    parts = exam_name.strip().split()
    if len(parts) >= 3:
        lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
    else:
        lines = ["", exam_name, ""]
    
    buf = build_excel(students, lines, ths)
    return send_file(buf, as_attachment=True, download_name=f"{exam_name}.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
