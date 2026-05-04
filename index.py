from flask import Flask, request, send_file, render_template_string, jsonify
import math
import io
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

def read_students(file_stream):
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    students = []
    for row in wb.active.iter_rows(values_only=True):
        if not row[0]: continue
        try:
            students.append((str(row[0]).strip(), float(row[1]), float(row[2]), float(row[3])))
        except: pass
    return students

def build_excel(students, exam_lines, ths):
    sorted_s = sorted(students, key=lambda x: -x[3])
    n = len(sorted_s)
    
    # 計算各等級人數
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for s in sorted_s:
        t = s[3]
        if t >= ths['th_app']: counts["A++"] += 1
        elif t >= ths['th_ap']: counts["A+"] += 1
        elif t >= ths['th_a']: counts["A"] += 1
        elif t >= ths['th_bpp']: counts["B++"] += 1

    rows_per_block = math.ceil((n + 2) / 3) 
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
        # 取得等級顏色
        g = ""
        if total >= ths['th_app']: g = "A++"
        elif total >= ths['th_ap']: g = "A+"
        elif total >= ths['th_a']: g = "A"
        elif total >= ths['th_bpp']: g = "B++"
        f = FILLS.get(g)
        for i, val in enumerate([name, sel, nonsel, total]):
            sc(ws, r, col + i, val, fill=f, border=all_thin())

    # 平均值 (自動補全剩餘空間)
    avg_pos = n
    b_avg, r_avg_start = avg_pos // rows_per_block, DATA_START + (avg_pos % rows_per_block)
    col_avg_base = b_avg * 4 + 1
    avg_vals = ["平均", 
                round(sum(s[1] for s in students)/n, 2), 
                round(sum(s[2] for s in students)/n, 2), 
                round(sum(s[3] for s in students)/n, 2)]
    
    for i, val in enumerate(avg_vals):
        curr_col = col_avg_base + i
        for fill_r in range(r_avg_start, FINAL_ROW + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        if FINAL_ROW > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=FINAL_ROW, end_column=curr_col)
        sc(ws, r_avg_start, curr_col, val, bold=True, border=all_thin())

    # 標題與人數方格
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font, tc.alignment = Font(name=FONT_NAME, bold=True, size=18), Alignment(horizontal="center", vertical="center", wrap_text=True)
    # 畫標題外框
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
# 進階漸層 UI
# ════════════════════════════════

HTML_TEMPLATE = '''
<!DOCTYPE html>
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
                radial-gradient(at 50% 100%, #0f172a 0, transparent 50%); /* 降低底部亮度 */
            background-attachment: fixed; /* 關鍵：讓漸層固定，不會因為頁面長短而亂跑 */
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
        .gradient-border {
            position: relative;
            background: rgba(15, 23, 42, 0.8);
            border-radius: 1rem;
        }
        .gradient-border::before {
            content: ""; position: absolute; inset: -1px; border-radius: 1rem; padding: 1px;
            background: linear-gradient(45deg, #6366f1, #a855f7, #ec4899);
            -webkit-mask: linear-gradient(#fff 0 0) content-box, linear-gradient(#fff 0 0);
            mask: linear-gradient(#fff 0 0) content-box, linear-gradient(#fff 0 0);
            -webkit-mask-composite: xor; mask-composite: exclude; pointer-events: none;
        }
        .btn-main {
            background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .btn-main:hover { transform: translateY(-2px); box-shadow: 0 0 20px rgba(168, 85, 247, 0.4); }
        .stat-card {
            background: linear-gradient(180deg, rgba(255,255,255,0.05) 0%, rgba(255,255,255,0) 100%);
            border: 1px solid rgba(255,255,255,0.1);
        }
    </style>
</head>
<body class="p-4 md:p-12 flex justify-center">
    <div class="max-w-4xl w-full space-y-8">
        <!-- Header -->
        <header class="text-center space-y-2">
            <h1 class="text-4xl font-extrabold tracking-tighter bg-gradient-to-r from-indigo-400 via-purple-400 to-pink-400 bg-clip-text text-transparent">
                SCORE REPORT GENERATOR
            </h1>
            <p class="text-slate-400 text-sm font-medium"></p>
        </header>

        <!-- 設定區域 -->
        <div class="glass-card rounded-2xl p-8 space-y-6">
            <form id="main-form" action="/generate" method="post" enctype="multipart/form-data" class="space-y-6">
                <div>
                    <label class="text-[10px] uppercase tracking-[0.2em] text-slate-500 font-bold mb-2 block">考試名稱</label>
                    <input type="text" name="exam_name" value="國三 金安模擬考 第六回" class="w-full bg-slate-950/50 border border-slate-800 rounded-lg px-4 py-3 focus:ring-2 focus:ring-indigo-500 outline-none transition-all">
                </div>

                <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">A++</label>
                        <input type="number" step="0.1" name="th_app" value="93.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">A+</label>
                        <input type="number" step="0.1" name="th_ap" value="85.7" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">A</label>
                        <input type="number" step="0.1" name="th_a" value="76.2" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                    <div class="space-y-1">
                        <label class="text-[10px] text-slate-500 font-bold">B++</label>
                        <input type="number" step="0.1" name="th_bpp" value="67.1" class="th-input w-full bg-slate-950/50 border border-slate-800 rounded-lg p-2 text-center">
                    </div>
                </div>

                <!-- 上傳區 -->
                <div id="dropzone" class="group border-2 border-dashed border-slate-800 rounded-xl p-10 text-center hover:border-indigo-500 hover:bg-indigo-500/5 transition-all cursor-pointer">
                    <input type="file" id="file-input" name="file" accept=".xlsx" class="hidden">
                    <div id="upload-content">
                        <div class="text-3xl mb-2 group-hover:scale-110 transition-transform">📁</div>
                        <p class="text-sm text-slate-500" id="file-status">點選或拖拽 XLSX 檔案至此</p>
                    </div>
                </div>

                <div id="success-bar" class="hidden bg-indigo-500/10 border border-indigo-500/20 py-3 px-4 rounded-lg flex items-center justify-between">
                    <span class="text-xs text-indigo-300 font-medium">✨ 讀取成功：<span id="st-count">0</span> 位成員</span>
                    <span class="text-[10px] px-2 py-0.5 bg-indigo-500/20 rounded-full text-indigo-400">READY</span>
                </div>

                <button type="submit" class="btn-main w-full py-4 rounded-xl font-bold text-white shadow-xl">
                    下載成績單
                </button>
            </form>
        </div>

        <!-- 數據摘要 -->
        <div class="space-y-4">
            <h3 class="text-xs font-bold text-slate-500 uppercase tracking-widest px-1"></h3>
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

            <!-- 橫向數據條 -->
            <div class="glass-card rounded-xl p-5 flex flex-wrap gap-6 justify-between items-center text-xs font-bold text-slate-400">
                <div class="flex gap-8">
                    <div class="flex flex-col">
                        <span class="text-[8px] text-slate-500">選擇平均</span>
                        <span id="sum-sel" class="text-slate-200">--</span>
                    </div>
                    <div class="flex flex-col">
                        <span class="text-[8px] text-slate-500">非選平均</span>
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

        dropzone.onclick = () => fileInput.click();

        async function updateDashboard() {
            const file = fileInput.files[0];
            if (!file) return;

            const formData = new FormData();
            formData.append('file', file);
            formData.append('th_app', document.getElementsByName('th_app')[0].value);
            formData.append('th_ap', document.getElementsByName('th_ap')[0].value);
            formData.append('th_a', document.getElementsByName('th_a')[0].value);
            formData.append('th_bpp', document.getElementsByName('th_bpp')[0].value);

            try {
                const res = await fetch('/analyze', { method: 'POST', body: formData });
                const d = await res.json();

                // 數值更新
                document.getElementById('sum-app').innerText = d.counts["A++"];
                document.getElementById('sum-ap').innerText = d.counts["A+"];
                document.getElementById('sum-a').innerText = d.counts["A"];
                document.getElementById('sum-bpp').innerText = d.counts["B++"];
                document.getElementById('sum-sel').innerText = d.avg_sel;
                document.getElementById('sum-nonsel').innerText = d.avg_nonsel;
                document.getElementById('sum-total').innerText = d.avg_total;
                document.getElementById('sum-count').innerText = d.count;
                document.getElementById('st-count').innerText = d.count;

                // UI 狀態切換
                document.getElementById('success-bar').classList.remove('hidden');
                document.getElementById('file-status').innerText = file.name;
                document.getElementById('file-status').classList.add('text-indigo-400', 'font-bold');
            } catch (e) { console.error(e); }
        }

        fileInput.onchange = updateDashboard;
        document.querySelectorAll('.th-input').forEach(i => i.onchange = updateDashboard);
    </script>
</body>
</html>
'''

@app.route('/')
def index(): return render_template_string(HTML_TEMPLATE)

@app.route('/analyze', methods=['POST'])
def analyze():
    file = request.files['file']
    ths = {k: float(request.form.get(k)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    students = read_students(io.BytesIO(file.read()))
    n = len(students)
    if n == 0: return jsonify({})
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for s in students:
        t = s[3]
        if t >= ths['th_app']: counts["A++"] += 1
        elif t >= ths['th_ap']: counts["A+"] += 1
        elif t >= ths['th_a']: counts["A"] += 1
        elif t >= ths['th_bpp']: counts["B++"] += 1
        
    return jsonify({
        "count": n,
        "avg_sel": round(sum(s[1] for s in students)/n, 2),
        "avg_nonsel": round(sum(s[2] for s in students)/n, 2),
        "avg_total": round(sum(s[3] for s in students)/n, 2),
        "counts": counts
    })

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files['file']
    exam_name = request.form.get('exam_name', '')
    ths = {k: float(request.form.get(k)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    students = read_students(io.BytesIO(file.read()))
    
    parts = exam_name.strip().split()
    lines = [parts[0], " ".join(parts[1:-1]), parts[-1]] if len(parts) >= 3 else ["", exam_name, ""]
    
    buf = build_excel(students, lines, ths)
    return send_file(buf, as_attachment=True, download_name=f"{exam_name}.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
