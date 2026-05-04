from flask import Flask, request, send_file, render_template_string, jsonify
import math
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ════════════════════════════════
# 核心 Excel 邏輯 (平均格數與間隔優化)
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
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name=FONT_NAME, bold=bold, size=size)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if border: c.border = border
    if fill:   c.fill   = fill
    return c

def read_students(file_stream):
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    students = []
    for row in wb.active.iter_rows(values_only=True):
        if not row[0]: continue
        try:
            students.append((str(row[0]).strip(), float(row[1]), float(row[2]), float(row[3])))
        except: pass
    return students

def build_report(students, exam_lines, th_app, th_ap, th_a, th_bpp):
    sorted_s = sorted(students, key=lambda x: -x[3])
    n = len(sorted_s)
    
    # 統計數據
    avg_sel = round(sum(s for _, s, _, _ in sorted_s) / n, 2)
    avg_nonsel = round(sum(ns for _, _, ns, _ in sorted_s) / n, 2)
    avg_total = round(sum(t for _, _, _, t in sorted_s) / n, 2)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    def get_g(t):
        if t >= th_app: return "A++"
        if t >= th_ap:  return "A+"
        if t >= th_a:   return "A"
        if t >= th_bpp: return "B++"
        return ""

    rows_per_block = math.ceil((n + 2) / 3) 
    HEADER_ROW, DATA_START = 1, 2
    FINAL_ROW = DATA_START + rows_per_block - 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

    # 設定欄寬
    for b in range(3):
        base = b * 4 + 1
        for offset, w in enumerate([9, 7.5, 7.5, 7.5]):
            ws.column_dimensions[get_column_letter(base + offset)].width = w
    ws.column_dimensions["M"].width = 0.4
    for col_let in ["N", "O", "P"]: ws.column_dimensions[col_let].width = 7

    # 填寫數據
    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, HEADER_ROW, base + i, h, border=all_thin())

    for idx, (name, sel, nonsel, total) in enumerate(sorted_s):
        b, r = idx // rows_per_block, DATA_START + (idx % rows_per_block)
        col, g = b * 4 + 1, get_g(total)
        if g in counts: counts[g] += 1
        f = FILLS[g]
        for i, val in enumerate([name, sel, nonsel, total]):
            sc(ws, r, col + i, val, fill=f, border=all_thin())

    # 平均值邏輯 (自動延伸到底端)
    avg_pos = n
    b_avg, r_avg_start = avg_pos // rows_per_block, DATA_START + (avg_pos % rows_per_block)
    col_avg_base, avg_vals = b_avg * 4 + 1, ["平均", avg_sel, avg_nonsel, avg_total]
    
    for i, val in enumerate(avg_vals):
        curr_col = col_avg_base + i
        for fill_r in range(r_avg_start, FINAL_ROW + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        if FINAL_ROW > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=FINAL_ROW, end_column=curr_col)
        sc(ws, r_avg_start, curr_col, val, bold=True, border=all_thin())

    # 補全格線
    for b in range(3):
        base = b * 4 + 1
        for r in range(DATA_START, FINAL_ROW + 1):
            if not ws.cell(row=r, column=base).border:
                for i in range(4): sc(ws, r, base + i, "", border=all_thin())

    # 標題方格 (間隔三格)
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font, tc.alignment = Font(name=FONT_NAME, bold=True, size=18), Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    # 人數摘要
    visible = [(g, counts[g]) for g in ["A++", "A+", "A", "B++"]]
    GRADE_R1 = TITLE_R2 + 4 
    for i, (g, cnt) in enumerate(visible):
        row = GRADE_R1 + i
        f = FILLS[g]
        for col, val in [(14, g), (15, cnt)]:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(name=FONT_NAME, bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = outer_med(row, col, GRADE_R1, 14, GRADE_R1 + 3, 15)
            if f: cell.fill = f

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ════════════════════════════════
# 網頁 UI 設計 (完全復刻圖片樣式)
# ════════════════════════════════

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>成績報表產生器</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
        body { background: radial-gradient(circle at center, #1e293b 0%, #0f172a 100%); font-family: 'Noto Sans TC', sans-serif; color: #e2e8f0; }
        .glass-panel { background: rgba(30, 41, 59, 0.7); backdrop-filter: blur(10px); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 12px; }
        .input-dark { background: #0f172a; border: 1px solid #334155; border-radius: 8px; padding: 10px 15px; width: 100%; color: white; outline: none; }
        .input-dark:focus { border-color: #6366f1; }
        .btn-gradient-purple { background: linear-gradient(90deg, #818cf8 0%, #a78bfa 100%); transition: opacity 0.2s; color: white; font-weight: bold; }
        .btn-gradient-blue { background: linear-gradient(90deg, #60a5fa 0%, #818cf8 100%); transition: opacity 0.2s; color: white; font-weight: bold; }
        .btn-gradient-purple:hover, .btn-gradient-blue:hover { opacity: 0.9; }
        .step-btn { background: #1e293b; border: 1px solid #334155; padding: 4px 10px; border-radius: 4px; color: #818cf8; font-weight: bold; }
        .step-btn:hover { background: #334155; }
        .summary-card { padding: 20px; border-radius: 12px; text-align: center; }
        .grade-a-pp { background: rgba(59, 130, 246, 0.2); border: 1px solid rgba(59, 130, 246, 0.3); }
        .grade-a-p { background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); }
        .grade-a { background: rgba(245, 158, 11, 0.1); border: 1px solid rgba(245, 158, 11, 0.3); }
        .grade-b-pp { background: rgba(239, 68, 68, 0.1); border: 1px solid rgba(239, 68, 68, 0.3); }
    </style>
</head>
<body class="min-h-screen p-8 flex flex-col items-center">

    <div class="max-w-4xl w-full space-y-10">
        
        <!-- Header -->
        <div class="text-center">
            <h1 class="text-3xl font-bold flex items-center justify-center gap-3">
                <span class="text-4xl">📊</span> 成績報表產生器
            </h1>
            <p class="text-slate-400 mt-2 text-sm">上傳成績 Excel，自動排名並輸出格式化報表</p>
        </div>

        <!-- 考試設定 -->
        <div class="glass-panel p-8 relative">
            <div class="flex items-center gap-2 mb-6 text-sm font-bold text-slate-300 uppercase tracking-widest">
                <span>⚙️ 考試設定</span>
            </div>
            <form id="main-form" action="/generate" method="post" enctype="multipart/form-data" class="space-y-8">
                <div>
                    <label class="block text-xs text-slate-400 mb-2">考試名稱 (以空格分三段)</label>
                    <input type="text" name="exam_name" value="國三 金安模擬考 第六回" class="input-dark">
                </div>

                <div class="grid grid-cols-4 gap-6">
                    <!-- 門檻調整組件 -->
                    <div class="space-y-2">
                        <label class="text-xs text-slate-400 block text-center">A++ 門檻</label>
                        <div class="flex items-center gap-2">
                            <input type="number" step="0.1" name="th_app" id="th_app" value="93.2" class="input-dark text-center py-2 px-1 text-sm">
                        </div>
                    </div>
                    <div class="space-y-2">
                        <label class="text-xs text-slate-400 block text-center">A+ 門檻</label>
                        <div class="flex items-center gap-2">
                            <input type="number" step="0.1" name="th_ap" id="th_ap" value="85.7" class="input-dark text-center py-2 px-1 text-sm">
                        </div>
                    </div>
                    <div class="space-y-2">
                        <label class="text-xs text-slate-400 block text-center">A 門檻</label>
                        <div class="flex items-center gap-2">
                            <input type="number" step="0.1" name="th_a" id="th_a" value="76.2" class="input-dark text-center py-2 px-1 text-sm">
                        </div>
                    </div>
                    <div class="space-y-2">
                        <label class="text-xs text-slate-400 block text-center">B++ 門檻</label>
                        <div class="flex items-center gap-2">
                            <input type="number" step="0.1" name="th_bpp" id="th_bpp" value="67.1" class="input-dark text-center py-2 px-1 text-sm">
                        </div>
                    </div>
                </div>

                <!-- 上傳檔案 -->
                <div class="space-y-3">
                    <label class="text-xs text-slate-400 font-bold flex items-center gap-2">📁 上傳成績檔案</label>
                    <div id="dropzone" class="border-2 border-dashed border-slate-700 rounded-xl p-8 flex items-center gap-4 transition-all hover:border-indigo-500 cursor-pointer">
                        <div id="file-preview" class="hidden flex items-center gap-3 bg-slate-800 p-3 rounded-lg border border-slate-600">
                            <span class="text-2xl">📄</span>
                            <div class="text-left">
                                <p id="file-name" class="text-xs font-bold text-slate-200">s2.xlsx</p>
                                <p id="file-size" class="text-[10px] text-slate-500">10.3KB</p>
                            </div>
                            <button type="button" class="ml-2 text-slate-500 hover:text-white">✕</button>
                        </div>
                        <div id="placeholder-text" class="flex items-center gap-3">
                            <span class="text-xl text-slate-500">+</span>
                            <span class="text-xs text-slate-500">點擊選取 Excel 檔案</span>
                        </div>
                        <input type="file" id="file-input" name="file" accept=".xlsx" class="hidden">
                    </div>
                </div>

                <!-- 成功提示 (上傳後才顯示) -->
                <div id="success-msg" class="hidden bg-emerald-900/30 border border-emerald-500/50 p-4 rounded-xl flex items-center gap-3">
                    <span class="text-emerald-400">✅</span>
                    <span class="text-sm text-emerald-200">成功讀取 <span id="student-count">38</span> 位學生資料</span>
                </div>

                <!-- 按鈕區 -->
                <div class="space-y-4">
                    <button type="submit" class="w-full py-4 rounded-xl btn-gradient-purple flex items-center justify-center gap-2">
                        🚀 產生報表
                    </button>
                    <button id="download-btn" disabled class="w-full py-4 rounded-xl btn-gradient-blue flex items-center justify-center gap-2 opacity-50 cursor-not-allowed">
                        ⬇️ 下載 Excel 報表
                    </button>
                </div>
            </form>
        </div>

        <!-- 報表摘要 -->
        <div id="summary-section" class="space-y-6">
            <div class="flex items-center gap-2 text-sm font-bold text-slate-300 uppercase tracking-widest">
                <span>📊 報表摘要</span>
            </div>
            <div class="grid grid-cols-4 gap-6">
                <div class="summary-card grade-a-pp">
                    <p class="text-[10px] font-bold text-blue-400">A++</p>
                    <p id="stat-app" class="text-3xl font-black mt-2">--</p>
                </div>
                <div class="summary-card grade-a-p">
                    <p class="text-[10px] font-bold text-emerald-400">A+</p>
                    <p id="stat-ap" class="text-3xl font-black mt-2">--</p>
                </div>
                <div class="summary-card grade-a">
                    <p class="text-[10px] font-bold text-amber-400">A</p>
                    <p id="stat-a" class="text-3xl font-black mt-2">--</p>
                </div>
                <div class="summary-card grade-b-pp">
                    <p class="text-[10px] font-bold text-rose-400">B++</p>
                    <p id="stat-bpp" class="text-3xl font-black mt-2">--</p>
                </div>
            </div>

            <!-- 底部橫向統計 -->
            <div class="glass-panel p-4 flex items-center justify-between text-xs font-bold text-slate-400 px-8">
                <div class="flex gap-6">
                    <span class="flex items-center gap-2">📐 選擇平均 <span id="avg-sel" class="text-slate-200">--</span></span>
                    <span class="flex items-center gap-2">非選平均 <span id="avg-nonsel" class="text-slate-200">--</span></span>
                    <span class="flex items-center gap-2">總分平均 <span id="avg-total" class="text-slate-200">--</span></span>
                </div>
                <div class="flex items-center gap-2">共 <span id="total-count" class="text-indigo-400">--</span> 人</div>
            </div>
        </div>
    </div>

    <script>
        const fileInput = document.getElementById('file-input');
        const dropzone = document.getElementById('dropzone');
        const filePreview = document.getElementById('file-preview');
        const placeholder = document.getElementById('placeholder-text');
        const successMsg = document.getElementById('success-msg');
        
        // 點擊上傳
        dropzone.onclick = () => fileInput.click();

        fileInput.onchange = function(e) {
            const file = e.target.files[0];
            if (file) {
                document.getElementById('file-name').innerText = file.name;
                document.getElementById('file-size').innerText = (file.size / 1024).toFixed(1) + 'KB';
                filePreview.classList.remove('hidden');
                placeholder.classList.add('hidden');
                
                // 模擬讀取 (實際可透過 AJAX 獲取更精準數據，這裡為了示範先顯示 UI)
                successMsg.classList.remove('hidden');
            }
        };

        // 產生報表動作 (這裡簡化為直接提交表單)
        const form = document.getElementById('main-form');
        form.onsubmit = function() {
            // 下載按鈕在上傳成功後啟用
            const dlBtn = document.getElementById('download-btn');
            dlBtn.classList.remove('opacity-50', 'cursor-not-allowed');
            dlBtn.disabled = false;
        };
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files['file']
    if not file: return "No file", 400
    
    exam_name = request.form.get('exam_name', '')
    params = {
        'th_app': float(request.form.get('th_app', 93.2)),
        'th_ap':  float(request.form.get('th_ap', 85.7)),
        'th_a':   float(request.form.get('th_a', 76.2)),
        'th_bpp': float(request.form.get('th_bpp', 67.1))
    }
    
    students = read_students(io.BytesIO(file.read()))
    parts = exam_name.strip().split()
    if len(parts) >= 3:
        lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
    elif len(parts) == 2:
        lines = [parts[0], "", parts[1]]
    else:
        lines = ["", exam_name.strip(), ""]

    report_buf = build_report(students, lines, **params)
    return send_file(
        report_buf, 
        as_attachment=True, 
        download_name=f"{exam_name}.xlsx", 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

def handler(event, context):
    return app(event, context)

if __name__ == "__main__":
    app.run(debug=True)
