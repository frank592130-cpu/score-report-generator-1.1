from flask import Flask, request, send_file, render_template_string
import math
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ════════════════════════════════
# 原始 Excel 核心邏輯 (功能 100% 保留)
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

def get_grade(total, th_app, th_ap, th_a, th_bpp):
    if total >= th_app: return "A++"
    if total >= th_ap:  return "A+"
    if total >= th_a:   return "A"
    if total >= th_bpp: return "B++"
    return ""

def read_students(file_stream):
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    students = []
    for row in wb.active.iter_rows(values_only=True):
        if not row[0]: continue
        try:
            students.append((str(row[0]).strip(), float(row[1]), float(row[2]), float(row[3])))
        except (TypeError, ValueError): pass
    return students

def build_report(students, exam_lines, th_app, th_ap, th_a, th_bpp):
    sorted_s = sorted(students, key=lambda x: -x[3])
    n = len(sorted_s)
    avg_sel = round(sum(s for _, s, _, _ in sorted_s) / n, 2)
    avg_nonsel = round(sum(ns for _, _, ns, _ in sorted_s) / n, 2)
    avg_total = round(sum(t for _, _, _, t in sorted_s) / n, 2)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for _, _, _, t in sorted_s:
        g = get_grade(t, th_app, th_ap, th_a, th_bpp)
        if g in counts: counts[g] += 1

    rows_per_block = math.ceil((n + 1) / 3)
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
        col, g = b * 4 + 1, get_grade(total, th_app, th_ap, th_a, th_bpp)
        f = FILLS[g]
        for i, val in enumerate([name, sel, nonsel, total]):
            sc(ws, r, col + i, val, fill=f, border=all_thin())

    # 平均值處理
    avg_pos = n
    b_avg, r_avg_start = avg_pos // rows_per_block, DATA_START + (avg_pos % rows_per_block)
    col_avg, avg_vals = b_avg * 4 + 1, ["平均", avg_sel, avg_nonsel, avg_total]
    for i, val in enumerate(avg_vals):
        curr_col = col_avg + i
        for fill_r in range(r_avg_start, FINAL_ROW + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        if FINAL_ROW > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=FINAL_ROW, end_column=curr_col)
        sc(ws, r_avg_start, curr_col, val, bold=True, size=12, border=all_thin())

    # 補全空格
    for b in range(3):
        base = b * 4 + 1
        for r in range(DATA_START, FINAL_ROW + 1):
            if not ws.cell(row=r, column=base).border:
                for i in range(4): sc(ws, r, base + i, "", border=all_thin())

    # 標題方格
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font, tc.alignment = Font(name=FONT_NAME, bold=True, size=18), Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    # 人數方格與標題方格間隔三格保留 (TITLE_R2=7, +4 = 第11行開始)
    visible = [(g, counts[g]) for g in ["A++", "A+", "A", "B++"] if counts[g] > 0]
    GRADE_R1 = TITLE_R2 + 4 
    
    for i, (g, cnt) in enumerate(visible):
        row = GRADE_R1 + i
        f = FILLS[g]
        for col, val in [(14, g), (15, cnt)]:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(name=FONT_NAME, bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = outer_med(row, col, GRADE_R1, 14, GRADE_R1 + len(visible) - 1, 15)
            if f: cell.fill = f

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ════════════════════════════════
# 網頁 UI 設計 (深淺模式 + 漸層)
# ════════════════════════════════

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>成績報表產生器</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            darkMode: 'class',
        }
    </script>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Noto Sans TC', sans-serif; transition: background-color 0.4s ease; }
        .glass-card {
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            transition: all 0.4s ease;
        }
        /* 自訂漸層背景 */
        .bg-gradient-light { background: radial-gradient(circle at top left, #f3f4f6, #e0e7ff); }
        .bg-gradient-dark { background: radial-gradient(circle at top right, #0f172a, #1e1b4b); }
    </style>
</head>
<body class="min-h-screen flex items-center justify-center p-6 bg-slate-50 dark:bg-slate-900 transition-colors duration-500">
    
    <!-- 深淺模式切換按鈕 -->
    <button id="theme-toggle" class="absolute top-6 right-6 p-3 rounded-full bg-white/50 dark:bg-slate-800/50 shadow-lg hover:scale-110 transition-all text-slate-800 dark:text-yellow-400 backdrop-blur-md border border-slate-200 dark:border-slate-700">
        <svg id="theme-toggle-dark-icon" class="hidden w-5 h-5" fill="currentColor" viewBox="0 0 20 20"><path d="M17.293 13.293A8 8 0 016.707 2.707a8.001 8.001 0 1010.586 10.586z"></path></svg>
        <svg id="theme-toggle-light-icon" class="hidden w-5 h-5" fill="currentColor" viewBox="0 0 20 20"><path d="M10 2a1 1 0 011 1v1a1 1 0 11-2 0V3a1 1 0 011-1zm4 8a4 4 0 11-8 0 4 4 0 018 0zm-.464 4.95l.707.707a1 1 0 001.414-1.414l-.707-.707a1 1 0 00-1.414 1.414zm2.12-10.607a1 1 0 010 1.414l-.706.707a1 1 0 11-1.414-1.414l.707-.707a1 1 0 011.414 0zM17 11a1 1 0 100-2h-1a1 1 0 100 2h1zm-7 4a1 1 0 011 1v1a1 1 0 11-2 0v-1a1 1 0 011-1zM5.05 6.464A1 1 0 106.465 5.05l-.708-.707a1 1 0 00-1.414 1.414l.707.707zm1.414 8.486l-.707.707a1 1 0 01-1.414-1.414l.707-.707a1 1 0 011.414 1.414zM4 11a1 1 0 100-2H3a1 1 0 000 2h1z" fill-rule="evenodd" clip-rule="evenodd"></path></svg>
    </button>

    <div class="glass-card max-w-lg w-full p-8 rounded-3xl shadow-2xl bg-white/70 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700 relative overflow-hidden">
        <!-- 裝飾用漸層光暈 -->
        <div class="absolute -top-20 -left-20 w-40 h-40 bg-blue-400/30 dark:bg-blue-600/20 rounded-full blur-3xl pointer-events-none"></div>
        <div class="absolute -bottom-20 -right-20 w-40 h-40 bg-purple-400/30 dark:bg-purple-600/20 rounded-full blur-3xl pointer-events-none"></div>

        <div class="text-center mb-8 relative z-10">
            <div class="inline-block p-4 rounded-2xl bg-gradient-to-br from-blue-100 to-indigo-100 dark:from-blue-900/40 dark:to-indigo-900/40 mb-4 shadow-inner">
                <span class="text-4xl">📊</span>
            </div>
            <h1 class="text-3xl font-bold bg-gradient-to-r from-blue-600 to-purple-600 dark:from-blue-400 dark:to-purple-400 bg-clip-text text-transparent">成績報表產生器</h1>
            <p class="text-slate-500 dark:text-slate-400 mt-2 text-sm">自動計算排名，輸出專屬 Excel</p>
        </div>

        <form action="/generate" method="post" enctype="multipart/form-data" class="space-y-6 relative z-10">
            <div>
                <label class="block text-xs font-bold uppercase tracking-wider text-slate-600 dark:text-slate-400 mb-2">考試名稱</label>
                <input type="text" name="exam_name" placeholder="例如：國三 金安 模擬考" required
                    class="w-full bg-white dark:bg-slate-900/50 border border-slate-300 dark:border-slate-600 rounded-xl px-4 py-3 text-slate-800 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all shadow-sm">
            </div>

            <div class="grid grid-cols-2 gap-4">
                <div>
                    <label class="block text-xs font-bold text-slate-600 dark:text-slate-400 mb-2">A++ 門檻</label>
                    <input type="number" step="0.1" name="th_app" value="93.2" class="w-full bg-white dark:bg-slate-900/50 border border-slate-300 dark:border-slate-600 rounded-xl px-4 py-3 text-slate-800 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-sm transition-all">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-600 dark:text-slate-400 mb-2">A+ 門檻</label>
                    <input type="number" step="0.1" name="th_ap" value="85.7" class="w-full bg-white dark:bg-slate-900/50 border border-slate-300 dark:border-slate-600 rounded-xl px-4 py-3 text-slate-800 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-sm transition-all">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-600 dark:text-slate-400 mb-2">A 門檻</label>
                    <input type="number" step="0.1" name="th_a" value="76.2" class="w-full bg-white dark:bg-slate-900/50 border border-slate-300 dark:border-slate-600 rounded-xl px-4 py-3 text-slate-800 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-sm transition-all">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-600 dark:text-slate-400 mb-2">B++ 門檻</label>
                    <input type="number" step="0.1" name="th_bpp" value="67.1" class="w-full bg-white dark:bg-slate-900/50 border border-slate-300 dark:border-slate-600 rounded-xl px-4 py-3 text-slate-800 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 shadow-sm transition-all">
                </div>
            </div>

            <div class="relative group">
                <label class="block text-xs font-bold text-slate-600 dark:text-slate-400 mb-2">上傳資料檔案 (XLSX)</label>
                <div class="flex items-center justify-center w-full">
                    <label class="flex flex-col items-center justify-center w-full h-32 border-2 border-slate-300 dark:border-slate-600 border-dashed rounded-2xl cursor-pointer bg-slate-50/50 dark:bg-slate-900/30 hover:bg-slate-100 dark:hover:bg-slate-800/50 transition-all group-hover:border-blue-500">
                        <div class="flex flex-col items-center justify-center pt-5 pb-6">
                            <svg class="w-8 h-8 mb-3 text-slate-400 dark:text-slate-500 group-hover:text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
                            <p class="text-sm text-slate-500 dark:text-slate-400 group-hover:text-slate-700 dark:group-hover:text-slate-200">點擊選取或拖曳檔案</p>
                        </div>
                        <input type="file" name="file" accept=".xlsx" class="hidden" required />
                    </label>
                </div>
            </div>

            <button type="submit" class="w-full py-4 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-bold rounded-2xl shadow-lg shadow-blue-500/30 dark:shadow-blue-900/40 transform transition hover:-translate-y-0.5 active:scale-[0.98]">
                🚀 產生並下載 Excel 報表
            </button>
        </form>
    </div>

    <!-- 深淺模式 JS 邏輯 -->
    <script>
        const themeToggleBtn = document.getElementById('theme-toggle');
        const darkIcon = document.getElementById('theme-toggle-dark-icon');
        const lightIcon = document.getElementById('theme-toggle-light-icon');
        
        // 初始狀態檢查
        if (localStorage.getItem('color-theme') === 'dark' || (!('color-theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
            document.documentElement.classList.add('dark');
            document.body.classList.add('bg-gradient-dark');
            lightIcon.classList.remove('hidden');
        } else {
            document.documentElement.classList.remove('dark');
            document.body.classList.add('bg-gradient-light');
            darkIcon.classList.remove('hidden');
        }

        themeToggleBtn.addEventListener('click', function() {
            darkIcon.classList.toggle('hidden');
            lightIcon.classList.toggle('hidden');

            if (document.documentElement.classList.contains('dark')) {
                document.documentElement.classList.remove('dark');
                document.body.classList.replace('bg-gradient-dark', 'bg-gradient-light');
                localStorage.setItem('color-theme', 'light');
            } else {
                document.documentElement.classList.add('dark');
                document.body.classList.replace('bg-gradient-light', 'bg-gradient-dark');
                localStorage.setItem('color-theme', 'dark');
            }
        });
    </script>
</body>
</html>
'''

@app.route('/')
def index(): return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files['file']
    if not file: return "No file", 400
    exam_name = request.form.get('exam_name', '')
    params = {k: float(request.form.get(k)) for k in ['th_app', 'th_ap', 'th_a', 'th_bpp']}
    
    students = read_students(io.BytesIO(file.read()))
    parts = exam_name.strip().split()
    if len(parts) >= 3: lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
    elif len(parts) == 2: lines = [parts[0], "", parts[1]]
    else: lines = ["", exam_name.strip(), ""]

    report_buf = build_report(students, lines, **params)
    return send_file(report_buf, as_attachment=True, download_name=f"{exam_name}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def handler(event, context): return app(event, context)

if __name__ == "__main__": app.run()
