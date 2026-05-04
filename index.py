from flask import Flask, request, send_file, render_template_string
import math
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ════════════════════════════════
# 原始 Excel 樣式與功能邏輯 (完全保留)
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

    avg_sel    = round(sum(s  for _, s,  _, _ in sorted_s) / n, 2)
    avg_nonsel = round(sum(ns for _, _, ns, _ in sorted_s) / n, 2)
    avg_total  = round(sum(t  for _, _, _,  t in sorted_s) / n, 2)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for _, _, _, t in sorted_s:
        g = get_grade(t, th_app, th_ap, th_a, th_bpp)
        if g in counts: counts[g] += 1

    rows_per_block = math.ceil((n + 1) / 3)
    HEADER_ROW = 1
    DATA_START = 2
    FINAL_ROW = DATA_START + rows_per_block - 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

    for b in range(3):
        base = b * 4 + 1
        ws.column_dimensions[get_column_letter(base)].width   = 9
        ws.column_dimensions[get_column_letter(base+1)].width = 7.5
        ws.column_dimensions[get_column_letter(base+2)].width = 7.5
        ws.column_dimensions[get_column_letter(base+3)].width = 7.5
    
    ws.column_dimensions["M"].width = 0.4
    ws.column_dimensions["N"].width = 7
    ws.column_dimensions["O"].width = 7
    ws.column_dimensions["P"].width = 7

    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, HEADER_ROW, base + i, h, border=all_thin())

    for idx, (name, sel, nonsel, total) in enumerate(sorted_s):
        b = idx // rows_per_block
        r = DATA_START + (idx % rows_per_block)
        col = b * 4 + 1
        g = get_grade(total, th_app, th_ap, th_a, th_bpp)
        f = FILLS[g]
        sc(ws, r, col,     name,   fill=f, border=all_thin())
        sc(ws, r, col + 1, sel,    fill=f, border=all_thin())
        sc(ws, r, col + 2, nonsel, fill=f, border=all_thin())
        sc(ws, r, col + 3, total,  fill=f, border=all_thin())

    avg_pos = n
    b_avg = avg_pos // rows_per_block
    r_avg_start = DATA_START + (avg_pos % rows_per_block)
    r_avg_end = FINAL_ROW
    col_avg = b_avg * 4 + 1
    avg_vals = ["平均", avg_sel, avg_nonsel, avg_total]

    for i, val in enumerate(avg_vals):
        curr_col = col_avg + i
        for fill_r in range(r_avg_start, r_avg_end + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        if r_avg_end > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=r_avg_end, end_column=curr_col)
        sc(ws, r_avg_start, curr_col, val, bold=True, size=12, border=all_thin())

    for b in range(3):
        base = b * 4 + 1
        for r in range(DATA_START, FINAL_ROW + 1):
            if not ws.cell(row=r, column=base).border:
                for i in range(4):
                    sc(ws, r, base + i, "", border=all_thin())

    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font = Font(name=FONT_NAME, bold=True, size=18)
    tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    visible = [(g, counts[g]) for g in ["A++", "A+", "A", "B++"] if counts[g] > 0]
    GRADE_R1 = TITLE_R2 + 2
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
# 網頁路由處理
# ════════════════════════════════

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>成績報表產生器</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body { font-family: sans-serif; background: #0f172a; color: #f1f5f9; display: flex; justify-content: center; padding: 40px 20px; }
        .card { background: #1e293b; padding: 30px; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); max-width: 500px; width: 100%; }
        h2 { color: #38bdf8; margin-top: 0; }
        .form-group { margin-bottom: 15px; }
        label { display: block; font-size: 14px; margin-bottom: 5px; color: #94a3b8; }
        input[type="text"], input[type="number"], input[type="file"] { width: 100%; padding: 10px; border-radius: 6px; border: 1px solid #334155; background: #0f172a; color: white; box-sizing: border-box; }
        button { width: 100%; padding: 12px; background: #3b82f6; border: none; color: white; font-weight: bold; border-radius: 6px; cursor: pointer; margin-top: 10px; }
        button:hover { background: #2563eb; }
        .footer { margin-top: 20px; font-size: 12px; color: #64748b; text-align: center; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📊 成績報表產生器</h2>
        <form action="/generate" method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label>考試名稱 (空格分三段, 如: 國三 金安 模擬考)</label>
                <input type="text" name="exam_name" value="國三 金安 模擬考" required>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                <div class="form-group"><label>A++ 門檻</label><input type="number" step="0.1" name="th_app" value="93.2"></div>
                <div class="form-group"><label>A+ 門檻</label><input type="number" step="0.1" name="th_ap" value="85.7"></div>
                <div class="form-group"><label>A 門檻</label><input type="number" step="0.1" name="th_a" value="76.2"></div>
                <div class="form-group"><label>B++ 門檻</label><input type="number" step="0.1" name="th_bpp" value="67.1"></div>
            </div>
            <div class="form-group">
                <label>上傳 Excel 檔案</label>
                <input type="file" name="file" accept=".xlsx" required>
            </div>
            <button type="submit">🚀 產生並下載 Excel</button>
        </form>
        <div class="footer">Deploy on Vercel with Python Serverless</div>
    </div>
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
    th_app = float(request.form.get('th_app', 93.2))
    th_ap  = float(request.form.get('th_ap', 85.7))
    th_a   = float(request.form.get('th_a', 76.2))
    th_bpp = float(request.form.get('th_bpp', 67.1))

    # 讀取學生資料
    students = read_students(io.BytesIO(file.read()))

    # 處理標題行切割
    parts = exam_name.strip().split()
    if len(parts) >= 3: lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
    elif len(parts) == 2: lines = [parts[0], "", parts[1]]
    else: lines = ["", exam_name.strip(), ""]

    # 產生報表
    report_buf = build_report(students, lines, th_app, th_ap, th_a, th_bpp)

    return send_file(
        report_buf,
        as_attachment=True,
        download_name=f"{exam_name}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# 為了 Vercel 的啟動
if __name__ == "__main__":
    app.run()
