import json, tempfile, os
import streamlit as st
import pandas as pd
import numpy as np
import cv2
import fitz
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill

@st.cache_data
def load_default_template():
    with open("template.json", "r", encoding="utf-8") as f:
        return json.load(f)

DARK_THRESHOLD = 215
MIN_GAP_TO_2ND = 18

def find_squares(gray):
    H, W = gray.shape[:2]
    area_total = H * W
    _, th = cv2.threshold(gray, 100, 255, cv2.THRESH_BINARY_INV)
    contours, _ = cv2.findContours(th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    squares = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        area = w * h
        if (area_total * 0.00012) < area < (area_total * 0.0020) and 0.7 < w / h < 1.35:
            squares.append((x + w / 2.0, y + h / 2.0, w, h))
    return squares

def detect_corners(gray):
    H, W = gray.shape[:2]
    squares = find_squares(gray)
    corners = {'TL': None, 'TR': None, 'BL': None, 'BR': None}
    best = {'TL': 1e18, 'TR': 1e18, 'BL': 1e18, 'BR': 1e18}
    for (cx, cy, w, h) in squares:
        d = {'TL': cx**2 + cy**2, 'TR': (W - cx)**2 + cy**2, 'BL': cx**2 + (H - cy)**2, 'BR': (W - cx)**2 + (H - cy)**2}
        for key, val in d.items():
            if val < best[key]:
                best[key] = val
                corners[key] = (cx, cy)
    return corners

def warp_to_template(page_img, template):
    gray = cv2.cvtColor(page_img, cv2.COLOR_BGR2GRAY)
    corners = detect_corners(gray)
    if any(v is None for v in corners.values()):
        return None, "Không tìm thấy đủ 4 điểm neo ở góc phiếu."
    src = np.array([corners['TL'], corners['TR'], corners['BL'], corners['BR']], dtype=np.float32)
    tc = template['corners']
    dst = np.array([tc['TL'], tc['TR'], tc['BL'], tc['BR']], dtype=np.float32)
    M = cv2.getPerspectiveTransform(src, dst)
    W, H = template['page_size']
    return cv2.warpPerspective(page_img, M, (int(W), int(H)), flags=cv2.INTER_LINEAR, borderValue=(255, 255, 255)), None

def circle_darkness(gray, x, y, r):
    x, y, r = int(round(x)), int(round(y)), int(r)
    H, W = gray.shape[:2]
    patch = gray[max(0, y-r):min(H, y+r+1), max(0, x-r):min(W, x+r+1)]
    yy, xx = np.ogrid[max(0, y-r):min(H, y+r+1), max(0, x-r):min(W, x+r+1)]
    mask = (xx - x)**2 + (yy - y)**2 <= r*r
    return float(patch[mask].mean()) if mask.sum() > 0 else 255.0

def read_answer(gray, options_dict, radius):
    vals = {opt: circle_darkness(gray, xy[0], xy[1], radius) for opt, xy in options_dict.items()}
    ordered = sorted(vals.items(), key=lambda kv: kv[1])
    darkest_opt, darkest_val = ordered[0]
    second_val = ordered[1][1] if len(ordered) > 1 else 255
    if darkest_val > DARK_THRESHOLD: return None, False
    if (second_val - darkest_val) < MIN_GAP_TO_2ND and second_val <= DARK_THRESHOLD: return None, True
    return darkest_opt, False

def read_digit_column(gray, rows, col_x, radius):
    vals = [circle_darkness(gray, col_x, y, radius) for y in rows]
    idx = int(np.argmin(vals))
    if vals[idx] > DARK_THRESHOLD:
        return None
    ordered = sorted(vals)
    if len(ordered) > 1 and (ordered[1] - ordered[0]) < MIN_GAP_TO_2ND and ordered[1] <= DARK_THRESHOLD:
        return None
    return idx

def read_stt(gray, stt_template, radius):
    if not stt_template:
        return None
    cols = stt_template.get("cols", [])
    if not cols and "col" in stt_template:
        cols = [stt_template["col"]]
    rows = stt_template.get("rows", [])
    digits = []
    for col_x in cols:
        d = read_digit_column(gray, rows, col_x, radius)
        if d is None:
            return None
        digits.append(str(d))
    return "".join(digits)

def read_made(gray, made_template, radius):
    if not made_template: return None
    digits = [read_digit_column(gray, r, c, radius) for c, r in zip(made_template['cols'], made_template['rows_per_col'])]
    return "".join(str(d) for d in digits) if None not in digits else None

def grade_page(page_img, template, answer_key, n_questions):
    warped, err = warp_to_template(page_img, template)
    result = {"ok": err is None, "error": err, "stt": None, "made": None,
              "dung": 0, "sai": 0, "bo_trong": 0, "khong_hop_le": 0, "diem": None, "name_crop": None}
    if err: return result

    gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
    radius = 12

    result["stt"] = read_stt(gray, template.get("stt_template"), radius)
    result["made"] = read_made(gray, template.get("made_template"), radius)

    dung = sai = bo_trong = khong_hop_le = 0
    bubble_radius = template.get("bubble_radius", 15) - 3
    for q in range(1, n_questions + 1):
        opts = template["questions"].get(str(q))
        if not opts: continue
        ans, is_multi = read_answer(gray, opts, bubble_radius)
        correct = answer_key.get(q)
        if is_multi: khong_hop_le += 1
        elif ans is None: bo_trong += 1
        elif correct is not None and ans == correct: dung += 1
        else: sai += 1

    result.update({"dung": dung, "sai": sai, "bo_trong": bo_trong, "khong_hop_le": khong_hop_le,
                   "diem": round(dung / n_questions * 100, 2) if n_questions else None})
    x0r, y0r, x1r, y1r = template["name_crop_rel"]
    W, H = template["page_size"]
    result["name_crop"] = warped[int(y0r*H):int(y1r*H), int(x0r*W):int(x1r*W)]
    return result

def export_excel(results, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ket qua"
    headers = ["Trang", "Ảnh (Tên/Lớp)", "STT", "Mã đề", "Số câu đúng", "Số câu sai", "Bỏ trống/Lỗi", "Điểm (/100)", "Ghi chú"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 30

    tmp_files = []
    for i, r in enumerate(results, start=2):
        ws.row_dimensions[i].height = 70
        ws.cell(row=i, column=1, value=r["page"])
        ws.cell(row=i, column=3, value=r["stt"] if r["stt"] is not None else "")
        ws.cell(row=i, column=4, value=r["made"] if r["made"] else "")
        ws.cell(row=i, column=5, value=r["dung"])
        ws.cell(row=i, column=6, value=r["sai"])
        ws.cell(row=i, column=7, value=r["bo_trong"] + r["khong_hop_le"])
        ws.cell(row=i, column=8, value=r["diem"])
        ws.cell(row=i, column=9, value=r["error"] or "")

        if r["name_crop"] is not None:
            fd, tmp_path = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            cv2.imwrite(tmp_path, r["name_crop"])
            tmp_files.append(tmp_path)
            img = XLImage(tmp_path)
            img.width = int(img.width * (90 / img.height))
            img.height = 90
            ws.add_image(img, f"B{i}")

    wb.save(out_path)
    for p in tmp_files:
        try: os.remove(p)
        except OSError: pass

# --- Streamlit UI ---
st.set_page_config(page_title="Chấm Thi Trắc Nghiệm Tự Động", layout="wide")
st.title("📝 Hệ Thống Chấm Điểm Trắc Nghiệm Tự Động")

with st.sidebar:
    st.header("⚙️ Cấu hình đề thi")
    ans_file = st.file_uploader("File đáp án (Excel/CSV)", type=["xlsx", "csv"])
    n_questions_input = st.number_input("Số câu hỏi bài thi", min_value=1, max_value=120, value=80)

pdf_file = st.file_uploader("📥 Tải lên file PDF bài làm", type=["pdf"])

if st.button("🚀 Bắt đầu chấm điểm", type="primary"):
    if not pdf_file or not ans_file:
        st.error("Vui lòng tải lên cả file PDF bài làm và file đáp án.")
    else:
        df_ans = pd.read_csv(ans_file, header=None) if ans_file.name.endswith(".csv") else pd.read_excel(ans_file, header=0)
        df_ans = df_ans.iloc[:, :2].dropna()
        df_ans.columns = ["cau", "dapan"]
        answer_key = {int(r["cau"]): str(r["dapan"]).strip().upper() for _, r in df_ans.iterrows() if str(r["dapan"]).strip().upper() in ("A","B","C","D")}

        template = load_default_template()
        n_questions = n_questions_input or max(answer_key.keys())

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(pdf_file.read())
            tmp_pdf_path = tmp_pdf.name

        doc = fitz.open(tmp_pdf_path)
        total_pages = len(doc)
        progress_bar = st.progress(0)
        status_text = st.empty()

        results = []
        for i, page in enumerate(doc):
            pw = page.rect.width
            zoom = template["page_size"][0] / pw if pw else 1.0
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            img = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)

            res = grade_page(img, template, answer_key, n_questions)
            res["page"] = i + 1
            results.append(res)
            
            progress_bar.progress((i + 1) / total_pages)
            status_text.text(f"Đang chấm trang {i + 1}/{total_pages}...")

        os.remove(tmp_pdf_path)
        out_excel = "ket_qua_cham.xlsx"
        export_excel(results, out_excel)
        status_text.success("✅ Đã hoàn tất chấm điểm!")

        summary_data = []
        for r in results:
            summary_data.append({
                "Trang": r["page"],
                "STT": r["stt"] or "Không nhận diện được",
                "Mã đề": r["made"] or "-",
                "Số câu đúng": r["dung"],
                "Số câu sai": r["sai"],
                "Lỗi/Trống": r["bo_trong"] + r["khong_hop_le"],
                "Điểm": r["diem"],
                "Trạng thái": "Thành công" if r["ok"] else r["error"]
            })
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

        with open(out_excel, "rb") as f:
            st.download_button("📥 Tải về file Excel kết quả chi tiết", data=f, file_name="Ket_qua_cham_trac_nghiem.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")