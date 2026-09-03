import os
import io
import math
import re
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk

# --- THƯ VIỆN CÀO DỮ LIỆU ---
from bs4 import BeautifulSoup
from curl_cffi import requests

# --- THƯ VIỆN XUẤT PDF VÀ XỬ LÝ ẢNH ---
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A5, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import simpleSplit, ImageReader 

try:
    import fitz  # PyMuPDF dùng để convert PDF sang PNG/JPG siêu nét và Review
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# Bật tính năng nhận diện DPI trên Windows
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

root = tk.Tk()
try: root.tk.call('tk', 'scaling', root.winfo_fpixels('1i')/72.0)
except: pass
root.withdraw()

# ==========================================
# CLASS TOOLTIP (HIỂN THỊ NHÃN KHI RÊ CHUỘT)
# ==========================================
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.widget.bind("<Enter>", self.showtip)
        self.widget.bind("<Leave>", self.hidetip)

    def showtip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 1
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                      background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                      font=("Arial", 9))
        label.pack(ipadx=3, ipady=2)

    def hidetip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# ==========================================
# CÁC HÀM CÀO DỮ LIỆU
# ==========================================
def parse_measurement(meas_str):
    if not meas_str: return "", "", ""
    cleaned = meas_str.lower().replace("mm", "").replace("x", " ").replace("-", " ")
    parts = [p for p in cleaned.split() if p.replace(".", "", 1).isdigit()]
    s1 = parts[0] if len(parts) > 0 else ""
    s2 = parts[1] if len(parts) > 1 else ""
    s3 = parts[2] if len(parts) > 2 else ""
    return s1, s2, s3

def get_p_span_value(soup, keyword, exclude=None):
    for p in soup.find_all("p"):
        clean_text = " ".join(p.text.split()).upper()
        if exclude and exclude.upper() in clean_text: continue
        if keyword.upper() in clean_text:
            span = p.find("span")
            if span: return span.text.strip()
    return ""

def get_table_dict(soup):
    data_dict = {}
    table = soup.find("table", class_=re.compile(r"bg-nen|daquy"))
    if not table:
        for t in soup.find_all("table"):
            if t.find("td", class_=re.compile(r"xanh2")):
                table = t
                break
    if not table: return data_dict
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) >= 2:
            label = " ".join(tds[0].text.split()).upper()
            val = tds[1].text.strip()
            if label and val: data_dict[label] = val
    return data_dict

def get_table_comments(soup):
    table = soup.find("table", class_=re.compile(r"bg-nen|daquy"))
    if not table: table = soup.find("table")
    if not table: return ""
    for td in table.find_all("td"):
        clean_text = " ".join(td.text.split()).upper()
        if "COMMENTS" in clean_text or "CHÚ THÍCH" in clean_text:
            span_xanh = td.find("span", class_="xanh")
            if span_xanh and span_xanh.text.strip(): return span_xanh.text.strip()
            text_parts = []
            for elem in td.children:
                if elem.name == "strong": continue
                if hasattr(elem, "text"): text_parts.append(elem.text.strip())
                elif isinstance(elem, str): text_parts.append(elem.strip())
            return " ".join(filter(None, text_parts)).strip()
    return ""

def find_in_dict(data_dict, *keywords):
    for key, val in data_dict.items():
        for kw in keywords:
            if kw.upper() in key: return val
    return ""

def scrape_pnj_item(session, pnj_code, prod_type, max_retries=3):
    pnj_code = str(int(pnj_code)) if isinstance(pnj_code, (int, float)) else str(pnj_code).strip()
    prod_type_str = str(prod_type).strip().lower()

    if "kc" in prod_type_str or "diamond" in prod_type_str: type_code = 1
    elif "dm" in prod_type_str or "gem" in prod_type_str: type_code = 2
    elif "nt" in prod_type_str or "pearl" in prod_type_str: type_code = 4
    else: return None

    url = f"https://pnjlab.com.vn/product-result/?ProductID={pnj_code}&type={type_code}"

    print(f"\n--- ĐANG KÉO DỮ LIỆU PNJ: {pnj_code} (Loại: {prod_type_str.upper()}) ---")
    for attempt in range(1, max_retries + 1):
        try:
            response = session.get(url, impersonate="chrome120", timeout=15)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                break
            elif response.status_code == 403:
                time.sleep(3)
            else:
                return None
        except Exception:
            time.sleep(2)
    else:
        return None

    row_data = {
        "Mã PNJ": pnj_code, "Loại sản phẩm": prod_type_str.upper(), "Tên đá/ Loại ngọc": "", "Weight (carat)": "",
        "size_1": "", "size_2": "", "size_3": "", "Màu sắc(Color)": "",
        "MSC_GIA": "", "MSC_PNJ": "0", "Đặc điểm": "", "Cut Grade": "",
        "Shape & Cut": "", "Clarity/Transparency": "", "Polish/Luster": "",
        "Symmetry/Surface": "", "Chú thích": ""
    }

    if type_code == 1:
        row_data["Tên đá/ Loại ngọc"] = "Diamond"
        carat_raw = get_p_span_value(soup, "CARAT WEIGHT") or get_p_span_value(soup, "carat")
        row_data["Weight (carat)"] = carat_raw.lower().replace("carat", "").replace("cts", "").strip()

        meas_raw = get_p_span_value(soup, "Measurement", exclude="MARGIN")
        s1, s2, s3 = parse_measurement(meas_raw)
        row_data["size_1"], row_data["size_2"], row_data["size_3"] = s1, s2, s3

        row_data["Màu sắc(Color)"] = get_p_span_value(soup, "COLOR GRADE")
        row_data["Cut Grade"] = get_p_span_value(soup, "CUT GRADE")
        row_data["Clarity/Transparency"] = get_p_span_value(soup, "CLARITY GRADE")
        row_data["Shape & Cut"] = get_p_span_value(soup, "Shape & Cut") or get_p_span_value(soup, "Shape")
        row_data["Polish/Luster"] = get_p_span_value(soup, "Polish")
        row_data["Symmetry/Surface"] = get_p_span_value(soup, "Symmetry")

        ins_val = get_p_span_value(soup, "Inscription")
        row_data["MSC_PNJ"] = ins_val if ins_val else "0"

        comments_div = soup.find(lambda t: t.name == "div" and "COMMENTS" in t.text.upper() and "title-c" in t.get("class", []))
        if comments_div:
            content_div = comments_div.find_next_sibling("div", class_="content")
            if content_div:
                comment_text = content_div.text.strip()
                row_data["Chú thích"] = comment_text
                match = re.search(r"GIA\s*(\d+)", comment_text)
                if match: row_data["MSC_GIA"] = match.group(1)

        char_div = soup.find(lambda t: t.name == "div" and "CLARITY CHARACTERISTICS" in t.text.upper() and "title-c" in t.get("class", []))
        if char_div:
            content_div = char_div.find_next_sibling("div", class_="content")
            if content_div and content_div.find("p"): row_data["Đặc điểm"] = content_div.find("p").text.strip()

    elif type_code in [2, 4]:
        td_dict = get_table_dict(soup)
        raw_w = find_in_dict(td_dict, "WEIGHT", "KHỐI LƯỢNG")
        row_data["Weight (carat)"] = raw_w.lower().replace("cts", "").replace("carat", "").strip()

        raw_m = find_in_dict(td_dict, "MEASUREMENT", "KÍCH THƯỚC")
        s1, s2, s3 = parse_measurement(raw_m)
        row_data["size_1"], row_data["size_2"], row_data["size_3"] = s1, s2, s3

        row_data["Màu sắc(Color)"] = find_in_dict(td_dict, "COLOR", "MÀU SẮC")
        row_data["Chú thích"] = get_table_comments(soup)

        if type_code == 2:  
            row_data["Tên đá/ Loại ngọc"] = find_in_dict(td_dict, "VARIETY", "TÊN ĐÁ")
            row_data["Đặc điểm"] = find_in_dict(td_dict, "GROUP/SPECIES", "SPECIES", "GROUP", "TÊN NHÓM")
            row_data["Shape & Cut"] = find_in_dict(td_dict, "SHAPE - CUTTING STYLE", "SHAPE", "DẠNG CẮT MÀI")
            row_data["Clarity/Transparency"] = find_in_dict(td_dict, "TRANSPARENCY", "ĐỘ TRONG")
        elif type_code == 4:  
            row_data["Tên đá/ Loại ngọc"] = find_in_dict(td_dict, "PEARL", "LOẠI NGỌC")
            row_data["Đặc điểm"] = find_in_dict(td_dict, "ENVIROMENT", "ENVIRONMENT", "MÔI TRƯỜNG")
            row_data["Shape & Cut"] = find_in_dict(td_dict, "SHARP", "SHAPE", "HÌNH DẠNG")
            row_data["Polish/Luster"] = find_in_dict(td_dict, "LUSTER", "ĐỘ BÓNG")
            row_data["Symmetry/Surface"] = find_in_dict(td_dict, "SURFACE", "BỀ MẶT")

    print("✅ Kéo dữ liệu thành công!")
    return row_data

# ==========================================
# CÁC HÀM GIAO DIỆN CHỌN FILE
# ==========================================
def create_round_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    points = [x1+radius, y1, x1+radius, y1, x2-radius, y1, x2-radius, y1, x2, y1, x2, y1+radius, x2, y1+radius, x2, y2-radius, x2, y2-radius, x2, y2, x2-radius, y2, x2-radius, y2, x1+radius, y2, x1+radius, y2, x1, y2, x1, y2-radius, x1, y2-radius, x1, y1+radius, x1, y1+radius, x1, y1]
    return canvas.create_polygon(points, smooth=True, **kwargs)

def center_window(window, pad_w=60, pad_h=60):
    window.update_idletasks()
    w = window.winfo_reqwidth() + pad_w
    h = window.winfo_reqheight() + pad_h
    x = (window.winfo_screenwidth() // 2) - (w // 2)
    y = (window.winfo_screenheight() // 2) - (h // 2)
    window.geometry(f'{w}x{h}+{x}+{y}')

def show_dropdown_ui(image_files):
    result = {"file": None, "so_tn": "", "gia_no": "", "loai_da": "kc"}
    window = tk.Toplevel(root)
    window.title("Chọn ảnh (Menu thả xuống)")

    tk.Label(window, text="Vui lòng chọn file ảnh từ danh sách:", font=("Arial", 10, "bold")).pack(pady=(15, 5))
    selected_file = tk.StringVar()
    dropdown = ttk.Combobox(window, textvariable=selected_file, values=image_files, state="readonly", width=35)
    dropdown.current(0)
    dropdown.pack()

    tk.Label(window, text="Loại sản phẩm (Chọn 1 loại):", font=("Arial", 10, "bold")).pack(pady=(15, 2))
    type_frame = tk.Frame(window)
    type_frame.pack()
    loai_da_var = tk.StringVar(value="kc")
    tk.Radiobutton(type_frame, text="Kim Cương", variable=loai_da_var, value="kc").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(type_frame, text="Đá Màu", variable=loai_da_var, value="dm").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(type_frame, text="Ngọc Trai", variable=loai_da_var, value="nt").pack(side=tk.LEFT, padx=5)

    tk.Label(window, text="Số TN (Mã PNJ để lấy data):", font=("Arial", 10)).pack(pady=(10, 2))
    entry_tn = tk.Entry(window, width=38, font=("Arial", 10))
    entry_tn.pack()

    tk.Label(window, text="GIA No. (Tùy chọn):", font=("Arial", 10)).pack(pady=(10, 2))
    entry_gia = tk.Entry(window, width=38, font=("Arial", 10))
    entry_gia.pack()

    def on_ok():
        result["file"] = selected_file.get()
        result["so_tn"] = entry_tn.get()
        result["gia_no"] = entry_gia.get()
        result["loai_da"] = loai_da_var.get()
        window.destroy()

    def on_cancel():
        result["file"] = "back" 
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_cancel) 
    btn_frame = tk.Frame(window)
    btn_frame.pack(pady=20)
    tk.Button(btn_frame, text="✔ Mở Ảnh", command=on_ok, width=12, bg="#d4edda", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="↩ Trở về", command=on_cancel, width=12, bg="#f8d7da", font=("Arial", 10)).pack(side=tk.RIGHT, padx=10)
    
    center_window(window)
    window.wait_window()
    return result

def show_listbox_ui(image_files):
    result = {"file": None, "so_tn": "", "gia_no": "", "loai_da": "kc"}
    window = tk.Toplevel(root)
    window.title("Chọn ảnh (Danh sách cuộn)")

    tk.Label(window, text="Chọn file ảnh:", font=("Arial", 10, "bold")).pack(pady=(15, 5))
    
    frame = tk.Frame(window)
    frame.pack(padx=20, fill=tk.BOTH, expand=True)
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, width=40, height=6, font=("Arial", 10))
    for item in image_files: listbox.insert(tk.END, item)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)
    listbox.selection_set(0)

    tk.Label(window, text="Loại sản phẩm (Chọn 1 loại):", font=("Arial", 10, "bold")).pack(pady=(15, 2))
    type_frame = tk.Frame(window)
    type_frame.pack()
    loai_da_var = tk.StringVar(value="kc")
    tk.Radiobutton(type_frame, text="Kim Cương", variable=loai_da_var, value="kc").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(type_frame, text="Đá Màu", variable=loai_da_var, value="dm").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(type_frame, text="Ngọc Trai", variable=loai_da_var, value="nt").pack(side=tk.LEFT, padx=5)

    tk.Label(window, text="Số TN (Mã PNJ để lấy data):", font=("Arial", 10)).pack(pady=(10, 2))
    entry_tn = tk.Entry(window, width=38, font=("Arial", 10))
    entry_tn.pack()

    tk.Label(window, text="GIA No. (Tùy chọn):", font=("Arial", 10)).pack(pady=(10, 2))
    entry_gia = tk.Entry(window, width=38, font=("Arial", 10))
    entry_gia.pack()

    def on_ok(event=None):
        selection = listbox.curselection()
        if selection:
            result["file"] = listbox.get(selection[0])
            result["so_tn"] = entry_tn.get()
            result["gia_no"] = entry_gia.get()
            result["loai_da"] = loai_da_var.get()
            window.destroy()

    def on_cancel():
        result["file"] = "back" 
        window.destroy()

    listbox.bind("<Double-Button-1>", on_ok)
    window.protocol("WM_DELETE_WINDOW", on_cancel) 
    btn_frame = tk.Frame(window)
    btn_frame.pack(pady=20)
    tk.Button(btn_frame, text="✔ Mở Ảnh", command=on_ok, width=12, bg="#d4edda", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="↩ Trở về", command=on_cancel, width=12, bg="#f8d7da", font=("Arial", 10)).pack(side=tk.RIGHT, padx=10)
    
    center_window(window)
    window.wait_window()
    return result

def get_target_file(target_dir):
    image_files = [f for f in os.listdir(target_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.jfif'))]
    if not image_files:
        messagebox.showerror("Lỗi", "Không tìm thấy file ảnh nào trong thư mục!")
        return None

    while True: 
        style_result = {"style": "exit"}
        choose_win = tk.Toplevel(root)
        choose_win.title("Menu Chính")

        tk.Label(choose_win, text="Bạn muốn chọn ảnh theo kiểu nào?", font=("Arial", 11, "bold")).pack(pady=20)
        def set_style(s):
            style_result["style"] = s
            choose_win.destroy()
        def on_close():
            style_result["style"] = "exit"
            choose_win.destroy()
            
        choose_win.protocol("WM_DELETE_WINDOW", on_close)
        btn_frame = tk.Frame(choose_win)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="1. Menu thả xuống", command=lambda: set_style("dropdown"), width=18, bg="#cce5ff", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="2. Danh sách cuộn", command=lambda: set_style("listbox"), width=18, bg="#fff3cd", font=("Arial", 10)).pack(side=tk.RIGHT, padx=10)
        tk.Button(choose_win, text="✖ Đóng", command=on_close, width=12, bg="#f8d7da", font=("Arial", 10)).pack(pady=15)
        
        center_window(choose_win)
        choose_win.wait_window()

        style = style_result["style"]
        if style == "exit": return None 
        data = None
        if style == "dropdown": data = show_dropdown_ui(image_files)
        elif style == "listbox": data = show_listbox_ui(image_files)
        if data and data["file"] == "back": continue 
        else: return data

def generate_zigzag(points, step, amplitude):
    resampled = [points[0]]
    for p in points[1:]:
        last = resampled[-1]
        if math.hypot(p[0]-last[0], p[1]-last[1]) >= step: resampled.append(p)
    if len(resampled) < 2: return []
    zz, sign = [], 1
    for i in range(len(resampled)-1):
        x1, y1 = resampled[i]
        x2, y2 = resampled[i+1]
        zz.extend([x1, y1])
        dx, dy = x2 - x1, y2 - y1
        dist = math.hypot(dx, dy)
        if dist == 0: continue
        nx, ny = -dy/dist, dx/dist
        mx = x1 + dx/2 + nx * amplitude * sign
        my = y1 + dy/2 + ny * amplitude * sign
        zz.extend([mx, my])
        sign *= -1
    zz.extend([resampled[-1][0], resampled[-1][1]])
    return zz

# ==========================================
# KHỞI CHẠY BẢN VẼ VÀ RENDER THÔNG TIN
# ==========================================
def run_mapping_tool(image_path, so_tn, gia_no, loai_da, gem_data):
    mapping_win = tk.Toplevel(root)
    mapping_win.title(f"Bản Đồ Tỳ Vết - {os.path.basename(image_path)}")
    
    actual_photos = []
    screen_w = mapping_win.winfo_screenwidth()
    screen_h = mapping_win.winfo_screenheight()
    available_canvas_h = screen_h - 350 

    try:
        img = Image.open(image_path)
        orig_w, orig_h = img.size
        img_ratio = orig_h / float(orig_w) 
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể mở ảnh: {e}")
        mapping_win.destroy()
        return

    img_w = int(screen_w * 0.42) 
    img_w = max(500, min(img_w, 900)) 
    img_h = int(img_w * img_ratio) 
    
    if (img_h + 160) > available_canvas_h:
        img_h = available_canvas_h - 160
        img_w = int(img_h / img_ratio)

    if gem_data:
        info_w = max(380, int(screen_w * 0.3))
        canvas_w = img_w + info_w
    else:
        info_w = 60 
        canvas_w = img_w + info_w

    canvas_h = img_h + 160
    
    img = img.resize((img_w, int(img_h)), Image.Resampling.LANCZOS)
    photo = ImageTk.PhotoImage(img)

    mapping_win.update_idletasks()
    win_w = canvas_w + 40
    win_h = canvas_h + 230 
    
    x = (screen_w // 2) - (win_w // 2)
    y = (screen_h // 2) - (win_h // 2) - 20
    if y < 0: y = 10
        
    mapping_win.geometry(f'{int(win_w)}x{int(win_h)}+{int(x)}+{int(y)}')

    toolbar_bottom = tk.Frame(mapping_win)
    toolbar_bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(2, 10))

    toolbar_top = tk.Frame(mapping_win)
    toolbar_top.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5, 2))

    canvas = tk.Canvas(mapping_win, width=canvas_w, height=canvas_h, bg="white", cursor="none") 
    canvas.pack(side=tk.TOP, pady=(10, 0), expand=True)

    margin, padding = 15, 25
    canvas.create_rectangle(margin, margin, canvas_w - margin, canvas_h - margin, dash=(6, 4), outline="gray")
    create_round_rect(canvas, padding, padding, canvas_w - padding, canvas_h - padding, radius=20, outline="black", fill="", width=1.5)
    
    divider_x = padding + img_w + 20
    if gem_data:
        canvas.create_line(divider_x, padding, divider_x, canvas_h - padding, fill="black", width=1.5)

    ui_cx = padding + 10 + img_w / 2.0
    ui_cy = canvas_h / 2.0

    img_x_start = padding + 10
    img_y_start = (canvas_h / 2) - (img_h / 2)
    canvas.create_rectangle(img_x_start, img_y_start, img_x_start + img_w, img_y_start + img_h, outline="#666666", width=1.5, dash=(4, 2))

    canvas.create_image(padding + 10, canvas_h // 2, anchor=tk.W, image=photo)
    canvas.image = photo 
    
    text_top = f"Số TN: {so_tn}                      GIA No.: {gia_no}"
    canvas.create_text(padding + 20, padding + 20, text=text_top, anchor=tk.W, font=("Arial", 12, "bold"))
    
    txt_ghi_chu = "Ghi chú: " + "." * 90
    canvas.create_text(padding + 20, canvas_h - padding - 25, text=txt_ghi_chu, anchor=tk.W, font=("Arial", 12), tags="note_display")

    info_lines = []
    if gem_data:
        title_x = divider_x + 20
        start_y = padding + 30
        
        canvas.create_text(title_x, start_y, text="THÔNG SỐ ĐÁ QUÝ", anchor=tk.NW, font=("Arial", 13, "bold"), fill="#003366")
        start_y += 35
        
        s1, s2, s3 = gem_data.get('size_1', ''), gem_data.get('size_2', ''), gem_data.get('size_3', '')
        dims = [s for s in [s1, s2, s3] if s]
        kich_thuoc = " x ".join(dims) + " mm" if dims else ""

        info_lines = [
            ("Số TN (Mã PNJ):", gem_data.get('Mã PNJ', '')),
            ("GIA No.:", gia_no if gia_no else gem_data.get('MSC_GIA', '')),
            ("Loại đá:", gem_data.get('Tên đá/ Loại ngọc', '')),
            ("Kích thước:", kich_thuoc),
            ("Trọng lượng:", f"{gem_data.get('Weight (carat)', '')} ct" if gem_data.get('Weight (carat)') else ""),
            ("Hình dạng:", gem_data.get('Shape & Cut', '')),
            ("Màu sắc:", gem_data.get('Màu sắc(Color)', '')),
            ("Độ tinh khiết:", gem_data.get('Clarity/Transparency', '')),
            ("Chất lượng cắt:", gem_data.get('Cut Grade', '')),
            ("Độ bóng:", gem_data.get('Polish/Luster', '')),
            ("Độ đối xứng:", gem_data.get('Symmetry/Surface', '')),
            ("Đặc điểm:", gem_data.get('Đặc điểm', '')),
            ("Chú thích:", gem_data.get('Chú thích', ''))
        ]
        
        temp_canvas = tk.Canvas()
        max_title_width = 0
        for t_str, val in info_lines:
            bb = temp_canvas.bbox(temp_canvas.create_text(0, 0, text=t_str, font=("Arial", 10, "bold")))
            if bb and (bb[2] - bb[0]) > max_title_width:
                max_title_width = bb[2] - bb[0]
                
        value_x = title_x + max_title_width + 10
        max_val_w = canvas_w - margin - value_x - 10 
        
        for title, val in info_lines:
            canvas.create_text(title_x, start_y, text=title, anchor=tk.NW, font=("Arial", 10, "bold"))
            val_id = canvas.create_text(value_x, start_y, text=val, anchor=tk.NW, font=("Arial", 10), width=max_val_w)
            
            bbox = canvas.bbox(val_id)
            val_h = (bbox[3] - bbox[1]) if bbox else 15
            start_y += max(20, val_h + 8)

    # --- LOGIC CÔNG CỤ VẼ TỲ VẾT ---
    current_tool = tk.StringVar(value="pen_blue")
    size_scale = tk.IntVar(value=4)
    angle_scale = tk.IntVar(value=0)
    note_var = tk.StringVar(value="") 
    
    draw_points, temp_items = [], []
    is_drawing = False
    history, current_action = [], []

    def update_note_text(*args):
        val = note_var.get().strip()
        if val: canvas.itemconfig("note_display", text=f"Ghi chú: {val}")
        else: canvas.itemconfig("note_display", text="Ghi chú: " + "." * 90)
            
    note_var.trace_add("write", update_note_text)
    def get_sz(): return size_scale.get()
    def get_angle(): return angle_scale.get()

    def draw_gia_symbol(cvs, tool, x, y, sz, angle_deg, tag_name):
        items = []
        r = max(2, sz)
        rad = math.radians(angle_deg)
        
        def rot(pts):
            rotated = []
            for px, py in pts:
                nx = x + px * math.cos(rad) - py * math.sin(rad)
                ny = y + px * math.sin(rad) + py * math.cos(rad)
                rotated.extend([nx, ny])
            return rotated

        def oval_pts(rx, ry, n=36):
            return [(rx * math.cos(2 * math.pi * i / n), ry * math.sin(2 * math.pi * i / n)) for i in range(n)]

        if tool == 'knot':
            items.append(cvs.create_polygon(*rot(oval_pts(r, r/1.5)), outline="green", fill="", width=1.5, smooth=True, tags=tag_name))
            items.append(cvs.create_polygon(*rot(oval_pts(r*0.6, r*0.4)), outline="red", fill="", width=1.5, smooth=True, tags=tag_name))
        elif tool == 'cavity':
            items.append(cvs.create_polygon(*rot(oval_pts(r, r/1.5)), outline="green", fill="", width=1.5, smooth=True, tags=tag_name))
            items.append(cvs.create_line(*rot([(-r*0.5, -r*0.4), (-r*0.1, r*0.4)]), fill="red", width=1.5, tags=tag_name))
            items.append(cvs.create_line(*rot([(-r*0.1, -r*0.4), (r*0.3, r*0.4)]), fill="red", width=1.5, tags=tag_name))
            items.append(cvs.create_line(*rot([(r*0.3, -r*0.4), (r*0.7, r*0.2)]), fill="red", width=1.5, tags=tag_name))
        elif tool == 'etch_channel':
            items.append(cvs.create_polygon(*rot([(-r, -r), (r, -r), (r, r), (-r, r)]), outline="red", fill="", width=1.5, tags=tag_name))
            items.append(cvs.create_polygon(*rot([(-r/2, -r/2), (r/2, -r/2), (r/2, r/2), (-r/2, r/2)]), outline="green", fill="", width=1.5, tags=tag_name))
        elif tool == 'indented_natural':
            items.append(cvs.create_line(*rot([(-r, -r/2), (0, -r*1.5), (r, -r/2)]), fill="red", width=1.5, joinstyle=tk.MITER, tags=tag_name))
            items.append(cvs.create_line(*rot([(-r, 0), (0, -r), (r, 0)]), fill="green", width=1.5, joinstyle=tk.MITER, tags=tag_name))
        elif tool == 'natural':
            items.append(cvs.create_line(*rot([(-r, 0), (0, -r), (r, 0)]), fill="green", width=1.5, joinstyle=tk.MITER, tags=tag_name))
        elif tool == 'extra_facet':
            items.append(cvs.create_line(*rot([(-r, 0), (0, -r), (r, 0)]), fill="black", width=1.5, joinstyle=tk.MITER, tags=tag_name))
        elif tool == 'chip':
            items.append(cvs.create_line(*rot([(-r, 0), (0, -r), (r, 0)]), fill="red", width=1.5, joinstyle=tk.MITER, tags=tag_name))
        elif tool == 'bruise':
            items.append(cvs.create_line(*rot([(-r/2, -r/2), (r/2, r/2)]), fill="red", width=1.5, tags=tag_name))
            items.append(cvs.create_line(*rot([(-r/2, r/2), (r/2, -r/2)]), fill="red", width=1.5, tags=tag_name))
        elif tool == 'crystal':
            items.append(cvs.create_polygon(*rot(oval_pts(r, r)), outline="red", fill="", width=1.5, smooth=True, tags=tag_name))
        elif tool == 'pinpoint':
            items.append(cvs.create_polygon(*rot(oval_pts(1.5, 1.5, 12)), outline="red", fill="red", tags=tag_name))
        elif tool == 'laser_drill':
            items.append(cvs.create_polygon(*rot(oval_pts(r, r)), outline="green", fill="", width=1.5, smooth=True, tags=tag_name))
            items.append(cvs.create_polygon(*rot(oval_pts(1.5, 1.5, 12)), outline="red", fill="red", tags=tag_name))
        elif tool == 'cloud':
            items.append(cvs.create_polygon(*rot(oval_pts(r, r)), outline="red", fill="", width=1.5, dash=(1, 3), smooth=True, tags=tag_name))

        return items

    gia_click_tools = ['crystal', 'pinpoint', 'laser_drill', 'knot', 'cavity', 'etch_channel', 'bruise', 'chip', 'natural', 'indented_natural', 'extra_facet', 'cloud']

    def on_hover(event):
        canvas.delete("preview")
        if is_drawing or event.x > divider_x: return  
        tool, sz, angle_deg = current_tool.get(), get_sz(), get_angle()
        x, y = event.x, event.y
        r = max(2, sz)

        canvas.create_line(x-5, y, x+5, y, fill="gray", tags="preview")
        canvas.create_line(x, y-5, x, y+5, fill="gray", tags="preview")

        if tool in ['P', 'G']: 
            canvas.create_text(x, y, text=tool, font=("Arial", sz*3, "bold"), fill="#aaaaaa", angle=angle_deg, tags="preview")
        elif tool == 'V':
            rad, spread, l = math.radians(angle_deg), math.radians(35), sz * 2.5
            dx1, dy1 = l * math.sin(rad - spread), -l * math.cos(rad - spread)
            dx2, dy2 = l * math.sin(rad + spread), -l * math.cos(rad + spread)
            canvas.create_line(x, y, x + dx1, y + dy1, fill="#ff8888", width=2, capstyle=tk.ROUND, tags="preview")
            canvas.create_line(x, y, x + dx2, y + dy2, fill="#ff8888", width=2, capstyle=tk.ROUND, tags="preview")
        elif tool == '+':
            rad = math.radians(angle_deg)
            l = sz * 1.5
            dx1, dy1 = l * math.cos(rad), l * math.sin(rad)
            dx2, dy2 = -l * math.sin(rad), l * math.cos(rad)
            canvas.create_line(x-dx1, y-dy1, x+dx1, y+dy1, fill="#ff8888", width=2, tags="preview")
            canvas.create_line(x-dx2, y-dy2, x+dx2, y+dy2, fill="#ff8888", width=2, tags="preview")
        elif tool == 'dot_red':
            canvas.create_oval(x-r, y-r, x+r, y+r, outline="#ff8888", width=2, tags="preview")
            
        elif tool in gia_click_tools:
            draw_gia_symbol(canvas, tool, x, y, sz, angle_deg, "preview")
            
        elif tool in ['pen_blue', 'line_red', 'zigzag', 'needle', 'twinning_wisp']:
            canvas.create_oval(x-r, y-r, x+r, y+r, outline="gray", dash=(2,2), tags="preview")
        elif tool == 'feather':
            canvas.create_oval(x-r, y-r, x+r, y+r, outline="red", dash=(2,2), tags="preview")
        elif tool == 'eraser':
            r = sz * 2
            canvas.create_rectangle(x-r, y-r, x+r, y+r, outline="gray", dash=(2,2), tags="preview")

    def on_click(event):
        nonlocal is_drawing, current_action
        canvas.focus_set()
        if event.x > divider_x: return 
        is_drawing = True
        canvas.delete("preview") 
        current_action = [] 
        
        tool, sz, angle_deg = current_tool.get(), get_sz(), get_angle()
        x, y = event.x, event.y
        r = max(2, sz)

        if tool == 'P': history.append([canvas.create_text(x, y, text="P", font=("Arial", sz*3, "bold"), fill="black", angle=angle_deg, tags="drawing")])
        elif tool == 'G': history.append([canvas.create_text(x, y, text="G", font=("Arial", sz*3, "bold"), fill="black", angle=angle_deg, tags="drawing")])
        
        elif tool == 'V':
            rad, spread, l = math.radians(angle_deg), math.radians(35), sz * 2.5
            dx1, dy1 = l * math.sin(rad - spread), -l * math.cos(rad - spread)
            dx2, dy2 = l * math.sin(rad + spread), -l * math.cos(rad + spread)
            history.append([canvas.create_line(x, y, x + dx1, y + dy1, fill="red", width=2, capstyle=tk.ROUND, tags="drawing"),
                            canvas.create_line(x, y, x + dx2, y + dy2, fill="red", width=2, capstyle=tk.ROUND, tags="drawing")])
        elif tool == '+':
            rad = math.radians(angle_deg)
            l = sz * 1.5
            dx1, dy1 = l * math.cos(rad), l * math.sin(rad)
            dx2, dy2 = -l * math.sin(rad), l * math.cos(rad)
            history.append([canvas.create_line(x-dx1, y-dy1, x+dx1, y+dy1, fill="red", width=2, tags="drawing"),
                            canvas.create_line(x-dx2, y-dy2, x+dx2, y+dy2, fill="red", width=2, tags="drawing")])
        elif tool == 'dot_red':
            history.append([canvas.create_oval(x-r, y-r, x+r, y+r, fill="red", outline="red", tags="drawing")])

        elif tool in gia_click_tools and tool != 'cloud':
            items = draw_gia_symbol(canvas, tool, x, y, sz, angle_deg, "drawing")
            if items: history.append(items)
        elif tool == 'cloud':
            items = draw_gia_symbol(canvas, tool, x, y, sz, angle_deg, "drawing")
            current_action.extend(items)
            draw_points.clear()
            draw_points.append((x, y))
        
        elif tool in ['pen_blue', 'line_red', 'zigzag', 'needle', 'feather', 'twinning_wisp']:
            draw_points.clear()
            draw_points.append((x, y))
            if tool == 'pen_blue':
                current_action.append(canvas.create_oval(x-r/1.5, y-r/1.5, x+r/1.5, y+r/1.5, fill="blue", outline="blue", tags="drawing"))
        
        elif tool == 'eraser':
            r = sz * 2
            items = canvas.find_overlapping(x - r, y - r, x + r, y + r)
            for item in items:
                if "drawing" in canvas.gettags(item): canvas.delete(item)

    def on_drag(event):
        tool, sz = current_tool.get(), get_sz()
        if event.x > divider_x: return
        
        if tool == 'pen_blue' and draw_points:
            last_x, last_y = draw_points[-1]
            current_action.append(canvas.create_line(last_x, last_y, event.x, event.y, fill="blue", width=max(1, sz), capstyle=tk.ROUND, smooth=True, tags="drawing"))
            draw_points.append((event.x, event.y))
            
        elif tool in ['line_red', 'needle'] and draw_points:
            for item in temp_items: canvas.delete(item)
            temp_items.clear()
            temp_items.append(canvas.create_line(draw_points[0][0], draw_points[0][1], event.x, event.y, fill="red", width=max(1, sz), tags=("drawing", "temp_line")))
            
        elif tool == 'zigzag' and draw_points:
            last_x, last_y = draw_points[-1]
            temp_items.append(canvas.create_line(last_x, last_y, event.x, event.y, fill="pink", width=1, dash=(2,2), tags=("drawing", "temp_zig")))
            draw_points.append((event.x, event.y))
            
        elif tool in ['cloud', 'feather', 'twinning_wisp'] and draw_points:
            if tool == 'cloud' and len(draw_points) == 1 and current_action:
                for item in current_action: canvas.delete(item)
                current_action.clear()
                
            last_x, last_y = draw_points[-1]
            dash_pattern = (1, 3) if tool == 'cloud' else None
            current_action.append(canvas.create_line(last_x, last_y, event.x, event.y, fill="red", width=max(1, sz), dash=dash_pattern, capstyle=tk.ROUND, smooth=True, tags="drawing"))
            draw_points.append((event.x, event.y))
            
        elif tool == 'eraser':
            r = sz * 2
            items = canvas.find_overlapping(event.x - r, event.y - r, event.x + r, event.y + r)
            for item in items:
                if "drawing" in canvas.gettags(item): canvas.delete(item)

    def on_release(event):
        nonlocal is_drawing, current_action
        tool, sz = current_tool.get(), get_sz()
        
        if tool == 'zigzag':
            for item in temp_items: canvas.delete(item)
            temp_items.clear()
            if len(draw_points) > 1:
                zz = generate_zigzag(draw_points, step=sz*2 + 2, amplitude=sz*1.5)
                if zz: history.append([canvas.create_line(*zz, fill="red", width=max(1, int(sz*0.8)), joinstyle=tk.MITER, tags="drawing")])
            draw_points.clear()
            
        elif tool in ['line_red', 'needle']:
            for item in temp_items: canvas.delete(item)
            temp_items.clear()
            if len(draw_points) > 0:
                history.append([canvas.create_line(draw_points[0][0], draw_points[0][1], event.x, event.y, fill="red", width=max(1, sz), tags="drawing")])
            draw_points.clear()
            
        elif tool == 'twinning_wisp':
            if current_action:
                step = 10
                resampled = [draw_points[0]]
                for p in draw_points[1:]:
                    if math.hypot(p[0]-resampled[-1][0], p[1]-resampled[-1][1]) >= step:
                        resampled.append(p)
                for i in range(len(resampled)-1):
                    x1, y1 = resampled[i]
                    x2, y2 = resampled[i+1]
                    dx, dy = x2 - x1, y2 - y1
                    dist = math.hypot(dx, dy)
                    if dist == 0: continue
                    nx, ny = -dy/dist, dx/dist
                    cx, cy = x1 + dx/2, y1 + dy/2
                    l = sz * 1.5
                    current_action.append(canvas.create_line(cx - nx*l, cy - ny*l, cx + nx*l, cy + ny*l, fill="red", width=max(1, sz), tags="drawing"))
                history.append(current_action)
            current_action = []
            draw_points.clear()
            
        elif tool in ['feather', 'cloud', 'pen_blue']:
            if current_action: history.append(current_action)
            current_action = []
            draw_points.clear()
            
        is_drawing = False
        on_hover(event) 

    def undo_action(event=None):
        if isinstance(mapping_win.focus_get(), tk.Entry): return 
        if history:
            for item in history.pop():
                try: canvas.delete(item)
                except: pass

    mapping_win.bind("<Control-z>", undo_action)
    mapping_win.bind("<Control-Z>", undo_action)
    canvas.bind("<Motion>", on_hover)
    canvas.bind("<Leave>", lambda e: canvas.delete("preview"))
    canvas.bind("<Button-1>", on_click)        
    canvas.bind("<B1-Motion>", on_drag)        
    canvas.bind("<ButtonRelease-1>", on_release)  

    # --- HỘP THOẠI LỰA CHỌN ẢNH THỰC TẾ (CÓ THUMBNAILS, TEXT WRAP VÀ TOOLTIP) ---
    def select_photos_folder():
        nonlocal actual_photos
        folder = filedialog.askdirectory(parent=mapping_win, title="Chọn thư mục chứa ảnh chụp tỳ vết")
        if not folder: return

        valid_exts = ('.png', '.jpg', '.jpeg', '.jfif')
        all_files = []
        for root_dir, dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(valid_exts):
                    all_files.append(os.path.join(root_dir, f))

        if not all_files:
            messagebox.showinfo("Thông báo", "Không tìm thấy ảnh nào trong thư mục này!")
            return

        select_win = tk.Toplevel(mapping_win)
        select_win.title(f"Lựa Chọn Hình Ảnh ({len(all_files)} ảnh được tìm thấy)")
        select_win.geometry("750x550")
        select_win.transient(mapping_win)
        select_win.grab_set()

        main_frame = tk.Frame(select_win)
        main_frame.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)

        canvas_scroll = tk.Canvas(main_frame)
        canvas_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas_scroll.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas_scroll.configure(yscrollcommand=scrollbar.set)
        canvas_scroll.bind('<Configure>', lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all")))

        grid_frame = tk.Frame(canvas_scroll)
        canvas_scroll.create_window((0, 0), window=grid_frame, anchor="nw")

        def _on_mousewheel(event):
            canvas_scroll.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas_scroll.bind_all("<MouseWheel>", _on_mousewheel)

        checkbox_vars = []
        select_win.thumbnails = []

        cols = 4
        for idx, img_path in enumerate(all_files):
            row = idx // cols
            col = idx % cols
            
            cell = tk.Frame(grid_frame, borderwidth=1, relief="solid", padx=5, pady=5)
            cell.grid(row=row, column=col, padx=10, pady=10)
            
            name = os.path.basename(img_path)
            
            # --- HIỂN THỊ HÌNH ẢNH KÈM TOOLTIP ---
            try:
                img_pil = Image.open(img_path)
                img_pil.thumbnail((120, 120))
                photo_thumb = ImageTk.PhotoImage(img_pil)
                select_win.thumbnails.append(photo_thumb)
                lbl_img = tk.Label(cell, image=photo_thumb)
                lbl_img.pack()
                ToolTip(lbl_img, name) # Gắn Tooltip cho hình ảnh
            except:
                lbl_err = tk.Label(cell, text="[Lỗi file ảnh]", width=15, height=5, fg="red")
                lbl_err.pack()
                ToolTip(lbl_err, name) # Gắn Tooltip cho nhãn lỗi nếu có
                
            var = tk.BooleanVar(value=True)
            checkbox_vars.append((var, img_path))
            
            # --- XỬ LÝ TEXT WRAP VÀ CẮT CHỮ KÈM TOOLTIP ---
            # Với font Arial size 9 và wraplength 110px, 1 dòng khoảng ~16 ký tự. 2 dòng ~32 ký tự.
            display_name = name if len(name) <= 32 else name[:29] + "..."
            chk = tk.Checkbutton(cell, text=display_name, variable=var, font=("Arial", 9), wraplength=110, justify="center")
            chk.pack()
            ToolTip(chk, name) # Gắn Tooltip cho phần chữ Checkbox

        btn_frame = tk.Frame(select_win)
        btn_frame.pack(fill=tk.X, pady=10)

        def select_all():
            for v, _ in checkbox_vars: v.set(True)
        def deselect_all():
            for v, _ in checkbox_vars: v.set(False)

        def on_confirm():
            actual_photos.clear()
            for v, path in checkbox_vars:
                if v.get(): actual_photos.append(path)
            lbl_photo_count.config(text=f"{len(actual_photos)} ảnh")
            canvas_scroll.unbind_all("<MouseWheel>")
            select_win.destroy()

        tk.Button(btn_frame, text="Chọn tất cả", command=select_all, width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Bỏ chọn tất cả", command=deselect_all, width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="✔ Lưu lựa chọn", command=on_confirm, bg="#d4edda", font=("Arial", 10, "bold"), width=15).pack(side=tk.RIGHT, padx=20)

        center_window(select_win, 0, 0)

    # ==========================================
    # HÀM TẠO NHÁP PDF VÀO RAM
    # ==========================================
    def generate_pdf_buffer():
        buffer = io.BytesIO()
        try:
            c = pdf_canvas.Canvas(buffer, pagesize=landscape(A5))
            w, h = landscape(A5) 
            
            try:
                pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
                font_name = 'Arial'
                try:
                    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
                    font_bold = 'Arial-Bold'
                except: font_bold = 'Arial'
            except:
                font_name = 'Helvetica'
                font_bold = 'Helvetica-Bold'
                
            margin_pdf = 10 * mm
            
            if gem_data:
                map_box_h = 55 * mm  
                map_box_y = h - margin_pdf - map_box_h
            else:
                map_box_h = h - 2 * margin_pdf  
                map_box_y = margin_pdf

            left_w = (w / 2) - margin_pdf
            map_box_x = margin_pdf
            map_box_w = left_w
            
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0, 0, 0)
            c.roundRect(map_box_x, map_box_y, map_box_w, map_box_h, 2*mm, stroke=1, fill=0)
            
            c.setFont(font_bold, 8)
            c.drawString(map_box_x + 3*mm, map_box_y + map_box_h - 4*mm, f"Số TN: {so_tn}")
            c.drawString(map_box_x + map_box_w/2, map_box_y + map_box_h - 4*mm, f"GIA No.: {gia_no}")
            
            note_text = note_var.get().strip()
            if not note_text: note_text = "..................................................."
            c.setFont(font_name, 8)
            c.drawString(map_box_x + 3*mm, map_box_y + 3*mm, f"Ghi chú: {note_text}")
            
            pdf_img_w = map_box_w - 6*mm
            pdf_img_h = pdf_img_w * img_ratio
            
            if pdf_img_h > (map_box_h - 10*mm):
                pdf_img_h = map_box_h - 10*mm
                pdf_img_w = pdf_img_h / img_ratio
            
            pdf_img_x = map_box_x + (map_box_w - pdf_img_w) / 2
            pdf_img_y = map_box_y + (map_box_h - pdf_img_h) / 2
            
            try:
                map_pil = Image.open(image_path)
                if map_pil.mode != "RGB": map_pil = map_pil.convert("RGB")
                c.drawImage(ImageReader(map_pil), pdf_img_x, pdf_img_y, width=pdf_img_w, height=pdf_img_h, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                print(f"Lỗi ảnh gốc: {e}")
            
            pdf_cx = pdf_img_x + pdf_img_w / 2
            pdf_cy = pdf_img_y + pdf_img_h / 2
            scale_to_pdf = pdf_img_w / img_w 
            
            for item in canvas.find_withtag("drawing"):
                ctype = canvas.type(item)
                color_source = canvas.itemcget(item, "outline") if ctype in ["oval", "rectangle", "polygon"] else canvas.itemcget(item, "fill")
                if color_source == "#ff8888" or color_source == "red":
                    c.setStrokeColorRGB(1, 0, 0)
                    c.setFillColorRGB(1, 0, 0)
                elif color_source == "green":
                    c.setStrokeColorRGB(0, 0.6, 0)
                    c.setFillColorRGB(0, 0.6, 0)
                elif color_source == "blue":
                    c.setStrokeColorRGB(0, 0, 1)
                    c.setFillColorRGB(0, 0, 1)
                else:
                    c.setStrokeColorRGB(0, 0, 0)
                    c.setFillColorRGB(0, 0, 0)
                
                try:
                    dash = canvas.itemcget(item, "dash")
                    if dash: 
                        if "1" in str(dash) and "3" in str(dash): c.setDash([1.5, 2.5], 0)
                        else: c.setDash([2, 2], 0)
                    else: c.setDash([], 0)
                except:
                    c.setDash([], 0)
                    
                if ctype == "line":
                    coords = canvas.coords(item)
                    p = c.beginPath()
                    px = pdf_cx + (coords[0] - ui_cx) * scale_to_pdf
                    py = pdf_cy - (coords[1] - ui_cy) * scale_to_pdf
                    p.moveTo(px, py)
                    for i in range(2, len(coords), 2):
                        px = pdf_cx + (coords[i] - ui_cx) * scale_to_pdf
                        py = pdf_cy - (coords[i+1] - ui_cy) * scale_to_pdf
                        p.lineTo(px, py)
                    try: width = float(canvas.itemcget(item, "width")) * scale_to_pdf
                    except: width = 1.0 * scale_to_pdf
                    c.setLineWidth(max(0.5, width))
                    c.drawPath(p, stroke=1, fill=0)
                    
                elif ctype == "oval":
                    coords = canvas.coords(item)
                    cx_val = (coords[0] + coords[2]) / 2
                    cy_val = (coords[1] + coords[3]) / 2
                    r_val = abs(coords[2] - coords[0]) / 2 * scale_to_pdf
                    px = pdf_cx + (cx_val - ui_cx) * scale_to_pdf
                    py = pdf_cy - (cy_val - ui_cy) * scale_to_pdf
                    
                    is_filled = 1 if canvas.itemcget(item, "fill") else 0
                    try: width = float(canvas.itemcget(item, "width")) * scale_to_pdf
                    except: width = 1.0 * scale_to_pdf
                    c.setLineWidth(max(0.5, width))
                    c.circle(px, py, r_val, stroke=1, fill=is_filled)
                
                elif ctype == "polygon":
                    coords = canvas.coords(item)
                    p = c.beginPath()
                    px = pdf_cx + (coords[0] - ui_cx) * scale_to_pdf
                    py = pdf_cy - (coords[1] - ui_cy) * scale_to_pdf
                    p.moveTo(px, py)
                    for i in range(2, len(coords), 2):
                        px = pdf_cx + (coords[i] - ui_cx) * scale_to_pdf
                        py = pdf_cy - (coords[i+1] - ui_cy) * scale_to_pdf
                        p.lineTo(px, py)
                    p.closePath()
                    
                    is_filled = 1 if canvas.itemcget(item, "fill") else 0
                    try: width = float(canvas.itemcget(item, "width")) * scale_to_pdf
                    except: width = 1.0 * scale_to_pdf
                    c.setLineWidth(max(0.5, width))
                    c.drawPath(p, stroke=1, fill=is_filled)
                
                elif ctype == "rectangle":
                    coords = canvas.coords(item)
                    px = pdf_cx + (coords[0] - ui_cx) * scale_to_pdf
                    py = pdf_cy - (coords[3] - ui_cy) * scale_to_pdf 
                    rw = abs(coords[2] - coords[0]) * scale_to_pdf
                    rh = abs(coords[3] - coords[1]) * scale_to_pdf
                    c.setLineWidth(0.5)
                    c.rect(px, py, rw, rh, stroke=1, fill=0)
                    
                elif ctype == "text":
                    coords = canvas.coords(item)
                    px = pdf_cx + (coords[0] - ui_cx) * scale_to_pdf
                    py = pdf_cy - (coords[1] - ui_cy) * scale_to_pdf
                    text = canvas.itemcget(item, "text")
                    font_str = canvas.itemcget(item, "font")
                    
                    try: angle_val = float(canvas.itemcget(item, "angle"))
                    except: angle_val = 0.0

                    f_size = 12
                    match = re.search(r'\d+', font_str)
                    if match: f_size = int(match.group()) * scale_to_pdf * 0.8
                    c.setFont(font_bold, max(6, f_size))
                    
                    if angle_val != 0:
                        c.saveState()
                        c.translate(px, py)
                        c.rotate(angle_val)
                        c.drawCentredString(0, -f_size*0.3, text)
                        c.restoreState()
                    else:
                        c.drawCentredString(px, py - f_size*0.3, text)

            if gem_data:
                title_x = margin_pdf
                val_x = margin_pdf + 30 * mm 
                pdf_max_val_w = left_w - val_x - 5*mm 
                
                start_y = map_box_y - 7 * mm 
                line_spacing = 4.5 * mm
                
                c.setFont(font_bold, 11)
                c.setFillColorRGB(0, 0, 0)
                c.drawString(title_x, start_y, "THÔNG SỐ ĐÁ QUÝ" if font_name == 'Arial' else "THONG SO DA QUY")
                start_y -= 7 * mm
                
                for line in info_lines:
                    if isinstance(line, tuple): 
                        t, v = line[0], line[1]
                    else:
                        parts = line.split(":", 1)
                        t = parts[0] + ":"
                        v = parts[1].strip() if len(parts) > 1 else ""

                    c.setFont(font_bold, 9)
                    c.drawString(title_x, start_y, t)
                    c.setFont(font_name, 9)
                    
                    if not v:
                        start_y -= line_spacing
                        continue
                        
                    wrapped_lines = simpleSplit(v, font_name, 9, pdf_max_val_w)
                    for ln in wrapped_lines:
                        c.drawString(val_x, start_y, ln)
                        start_y -= line_spacing

            right_x = (w / 2) + 5 * mm
            right_y = 10 * mm
            right_w = (w / 2) - 15 * mm
            right_h = h - 20 * mm
            
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0.8, 0.8, 0.8)
            c.setDash([], 0)
            c.rect(right_x, right_y, right_w, right_h, stroke=1, fill=0)
            
            n = len(actual_photos)
            if n > 0:
                if n == 1: c_cols, c_rows = 1, 1
                elif n == 2: c_cols, c_rows = 1, 2
                elif n <= 4: c_cols, c_rows = 2, 2
                elif n <= 6: c_cols, c_rows = 2, 3
                elif n <= 8: c_cols, c_rows = 2, 4
                elif n <= 9: c_cols, c_rows = 3, 3
                else: c_cols, c_rows = 3, math.ceil(n/3)
                
                cell_w = right_w / c_cols
                cell_h = right_h / c_rows
                p = 2 * mm 
                
                for idx, photo_path in enumerate(actual_photos):
                    r_idx = idx // c_cols
                    c_idx = idx % c_cols
                    cx = right_x + c_idx * cell_w
                    cy = right_y + right_h - (r_idx + 1) * cell_h
                    
                    try:
                        pic_pil = Image.open(photo_path)
                        if pic_pil.mode != "RGB": pic_pil = pic_pil.convert("RGB")
                        c.drawImage(ImageReader(pic_pil), cx + p, cy + p, width=cell_w - 2*p, height=cell_h - 2*p, preserveAspectRatio=True, mask='auto')
                    except Exception as e:
                        print(f"Lỗi rửa ảnh đính kèm: {e}")
            else:
                c.setFont(font_name, 10)
                c.setFillColorRGB(0.5, 0.5, 0.5)
                c.drawCentredString(right_x + right_w/2, right_y + right_h/2, "Không có hình ảnh thực tế")
            
            c.save()
            buffer.seek(0)
            return buffer
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi tạo nháp PDF: {e}")
            return None

    def show_preview_window(pdf_buffer, folder, file_name, file_type, parent_dialog):
        preview_win = tk.Toplevel(mapping_win)
        preview_win.title("Xem Trước Báo Cáo")
        preview_win.geometry("900x650")
        preview_win.transient(mapping_win)
        preview_win.grab_set()
        
        try:
            pdf_bytes = pdf_buffer.getvalue()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=120) 
            img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            photo = ImageTk.PhotoImage(img_pil)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi hiển thị bản xem trước: {e}", parent=preview_win)
            return

        lbl_img = tk.Label(preview_win, image=photo)
        lbl_img.image = photo 
        lbl_img.pack(pady=10, expand=True)
        
        btn_frame = tk.Frame(preview_win)
        btn_frame.pack(fill=tk.X, pady=10)
        
        def on_save():
            try:
                final_ext = ".pdf" if file_type == "PDF" else f".{file_type.lower()}"
                out_final_name = os.path.join(folder, file_name + final_ext)
                
                if file_type == "PDF":
                    with open(out_final_name, "wb") as f:
                        f.write(pdf_bytes)
                else:
                    pix_high = page.get_pixmap(dpi=300)
                    pix_high.save(out_final_name)
                    
                doc.close()
                preview_win.destroy()
                if parent_dialog: parent_dialog.destroy()
                messagebox.showinfo("Thành công", f"Đã xuất báo cáo đồ giải thành công tại:\n{out_final_name}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi lưu file: {e}", parent=preview_win)
        
        def on_back():
            doc.close()
            preview_win.destroy()
            
        tk.Button(btn_frame, text="✖ Trở lại (Sửa đổi)", command=on_back, bg="#f8d7da", width=20, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=50)
        tk.Button(btn_frame, text="✔ Chốt & Lưu file", command=on_save, bg="#d4edda", width=20, font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=50)
        
        center_window(preview_win)

    def open_save_dialog():
        folder = filedialog.askdirectory(parent=mapping_win, title="Bước 1: Chọn thư mục lưu file")
        if not folder: return

        save_win = tk.Toplevel(mapping_win)
        save_win.title("Lưu Báo Cáo")
        save_win.transient(mapping_win)
        save_win.grab_set()

        tk.Label(save_win, text="Thư mục:", font=("Arial", 9, "bold")).pack(pady=(15, 0))
        tk.Label(save_win, text=folder, font=("Arial", 8), fg="gray", wraplength=350).pack(pady=(0, 10))

        tk.Label(save_win, text="Tên file:", font=("Arial", 10, "bold")).pack()
        name_var = tk.StringVar(value=f"{so_tn}_Report" if so_tn else "GIA_Report")
        tk.Entry(save_win, textvariable=name_var, width=35, font=("Arial", 10)).pack(pady=5)

        tk.Label(save_win, text="Loại file:", font=("Arial", 10, "bold")).pack()
        fmt_var = tk.StringVar(value="PDF")
        cb = ttk.Combobox(save_win, textvariable=fmt_var, values=["PDF", "PNG", "JPG"], state="readonly", width=32)
        cb.pack(pady=5)

        def on_preview():
            fn = name_var.get().strip()
            fmt = fmt_var.get()
            if not fn:
                messagebox.showerror("Lỗi", "Vui lòng nhập tên file!", parent=save_win)
                return
            if not HAS_FITZ:
                messagebox.showerror("Thiếu thư viện", "Vui lòng mở Terminal chạy lệnh:\n\npip install pymupdf\n\nđể sử dụng tính năng xem trước và xuất ảnh.", parent=save_win)
                return

            pdf_buffer = generate_pdf_buffer()
            if pdf_buffer:
                show_preview_window(pdf_buffer, folder, fn, fmt, save_win)

        tk.Button(save_win, text="👁 Tiếp tục (Xem trước)", command=on_preview, bg="#d4edda", width=25, font=("Arial", 10, "bold")).pack(pady=15)

        save_win.update_idletasks()
        w = save_win.winfo_reqwidth() + 40
        h = save_win.winfo_reqheight() + 20
        x = mapping_win.winfo_x() + (mapping_win.winfo_width()//2) - (w//2)
        y = mapping_win.winfo_y() + (mapping_win.winfo_height()//2) - (h//2)
        save_win.geometry(f"{w}x{h}+{x}+{y}")

    # --- SETUP THANH TOOLBAR ---
    def update_preview(val):
        x, y = mapping_win.winfo_pointerxy()
        widget = mapping_win.winfo_containing(x, y)
        if widget == canvas:
            class DummyEvent: pass
            e = DummyEvent()
            e.x = x - canvas.winfo_rootx()
            e.y = y - canvas.winfo_rooty()
            on_hover(e)

    group3 = tk.LabelFrame(toolbar_top, text="Điều chỉnh")
    group3.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Label(group3, text="Cỡ:").grid(row=0, column=0, sticky="e")
    tk.Scale(group3, from_=2, to_=12, variable=size_scale, orient=tk.HORIZONTAL, length=80, showvalue=0, command=update_preview).grid(row=0, column=1)
    tk.Label(group3, text="Góc:").grid(row=1, column=0, sticky="e")
    tk.Scale(group3, from_=0, to_=360, variable=angle_scale, orient=tk.HORIZONTAL, length=80, showvalue=0, command=update_preview).grid(row=1, column=1)

    group1 = tk.LabelFrame(toolbar_top, text="Ký hiệu & Chấu")
    group1.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Radiobutton(group1, text="Đóng: P", variable=current_tool, value="P", indicatoron=0, width=8, bg="#e2e3e5").grid(row=0, column=0, padx=2, pady=2)
    tk.Radiobutton(group1, text="Đóng: G", variable=current_tool, value="G", indicatoron=0, width=8, bg="#e2e3e5").grid(row=0, column=1, padx=2, pady=2)
    tk.Radiobutton(group1, text="Chấu (Xanh)", variable=current_tool, value="pen_blue", indicatoron=0, width=12, bg="#cce5ff").grid(row=1, column=0, columnspan=2, padx=2, pady=2, sticky="we")

    group_old_red = tk.LabelFrame(toolbar_top, text="Các Loại Tỳ Vết (Đỏ)")
    group_old_red.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Radiobutton(group_old_red, text="Cấn: V", variable=current_tool, value="V", indicatoron=0, width=6, bg="#f8d7da").grid(row=0, column=0, padx=1, pady=2)
    tk.Radiobutton(group_old_red, text="Cấn: +", variable=current_tool, value="+", indicatoron=0, width=6, bg="#f8d7da").grid(row=0, column=1, padx=1, pady=2)
    tk.Radiobutton(group_old_red, text="Mẻ (Chấm)", variable=current_tool, value="dot_red", indicatoron=0, width=14, bg="#f8d7da").grid(row=0, column=2, padx=1, pady=2)
    tk.Radiobutton(group_old_red, text="Cấn: Nét thẳng", variable=current_tool, value="line_red", indicatoron=0, width=14, bg="#f8d7da").grid(row=1, column=0, columnspan=2, padx=1, pady=2, sticky="we")
    tk.Radiobutton(group_old_red, text="Nhăm (Zigzag)", variable=current_tool, value="zigzag", indicatoron=0, width=14, bg="#f8d7da").grid(row=1, column=2, padx=1, pady=2, sticky="we")

    group_gia = tk.LabelFrame(toolbar_top, text="Ký Hiệu Tỳ Vết (GIA Standard)")
    group_gia.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    
    gia_tools = [
        ("Crystal", "crystal", "red"), ("Pinpoint", "pinpoint", "red"), ("Needle", "needle", "red"), ("Cloud", "cloud", "red"), ("Feather", "feather", "red"),
        ("Tw. Wisp", "twinning_wisp", "red"), ("Cavity", "cavity", "green"), ("Bruise", "bruise", "red"), ("Chip", "chip", "red"), ("Laser DH", "laser_drill", "green"),
        ("Knot", "knot", "green"), ("Etch Ch.", "etch_channel", "green"), ("Ind. Nat.", "indented_natural", "green"), ("Natural", "natural", "green"), ("Ex. Facet", "extra_facet", "black")
    ]
    for i, (txt, val, col) in enumerate(gia_tools):
        row, column = divmod(i, 5)
        tk.Radiobutton(group_gia, text=txt, variable=current_tool, value=val, indicatoron=0, width=8, fg=col, font=("Arial", 8, "bold")).grid(row=row, column=column, padx=1, pady=1)

    group4 = tk.LabelFrame(toolbar_top, text="Công cụ xóa")
    group4.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Radiobutton(group4, text="🧽 Tẩy nét vẽ", variable=current_tool, value="eraser", indicatoron=0, width=12, bg="#e2e3e5").pack(pady=2, padx=5, fill=tk.X)
    tk.Button(group4, text="✖ Xóa tất cả", command=lambda: [canvas.delete("drawing"), history.clear()], bg="#f8d7da", width=12).pack(pady=2, padx=5, fill=tk.X)

    group5 = tk.LabelFrame(toolbar_bottom, text="Ghi chú & Hình thực tế")
    group5.pack(fill=tk.BOTH, expand=True, padx=5)
    
    note_entry = tk.Entry(group5, textvariable=note_var, font=("Arial", 11))
    note_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 5), pady=5)
    
    btn_photos = tk.Button(group5, text="📸 Thư mục ảnh", command=select_photos_folder, bg="#cce5ff", font=("Arial", 9))
    btn_photos.pack(side=tk.LEFT, padx=(5, 2), pady=5, fill=tk.Y)
    
    lbl_photo_count = tk.Label(group5, text="0 ảnh", font=("Arial", 9, "italic"), fg="gray")
    lbl_photo_count.pack(side=tk.LEFT, padx=(0, 10))
    
    btn_save = tk.Button(group5, text="💾 Lưu file...", command=open_save_dialog, bg="#d4edda", font=("Arial", 10, "bold"), width=15)
    btn_save.pack(side=tk.RIGHT, padx=(2, 5), pady=5, fill=tk.Y)

    mapping_win.wait_window()

def main():
    target_dir = filedialog.askdirectory(title="Bước 1: Chọn thư mục chứa các file ảnh (.png, .jpg)")
    if not target_dir: return

    while True:
        data = get_target_file(target_dir)
        if not data: return 
            
        img_path = os.path.join(target_dir, data["file"])
        so_tn = data.get("so_tn", "")
        gia_no = data.get("gia_no", "")
        loai_da = data.get("loai_da", "kc")
        
        gem_data = {}
        if so_tn:
            session = requests.Session()
            scraped_data = scrape_pnj_item(session, so_tn, loai_da)
            if not scraped_data:
                messagebox.showerror("Lỗi kéo dữ liệu", "Không thể lấy thông tin hoặc mã Số TN không tồn tại!\nVui lòng kiểm tra lại mạng hoặc nhập lại Số TN.")
                continue 
            gem_data = scraped_data
        
        run_mapping_tool(img_path, so_tn, gia_no, loai_da, gem_data)
        break

if __name__ == "__main__":
    main()