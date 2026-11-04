import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime
import json
import os
import subprocess 
import threading
import re
import math
import io
import sys

# Bật tính năng nhận diện DPI trên Windows để chữ luôn sắc nét
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

try:
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt Pillow bằng lệnh: pip install pillow")

try:
    from tkcalendar import DateEntry
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False

# --- CÁC THƯ VIỆN PHỤC VỤ TAB 7 (CÀO WEB & XUẤT PDF) ---
try:
    from bs4 import BeautifulSoup
    from curl_cffi import requests as curl_requests
except ImportError:
    pass

try:
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.lib.pagesizes import A5, landscape
    from reportlab.lib.units import mm
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    from reportlab.lib.utils import simpleSplit, ImageReader
except ImportError:
    pass

try:
    import pymupdf as fitz  
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# --- CẤU HÌNH ĐƯỜNG DẪN TỐI ƯU ---
IMAGE_DIR = r"D:\OneDrive-NamTL-CAF\OneDrive - pnj.com.vn\PNJ Onedrive -NamTL\product\hinh-anh\HINH ANH SAN PHAM CAO"
TRACKING_FILE = r"D:\nam.tl01\nam.tl01\product\CAF\SCBH-PH\data\tracking-SCBH.xlsx"
DEFAULT_EXPORT_DIR = r"D:\nam.tl01\nam.tl01\product\CAF\SCBH-PH\data"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(CURRENT_DIR, "form_data.json")
INDEX_FILE = os.path.join(CURRENT_DIR, "image_index.json") 
REF_INDEX_FILE = os.path.join(CURRENT_DIR, "ref_image_index.json")

# --- DỮ LIỆU MẶC ĐỊNH ---
default_data = {
    "categories": {
        "Loại hàng": ["Hàng mua lại TKD", "Hàng trưng bày", "Hàng bảo hành", "Hàng phản hồi", "Hàng tách đá", "Hủy nấu"],
        "Loại lỗi": ["Đá trầy (cạnh, mặt)", "Đá nhăm (cạnh, mặt)", "Đá cấn (cạnh, mặt)", "Đá rớt", "Thiếu giấy kiểm định"],
        "Người kiểm": ["E0225574-TRẦN LÊ NAM"]
    },
    "tab2_rows": [],
    "last_nguoi_kiem": "E0225574-TRẦN LÊ NAM",
    "image_search_dir": "",
    "ref_image_dir": ""
}

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k in default_data:
                    if k not in data: data[k] = default_data[k]
                return data
        except Exception:
            pass
    return default_data

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

app_data = load_data()

image_index_list = []
def load_image_index():
    global image_index_list
    if os.path.exists(INDEX_FILE):
        try:
            with open(INDEX_FILE, "r", encoding="utf-8") as f:
                image_index_list = json.load(f)
        except:
            image_index_list = []
load_image_index()

ref_image_index_list = []
def load_ref_image_index():
    global ref_image_index_list
    if os.path.exists(REF_INDEX_FILE):
        try:
            with open(REF_INDEX_FILE, "r", encoding="utf-8") as f:
                ref_image_index_list = json.load(f)
        except:
            ref_image_index_list = []
load_ref_image_index()

class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list)
        self['values'] = self._completion_list
        self._timer = None 
        self.bind('<KeyRelease>', self.handle_keyrelease)
        
    def handle_keyrelease(self, event):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab'): return
        if self._timer: self.after_cancel(self._timer)
        self._timer = self.after(400, self.filter_list)

    def filter_list(self):
        value = self.get()
        if value == '':
            self['values'] = self._completion_list
        else:
            search_terms = value.lower().split()
            data = [item for item in self._completion_list if all(term in item.lower() for term in search_terms)]
            self['values'] = data
            try:
                self.tk.eval(f'ttk::combobox::Post {self}')
                self.icursor(tk.END)
                self.selection_clear()
            except tk.TclError:
                pass

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
                table = t; break
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
    for attempt in range(1, max_retries + 1):
        try:
            response = session.get(url, impersonate="chrome120", timeout=15)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                break
            elif response.status_code == 403: time.sleep(1)
            else: return None
        except: time.sleep(1)
    else: return None

    row_data = {
        "Mã PNJ": pnj_code, "Loại sản phẩm": prod_type_str.upper(), "Tên đá/ Loại ngọc": "", "Weight (carat)": "",
        "size_1": "", "size_2": "", "size_3": "", "Màu sắc(Color)": "", "MSC_GIA": "", "MSC_PNJ": "", "Đặc điểm": "", "Cut Grade": "",
        "Shape & Cut": "", "Clarity/Transparency": "", "Polish/Luster": "", "Symmetry/Surface": "", "Chú thích": ""
    }

    if type_code == 1:
        row_data["Tên đá/ Loại ngọc"] = "Diamond"
        carat_raw = get_p_span_value(soup, "CARAT WEIGHT") or get_p_span_value(soup, "carat")
        row_data["Weight (carat)"] = carat_raw.lower().replace("carat", "").replace("cts", "").strip()
        s1, s2, s3 = parse_measurement(get_p_span_value(soup, "Measurement", exclude="MARGIN"))
        row_data["size_1"], row_data["size_2"], row_data["size_3"] = s1, s2, s3
        row_data["Màu sắc(Color)"] = get_p_span_value(soup, "COLOR GRADE")
        row_data["Cut Grade"] = get_p_span_value(soup, "CUT GRADE")
        row_data["Clarity/Transparency"] = get_p_span_value(soup, "CLARITY GRADE")
        row_data["Shape & Cut"] = get_p_span_value(soup, "Shape & Cut") or get_p_span_value(soup, "Shape")
        row_data["Polish/Luster"] = get_p_span_value(soup, "Polish")
        row_data["Symmetry/Surface"] = get_p_span_value(soup, "Symmetry")
        
        ins_val = get_p_span_value(soup, "Inscription")
        row_data["MSC_PNJ"] = ins_val if ins_val else ""

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
        row_data["Weight (carat)"] = find_in_dict(td_dict, "WEIGHT", "KHỐI LƯỢNG").lower().replace("cts", "").replace("carat", "").strip()
        s1, s2, s3 = parse_measurement(find_in_dict(td_dict, "MEASUREMENT", "KÍCH THƯỚC"))
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
    return row_data

def generate_zigzag(points, step, amplitude):
    resampled = [points[0]]
    for p in points[1:]:
        if math.hypot(p[0]-resampled[-1][0], p[1]-resampled[-1][1]) >= step: resampled.append(p)
    if len(resampled) < 2: return []
    zz, sign = [], 1
    for i in range(len(resampled)-1):
        x1, y1 = resampled[i]; x2, y2 = resampled[i+1]
        zz.extend([x1, y1])
        dx, dy = x2 - x1, y2 - y1
        dist = math.hypot(dx, dy)
        if dist == 0: continue
        nx, ny = -dy/dist, dx/dist
        mx, my = x1 + dx/2 + nx * amplitude * sign, y1 + dy/2 + ny * amplitude * sign
        zz.extend([mx, my])
        sign *= -1
    zz.extend([resampled[-1][0], resampled[-1][1]])
    return zz

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

def main():
    root = tk.Tk()
    root.title("FORM KIỂM TRA ĐÁ SCBH CAF - V74 (Tối ưu Bố cục Trục Tra Cứu)")
    
    try:
        root.state('zoomed')
    except:
        root.attributes('-zoomed', True) 

    # --- KHAI BÁO CÁC HÀM XỬ LÝ ẢNH CHUNG (Tab 1 & Tab 6 dùng chung) ---
    def open_image_location(filepath):
        try: 
            subprocess.run(['explorer', '/select,', os.path.normpath(filepath)])
        except Exception as e: 
            messagebox.showerror("Lỗi", f"Không thể mở thư mục:\n{e}")

    def open_hd_image(filepath):
        top = tk.Toplevel(root)
        top.title("🔍 Chế độ Soi Ảnh HD")
        try: top.state('zoomed')
        except: top.attributes('-zoomed', True)
        
        tb_frame = tk.Frame(top, bg="#e2e3e5", bd=1, relief="raised")
        tb_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Button(tb_frame, text="📂 MỞ THƯ MỤC GỐC", command=lambda: open_image_location(filepath), font=("Arial", 10, "bold"), bg="#fff3cd").pack(side=tk.RIGHT, padx=10, pady=5)
        tk.Label(tb_frame, text=f"Đang xem: {os.path.basename(filepath)}", bg="#e2e3e5", font=("Arial", 10, "italic")).pack(side=tk.LEFT, padx=10, pady=5)

        img_canvas = tk.Canvas(top, bg="#333333")
        v_sb = ttk.Scrollbar(top, orient="vertical", command=img_canvas.yview)
        h_sb = ttk.Scrollbar(top, orient="horizontal", command=img_canvas.xview)
        img_canvas.configure(yscrollcommand=v_sb.set, xscrollcommand=h_sb.set)
        
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        h_sb.pack(side=tk.BOTTOM, fill=tk.X)
        img_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        try:
            pil_img = Image.open(filepath)
            img_state = {'scale': 1.0, 'tk_img': None, 'id': None}
            def render_img():
                resized = pil_img.resize((int(pil_img.width * img_state['scale']), int(pil_img.height * img_state['scale'])), Image.Resampling.LANCZOS)
                img_state['tk_img'] = ImageTk.PhotoImage(resized)
                if img_state['id']: img_canvas.delete(img_state['id'])
                img_state['id'] = img_canvas.create_image(0, 0, anchor="nw", image=img_state['tk_img'])
                img_canvas.configure(scrollregion=img_canvas.bbox("all"))
            render_img()
            img_canvas.bind("<MouseWheel>", lambda e: [img_state.update({'scale': img_state['scale'] * (1.1 if e.delta > 0 else 1/1.1)}), render_img()])
        except Exception as e: 
            messagebox.showerror("Lỗi", f"Không đọc được file:\n{e}")
            top.destroy()

    # --- THIẾT KẾ MÀU SẮC THEO BRAND GUIDELINES ---
    style = ttk.Style()
    style.theme_use('default')
    style.configure("TNotebook", background="#f0f0f0", borderwidth=0)
    style.configure("TNotebook.Tab", background="#e2e3e5", foreground="black", padding=[15, 5], font=("Arial", 10, "bold"))
    style.map("TNotebook.Tab", background=[("selected", "#005b52")], foreground=[("selected", "#f1d09f")])

    header_frame = tk.Frame(root, bg="#005b52")
    header_frame.pack(side=tk.TOP, fill=tk.X)
    
    lbl_logo = tk.Label(header_frame, text="C A F", font=("Arial", 18, "bold", "italic"), bg="#005b52", fg="#f1d09f")
    lbl_logo.pack(side=tk.LEFT, padx=(20, 5), pady=10)
    
    lbl_sep = tk.Label(header_frame, text="│", font=("Arial", 16), bg="#005b52", fg="#f1d09f")
    lbl_sep.pack(side=tk.LEFT, padx=5, pady=10)
    
    lbl_title = tk.Label(header_frame, text="K I Ể M  T R A  Đ Á  S C B H", font=("Arial", 14, "bold"), bg="#005b52", fg="#f1d09f")
    lbl_title.pack(side=tk.LEFT, padx=5, pady=10)
    
    def show_hotkeys(event=None):
        msg = """--- DANH SÁCH PHÍM TẮT (HOTKEYS) ---

[ ĐIỀU HƯỚNG CHUNG ]
• Alt + 1 đến 7 : Chuyển đổi nhanh giữa các Tab.
• F1 : Mở bảng tra cứu phím tắt này.
• Ctrl + S : Chạy lệnh Lưu/Xuất File thông minh 
  (Đang ở Tab 2 -> Xuất Excel | Đang ở Tab 7 -> Xuất PDF).

[ TAB 1: NHẬP LIỆU ]
• Enter : Tải ảnh gốc (khi đang ở ô Mã SP) hoặc Chuyển sang ô tiếp theo.
• Ctrl + N : Nhập dữ liệu nhanh (Đẩy sang Tab 2) mà không cần click chuột.
• Ctrl + PageDown : Mở ảnh đang chọn sang Tab Bản đồ tỳ vết (Tab 7).

[ TAB 2: DỮ LIỆU TẠM ]
• Delete : Xóa nhanh (các) dòng đang chọn.
• Ctrl + C : Copy (các) dòng đang chọn.

[ TAB 7: BẢN ĐỒ TỲ VẾT ]
• Space (Khoảng trắng) : Kích hoạt cào dữ liệu nhanh.
• Ctrl + Z : Hoàn tác (Undo) nét vẽ cuối cùng.
"""
        messagebox.showinfo("Tra Cứu Phím Tắt", msg)
        
    btn_help = tk.Button(header_frame, text="⌨️ Phím tắt (F1)", font=("Arial", 9, "bold"), bg="#004a43", fg="#f1d09f", relief="flat", activebackground="#003d36", activeforeground="white", command=show_hotkeys)
    btn_help.pack(side=tk.RIGHT, padx=20, pady=15)
    lbl_subtitle = tk.Label(header_frame, text="Form kiểm tra đá SCBH CAF: nhập liệu, lịch sử, thống kê, tìm ảnh lỗi", font=("Arial", 10), bg="#005b52", fg="#f1d09f")
    lbl_subtitle.pack(side=tk.RIGHT, padx=10, pady=15)

    bottom_frame = tk.Frame(root, bg="#005b52", bd=1, relief="ridge")
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

    notebook = ttk.Notebook(root)
    notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

    tab1_frame = tk.Frame(notebook, bg="#f0f0f0")
    tab2_frame = tk.Frame(notebook, bg="#f0f0f0")
    tab3_frame = tk.Frame(notebook, bg="#f0f0f0")
    tab4_frame = tk.Frame(notebook, bg="#f0f0f0") 
    tab5_frame = tk.Frame(notebook, bg="#f0f0f0") 
    tab6_frame = tk.Frame(notebook, bg="#f0f0f0") 
    tab7_frame = tk.Frame(notebook, bg="#f0f0f0") 

    notebook.add(tab1_frame, text=" 📝 TAB 1: NHẬP LIỆU ")
    notebook.add(tab2_frame, text=" 🗂️ TAB 2: DỮ LIỆU TẠM ")
    notebook.add(tab3_frame, text=" ⚙️ TAB 3: DANH MỤC ")
    notebook.add(tab4_frame, text=" 🔍 TAB 4: TÌM LỊCH SỬ ")
    notebook.add(tab5_frame, text=" 📊 TAB 5: THỐNG KÊ ")
    notebook.add(tab6_frame, text=" 🖼️ TAB 6: TÌM ẢNH LỖI ")
    notebook.add(tab7_frame, text=" 🗺️ TAB 7: BẢN ĐỒ TỲ VẾT ")

    for i, tab in enumerate([tab1_frame, tab2_frame, tab3_frame, tab4_frame, tab5_frame, tab6_frame, tab7_frame]):
        root.bind(f'<Alt-{i+1}>', lambda e, t=tab: notebook.select(t))
        root.bind(f'<Alt-Key-{i+1}>', lambda e, t=tab: notebook.select(t))

    root.bind('<F1>', show_hotkeys)

    t7_state = {'image_path': None, 'img_w': 0, 'img_h': 0, 'ui_cx': 800, 'ui_cy': 600, 'img_ratio': 1.0, 'actual_photos': [], 'has_scraped': False}
    t7_draw_state = {'is_drawing': False, 'current_action': [], 'history': [], 'draw_points': [], 'temp_items': []}

    t7_so_tn_var = tk.StringVar()
    t7_gia_no_var = tk.StringVar()
    t7_note_var = tk.StringVar()
    t7_copy_string_var = tk.StringVar() 

    # ==========================================
    # TAB 1: THIẾT KẾ GIAO DIỆN NHẬP LIỆU
    # ==========================================
    col_container = tk.Frame(tab1_frame, bg="#f0f0f0")
    col_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    col_container.columnconfigure(0, weight=2) # Khu nhập liệu và Ảnh gốc
    col_container.columnconfigure(1, weight=1) # Cột phải Tra cứu (Lịch sử + Tham khảo)
    col_container.rowconfigure(0, weight=1)
    
    # --- CỘT TRÁI: KHU VỰC NHẬP LIỆU & ẢNH GỐC ---
    left_big_frame = tk.LabelFrame(col_container, text="Thông tin Sản phẩm", font=("Arial", 10, "bold"), bg="#f0f0f0")
    left_big_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    left_big_frame.columnconfigure(0, weight=1)
    left_big_frame.columnconfigure(1, weight=1)
    left_big_frame.rowconfigure(0, weight=1)

    input_frame = tk.Frame(left_big_frame, bg="#f0f0f0")
    input_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    
    img_container = tk.Frame(left_big_frame, bg="#f0f0f0")
    img_container.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

    fields_left = ["Batch", "Material/Mã SP", "Plant", "Loại hàng", "Loại lỗi", "Tình trạng đá", "Ghi chú", "Ghi chú tham chiếu", "Hóa đơn", "Số lượng", "Ngày", "Số TN", "Số GIA", "Người kiểm"]
    entries = {} 

    for i, field in enumerate(fields_left):
        align = "nw" if field in ["Ghi chú", "Ghi chú tham chiếu"] else "w"
        tk.Label(input_frame, text=f"{field}:", font=("Arial", 9), bg="#f0f0f0").grid(row=i, column=0, sticky=align, padx=10, pady=4)
        
        if field in ["Loại hàng", "Loại lỗi", "Người kiểm"]:
            cb = AutocompleteCombobox(input_frame, width=28, font=("Arial", 10))
            cb.grid(row=i, column=1, padx=10, pady=4, sticky="ew")
            entries[field] = cb
        elif field == "Ngày":
            entry = tk.Entry(input_frame, width=28, font=("Arial", 10), relief="solid", bd=1)
            entry.insert(0, datetime.today().strftime('%d/%m/%Y'))
            entry.grid(row=i, column=1, padx=10, pady=4, sticky="ew")
            entries[field] = entry
        elif field == "Số lượng":
            entry = tk.Entry(input_frame, width=28, font=("Arial", 10), relief="solid", bd=1)
            entry.insert(0, "1")
            entry.grid(row=i, column=1, padx=10, pady=4, sticky="ew")
            entries[field] = entry
        elif field == "Tình trạng đá":
            radio_frame = tk.Frame(input_frame, bg="#f0f0f0")
            radio_frame.grid(row=i, column=1, sticky="w", padx=10, pady=4)
            tinh_trang_var = tk.StringVar(value="ĐẠT")
            tk.Radiobutton(radio_frame, text="ĐẠT", variable=tinh_trang_var, value="ĐẠT", bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 10))
            tk.Radiobutton(radio_frame, text="KHÔNG ĐẠT", variable=tinh_trang_var, value="KHÔNG ĐẠT", bg="#f0f0f0").pack(side=tk.LEFT)
            entries[field] = tinh_trang_var
        elif field in ["Ghi chú", "Ghi chú tham chiếu"]:
            h = 3 if field == "Ghi chú" else 2
            txt_ghichu = tk.Text(input_frame, width=28, height=h, font=("Arial", 10), relief="solid", bd=1)
            txt_ghichu.grid(row=i, column=1, padx=10, pady=4, sticky="ew")
            entries[field] = txt_ghichu
        else:
            entry = tk.Entry(input_frame, width=28, font=("Arial", 10), relief="solid", bd=1)
            entry.grid(row=i, column=1, padx=10, pady=4, sticky="ew")
            entries[field] = entry
            
    input_frame.columnconfigure(1, weight=1)

    img_frame = tk.LabelFrame(img_container, text="Hình ảnh Sản phẩm", bg="#f0f0f0")
    img_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    lbl_img = tk.Label(img_frame, text="Chưa tải ảnh\n(Nhập Mã SP và nhấn Enter)", bg="white", relief="sunken")
    lbl_img.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Khai báo trước biến để dùng trong call_tab7_mapping
    ref_code_var = tk.StringVar()

    # --- CỘT PHẢI: LỊCH SỬ & KHO ẢNH THAM KHẢO (ĐÃ QUY HOẠCH LẠI BỐ CỤC) ---
    right_container = tk.Frame(col_container, bg="#f0f0f0")
    right_container.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
    
    # 1. Khung Lịch Sử (Nửa trên)
    history_frame = tk.LabelFrame(right_container, text="Thông tin lịch sử SP", font=("Arial", 10, "bold"), bg="#f0f0f0")
    history_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 5))
    
    txt_history = tk.Text(history_frame, font=("Arial", 10), relief="sunken", bd=1, wrap=tk.WORD, state=tk.DISABLED, bg="#f9f9f9")
    txt_history.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 2. Khung Kho Ảnh Tham Khảo (Nửa dưới)
    ref_frame = tk.LabelFrame(right_container, text="📁 Kho Ảnh Tham Khảo", font=("Arial", 9, "bold"), bg="#f0f0f0")
    ref_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=(5, 0))

    ref_config_frame = tk.Frame(ref_frame, bg="#f0f0f0")
    ref_config_frame.pack(fill=tk.X, padx=5, pady=2)
    
    tk.Label(ref_config_frame, text="Thư mục:", bg="#f0f0f0", font=("Arial", 8)).pack(side=tk.LEFT)
    ent_ref_dir = tk.Entry(ref_config_frame, font=("Arial", 9))
    ent_ref_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
    ent_ref_dir.insert(0, app_data.get("ref_image_dir", ""))
    
    def choose_ref_dir():
        d = filedialog.askdirectory(title="Chọn Thư Mục Ảnh Tham Khảo")
        if d:
            ent_ref_dir.delete(0, tk.END)
            ent_ref_dir.insert(0, d)
            app_data["ref_image_dir"] = d
            save_data(app_data)
            
    tk.Button(ref_config_frame, text="Chọn", command=choose_ref_dir, font=("Arial", 8), bg="#e2e3e5").pack(side=tk.LEFT, padx=2)
    
    def sync_ref_dir():
        folder_path = ent_ref_dir.get().strip()
        if not os.path.exists(folder_path): 
            return messagebox.showerror("Lỗi", "Đường dẫn không tồn tại.")
            
        btn_sync_ref.config(state=tk.DISABLED, text="⏳...")
        def task():
            global ref_image_index_list
            temp_list = []
            for r, _, files in os.walk(folder_path):
                for f in files:
                    if f.lower().endswith(('.jpg', '.jpeg', '.png', '.jfif')): 
                        temp_list.append(os.path.join(r, f))
            ref_image_index_list = temp_list
            with open(REF_INDEX_FILE, "w", encoding="utf-8") as f: 
                json.dump(ref_image_index_list, f, ensure_ascii=False)
            root.after(0, lambda: [btn_sync_ref.config(state=tk.NORMAL, text="Đồng bộ"), messagebox.showinfo("Hoàn tất", f"Đã quét {len(ref_image_index_list)} ảnh tham khảo!")])
            
        threading.Thread(target=task, daemon=True).start()
        
    btn_sync_ref = tk.Button(ref_config_frame, text="Đồng bộ", command=sync_ref_dir, font=("Arial", 8), bg="#d4edda")
    btn_sync_ref.pack(side=tk.LEFT, padx=2)
    
    ref_search_frame = tk.Frame(ref_frame, bg="#f0f0f0")
    ref_search_frame.pack(fill=tk.X, padx=5, pady=2)
    tk.Label(ref_search_frame, text="Mã tìm ảnh:", bg="#f0f0f0", font=("Arial", 8, "bold")).pack(side=tk.LEFT)
    ent_ref_code = tk.Entry(ref_search_frame, textvariable=ref_code_var, font=("Arial", 9), width=25)
    ent_ref_code.pack(side=tk.LEFT, padx=5)
    tk.Label(ref_search_frame, text="(Ưu tiên mã này. Nếu trống sẽ tự dùng Mã SP)", bg="#f0f0f0", font=("Arial", 8, "italic"), fg="gray").pack(side=tk.LEFT)
    
    ref_gallery = tk.Frame(ref_frame, bg="white")
    ref_gallery.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    ref_gallery.pack_propagate(False)
    
    ref_canvas = tk.Canvas(ref_gallery, bg="white", height=115)
    ref_hbar = ttk.Scrollbar(ref_gallery, orient="horizontal", command=ref_canvas.xview)
    ref_canvas.configure(xscrollcommand=ref_hbar.set)
    ref_hbar.pack(side=tk.BOTTOM, fill=tk.X)
    ref_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    ref_inner = tk.Frame(ref_canvas, bg="white")
    ref_canvas.create_window((0, 0), window=ref_inner, anchor="nw")
    ref_inner.bind("<Configure>", lambda e: ref_canvas.configure(scrollregion=ref_canvas.bbox("all")))
    
    ref_thumbnail_refs = []

    # --- SỢI DÂY LIÊN KẾT: TỪ TAB 1 -> TAB 7 ---
    def get_actual_image_path(ma_sp):
        for ext in ['.jpg', '.jpeg', '.png', '.jfif']:
            p = os.path.join(IMAGE_DIR, f"{ma_sp}{ext}")
            if os.path.exists(p): return p
        return None

    def call_tab7_mapping(event=None):
        ma_sp = entries["Material/Mã SP"].get().strip()
        so_tn_raw = entries["Số TN"].get().strip()
        so_gia = entries["Số GIA"].get().strip()
        raw_batch = entries["Batch"].get().strip()

        if not ma_sp and not so_tn_raw:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập Mã SP hoặc Số TN để có thể tạo bản vẽ!")
            return

        img_path = get_actual_image_path(ma_sp) if ma_sp else None
        
        if not img_path:
            messagebox.showinfo("Thông báo", "Không tìm thấy ảnh gốc trong thư mục. Mở Bản Đồ Tỳ Vết với khung giấy trắng.")

        notebook.select(6)

        try:
            t7_canvas.focus_set()
        except:
            pass

        tn_list = [x.strip() for x in re.split(r'[\s,;\+]+', so_tn_raw) if x.strip()]
        
        t7_so_tn['values'] = tn_list
        if tn_list: 
            t7_so_tn.current(0) 
        else: 
            t7_so_tn_var.set("")
            
        t7_gia_no_var.set(so_gia)

        real_batch = raw_batch.split("/")[-1] if "/" in raw_batch else raw_batch
        tn1 = tn_list[0] if len(tn_list) > 0 else ""
        tn2 = tn_list[1] if len(tn_list) > 1 else ""
        
        copy_str = f"{real_batch}_{ma_sp}_{tn1}_{tn2}_"
        t7_copy_string_var.set(copy_str)

        t7_load_image(img_path)

        if tn_list: 
            t7_loai_da.set("kc")

    # Nút bấm Mở sang Tab 7 nằm dưới Ảnh SP thực tế
    btn_map = tk.Button(img_container, text="➡ MỞ ẢNH NÀY SANG TAB BẢN ĐỒ TỲ VẾT", command=call_tab7_mapping, font=("Arial", 10, "bold"), bg="#cce5ff", height=2)
    btn_map.pack(fill=tk.X, padx=5, pady=(5, 10))

    # ==========================================
    # CÁC HÀM XỬ LÝ TAB 1
    # ==========================================
    def load_product_image(event=None):
        ma_sp = entries["Material/Mã SP"].get().strip()
        
        # 1. Tải ảnh chính cho khung lớn
        if not ma_sp: 
            lbl_img.config(image='', text="Chưa nhập Mã SP")
        else:
            img_path = get_actual_image_path(ma_sp)
            if img_path:
                try:
                    img = Image.open(img_path)
                    img.thumbnail((350, 350), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    lbl_img.config(image=photo, text="")
                    lbl_img.image = photo 
                except Exception as e: 
                    lbl_img.config(image='', text=f"Lỗi mở ảnh: {e}")
            else: 
                lbl_img.config(image='', text=f"Không tìm thấy ảnh:\n{ma_sp}.(jpg/png/jfif)")
        
        # 2. Quét & Tải ảnh tham khảo vào Mini Gallery
        for widget in ref_inner.winfo_children(): widget.destroy()
        ref_thumbnail_refs.clear()
        
        search_code = ref_code_var.get().strip() or ma_sp
        
        if search_code:
            matched = []
            search_lower = search_code.lower()
            for p in ref_image_index_list:
                fname = os.path.basename(p).lower()
                if search_lower in fname:
                    matched.append(p)
            
            if matched:
                for idx, path in enumerate(matched[:20]):
                    try:
                        img = Image.open(path)
                        img.thumbnail((100, 100), Image.Resampling.LANCZOS)
                        photo = ImageTk.PhotoImage(img)
                        ref_thumbnail_refs.append(photo)
                        
                        frm = tk.Frame(ref_inner, bg="white", bd=1, relief="solid", padx=2, pady=2)
                        frm.grid(row=0, column=idx, padx=5, pady=5)
                        
                        lbl = tk.Label(frm, image=photo, bg="white", cursor="hand2")
                        lbl.pack()
                        
                        fname_disp = os.path.basename(path)
                        if len(fname_disp) > 15: fname_disp = fname_disp[:12] + "..."
                        tk.Label(frm, text=fname_disp, bg="white", font=("Arial", 7)).pack()
                        
                        lbl.bind("<Double-Button-1>", lambda e, p=path: open_hd_image(p))
                        lbl.bind("<Button-3>", lambda e, p=path: open_image_location(p))
                    except: pass
            else:
                tk.Label(ref_inner, text="(Không có ảnh tham khảo)", bg="white", font=("Arial", 9, "italic"), fg="gray").grid(row=0, column=0, padx=10, pady=40)
        
        if event and event.keysym == 'Return': 
            event.widget.tk_focusNext().focus()
            return "break"

    def load_history(event=None):
        raw_batch = entries["Batch"].get().strip()
        real_batch = raw_batch.split("/")[-1] if "/" in raw_batch else raw_batch
        if not raw_batch: return
        txt_history.config(state=tk.NORMAL)
        txt_history.delete("1.0", tk.END)
        txt_history.insert(tk.END, f"🔍 Đang tìm lịch sử cho Batch: {real_batch}...\n\n")
        root.update()
        if not os.path.exists(TRACKING_FILE): 
            txt_history.insert(tk.END, f"❌ Không tìm thấy file Excel tracking:\n{TRACKING_FILE}")
            txt_history.config(state=tk.DISABLED)
            return
            
        try:
            import pandas as pd
            df = pd.read_excel(TRACKING_FILE)
            mask = df.astype(str).apply(lambda x: x.str.contains(real_batch, case=False, na=False)).any(axis=1)
            results = df[mask]
            txt_history.delete("1.0", tk.END)
            if results.empty: 
                txt_history.insert(tk.END, f"Không có dữ liệu lịch sử nào cho Batch:\n{real_batch}")
            else:
                txt_history.insert(tk.END, f"Đã tìm thấy {len(results)} kết quả:\n")
                for idx, row in results.iterrows():
                    txt_history.insert(tk.END, "-" * 35 + "\n")
                    for col in df.columns:
                        val = row[col]
                        if pd.notna(val) and str(val).strip(): 
                            txt_history.insert(tk.END, f"• {col}: {val}\n")
                txt_history.insert(tk.END, "-" * 35 + "\n")
        except Exception as e: 
            txt_history.insert(tk.END, f"❌ Lỗi đọc file Excel:\n{e}")
        finally: 
            txt_history.config(state=tk.DISABLED)

    def on_batch_return(event):
        load_history()
        event.widget.tk_focusNext().focus()
        return "break"

    entries["Material/Mã SP"].bind("<Return>", load_product_image)
    entries["Material/Mã SP"].bind("<FocusOut>", lambda e: load_product_image())
    
    ent_ref_code.bind("<Return>", load_product_image)
    ent_ref_code.bind("<FocusOut>", lambda e: load_product_image())
    
    entries["Batch"].bind("<Return>", on_batch_return)
    entries["Batch"].bind("<FocusOut>", lambda e: load_history())
    
    for key, widget in entries.items():
        if key not in ["Material/Mã SP", "Batch"] and isinstance(widget, (tk.Entry, ttk.Combobox)): 
            widget.bind("<Return>", lambda e: [e.widget.tk_focusNext().focus(), "break"])

    def update_dropdowns():
        entries["Loại hàng"].set_completion_list(app_data["categories"]["Loại hàng"])
        entries["Loại lỗi"].set_completion_list(app_data["categories"]["Loại lỗi"])
        entries["Người kiểm"].set_completion_list(app_data["categories"]["Người kiểm"])
        last_nk = app_data.get("last_nguoi_kiem")
        if last_nk and last_nk in app_data["categories"]["Người kiểm"]: 
            entries["Người kiểm"].set(last_nk)
        elif app_data["categories"]["Người kiểm"]: 
            entries["Người kiểm"].set(app_data["categories"]["Người kiểm"][0])
            
    update_dropdowns()

    # ==========================================
    # TAB 2: THIẾT KẾ BẢNG DỮ LIỆU TẠM & BỘ LỌC
    # ==========================================
    t2_filter_frame = tk.Frame(tab2_frame, bg="#f0f0f0")
    t2_filter_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
    tk.Label(t2_filter_frame, text="🔍 Lọc nhanh dữ liệu:", bg="#f0f0f0", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
    
    ent_t2_filter = tk.Entry(t2_filter_frame, font=("Arial", 10), width=30)
    ent_t2_filter.pack(side=tk.LEFT, padx=5)

    def apply_t2_filter(e=None):
        kw = ent_t2_filter.get().strip().lower()
        for item in tree.get_children(): 
            tree.delete(item)
        for row in app_data["tab2_rows"]:
            if not kw or any(kw in str(cell).lower() for cell in row):
                tree.insert("", tk.END, values=row)

    ent_t2_filter.bind("<Return>", apply_t2_filter)
    tk.Button(t2_filter_frame, text="Lọc", command=apply_t2_filter, bg="#cce5ff", width=8).pack(side=tk.LEFT, padx=2)
    tk.Button(t2_filter_frame, text="✖ Hủy", command=lambda: [ent_t2_filter.delete(0, tk.END), apply_t2_filter()], bg="#e2e3e5", width=8).pack(side=tk.LEFT, padx=2)

    columns = ("Batch", "Mã SP", "Plant", "Số lượng", "Loại hàng", "Lỗi", "Ngày", "Tình trạng", "Số TN", "Số GIA", "Ghi chú", "Người kiểm", "Hóa đơn")
    tree_frame = tk.Frame(tab2_frame)
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    tree_scroll = tk.Scrollbar(tree_frame)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
    tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set, xscrollcommand=tree_scroll_x.set)
    tree.pack(fill=tk.BOTH, expand=True)
    tree_scroll.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)
    
    for col in columns: 
        tree.heading(col, text=col)
        tree.column(col, width=90, anchor=tk.CENTER)
        
    tree.column("Mã SP", width=120)
    tree.column("Số TN", width=100)
    tree.column("Số GIA", width=100)
    tree.column("Ghi chú", width=180)
    tree.column("Người kiểm", width=120)
    tree.column("Hóa đơn", width=100)
    
    for row in app_data["tab2_rows"]: 
        tree.insert("", tk.END, values=row)

    popup_menu = tk.Menu(tree_frame, tearoff=0, font=("Arial", 10))
    def show_popup(event):
        item_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not item_id or not col_id: return
        tree.selection_set(item_id)
        col_index = int(col_id.replace('#', '')) - 1
        col_name = columns[col_index]
        row_values = tree.item(item_id, 'values')
        cell_value = row_values[col_index]
        
        popup_menu.delete(0, tk.END)
        popup_menu.add_command(label=f"📝 Copy ô này: '{cell_value}'", command=lambda: [root.clipboard_clear(), root.clipboard_append(str(cell_value))])
        def copy_column():
            col_data = [col_name]
            [col_data.append(str(tree.item(child, 'values')[col_index])) for child in tree.get_children()]
            root.clipboard_clear()
            root.clipboard_append("\n".join(col_data))
            
        popup_menu.add_command(label=f"📐 Copy toàn bộ cột '{col_name}'", command=copy_column)
        popup_menu.add_command(label=f"📃 Copy cả dòng này (Tất cả cột)", command=lambda: [root.clipboard_clear(), root.clipboard_append("\t".join(str(v) for v in row_values))])
        popup_menu.add_command(label=f"✂️ Copy dòng (Từ Batch đến Ghi chú)", command=lambda: [root.clipboard_clear(), root.clipboard_append("\t".join(str(v) for v in row_values[:columns.index("Ghi chú") + 1]))])
        popup_menu.post(event.x_root, event.y_root)
        
    tree.bind("<Button-3>", show_popup)

    def push_to_tab2():
        ma_sp = entries["Material/Mã SP"].get().strip()
        if not ma_sp: return messagebox.showwarning("Cảnh báo", "Vui lòng nhập Mã SP!")
        raw_batch = entries["Batch"].get().strip()
        real_batch = raw_batch.split("/")[-1] if "/" in raw_batch else raw_batch
        data_row = (real_batch or "none", ma_sp, entries["Plant"].get().strip() or "none", entries["Số lượng"].get().strip() or "1", entries["Loại hàng"].get().strip() or "none", entries["Loại lỗi"].get().strip() or "none", entries["Ngày"].get().strip() or "none", entries["Tình trạng đá"].get().strip() or "none", entries["Số TN"].get().strip() or "none", entries["Số GIA"].get().strip() or "none", entries["Ghi chú"].get("1.0", tk.END).strip() or "none", entries["Người kiểm"].get().strip() or "none", entries["Hóa đơn"].get().strip() or "0")
        
        app_data["tab2_rows"].append(data_row)
        app_data["last_nguoi_kiem"] = entries["Người kiểm"].get()
        save_data(app_data)
        update_tong_sl()
        
        apply_t2_filter()
        
        for field in ["Batch", "Material/Mã SP", "Plant", "Loại hàng", "Loại lỗi", "Số TN", "Số GIA", "Hóa đơn"]: 
            entries[field].delete(0, tk.END)
            if isinstance(entries[field], ttk.Combobox): entries[field].set('')
            
        entries["Ghi chú"].delete("1.0", tk.END)
        entries["Số lượng"].delete(0, tk.END)
        entries["Số lượng"].insert(0, "1")
        lbl_img.config(image='', text="Chưa tải ảnh\n(Nhập Mã SP và nhấn Enter)")
        txt_history.config(state=tk.NORMAL)
        txt_history.delete("1.0", tk.END)
        txt_history.config(state=tk.DISABLED)
        
        # Reset luôn thư viện ảnh tham khảo
        for widget in ref_inner.winfo_children(): widget.destroy()
        ref_code_var.set("")
        
        entries["Batch"].focus() 

    def xoa_dong_tab2():
        selected = tree.selection()
        if not selected: return
        for sel in selected:
            values_str = tuple(str(v) for v in tree.item(sel, "values"))
            for idx, row in enumerate(app_data["tab2_rows"]):
                if tuple(str(v) for v in row) == values_str:
                    app_data["tab2_rows"].pop(idx)
                    break
            tree.delete(sel)
        save_data(app_data)
        update_tong_sl()

    def sua_dong_tab2():
        selected = tree.selection()
        if not selected: return messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 dòng để chỉnh sửa!")
        if len(selected) > 1: return messagebox.showwarning("Cảnh báo", "Chỉ được chọn 1 dòng để chỉnh sửa mỗi lần!")
        
        item = selected[0]
        values = tree.item(item, "values")
        
        edit_win = tk.Toplevel(root)
        edit_win.title("Chỉnh sửa dữ liệu")
        edit_win.geometry("450x600")
        edit_win.grab_set() 
        tk.Label(edit_win, text="Chỉnh sửa thông tin bên dưới và bấm Lưu:", font=("Arial", 10, "bold")).pack(pady=10)
        
        edit_entries = {}
        for i, col in enumerate(columns):
            frame = tk.Frame(edit_win)
            frame.pack(fill=tk.X, padx=20, pady=5)
            tk.Label(frame, text=col + ":", width=12, anchor="w", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
            ent = tk.Entry(frame, font=("Arial", 10))
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
            ent.insert(0, str(values[i]))
            edit_entries[col] = ent
            
        def save_edit():
            new_vals = [edit_entries[col].get().strip() for col in columns]
            tree.item(item, values=new_vals)
            
            values_str = tuple(str(v) for v in values)
            for idx, row in enumerate(app_data["tab2_rows"]):
                if tuple(str(v) for v in row) == values_str:
                    app_data["tab2_rows"][idx] = new_vals
                    break
                    
            save_data(app_data)
            update_tong_sl()
            edit_win.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật dữ liệu dòng được chọn!")
            
        tk.Button(edit_win, text="💾 LƯU THAY ĐỔI", command=save_edit, bg="#d4edda", font=("Arial", 10, "bold"), width=15).pack(pady=20)

    def copy_data_tab2(event=None):
        selected_items = tree.selection()
        items_to_copy = selected_items if selected_items else tree.get_children()
        if not items_to_copy: 
            if not event: messagebox.showwarning("Trống", "Không có dữ liệu để copy!")
            return
            
        lines = ["\t".join(columns)]
        for item in items_to_copy: 
            lines.append("\t".join(str(v) for v in tree.item(item, "values")))
            
        root.clipboard_clear()
        root.clipboard_append("\n".join(lines))
        if not event: 
            messagebox.showinfo("Thành công", f"Đã copy {'TOÀN BỘ bảng' if not selected_items else f'{len(selected_items)} dòng'}!")
            
    tree.bind("<Control-c>", copy_data_tab2)
    tree.bind("<Control-C>", copy_data_tab2)

    def export_excel():
        if not app_data["tab2_rows"]: return messagebox.showwarning("Trống", "Không có dữ liệu nào ở Tab 2 để xuất!")
        target_file = filedialog.askopenfilename(title="Chọn file Excel để cập nhật dữ liệu", initialdir=DEFAULT_EXPORT_DIR, filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not target_file: return 
        try:
            import openpyxl
            wb = openpyxl.load_workbook(target_file)
            if "source" not in wb.sheetnames: 
                wb.close()
                return messagebox.showerror("Lỗi", f"Không tìm thấy sheet mang tên 'source' trong file:\n{os.path.basename(target_file)}")
                
            sheet = wb["source"]
            last_stt = 0
            for r in range(sheet.max_row, 0, -1):
                val = sheet.cell(row=r, column=1).value
                if isinstance(val, int): 
                    last_stt = val
                    break
                elif str(val).isdigit(): 
                    last_stt = int(val)
                    break
                    
            for row_data in app_data["tab2_rows"]:
                real_batch, ma_sp, plant, so_luong, loai_hang, loai_loi, ngay, tinh_trang, so_tn, so_gia, ghi_chu, nguoi_kiem, hoa_don = row_data
                last_stt += 1
                loai_loi = "none" if not loai_loi.strip() or loai_loi.lower() == "none" else loai_loi
                nguoi_kiem = nguoi_kiem.split("-")[0].strip() if "-" in nguoi_kiem and nguoi_kiem.lower() != "none" else nguoi_kiem
                tt_short = "Đ" if tinh_trang.strip().upper() == "ĐẠT" else ("KĐ" if tinh_trang.strip().upper() == "KHÔNG ĐẠT" else tinh_trang)
                sl_val = int(so_luong) if str(so_luong).isdigit() else so_luong
                
                try: 
                    hd_val = int(hoa_don) if str(hoa_don).isdigit() else float(hoa_don)
                except ValueError: 
                    hd_val = 0
                    
                ngay_obj = ngay
                is_valid_date = False
                if ngay.lower() != "none":
                    try: 
                        ngay_obj = datetime.strptime(ngay.strip(), "%d/%m/%Y").date()
                        is_valid_date = True
                    except: pass 
                    
                sheet.append([last_stt, real_batch, ma_sp, plant, sl_val, loai_hang, loai_loi, ngay_obj, tt_short, ghi_chu, so_tn, so_gia, nguoi_kiem, hd_val])
                if is_valid_date: 
                    sheet.cell(row=sheet.max_row, column=8).number_format = 'DD/MM/YYYY'
                    
            wb.save(target_file)
            wb.close()
            for item in tree.get_children(): tree.delete(item)
            app_data["tab2_rows"].clear()
            save_data(app_data)
            update_tong_sl()
            messagebox.showinfo("Thành công", f"Đã ghi nối thành công toàn bộ dữ liệu vào file:\n{os.path.basename(target_file)}")
        except PermissionError: 
            messagebox.showerror("Lỗi Truy Cập", "Không thể ghi dữ liệu vì file Excel đang được mở.\nVui lòng đóng file Excel trước khi xuất!")
        except Exception as e: 
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi ghi file:\n{e}")

    # ==========================================
    # TAB 3: DANH MỤC 
    # ==========================================
    tk.Label(tab3_frame, text="Quản lý danh sách (Phím tắt: Alt+1, Alt+2, Alt+3 để chuyển Tab). Bạn có thể COPY TỪ EXCEL dán vào ô bên dưới.", font=("Arial", 10, "italic")).pack(pady=10)
    list_frame = tk.Frame(tab3_frame)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=10)
    list_frame.columnconfigure(0, weight=1)
    list_frame.columnconfigure(1, weight=1)
    list_frame.columnconfigure(2, weight=1)
    list_frame.rowconfigure(0, weight=1)

    def create_editable_table(parent, title, category_key, col_index):
        frame = tk.LabelFrame(parent, text=title, font=("Arial", 10, "bold"))
        frame.grid(row=0, column=col_index, sticky="nsew", padx=5, pady=5)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        
        cat_tree = ttk.Treeview(frame, columns=("Value",), show="headings", height=12)
        cat_tree.heading("Value", text="Tên danh mục")
        cat_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        for item in app_data["categories"][category_key]: 
            cat_tree.insert("", tk.END, values=(item,))
            
        input_frame = tk.Frame(frame)
        input_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        
        def add_item():
            raw_val = txt_new.get("1.0", tk.END).strip()
            added_any = False
            if not raw_val: return
            for line in raw_val.split('\n'):
                val = line.strip()
                if val and val not in app_data["categories"][category_key]:
                    cat_tree.insert("", tk.END, values=(val,))
                    app_data["categories"][category_key].append(val)
                    added_any = True
            if added_any: 
                save_data(app_data)
                update_dropdowns()
            txt_new.delete("1.0", tk.END)
                
        def del_item():
            selected = cat_tree.selection()
            if selected:
                for sel in selected:
                    val = cat_tree.item(sel)["values"][0]
                    cat_tree.delete(sel)
                    if val in app_data["categories"][category_key]: 
                        app_data["categories"][category_key].remove(val)
                save_data(app_data)
                update_dropdowns()
                
        tk.Button(input_frame, text="Xóa", bg="#f8d7da", command=del_item).pack(side=tk.RIGHT, padx=(5, 0))
        tk.Button(input_frame, text="Thêm", bg="#d4edda", command=add_item).pack(side=tk.RIGHT)
        txt_new = tk.Text(input_frame, height=2, font=("Arial", 9))
        txt_new.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

    create_editable_table(list_frame, "Loại hàng", "Loại hàng", 0)
    create_editable_table(list_frame, "Loại lỗi", "Loại lỗi", 1)
    create_editable_table(list_frame, "Người kiểm", "Người kiểm", 2)

    # ==========================================
    # TAB 4: TÌM KIẾM LỊCH SỬ ĐA CHIỀU
    # ==========================================
    search_input_frame = tk.LabelFrame(tab4_frame, text="Tiêu chí tìm kiếm", font=("Arial", 10, "bold"), bg="#f0f0f0")
    search_input_frame.pack(fill=tk.X, padx=10, pady=10)
    
    tk.Label(search_input_frame, text="Mã SP:", bg="#f0f0f0", font=("Arial", 10)).grid(row=0, column=0, padx=10, pady=10, sticky="e")
    search_ma_sp = tk.Entry(search_input_frame, width=25, font=("Arial", 10))
    search_ma_sp.grid(row=0, column=1, padx=5, pady=10)
    
    tk.Label(search_input_frame, text="Plant:", bg="#f0f0f0", font=("Arial", 10)).grid(row=0, column=2, padx=10, pady=10, sticky="e")
    search_plant = tk.Entry(search_input_frame, width=25, font=("Arial", 10))
    search_plant.grid(row=0, column=3, padx=5, pady=10)
    
    tk.Label(search_input_frame, text="Số TN:", bg="#f0f0f0", font=("Arial", 10)).grid(row=1, column=0, padx=10, pady=10, sticky="e")
    search_so_tn_tab4 = tk.Entry(search_input_frame, width=25, font=("Arial", 10))
    search_so_tn_tab4.grid(row=1, column=1, padx=5, pady=10)
    
    tk.Label(search_input_frame, text="Số GIA:", bg="#f0f0f0", font=("Arial", 10)).grid(row=1, column=2, padx=10, pady=10, sticky="e")
    search_so_gia_tab4 = tk.Entry(search_input_frame, width=25, font=("Arial", 10))
    search_so_gia_tab4.grid(row=1, column=3, padx=5, pady=10)
    
    btn_frame_tab4 = tk.Frame(search_input_frame, bg="#f0f0f0")
    btn_frame_tab4.grid(row=0, column=4, rowspan=2, padx=30, pady=10)
    
    btn_search_tab4 = tk.Button(btn_frame_tab4, text="🔍 TÌM KIẾM", font=("Arial", 10, "bold"), bg="#cce5ff", height=2, width=15)
    btn_search_tab4.pack(side=tk.TOP, pady=2)
    
    btn_clear_tab4 = tk.Button(btn_frame_tab4, text="🔄 LÀM MỚI", command=lambda: [e.delete(0, tk.END) for e in (search_ma_sp, search_plant, search_so_tn_tab4, search_so_gia_tab4)] + [txt_search_result.config(state=tk.NORMAL), txt_search_result.delete("1.0", tk.END), txt_search_result.config(state=tk.DISABLED)], font=("Arial", 9), bg="#e2e3e5", width=17)
    btn_clear_tab4.pack(side=tk.TOP, pady=2)
    
    result_frame_tab4 = tk.LabelFrame(tab4_frame, text="Kết quả tìm kiếm", font=("Arial", 10, "bold"), bg="#f0f0f0")
    result_frame_tab4.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    txt_search_result = tk.Text(result_frame_tab4, font=("Arial", 11), relief="sunken", bd=1, wrap=tk.WORD, state=tk.DISABLED, bg="#f9f9f9")
    txt_search_result.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def perform_search(event=None):
        vals = [e.get().strip() for e in (search_ma_sp, search_plant, search_so_tn_tab4, search_so_gia_tab4)]
        if not any(vals): return messagebox.showwarning("Cảnh báo", "Vui lòng nhập ít nhất 1 tiêu chí tìm kiếm!")
        
        txt_search_result.config(state=tk.NORMAL)
        txt_search_result.delete("1.0", tk.END)
        root.update()
        
        if not os.path.exists(TRACKING_FILE): 
            txt_search_result.insert(tk.END, "❌ Không tìm thấy file Excel tracking")
            txt_search_result.config(state=tk.DISABLED)
            return
            
        try:
            import pandas as pd
            df = pd.read_excel(TRACKING_FILE)
            mask = pd.Series(True, index=df.index)
            for v in vals:
                if v: mask = mask & df.astype(str).apply(lambda x: x.str.contains(v, case=False, na=False)).any(axis=1)
            results = df[mask]
            txt_search_result.delete("1.0", tk.END)
            if results.empty: 
                txt_search_result.insert(tk.END, "Không có dữ liệu.")
            else:
                txt_search_result.insert(tk.END, f"Đã tìm thấy {len(results)} kết quả phù hợp:\n")
                for _, row in results.iterrows():
                    txt_search_result.insert(tk.END, "-" * 60 + "\n")
                    for col in df.columns:
                        if pd.notna(row[col]) and str(row[col]).strip(): 
                            txt_search_result.insert(tk.END, f"• {col}: {row[col]}\n")
        except Exception as e: 
            txt_search_result.insert(tk.END, f"❌ Lỗi:\n{e}")
        finally: 
            txt_search_result.config(state=tk.DISABLED)
            
    btn_search_tab4.config(command=perform_search)
    for e in (search_ma_sp, search_plant, search_so_tn_tab4, search_so_gia_tab4): 
        e.bind("<Return>", perform_search)

    # ==========================================
    # TAB 5: DASHBOARD THỐNG KÊ 
    # ==========================================
    filter_frame = tk.LabelFrame(tab5_frame, text="Công cụ lọc dữ liệu", font=("Arial", 10, "bold"), bg="#f0f0f0")
    filter_frame.pack(fill=tk.X, padx=10, pady=5)
    tk.Label(filter_frame, text="Từ ngày:", bg="#f0f0f0", font=("Arial", 9)).grid(row=0, column=0, padx=10, pady=10)
    
    if HAS_CALENDAR:
        entry_tu_ngay = DateEntry(filter_frame, width=15, font=("Arial", 10), date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        entry_tu_ngay.delete(0, tk.END)
        tk.Label(filter_frame, text="Đến ngày:", bg="#f0f0f0", font=("Arial", 9)).grid(row=0, column=2, padx=10, pady=10)
        entry_den_ngay = DateEntry(filter_frame, width=15, font=("Arial", 10), date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        entry_den_ngay.delete(0, tk.END)
    else:
        entry_tu_ngay = tk.Entry(filter_frame, width=15, font=("Arial", 10))
        tk.Label(filter_frame, text="(Cài tkcalendar) Đến ngày:", bg="#f0f0f0", font=("Arial", 9, "italic")).grid(row=0, column=2, padx=10, pady=10)
        entry_den_ngay = tk.Entry(filter_frame, width=15, font=("Arial", 10))
        
    entry_tu_ngay.grid(row=0, column=1, padx=5, pady=10)
    entry_den_ngay.grid(row=0, column=3, padx=5, pady=10)
    
    tk.Button(filter_frame, text="✖ XÓA LỌC", command=lambda: [entry_tu_ngay.delete(0, tk.END), entry_den_ngay.delete(0, tk.END)], font=("Arial", 9, "bold"), bg="#ffcccc").grid(row=0, column=4, padx=5, pady=10)
    btn_load_dashboard = tk.Button(filter_frame, text="⚙️ XỬ LÝ DỮ LIỆU", font=("Arial", 10, "bold"), bg="#e2e3e5", height=1)
    btn_load_dashboard.grid(row=0, column=5, padx=20, pady=10)

    kpi_container = tk.Frame(tab5_frame, bg="#f0f0f0")
    kpi_container.pack(fill=tk.X, padx=10, pady=10)
    
    chart_btn_frame = tk.Frame(tab5_frame, bg="#f0f0f0")
    chart_btn_frame.pack(fill=tk.X, padx=10, pady=15)
    
    btn_hm = tk.Button(chart_btn_frame, text="🔥 MỞ BẢN ĐỒ NHIỆT", font=("Arial", 10, "bold"), bg="#ffcccc", height=2, width=25, state=tk.DISABLED)
    btn_hm.pack(side=tk.LEFT, padx=10)
    btn_line = tk.Button(chart_btn_frame, text="📈 MỞ ĐƯỜNG XU HƯỚNG", font=("Arial", 10, "bold"), bg="#cce5ff", height=2, width=28, state=tk.DISABLED)
    btn_line.pack(side=tk.LEFT, padx=10)
    btn_pareto = tk.Button(chart_btn_frame, text="📊 MỞ PHÂN TÍCH PARETO", font=("Arial", 10, "bold"), bg="#d4edda", height=2, width=25, state=tk.DISABLED)
    btn_pareto.pack(side=tk.LEFT, padx=10)
    
    lbl_instruction = tk.Label(tab5_frame, text="Vui lòng bấm nút '⚙️ XỬ LÝ DỮ LIỆU' trước.\nSau đó chọn các nút bên trên để mở từng biểu đồ phân tích chuyên sâu.", font=("Arial", 12, "italic"), fg="#666666", bg="#f0f0f0")
    lbl_instruction.pack(pady=40)

    filtered_data = {'df': None, 'cols': {}}

    def process_data_and_kpi():
        for widget in kpi_container.winfo_children(): widget.destroy()
        btn_hm.config(state=tk.DISABLED)
        btn_line.config(state=tk.DISABLED)
        btn_pareto.config(state=tk.DISABLED)

        try: import pandas as pd
        except ImportError: return messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt: pip install pandas matplotlib seaborn openpyxl")

        if not os.path.exists(TRACKING_FILE): return messagebox.showerror("Lỗi", "Chưa có file Excel (tracking-SCBH.xlsx) để thống kê.")

        try:
            df = pd.read_excel(TRACKING_FILE)
            if df.empty: return messagebox.showinfo("Trống", "File Excel chưa có dữ liệu.")
            
            def find_col(keywords):
                for col in df.columns:
                    if any(kw in str(col).lower() for kw in keywords): return col
                return None
                
            cols = {'ngay': find_col(['ngày', 'ngay']), 'nk': find_col(['người kiểm', 'nv kcs', 'kcs']), 'lh': find_col(['loại hàng', 'loai hang']),
                'sl': find_col(['sl', 'số lượng', 'so luong']), 'tt': find_col(['tình trạng', 'tinh trang']), 'loi': find_col(['lỗi', 'loi'])}
                
            tu_ngay = entry_tu_ngay.get().strip()
            den_ngay = entry_den_ngay.get().strip()
            if cols['ngay'] and (tu_ngay or den_ngay):
                df[cols['ngay']] = pd.to_datetime(df[cols['ngay']], errors='coerce')
                if tu_ngay:
                    try: df = df[df[cols['ngay']] >= pd.to_datetime(tu_ngay, format="%d/%m/%Y")]
                    except: pass
                if den_ngay:
                    try: df = df[df[cols['ngay']] <= pd.to_datetime(den_ngay, format="%d/%m/%Y")]
                    except: pass
                    
            if df.empty: return messagebox.showinfo("Thông báo", "Không có dữ liệu.")
            
            filtered_data['df'] = df
            filtered_data['cols'] = cols
            
            total_items = int(df[cols['sl']].sum()) if cols['sl'] and pd.api.types.is_numeric_dtype(df[cols['sl']]) else len(df)
            pass_rate = "0%"
            if cols['tt']:
                dat_mask = df[cols['tt']].astype(str).str.upper().str.contains('^Đ$|^ĐẠT$')
                total_valid = len(df[cols['tt']].dropna())
                if total_valid > 0: pass_rate = f"{(df[dat_mask].shape[0] / total_valid) * 100:.1f}%"
                
            top_defect = "Không có"
            if cols['loi']:
                defects = df[df[cols['loi']].astype(str).str.lower() != 'none'][cols['loi']].dropna()
                if not defects.empty: top_defect = defects.value_counts().idxmax()
                
            def make_kpi(parent, title, value, color):
                frm = tk.Frame(parent, bg=color, bd=2, relief="groove")
                frm.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
                tk.Label(frm, text=title, font=("Arial", 11, "bold"), bg=color).pack(pady=(10,0))
                tk.Label(frm, text=value, font=("Arial", 16, "bold"), bg=color, fg="#333333").pack(pady=(5,10))
                
            make_kpi(kpi_container, "Tổng Số Lượng (Năng suất)", str(total_items), "#d4edda")
            make_kpi(kpi_container, "Tỷ Lệ Đạt (Chất lượng)", pass_rate, "#cce5ff")
            make_kpi(kpi_container, "Lỗi Phổ Biến Nhất", str(top_defect), "#fff3cd")
            
            btn_hm.config(state=tk.NORMAL)
            btn_line.config(state=tk.NORMAL)
            btn_pareto.config(state=tk.NORMAL)
            lbl_instruction.config(text="Đã xử lý xong dữ liệu! Vui lòng chọn mở biểu đồ bên trên.", fg="green")
        except Exception as e: 
            messagebox.showerror("Lỗi", f"Lỗi xử lý dữ liệu:\n{e}")

    btn_load_dashboard.config(command=process_data_and_kpi)

    def open_chart_window(act):
        if filtered_data['df'] is None: return
        try:
            import pandas as pd
            import seaborn as sns
            import matplotlib.pyplot as plt
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
            import matplotlib.ticker as ticker
            
            df = filtered_data['df']
            cols = filtered_data['cols']
            
            top = tk.Toplevel(root)
            top.title(f"🔍 Phân Tích Chuyên Sâu - {act.upper()}")
            try: top.state('zoomed')
            except: top.attributes('-zoomed', True)
            
            ctrl_frame = tk.Frame(top, bg="#e2e3e5", bd=1, relief="raised")
            ctrl_frame.pack(side=tk.TOP, fill=tk.X)
            tk.Label(ctrl_frame, text="CÔNG CỤ ĐIỀU KHIỂN:", font=("Arial", 10, "bold"), bg="#e2e3e5").pack(side=tk.LEFT, padx=10)
            
            btn_zoom_in = tk.Button(ctrl_frame, text="➕ PHÓNG TO", font=("Arial", 9, "bold"), bg="#d4edda", width=15)
            btn_zoom_in.pack(side=tk.LEFT, padx=5, pady=5)
            btn_zoom_out = tk.Button(ctrl_frame, text="➖ THU NHỎ", font=("Arial", 9, "bold"), bg="#ffcccc", width=15)
            btn_zoom_out.pack(side=tk.LEFT, padx=5, pady=5)
            btn_save = tk.Button(ctrl_frame, text="💾 LƯU ẢNH", font=("Arial", 9, "bold"), bg="#cce5ff", width=15)
            btn_save.pack(side=tk.RIGHT, padx=10, pady=5)

            scroll_canvas = tk.Canvas(top, bg="white")
            v_bar = ttk.Scrollbar(top, orient="vertical", command=scroll_canvas.yview)
            h_bar = ttk.Scrollbar(top, orient="horizontal", command=scroll_canvas.xview)
            scroll_canvas.configure(yscrollcommand=v_bar.set, xscrollcommand=h_bar.set)
            
            v_bar.pack(side=tk.RIGHT, fill=tk.Y)
            h_bar.pack(side=tk.BOTTOM, fill=tk.X)
            scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            inner_frame = tk.Frame(scroll_canvas, bg="white")
            scroll_canvas.create_window((0, 0), window=inner_frame, anchor="nw")
            inner_frame.bind("<Configure>", lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))
            
            chart_state = {'scale': 1.0, 'fig': None}
            base_w, base_h = 12, 6 
            
            def draw_chart_on_canvas():
                for widget in inner_frame.winfo_children(): widget.destroy()
                plt.close('all')
                fig = Figure(dpi=100, layout='tight')
                fig.set_size_inches(base_w * chart_state['scale'], base_h * chart_state['scale'])
                ax = fig.add_subplot(111)
                chart_state['fig'] = fig
                
                if act == 'heatmap':
                    df_hm = df.copy()
                    df_hm[cols['nk']] = df_hm[cols['nk']].astype(str).str.split('-').str[0]
                    if cols['sl'] and pd.api.types.is_numeric_dtype(df_hm[cols['sl']]): 
                        pivot = pd.pivot_table(df_hm, values=cols['sl'], index=cols['nk'], columns=cols['lh'], aggfunc='sum', fill_value=0)
                    else: 
                        pivot = pd.crosstab(df_hm[cols['nk']], df_hm[cols['lh']])
                        
                    if not pivot.empty:
                        row_sums = pivot.sum(axis=1)
                        pivot = pivot.loc[row_sums.sort_values(ascending=False).index]
                        col_sums = pivot.sum(axis=0)
                        pivot = pivot[col_sums.sort_values(ascending=False).index]
                        sns.heatmap(pivot, annot=True, fmt='g', cmap='Reds', ax=ax, linewidths=.5, cbar_kws={'label': 'Số lượng SP'}, annot_kws={'rotation': 90})
                        ax.set_title('Bản đồ Nhiệt: Năng Suất', fontsize=14, fontweight='bold', pad=15)
                        ax.set_xlabel('Loại Hàng', fontsize=12)
                        ax.set_ylabel('Người Kiểm', fontsize=12)
                        for label in ax.get_xticklabels(): 
                            label.set_rotation(30)
                            label.set_ha('right')

                elif act == 'line':
                    df_line = df.copy()
                    df_line[cols['nk']] = df_line[cols['nk']].astype(str).str.split('-').str[0]
                    df_line['Date_Only'] = pd.to_datetime(df_line[cols['ngay']], errors='coerce').dt.date
                    if cols['sl'] and pd.api.types.is_numeric_dtype(df_line[cols['sl']]):
                        trend_df = df_line.groupby(['Date_Only', cols['nk']])[cols['sl']].sum().reset_index()
                        y_val = cols['sl']
                    else:
                        trend_df = df_line.groupby(['Date_Only', cols['nk']]).size().reset_index(name='Count')
                        y_val = 'Count'
                        
                    if not trend_df.empty:
                        sns.lineplot(data=trend_df, x='Date_Only', y=y_val, hue=cols['nk'], marker='o', ax=ax)
                        mean_val = trend_df[y_val].mean()
                        ax.axhline(mean_val, color='red', linestyle='--', linewidth=1.5, alpha=0.7, label=f'Trung bình: {mean_val:.1f}')
                        for idx, row in trend_df.iterrows(): 
                            ax.text(row['Date_Only'], row[y_val], str(int(row[y_val])), ha='center', va='bottom', fontsize=9, color='black')
                        ax.set_title('Biểu đồ: Năng Suất Theo Ngày', fontsize=14, fontweight='bold', pad=15)
                        ax.set_xlabel('Ngày', fontsize=12)
                        ax.set_ylabel('Số lượng SP', fontsize=12)
                        ax.legend()
                        fig.autofmt_xdate(rotation=30)

                elif act == 'pareto':
                    defects = df[df[cols['loi']].astype(str).str.lower() != 'none'][cols['loi']].dropna()
                    if not defects.empty:
                        defect_counts = defects.value_counts()
                        df_pareto = pd.DataFrame({'Loi': defect_counts.index, 'SoLuong': defect_counts.values})
                        df_pareto['TichLuy'] = df_pareto['SoLuong'].cumsum() / df_pareto['SoLuong'].sum() * 100
                        bars = ax.bar(df_pareto['Loi'], df_pareto['SoLuong'], color='#f8d7da', edgecolor='red')
                        ax.set_ylim(0, df_pareto['SoLuong'].max() * 1.15)
                        for i, bar in enumerate(bars):
                            yval = bar.get_height()
                            pct = df_pareto['TichLuy'].iloc[i]
                            ax.text(bar.get_x() + bar.get_width()/2.0, yval, f"{int(yval)}\n({pct:.1f}%)", ha='center', va='bottom', fontsize=9, color='black')
                            
                        ax2 = ax.twinx()
                        ax2.plot(df_pareto['Loi'], df_pareto['TichLuy'], color='blue', marker='D', ms=5, linewidth=2)
                        ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, loc: "{:,}%".format(int(x))))
                        ax.set_title('Biểu đồ Pareto: Phân Tích Lỗi', fontsize=14, fontweight='bold', pad=15)
                        ax.set_ylabel('Số lượng lỗi', fontsize=12)
                        ax2.set_ylabel('Tỷ lệ phần trăm tích lũy', fontsize=12)
                        ax.set_xticks(range(len(df_pareto['Loi'])))
                        ax.set_xticklabels(df_pareto['Loi'])
                        for label in ax.get_xticklabels(): 
                            label.set_rotation(35)
                            label.set_ha('right')

                canvas_agg = FigureCanvasTkAgg(fig, master=inner_frame)
                canvas_agg.draw()
                canvas_agg.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                inner_frame.update_idletasks()
                scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

            def on_zoom_in():
                if chart_state['scale'] < 3.0: 
                    chart_state['scale'] += 0.3
                    draw_chart_on_canvas()
                    
            def on_zoom_out():
                if chart_state['scale'] > 0.6: 
                    chart_state['scale'] -= 0.3
                    draw_chart_on_canvas()
                    
            def on_save():
                if chart_state['fig']:
                    filepath = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png")])
                    if filepath: 
                        chart_state['fig'].savefig(filepath, bbox_inches='tight')
                        messagebox.showinfo("Thành công", "Đã lưu ảnh biểu đồ!")

            btn_zoom_in.config(command=on_zoom_in)
            btn_zoom_out.config(command=on_zoom_out)
            btn_save.config(command=on_save)
            
            draw_chart_on_canvas()
            
        except Exception as e: 
            messagebox.showerror("Lỗi Vẽ Biểu Đồ", str(e))
            top.destroy()

    btn_hm.config(command=lambda: open_chart_window('heatmap'))
    btn_line.config(command=lambda: open_chart_window('line'))
    btn_pareto.config(command=lambda: open_chart_window('pareto'))

    # ==========================================
    # TAB 6: PIM SYSTEM - TÌM KIẾM HÌNH ẢNH LỖI 
    # ==========================================
    t6_top_frame = tk.LabelFrame(tab6_frame, text="Cấu hình & Đồng bộ Danh bạ Kho Ảnh", font=("Arial", 10, "bold"), bg="#f0f0f0")
    t6_top_frame.pack(fill=tk.X, padx=10, pady=5)
    tk.Label(t6_top_frame, text="Thư mục Kho ảnh:", bg="#f0f0f0", font=("Arial", 9)).grid(row=0, column=0, padx=10, pady=10)
    
    ent_search_dir = tk.Entry(t6_top_frame, width=50, font=("Arial", 10))
    ent_search_dir.insert(0, app_data.get("image_search_dir", ""))
    ent_search_dir.grid(row=0, column=1, padx=5, pady=10)

    def choose_search_dir():
        d = filedialog.askdirectory(title="Chọn Thư Mục")
        if d: 
            ent_search_dir.delete(0, tk.END)
            ent_search_dir.insert(0, d)
            app_data["image_search_dir"] = d
            save_data(app_data)
            
    tk.Button(t6_top_frame, text="📁 CHỌN THƯ MỤC", command=choose_search_dir, font=("Arial", 9)).grid(row=0, column=2, padx=5, pady=10)
    lbl_sync_status = tk.Label(t6_top_frame, text=f"Tổng số ảnh trong danh bạ hiện tại: {len(image_index_list)}", bg="#f0f0f0", font=("Arial", 9, "italic"), fg="blue")
    lbl_sync_status.grid(row=0, column=4, padx=20, pady=10)

    def run_sync_images():
        folder_path = ent_search_dir.get().strip()
        if not os.path.exists(folder_path): return messagebox.showerror("Lỗi", "Đường dẫn không tồn tại.")
        btn_sync.config(state=tk.DISABLED, text="⏳ ĐANG QUÉT...")
        lbl_sync_status.config(text="Đang quét thư mục... Vui lòng đợi!", fg="red")
        
        def task():
            global image_index_list; temp_list = []
            for r, _, files in os.walk(folder_path):
                for f in files:
                    if f.lower().endswith(('.jpg', '.jpeg', '.png', '.jfif')): 
                        temp_list.append(os.path.join(r, f))
            image_index_list = temp_list
            with open(INDEX_FILE, "w", encoding="utf-8") as f: 
                json.dump(image_index_list, f, ensure_ascii=False)
            root.after(0, lambda: [btn_sync.config(state=tk.NORMAL, text="🔄 ĐỒNG BỘ LẠI KHO ẢNH"), lbl_sync_status.config(text=f"✅ Đồng bộ thành công! Tìm thấy {len(image_index_list)} ảnh.", fg="green")])
            
        threading.Thread(target=task, daemon=True).start()

    btn_sync = tk.Button(t6_top_frame, text="🔄 ĐỒNG BỘ LẠI KHO ẢNH", command=run_sync_images, font=("Arial", 9, "bold"), bg="#d4edda")
    btn_sync.grid(row=0, column=3, padx=10, pady=10)

    t6_search_frame = tk.Frame(tab6_frame, bg="#f0f0f0")
    t6_search_frame.pack(fill=tk.X, padx=10, pady=(5,0))
    tk.Label(t6_search_frame, text="Từ ngày:", bg="#f0f0f0", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
    
    if HAS_CALENDAR:
        t6_ent_tu_ngay = DateEntry(t6_search_frame, width=12, font=("Arial", 10), date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        t6_ent_tu_ngay.pack(side=tk.LEFT, padx=5)
        t6_ent_tu_ngay.delete(0, tk.END)
        tk.Label(t6_search_frame, text="Đến ngày:", bg="#f0f0f0", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        t6_ent_den_ngay = DateEntry(t6_search_frame, width=12, font=("Arial", 10), date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        t6_ent_den_ngay.pack(side=tk.LEFT, padx=5)
        t6_ent_den_ngay.delete(0, tk.END)
    else:
        t6_ent_tu_ngay = tk.Entry(t6_search_frame, width=12, font=("Arial", 10))
        t6_ent_tu_ngay.pack(side=tk.LEFT, padx=5)
        tk.Label(t6_search_frame, text="Đến ngày:", bg="#f0f0f0", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        t6_ent_den_ngay = tk.Entry(t6_search_frame, width=12, font=("Arial", 10))
        t6_ent_den_ngay.pack(side=tk.LEFT, padx=5)
        
    tk.Button(t6_search_frame, text="✖ XÓA", command=lambda: [t6_ent_tu_ngay.delete(0, tk.END), t6_ent_den_ngay.delete(0, tk.END)], font=("Arial", 8, "bold"), bg="#ffcccc").pack(side=tk.LEFT, padx=5)
    tk.Label(t6_search_frame, text="Từ khóa:", bg="#f0f0f0", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(15, 5))
    ent_img_search = tk.Entry(t6_search_frame, width=30, font=("Arial", 12))
    ent_img_search.pack(side=tk.LEFT, padx=5)
    
    lbl_img_instruction = tk.Label(tab6_frame, text="💡 HƯỚNG DẪN THAO TÁC: Đưa chuột vào ảnh tìm được\n 👉 Click ĐÚP (Chuột trái) = Phóng to soi lỗi HD | 🖱️ Click CHUỘT PHẢI = Mở thẳng Folder gốc chứa ảnh", bg="#fff3cd", fg="#856404", font=("Arial", 10, "bold"), justify=tk.CENTER, bd=1, relief="solid")
    lbl_img_instruction.pack(fill=tk.X, padx=10, pady=10, ipady=5)

    gallery_container = tk.Frame(tab6_frame, bg="white", bd=1, relief="sunken")
    gallery_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
    gallery_canvas = tk.Canvas(gallery_container, bg="white")
    g_vbar = ttk.Scrollbar(gallery_container, orient="vertical", command=gallery_canvas.yview)
    gallery_canvas.configure(yscrollcommand=g_vbar.set)
    g_vbar.pack(side=tk.RIGHT, fill=tk.Y)
    gallery_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    gallery_inner = tk.Frame(gallery_canvas, bg="white")
    gallery_canvas.create_window((0, 0), window=gallery_inner, anchor="nw")
    gallery_inner.bind("<Configure>", lambda e: gallery_canvas.configure(scrollregion=gallery_canvas.bbox("all")))
    thumbnail_refs = [] 

    def open_image_location(filepath):
        try: 
            subprocess.run(['explorer', '/select,', os.path.normpath(filepath)])
        except Exception as e: 
            messagebox.showerror("Lỗi", f"Không thể mở thư mục:\n{e}")

    def open_hd_image(filepath):
        top = tk.Toplevel(root)
        top.title("🔍 Chế độ Soi Ảnh HD")
        try: top.state('zoomed')
        except: top.attributes('-zoomed', True)
        
        tb_frame = tk.Frame(top, bg="#e2e3e5", bd=1, relief="raised")
        tb_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Button(tb_frame, text="📂 MỞ THƯ MỤC GỐC", command=lambda: open_image_location(filepath), font=("Arial", 10, "bold"), bg="#fff3cd").pack(side=tk.RIGHT, padx=10, pady=5)
        tk.Label(tb_frame, text=f"Đang xem: {os.path.basename(filepath)}", bg="#e2e3e5", font=("Arial", 10, "italic")).pack(side=tk.LEFT, padx=10, pady=5)

        img_canvas = tk.Canvas(top, bg="#333333")
        v_sb = ttk.Scrollbar(top, orient="vertical", command=img_canvas.yview)
        h_sb = ttk.Scrollbar(top, orient="horizontal", command=img_canvas.xview)
        img_canvas.configure(yscrollcommand=v_sb.set, xscrollcommand=h_sb.set)
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        h_sb.pack(side=tk.BOTTOM, fill=tk.X)
        img_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        try:
            pil_img = Image.open(filepath)
            img_state = {'scale': 1.0, 'tk_img': None, 'id': None}
            def render_img():
                resized = pil_img.resize((int(pil_img.width * img_state['scale']), int(pil_img.height * img_state['scale'])), Image.Resampling.LANCZOS)
                img_state['tk_img'] = ImageTk.PhotoImage(resized)
                if img_state['id']: img_canvas.delete(img_state['id'])
                img_state['id'] = img_canvas.create_image(0, 0, anchor="nw", image=img_state['tk_img'])
                img_canvas.configure(scrollregion=img_canvas.bbox("all"))
            render_img()
            img_canvas.bind("<MouseWheel>", lambda e: [img_state.update({'scale': img_state['scale'] * (1.1 if e.delta > 0 else 1/1.1)}), render_img()])
        except Exception as e: 
            messagebox.showerror("Lỗi", f"Không đọc được file:\n{e}")
            top.destroy()

    def perform_image_search(event=None):
        start_date_int = 0; end_date_int = 99999999
        if t6_ent_tu_ngay.get().strip():
            try: start_date_int = int(datetime.strptime(t6_ent_tu_ngay.get().strip(), "%d/%m/%Y").strftime("%Y%m%d"))
            except: pass
        if t6_ent_den_ngay.get().strip():
            try: end_date_int = int(datetime.strptime(t6_ent_den_ngay.get().strip(), "%d/%m/%Y").strftime("%Y%m%d"))
            except: pass

        raw_kw = ent_img_search.get().strip().lower()
        keywords = raw_kw.split() if raw_kw else []
        
        for widget in gallery_inner.winfo_children(): widget.destroy()
        thumbnail_refs.clear()
        
        results = []
        for p in image_index_list:
            fname = os.path.basename(p)
            if start_date_int > 0 or end_date_int < 99999999:
                match = re.search(r'^(\d{8})', fname)
                if match:
                    if int(match.group(1)) < start_date_int or int(match.group(1)) > end_date_int: continue 
                else: continue 
            if keywords and not all(kw in p.lower() for kw in keywords): continue
            results.append(p)
        
        results.sort(reverse=True)
        MAX_RES = 50
        display_res = results[:MAX_RES]
        
        if not display_res: 
            return tk.Label(gallery_inner, text="Không tìm thấy ảnh nào khớp với các tiêu chí lọc.", font=("Arial", 12), bg="white").pack(pady=20)
            
        if len(results) > MAX_RES: 
            tk.Label(gallery_inner, text=f"⚠️ Tìm thấy {len(results)} ảnh. Đang hiển thị {MAX_RES} ảnh MỚI NHẤT.", fg="red", bg="white", font=("Arial", 10, "italic")).grid(row=0, column=0, columnspan=5, pady=5)
        
        row_idx = 1; col_idx = 0; COLUMNS = 4 
        for path in display_res:
            try:
                img = Image.open(path)
                img.thumbnail((250, 250), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                thumbnail_refs.append(photo)
                
                frm = tk.Frame(gallery_inner, bg="white", bd=1, relief="solid", padx=5, pady=5)
                frm.grid(row=row_idx, column=col_idx, padx=10, pady=10)
                
                lbl_img = tk.Label(frm, image=photo, bg="white", cursor="hand2")
                lbl_img.pack()
                
                fname = os.path.basename(path)
                tk.Label(frm, text=fname if len(fname) < 25 else fname[:22] + "...", bg="white", font=("Arial", 9)).pack(pady=(5,0))
                
                lbl_img.bind("<Double-Button-1>", lambda e, p=path: open_hd_image(p))
                def on_right_click(event, p=path):
                    ctx_menu = tk.Menu(root, tearoff=0, font=("Arial", 10))
                    ctx_menu.add_command(label="🔍 Xem ảnh HD (Phóng to)", command=lambda: open_hd_image(p))
                    ctx_menu.add_separator()
                    ctx_menu.add_command(label="📂 Mở thư mục chứa ảnh này", command=lambda: open_image_location(p))
                    ctx_menu.post(event.x_root, event.y_root)
                lbl_img.bind("<Button-3>", on_right_click)

                col_idx += 1
                if col_idx >= COLUMNS: 
                    col_idx = 0
                    row_idx += 1
            except: pass 

    tk.Button(t6_search_frame, text="🔍 TÌM ẢNH", command=perform_image_search, font=("Arial", 10, "bold"), bg="#cce5ff", width=15).pack(side=tk.LEFT, padx=10)
    ent_img_search.bind("<Return>", lambda e: perform_image_search())

    # ==========================================
    # TAB 7: BẢN ĐỒ TỲ VẾT
    # ==========================================
    t7_left = tk.Frame(tab7_frame, width=320, bg="#f0f0f0", bd=1, relief="ridge")
    t7_left.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
    t7_left.pack_propagate(False)
    
    t7_right = tk.Frame(tab7_frame, bg="white", bd=1, relief="sunken")
    t7_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

    # --- KHU VỰC CỘT TRÁI (THÔNG TIN & CÀO DATA) ---
    tk.Label(t7_left, text="1. NGUỒN DỮ LIỆU", font=("Arial", 10, "bold"), bg="#cce5ff").pack(fill=tk.X, pady=5)
    def t7_load_external():
        file_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg *.jpeg *.png *.jfif")])
        if file_path: t7_load_image(file_path)
    tk.Button(t7_left, text="📂 Tải ảnh ngoài từ máy tính", command=t7_load_external, bg="#e2e3e5", font=("Arial", 9)).pack(fill=tk.X, padx=10, pady=5)

    # --- BOX LƯU TẠM CHUỖI THÔNG TIN BATCH_MASP_TN ---
    tk.Label(t7_left, text="Chuỗi dữ liệu lưu tạm (Copy):", font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=10, pady=(5,0))
    f_cp = tk.Frame(t7_left, bg="#f0f0f0")
    f_cp.pack(fill=tk.X, padx=10, pady=2)
    t7_copy_entry = tk.Entry(f_cp, textvariable=t7_copy_string_var, font=("Arial", 9), state="readonly", readonlybackground="#e2e3e5")
    t7_copy_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    def btn_copy_to_clipboard():
        root.clipboard_clear()
        root.clipboard_append(t7_copy_string_var.get())
        messagebox.showinfo("Thành công", "Đã copy chuỗi dữ liệu!")
    tk.Button(f_cp, text="📋", command=btn_copy_to_clipboard, font=("Arial", 8)).pack(side=tk.RIGHT, padx=2)

    tk.Label(t7_left, text="Số TN (Chọn hoặc Nhập):", font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=10, pady=(10,0))
    t7_so_tn = ttk.Combobox(t7_left, textvariable=t7_so_tn_var, font=("Arial", 10))
    t7_so_tn.pack(fill=tk.X, padx=10, pady=2)

    tk.Label(t7_left, text="GIA No. (Tùy chọn):", font=("Arial", 9)).pack(anchor=tk.W, padx=10, pady=(5,0))
    t7_gia_no = tk.Entry(t7_left, textvariable=t7_gia_no_var, font=("Arial", 10))
    t7_gia_no.pack(fill=tk.X, padx=10, pady=2)

    t7_loai_da = tk.StringVar(value="kc")
    f_rd = tk.Frame(t7_left, bg="#f0f0f0")
    f_rd.pack(fill=tk.X, padx=10, pady=5)
    tk.Radiobutton(f_rd, text="Kim cương", variable=t7_loai_da, value="kc", bg="#f0f0f0").pack(side=tk.LEFT)
    tk.Radiobutton(f_rd, text="Đá màu", variable=t7_loai_da, value="dm", bg="#f0f0f0").pack(side=tk.LEFT)
    tk.Radiobutton(f_rd, text="Ngọc trai", variable=t7_loai_da, value="nt", bg="#f0f0f0").pack(side=tk.LEFT)

    def t7_run_scrape():
        stn = t7_so_tn_var.get().strip()
        lda = t7_loai_da.get()
        if not stn: return
        btn_t7_scrape.config(text="⏳ Đang cào dữ liệu...", state=tk.DISABLED)
        root.update()
        def task():
            try:
                session = curl_requests.Session()
                gem_data = scrape_pnj_item(session, stn, lda)
                root.after(0, lambda: t7_populate_gem_data(gem_data))
            except Exception as e: 
                root.after(0, lambda: messagebox.showerror("Lỗi", f"Lỗi cào dữ liệu:\n{e}"))
            finally: 
                root.after(0, lambda: btn_t7_scrape.config(text="⬇ CÀO DỮ LIỆU TỪ PNJLAB", state=tk.NORMAL))
        threading.Thread(target=task, daemon=True).start()

    btn_t7_scrape = tk.Button(t7_left, text="⬇ CÀO DỮ LIỆU TỪ PNJLAB", command=t7_run_scrape, bg="#d4edda", font=("Arial", 9, "bold"))
    btn_t7_scrape.pack(fill=tk.X, padx=10, pady=10)

    def t7_space_scrape(event):
        if notebook.index(notebook.select()) == 6:
            widget = root.focus_get()
            if isinstance(widget, (tk.Entry, tk.Text)) or type(widget) is ttk.Combobox:
                return
            t7_run_scrape()
            return "break"
            
    root.bind('<space>', t7_space_scrape)
    t7_so_tn.bind("<<ComboboxSelected>>", lambda e: root.focus_set())
    t7_so_tn.bind("<Return>", lambda e: [t7_run_scrape(), root.focus_set()])

    tk.Label(t7_left, text="2. THÔNG SỐ ĐÁ QUÝ", font=("Arial", 10, "bold"), bg="#cce5ff").pack(fill=tk.X, pady=(15,5))

    t7_gem_vars = {}
    t7_gem_entries = {}
    gem_fields = ["Trọng lượng", "Kích thước", "Hình dạng", "Màu sắc", "Độ tinh khiết", "Chất lượng cắt", "Độ bóng", "Độ đối xứng", "Đặc điểm", "Chú thích", "Khắc mã cạnh PNJ"]
    
    for f in gem_fields:
        fr = tk.Frame(t7_left, bg="#f0f0f0")
        fr.pack(fill=tk.X, padx=10, pady=3)
        tk.Label(fr, text=f+":", width=12, anchor=tk.W, font=("Arial", 9, "bold"), bg="#f0f0f0").pack(side=tk.LEFT)
        v = tk.StringVar()
        e = tk.Entry(fr, textvariable=v, state="readonly", bg="white", font=("Arial", 9), readonlybackground="white", relief="solid", bd=1)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True)
        t7_gem_vars[f] = v
        t7_gem_entries[f] = e

    def t7_populate_gem_data(gem_data):
        if not gem_data: return messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu cho mã Số TN này trên Web PNJLab!")
        
        for e in t7_gem_entries.values(): e.config(state="normal")
        
        t7_gem_vars['Trọng lượng'].set(f"{gem_data.get('Weight (carat)', '')} ct")
        dims = [s for s in [gem_data.get('size_1',''), gem_data.get('size_2',''), gem_data.get('size_3','')] if s]
        t7_gem_vars['Kích thước'].set(" x ".join(dims) + " mm" if dims else "")
        t7_gem_vars['Hình dạng'].set(gem_data.get('Shape & Cut', ''))
        t7_gem_vars['Màu sắc'].set(gem_data.get('Màu sắc(Color)', ''))
        t7_gem_vars['Độ tinh khiết'].set(gem_data.get('Clarity/Transparency', ''))
        t7_gem_vars['Chất lượng cắt'].set(gem_data.get('Cut Grade', ''))
        t7_gem_vars['Độ bóng'].set(gem_data.get('Polish/Luster', ''))
        t7_gem_vars['Độ đối xứng'].set(gem_data.get('Symmetry/Surface', ''))
        t7_gem_vars['Đặc điểm'].set(gem_data.get('Đặc điểm', ''))
        t7_gem_vars['Chú thích'].set(gem_data.get('Chú thích', ''))
        t7_gem_vars['Khắc mã cạnh PNJ'].set(gem_data.get('MSC_PNJ', ''))
        
        if gem_data.get('MSC_GIA'): 
            t7_gia_no_var.set(gem_data.get('MSC_GIA'))
        
        for e in t7_gem_entries.values(): e.config(state="readonly")

    # --- KHU VỰC CỘT PHẢI (DRAWING CANVAS & 2-ROW TOOLBAR CHỐNG TRÀN) ---
    t7_top_bar = tk.Frame(t7_right, bg="#e2e3e5", bd=1, relief="raised")
    t7_top_bar.pack(side=tk.TOP, fill=tk.X)
    
    t7_top_row1 = tk.Frame(t7_top_bar, bg="#e2e3e5")
    t7_top_row1.pack(side=tk.TOP, fill=tk.X, pady=2)
    t7_top_row2 = tk.Frame(t7_top_bar, bg="#e2e3e5")
    t7_top_row2.pack(side=tk.TOP, fill=tk.X, pady=2)

    t7_current_tool = tk.StringVar(value="pen_blue")
    t7_size_scale = tk.IntVar(value=4)
    t7_angle_scale = tk.IntVar(value=0)

    g_adj = tk.LabelFrame(t7_top_row1, text="Kích Cỡ & Góc độ", bg="#e2e3e5")
    g_adj.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Label(g_adj, text="Cỡ:", bg="#e2e3e5").grid(row=0, column=0, sticky="e")
    tk.Scale(g_adj, from_=2, to_=12, variable=t7_size_scale, orient=tk.HORIZONTAL, length=70, showvalue=0, bg="#e2e3e5").grid(row=0, column=1)
    tk.Label(g_adj, text="Góc:", bg="#e2e3e5").grid(row=1, column=0, sticky="e")
    tk.Scale(g_adj, from_=0, to_=360, variable=t7_angle_scale, orient=tk.HORIZONTAL, length=70, showvalue=0, bg="#e2e3e5").grid(row=1, column=1)

    g_basic = tk.LabelFrame(t7_top_row1, text="Cơ bản", bg="#e2e3e5")
    g_basic.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Radiobutton(g_basic, text="Đóng: P", variable=t7_current_tool, value="P", indicatoron=0, width=6, bg="white").grid(row=0, column=0, padx=1, pady=1)
    tk.Radiobutton(g_basic, text="Đóng: G", variable=t7_current_tool, value="G", indicatoron=0, width=6, bg="white").grid(row=0, column=1, padx=1, pady=1)
    tk.Radiobutton(g_basic, text="Bút xanh (Chấu)", variable=t7_current_tool, value="pen_blue", indicatoron=0, width=15, bg="#cce5ff").grid(row=1, column=0, columnspan=2, padx=1, pady=1)

    g_red = tk.LabelFrame(t7_top_row1, text="Tỳ vết đỏ", bg="#e2e3e5")
    g_red.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Radiobutton(g_red, text="Cấn: V", variable=t7_current_tool, value="V", indicatoron=0, width=6, bg="#f8d7da").grid(row=0, column=0, padx=1, pady=1)
    tk.Radiobutton(g_red, text="Cấn: +", variable=t7_current_tool, value="+", indicatoron=0, width=6, bg="#f8d7da").grid(row=0, column=1, padx=1, pady=1)
    tk.Radiobutton(g_red, text="Chấm đỏ", variable=t7_current_tool, value="dot_red", indicatoron=0, width=8, bg="#f8d7da").grid(row=0, column=2, padx=1, pady=1)
    tk.Radiobutton(g_red, text="Nét thẳng", variable=t7_current_tool, value="line_red", indicatoron=0, width=14, bg="#f8d7da").grid(row=1, column=0, columnspan=2, padx=1, pady=1)
    tk.Radiobutton(g_red, text="Zigzag", variable=t7_current_tool, value="zigzag", indicatoron=0, width=8, bg="#f8d7da").grid(row=1, column=2, padx=1, pady=1)

    g_del = tk.LabelFrame(t7_top_row1, text="Xóa", bg="#e2e3e5")
    g_del.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    tk.Radiobutton(g_del, text="🧽 Cục tẩy", variable=t7_current_tool, value="eraser", indicatoron=0, width=10, bg="white").pack(pady=2, padx=2)
    tk.Button(g_del, text="✖ Xóa tất cả", command=lambda: [t7_canvas.delete("drawing"), t7_draw_state['history'].clear()], bg="#f8d7da", width=10).pack(pady=2, padx=2)

    g_gia = tk.LabelFrame(t7_top_row2, text="GIA Standard", bg="#e2e3e5")
    g_gia.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    gia_tools = [
        ("Crystal", "crystal", "red"), ("Pinpoint", "pinpoint", "red"), ("Needle", "needle", "red"), ("Cloud", "cloud", "red"), ("Feather", "feather", "red"),
        ("Tw. Wisp", "twinning_wisp", "red"), ("Cavity", "cavity", "green"), ("Bruise", "bruise", "red"), ("Chip", "chip", "red"), ("Laser DH", "laser_drill", "green"),
        ("Knot", "knot", "green"), ("Etch Ch.", "etch_channel", "green"), ("Ind. Nat.", "indented_natural", "green"), ("Natural", "natural", "green"), ("Ex. Facet", "extra_facet", "black")
    ]
    for i, (txt, val, col) in enumerate(gia_tools):
        r, c = divmod(i, 8)
        tk.Radiobutton(g_gia, text=txt, variable=t7_current_tool, value=val, indicatoron=0, width=8, fg=col, font=("Arial", 8, "bold")).grid(row=r, column=c, padx=1, pady=1)

    t7_scroll_frame = tk.Frame(t7_right, bg="gray")
    t7_scroll_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    t7_canvas = tk.Canvas(t7_scroll_frame, bg="white", cursor="crosshair", scrollregion=(0,0,1600,1200))
    cv_vbar = ttk.Scrollbar(t7_scroll_frame, orient="vertical", command=t7_canvas.yview)
    cv_hbar = ttk.Scrollbar(t7_scroll_frame, orient="horizontal", command=t7_canvas.xview)
    t7_canvas.configure(yscrollcommand=cv_vbar.set, xscrollcommand=cv_hbar.set)
    
    cv_vbar.pack(side=tk.RIGHT, fill=tk.Y)
    cv_hbar.pack(side=tk.BOTTOM, fill=tk.X)
    t7_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    t7_bottom_bar = tk.Frame(t7_right, bg="#e2e3e5", bd=1, relief="raised")
    t7_bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)
    tk.Label(t7_bottom_bar, text="Ghi chú PDF:", bg="#e2e3e5", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=10)
    tk.Entry(t7_bottom_bar, textvariable=t7_note_var, width=50, font=("Arial", 10)).pack(side=tk.LEFT, padx=5, pady=8)

    # --- LIVE UPDATE HEADER/FOOTER CHO TAB 7 CANVAS ---
    def update_t7_canvas_texts(*args):
        stn = t7_so_tn_var.get().strip()
        gia = t7_gia_no_var.get().strip()
        t7_canvas.itemconfig("header_text", text=f"Số TN: {stn}                      GIA No.: {gia}")
        note = t7_note_var.get().strip()
        t7_canvas.itemconfig("note_display", text=f"Ghi chú: {note}" if note else "Ghi chú: " + "."*90)

    t7_so_tn_var.trace_add("write", update_t7_canvas_texts)
    t7_gia_no_var.trace_add("write", update_t7_canvas_texts)
    t7_note_var.trace_add("write", update_t7_canvas_texts)

    def t7_select_photos():
        folder = filedialog.askdirectory(title="Chọn thư mục chứa ảnh chụp thực tế đính kèm")
        if not folder: return
        all_files = [os.path.join(r, f) for r, d, files in os.walk(folder) for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.jfif'))]
        if not all_files: return messagebox.showinfo("Thông báo", "Không tìm thấy ảnh nào!")

        win = tk.Toplevel(root)
        win.title(f"Lựa Chọn Hình Ảnh ({len(all_files)} ảnh)")
        win.geometry("750x550")
        win.grab_set()
        
        mf = tk.Frame(win)
        mf.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)
        cv = tk.Canvas(mf)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        sb = tk.Scrollbar(mf, orient=tk.VERTICAL, command=cv.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.configure(yscrollcommand=sb.set)
        cv.bind('<Configure>', lambda e: cv.configure(scrollregion=cv.bbox("all")))
        
        gf = tk.Frame(cv)
        cv.create_window((0, 0), window=gf, anchor="nw")
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        c_vars = []
        win.thumbs = []
        for idx, p in enumerate(all_files):
            c = tk.Frame(gf, borderwidth=1, relief="solid", padx=5, pady=5)
            c.grid(row=idx//4, column=idx%4, padx=10, pady=10)
            name = os.path.basename(p)
            try:
                img = Image.open(p)
                img.thumbnail((120, 120))
                photo = ImageTk.PhotoImage(img)
                win.thumbs.append(photo)
                tk.Label(c, image=photo).pack()
            except: 
                tk.Label(c, text="[Lỗi file]", width=15, height=5, fg="red").pack()
            
            var = tk.BooleanVar(value=True)
            c_vars.append((var, p))
            tk.Checkbutton(c, text=name if len(name)<=32 else name[:29]+"...", variable=var, font=("Arial", 9), wraplength=110).pack()

        bf = tk.Frame(win)
        bf.pack(fill=tk.X, pady=10)
        tk.Button(bf, text="Chọn tất cả", command=lambda: [v.set(True) for v, _ in c_vars], width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(bf, text="Bỏ chọn tất cả", command=lambda: [v.set(False) for v, _ in c_vars], width=15).pack(side=tk.LEFT, padx=10)
        def on_ok():
            t7_state['actual_photos'] = [p for v, p in c_vars if v.get()]
            lbl_t7_photo_count.config(text=f"{len(t7_state['actual_photos'])} ảnh")
            cv.unbind_all("<MouseWheel>")
            win.destroy()
            
        tk.Button(bf, text="✔ Lưu lựa chọn", command=on_ok, bg="#d4edda", font=("Arial", 10, "bold"), width=15).pack(side=tk.RIGHT, padx=20)
        
        win.update_idletasks()
        w = win.winfo_reqwidth() + 60
        h = win.winfo_reqheight() + 60
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

    btn_t7_photos = tk.Button(t7_bottom_bar, text="📸 Đính kèm ảnh thực tế", command=t7_select_photos, bg="#cce5ff", font=("Arial", 9, "bold"))
    btn_t7_photos.pack(side=tk.LEFT, padx=15)
    lbl_t7_photo_count = tk.Label(t7_bottom_bar, text="0 ảnh", fg="blue", bg="#e2e3e5", font=("Arial", 9, "italic"))
    lbl_t7_photo_count.pack(side=tk.LEFT)

    def t7_generate_pdf():
        buffer = io.BytesIO()
        try:
            c = pdf_canvas.Canvas(buffer, pagesize=landscape(A5))
            w, h = landscape(A5) 
            
            try: 
                pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
                font_name = 'Arial'
                pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
                font_bold = 'Arial-Bold'
            except: 
                font_name = 'Helvetica'
                font_bold = 'Helvetica-Bold'
                
            margin_pdf = 10 * mm
            map_box_h = 55 * mm
            map_box_y = h - margin_pdf - map_box_h
            left_w = (w / 2) - margin_pdf
            map_box_x = margin_pdf
            map_box_w = left_w
            
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0, 0, 0)
            c.roundRect(map_box_x, map_box_y, map_box_w, map_box_h, 2*mm, stroke=1, fill=0)
            
            c.setFont(font_bold, 8)
            c.drawString(map_box_x + 3*mm, map_box_y + map_box_h - 4*mm, f"Số TN: {t7_so_tn_var.get().strip()}")
            c.drawString(map_box_x + map_box_w/2, map_box_y + map_box_h - 4*mm, f"GIA No.: {t7_gia_no_var.get().strip()}")
            
            note_text = t7_note_var.get().strip() or "..................................................."
            c.setFont(font_name, 8)
            c.drawString(map_box_x + 3*mm, map_box_y + 3*mm, f"Ghi chú: {note_text}")
            
            if t7_state['image_path'] and os.path.exists(t7_state['image_path']):
                pdf_img_w = map_box_w - 6*mm
                pdf_img_h = pdf_img_w * t7_state['img_ratio']
                if pdf_img_h > (map_box_h - 10*mm):
                    pdf_img_h = map_box_h - 10*mm
                    pdf_img_w = pdf_img_h / t7_state['img_ratio']
                
                pdf_img_x = map_box_x + (map_box_w - pdf_img_w) / 2
                pdf_img_y = map_box_y + (map_box_h - pdf_img_h) / 2
                
                try:
                    map_pil = Image.open(t7_state['image_path'])
                    if map_pil.mode != "RGB": map_pil = map_pil.convert("RGB")
                    c.drawImage(ImageReader(map_pil), pdf_img_x, pdf_img_y, width=pdf_img_w, height=pdf_img_h, preserveAspectRatio=True, mask='auto')
                except: pass
                
                pdf_cx = pdf_img_x + pdf_img_w / 2
                pdf_cy = pdf_img_y + pdf_img_h / 2
                scale_to_pdf = pdf_img_w / t7_state['img_w']
                
                for item in t7_canvas.find_withtag("drawing"):
                    ctype = t7_canvas.type(item)
                    try: 
                        color_source = t7_canvas.itemcget(item, "outline") if ctype in ["oval", "rectangle", "polygon"] else t7_canvas.itemcget(item, "fill")
                    except: 
                        color_source = "black"
                    
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
                        dash = t7_canvas.itemcget(item, "dash")
                        if dash: 
                            c.setDash([1.5, 2.5], 0) if "1" in str(dash) and "3" in str(dash) else c.setDash([2, 2], 0)
                        else: 
                            c.setDash([], 0)
                    except tk.TclError: 
                        c.setDash([], 0)
                        
                    if ctype == "line":
                        coords = t7_canvas.coords(item)
                        p = c.beginPath()
                        p.moveTo(pdf_cx + (coords[0] - t7_state['ui_cx']) * scale_to_pdf, pdf_cy - (coords[1] - t7_state['ui_cy']) * scale_to_pdf)
                        for i in range(2, len(coords), 2): 
                            p.lineTo(pdf_cx + (coords[i] - t7_state['ui_cx']) * scale_to_pdf, pdf_cy - (coords[i+1] - t7_state['ui_cy']) * scale_to_pdf)
                        try: width = float(t7_canvas.itemcget(item, "width")) * scale_to_pdf
                        except: width = 1.0 * scale_to_pdf
                        c.setLineWidth(max(0.5, width))
                        c.drawPath(p, stroke=1, fill=0)
                        
                    elif ctype == "oval":
                        coords = t7_canvas.coords(item)
                        px = pdf_cx + ((coords[0]+coords[2])/2 - t7_state['ui_cx']) * scale_to_pdf
                        py = pdf_cy - ((coords[1]+coords[3])/2 - t7_state['ui_cy']) * scale_to_pdf
                        c.circle(px, py, abs(coords[2]-coords[0])/2 * scale_to_pdf, stroke=1, fill=1 if t7_canvas.itemcget(item, "fill") else 0)
                    
                    elif ctype == "polygon":
                        coords = t7_canvas.coords(item)
                        p = c.beginPath()
                        p.moveTo(pdf_cx + (coords[0] - t7_state['ui_cx']) * scale_to_pdf, pdf_cy - (coords[1] - t7_state['ui_cy']) * scale_to_pdf)
                        for i in range(2, len(coords), 2): 
                            p.lineTo(pdf_cx + (coords[i] - t7_state['ui_cx']) * scale_to_pdf, pdf_cy - (coords[i+1] - t7_state['ui_cy']) * scale_to_pdf)
                        p.closePath()
                        c.drawPath(p, stroke=1, fill=1 if t7_canvas.itemcget(item, "fill") else 0)
                        
                    elif ctype == "text":
                        coords = t7_canvas.coords(item)
                        px = pdf_cx + (coords[0] - t7_state['ui_cx']) * scale_to_pdf
                        py = pdf_cy - (coords[1] - t7_state['ui_cy']) * scale_to_pdf
                        text = t7_canvas.itemcget(item, "text")
                        font_str = t7_canvas.itemcget(item, "font")
                        try: angle_val = float(t7_canvas.itemcget(item, "angle"))
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

            title_x = margin_pdf
            val_x = margin_pdf + 30 * mm
            pdf_max_val_w = left_w - val_x - 5*mm 
            start_y = map_box_y - 7 * mm
            line_spacing = 4.5 * mm
            
            c.setFont(font_bold, 11)
            c.setFillColorRGB(0, 0, 0)
            c.drawString(title_x, start_y, "THÔNG SỐ ĐÁ QUÝ" if font_name == 'Arial' else "THONG SO DA QUY")
            start_y -= 7 * mm
            
            for t_str, var in t7_gem_vars.items():
                val = var.get().strip()
                c.setFont(font_bold, 9)
                c.drawString(title_x, start_y, t_str)
                c.setFont(font_name, 9)
                if not val: 
                    start_y -= line_spacing
                    continue
                for ln in simpleSplit(val, font_name, 9, pdf_max_val_w): 
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
            
            n = len(t7_state['actual_photos'])
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
                
                for idx, p_path in enumerate(t7_state['actual_photos']):
                    try:
                        pic_pil = Image.open(p_path)
                        if pic_pil.mode != "RGB": pic_pil = pic_pil.convert("RGB")
                        c.drawImage(ImageReader(pic_pil), right_x + (idx%c_cols)*cell_w + p, right_y + right_h - (idx//c_cols + 1)*cell_h + p, width=cell_w - 2*p, height=cell_h - 2*p, preserveAspectRatio=True, mask='auto')
                    except: pass
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

    def t7_open_save_dialog():
        folder = filedialog.askdirectory(title="Bước 1: Chọn thư mục lưu Báo cáo")
        if not folder: return

        swin = tk.Toplevel(root)
        swin.title("Lưu Báo Cáo")
        swin.grab_set()
        
        swin.update_idletasks()
        w = swin.winfo_reqwidth() + 60
        h = swin.winfo_reqheight() + 60
        x = (swin.winfo_screenwidth() // 2) - (w // 2)
        y = (swin.winfo_screenheight() // 2) - (h // 2)
        swin.geometry(f'{w}x{h}+{x}+{y}')
        
        tk.Label(swin, text="Thư mục:", font=("Arial", 9, "bold")).pack(pady=(15, 0))
        tk.Label(swin, text=folder, font=("Arial", 8), fg="gray", wraplength=350).pack(pady=(0, 10))
        tk.Label(swin, text="Tên file:", font=("Arial", 10, "bold")).pack()
        name_var = tk.StringVar(value=f"{t7_so_tn_var.get()}_Report" if t7_so_tn_var.get() else "GIA_Report")
        tk.Entry(swin, textvariable=name_var, width=35, font=("Arial", 10)).pack(pady=5)
        tk.Label(swin, text="Định dạng xuất:", font=("Arial", 10, "bold")).pack()
        fmt_var = tk.StringVar(value="PDF")
        ttk.Combobox(swin, textvariable=fmt_var, values=["PDF", "PNG", "JPG"], state="readonly", width=32).pack(pady=5)

        def on_preview():
            fn = name_var.get().strip()
            fmt = fmt_var.get()
            if not fn: return messagebox.showerror("Lỗi", "Vui lòng nhập tên file!", parent=swin)
            if not HAS_FITZ: return messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt: pip install pymupdf\nđể có thể xuất ảnh.", parent=swin)

            pdf_buffer = t7_generate_pdf()
            if pdf_buffer:
                pwin = tk.Toplevel(swin)
                pwin.title("Xem Trước Báo Cáo")
                pwin.geometry("900x650")
                pwin.grab_set()
                
                pwin.update_idletasks()
                pw = pwin.winfo_reqwidth() + 60
                ph = pwin.winfo_reqheight() + 60
                px = (pwin.winfo_screenwidth() // 2) - (pw // 2)
                py = (pwin.winfo_screenheight() // 2) - (ph // 2)
                pwin.geometry(f'{pw}x{ph}+{px}+{py}')
                
                try:
                    pdf_bytes = pdf_buffer.getvalue()
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    pix = doc[0].get_pixmap(dpi=120) 
                    img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    photo = ImageTk.PhotoImage(img_pil)
                    lbl = tk.Label(pwin, image=photo)
                    lbl.image = photo
                    lbl.pack(pady=10, expand=True)
                except Exception as e:
                    return messagebox.showerror("Lỗi", f"Lỗi hiển thị: {e}", parent=pwin)
                    
                bf = tk.Frame(pwin)
                bf.pack(fill=tk.X, pady=10)
                
                def on_save_final():
                    try:
                        out = os.path.join(folder, fn + (".pdf" if fmt=="PDF" else f".{fmt.lower()}"))
                        if fmt == "PDF":
                            with open(out, "wb") as f: f.write(pdf_bytes)
                        else: 
                            doc[0].get_pixmap(dpi=300).save(out)
                        doc.close()
                        pwin.destroy()
                        swin.destroy()
                        messagebox.showinfo("Thành công", f"Đã xuất báo cáo đồ giải thành công tại:\n{out}")
                    except Exception as e: 
                        messagebox.showerror("Lỗi", f"Lỗi lưu file: {e}", parent=pwin)
                
                tk.Button(bf, text="✖ Trở lại (Sửa)", command=lambda: [doc.close(), pwin.destroy()], bg="#f8d7da", width=20, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=50)
                tk.Button(bf, text="✔ Chốt & Lưu file", command=on_save_final, bg="#d4edda", width=20, font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=50)

        tk.Button(swin, text="👁 Tiếp tục (Xem trước)", command=on_preview, bg="#d4edda", width=25, font=("Arial", 10, "bold")).pack(pady=15)

    btn_t7_save = tk.Button(t7_bottom_bar, text="💾 XUẤT BÁO CÁO", command=t7_open_save_dialog, bg="#d4edda", font=("Arial", 10, "bold"))
    btn_t7_save.pack(side=tk.RIGHT, padx=15, pady=5)

    def t7_load_image(img_path):
        t7_state['image_path'] = img_path
        t7_canvas.delete("all")
        t7_draw_state['history'].clear()
        
        if not img_path or not os.path.exists(img_path):
            img_w = 650
            img_h = 650
            t7_state['img_ratio'] = 1.0
            photo = None
        else:
            try:
                img = Image.open(img_path)
                orig_w, orig_h = img.size
                t7_state['img_ratio'] = orig_h / float(orig_w)
                img_w = 650
                img_h = int(img_w * t7_state['img_ratio'])
                img = img.resize((img_w, img_h), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                t7_canvas.image = photo 
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không đọc được ảnh:\n{e}")
                return

        t7_state['img_w'] = img_w
        t7_state['img_h'] = img_h

        padding = 25
        margin = 15
        canvas_w = img_w + padding * 2 + 20
        canvas_h = img_h + 160
        t7_state['ui_cx'] = padding + 10 + img_w / 2.0
        t7_state['ui_cy'] = canvas_h / 2.0
        t7_canvas.configure(scrollregion=(0, 0, canvas_w, canvas_h))

        t7_canvas.create_rectangle(margin, margin, canvas_w - margin, canvas_h - margin, dash=(6, 4), outline="gray", tags="bg_layer")
        t7_canvas.create_rectangle(padding, padding, canvas_w - padding, canvas_h - padding, outline="black", width=1.5, tags="bg_layer")

        img_x_start = padding + 10
        img_y_start = (canvas_h / 2) - (img_h / 2)
        t7_canvas.create_rectangle(img_x_start, img_y_start, img_x_start + img_w, img_y_start + img_h, outline="#666666", width=1.5, dash=(4, 2), tags="bg_layer")
        
        if photo:
            t7_canvas.create_image(t7_state['ui_cx'], t7_state['ui_cy'], anchor=tk.CENTER, image=photo, tags="bg_image")
        else:
            t7_canvas.create_text(t7_state['ui_cx'], t7_state['ui_cy'], text="(Không có ảnh sản phẩm)", font=("Arial", 14, "italic"), fill="gray", tags="bg_image")
            
        t7_canvas.create_text(padding + 20, padding + 20, text="", anchor=tk.W, font=("Arial", 12, "bold"), tags="header_text")
        t7_canvas.create_text(padding + 20, canvas_h - padding - 25, text="", anchor=tk.W, font=("Arial", 12), tags="note_display")
        update_t7_canvas_texts()

    def t7_on_hover(event):
        t7_canvas.delete("preview")
        if t7_draw_state['is_drawing']: return  
        tool, sz, angle_deg = t7_current_tool.get(), t7_size_scale.get(), t7_angle_scale.get()
        x = t7_canvas.canvasx(event.x)
        y = t7_canvas.canvasy(event.y)
        r = max(2, sz)
        
        t7_canvas.create_line(x-5, y, x+5, y, fill="gray", tags="preview")
        t7_canvas.create_line(x, y-5, x, y+5, fill="gray", tags="preview")
        
        if tool in ['P', 'G']: 
            t7_canvas.create_text(x, y, text=tool, font=("Arial", sz*3, "bold"), fill="#aaaaaa", angle=angle_deg, tags="preview")
        elif tool == 'V':
            rad, spread, l = math.radians(angle_deg), math.radians(35), sz * 2.5
            dx1, dy1 = l * math.sin(rad - spread), -l * math.cos(rad - spread)
            dx2, dy2 = l * math.sin(rad + spread), -l * math.cos(rad + spread)
            t7_canvas.create_line(x, y, x + dx1, y + dy1, fill="#ff8888", width=2, capstyle=tk.ROUND, tags="preview")
            t7_canvas.create_line(x, y, x + dx2, y + dy2, fill="#ff8888", width=2, capstyle=tk.ROUND, tags="preview")
        elif tool == '+':
            rad = math.radians(angle_deg); l = sz * 1.5
            dx1, dy1 = l * math.cos(rad), l * math.sin(rad)
            dx2, dy2 = -l * math.sin(rad), l * math.cos(rad)
            t7_canvas.create_line(x-dx1, y-dy1, x+dx1, y+dy1, fill="#ff8888", width=2, tags="preview")
            t7_canvas.create_line(x-dx2, y-dy2, x+dx2, y+dy2, fill="#ff8888", width=2, tags="preview")
        elif tool == 'dot_red': 
            t7_canvas.create_oval(x-r, y-r, x+r, y+r, outline="#ff8888", width=2, tags="preview")
        elif tool in [t[1] for t in gia_tools]: 
            draw_gia_symbol(t7_canvas, tool, x, y, sz, angle_deg, "preview")
        elif tool in ['pen_blue', 'line_red', 'zigzag', 'needle', 'twinning_wisp']: 
            t7_canvas.create_oval(x-r, y-r, x+r, y+r, outline="gray", dash=(2,2), tags="preview")
        elif tool == 'feather': 
            t7_canvas.create_oval(x-r, y-r, x+r, y+r, outline="red", dash=(2,2), tags="preview")
        elif tool == 'eraser': 
            t7_canvas.create_rectangle(x-sz*2, y-sz*2, x+sz*2, y+sz*2, outline="gray", dash=(2,2), tags="preview")

    def t7_on_click(event):
        t7_canvas.focus_set()
        t7_draw_state['is_drawing'] = True
        t7_canvas.delete("preview")
        t7_draw_state['current_action'] = [] 
        
        tool, sz, angle_deg = t7_current_tool.get(), t7_size_scale.get(), t7_angle_scale.get()
        x = t7_canvas.canvasx(event.x)
        y = t7_canvas.canvasy(event.y)
        r = max(2, sz)

        if tool == 'P': 
            t7_draw_state['history'].append([t7_canvas.create_text(x, y, text="P", font=("Arial", sz*3, "bold"), fill="black", angle=angle_deg, tags="drawing")])
        elif tool == 'G': 
            t7_draw_state['history'].append([t7_canvas.create_text(x, y, text="G", font=("Arial", sz*3, "bold"), fill="black", angle=angle_deg, tags="drawing")])
        elif tool == 'V':
            rad, spread, l = math.radians(angle_deg), math.radians(35), sz * 2.5
            dx1, dy1 = l * math.sin(rad - spread), -l * math.cos(rad - spread)
            dx2, dy2 = l * math.sin(rad + spread), -l * math.cos(rad + spread)
            item1 = t7_canvas.create_line(x, y, x + dx1, y + dy1, fill="red", width=2, capstyle=tk.ROUND, tags="drawing")
            item2 = t7_canvas.create_line(x, y, x + dx2, y + dy2, fill="red", width=2, capstyle=tk.ROUND, tags="drawing")
            t7_draw_state['history'].append([item1, item2])
        elif tool == '+':
            rad = math.radians(angle_deg); l = sz * 1.5
            dx1, dy1 = l * math.cos(rad), l * math.sin(rad)
            dx2, dy2 = -l * math.sin(rad), l * math.cos(rad)
            item1 = t7_canvas.create_line(x-dx1, y-dy1, x+dx1, y+dy1, fill="red", width=2, tags="drawing")
            item2 = t7_canvas.create_line(x-dx2, y-dy2, x+dx2, y+dy2, fill="red", width=2, tags="drawing")
            t7_draw_state['history'].append([item1, item2])
        elif tool == 'dot_red': 
            t7_draw_state['history'].append([t7_canvas.create_oval(x-r, y-r, x+r, y+r, fill="red", outline="red", tags="drawing")])
        elif tool in [t[1] for t in gia_tools] and tool != 'cloud':
            items = draw_gia_symbol(t7_canvas, tool, x, y, sz, angle_deg, "drawing")
            if items: t7_draw_state['history'].append(items)
        elif tool == 'cloud':
            items = draw_gia_symbol(t7_canvas, tool, x, y, sz, angle_deg, "drawing")
            t7_draw_state['current_action'].extend(items)
            t7_draw_state['draw_points'].clear()
            t7_draw_state['draw_points'].append((x, y))
        elif tool in ['pen_blue', 'line_red', 'zigzag', 'needle', 'feather', 'twinning_wisp']:
            t7_draw_state['draw_points'].clear()
            t7_draw_state['draw_points'].append((x, y))
            if tool == 'pen_blue': 
                t7_draw_state['current_action'].append(t7_canvas.create_oval(x-r/1.5, y-r/1.5, x+r/1.5, y+r/1.5, fill="blue", outline="blue", tags="drawing"))
        elif tool == 'eraser':
            for item in t7_canvas.find_overlapping(x - sz*2, y - sz*2, x + sz*2, y + sz*2):
                if "drawing" in t7_canvas.gettags(item): t7_canvas.delete(item)

    def t7_on_drag(event):
        tool, sz = t7_current_tool.get(), t7_size_scale.get()
        x = t7_canvas.canvasx(event.x)
        y = t7_canvas.canvasy(event.y)
        dp = t7_draw_state['draw_points']
        ca = t7_draw_state['current_action']
        ti = t7_draw_state['temp_items']
        
        if tool == 'pen_blue' and dp:
            last_x, last_y = dp[-1]
            item = t7_canvas.create_line(last_x, last_y, x, y, fill="blue", width=max(1, sz), capstyle=tk.ROUND, smooth=True, tags="drawing")
            ca.append(item)
            dp.append((x, y))
            
        elif tool in ['line_red', 'needle'] and dp:
            for i in ti: t7_canvas.delete(i)
            ti.clear()
            item = t7_canvas.create_line(dp[0][0], dp[0][1], x, y, fill="red", width=max(1, sz), tags=("drawing", "temp_line"))
            ti.append(item)
            
        elif tool == 'zigzag' and dp:
            last_x, last_y = dp[-1]
            item = t7_canvas.create_line(last_x, last_y, x, y, fill="pink", width=1, dash=(2,2), tags=("drawing", "temp_zig"))
            ti.append(item)
            dp.append((x, y))
            
        elif tool in ['cloud', 'feather', 'twinning_wisp'] and dp:
            if tool == 'cloud' and len(dp) == 1 and ca:
                for i in ca: t7_canvas.delete(i)
                ca.clear()
            last_x, last_y = dp[-1]
            dash_pattern = (1, 3) if tool == 'cloud' else ()
            
            new_line = t7_canvas.create_line(
                last_x, last_y, x, y, 
                fill="red", 
                width=max(1, sz), 
                dash=dash_pattern, 
                capstyle=tk.ROUND, 
                smooth=True, 
                tags="drawing"
            )
            ca.append(new_line)
            dp.append((x, y))
            
        elif tool == 'eraser':
            for item in t7_canvas.find_overlapping(x - sz*2, y - sz*2, x + sz*2, y + sz*2):
                if "drawing" in t7_canvas.gettags(item): t7_canvas.delete(item)

    def t7_on_release(event):
        t7_draw_state['is_drawing'] = False
        tool, sz = t7_current_tool.get(), t7_size_scale.get()
        x = t7_canvas.canvasx(event.x)
        y = t7_canvas.canvasy(event.y)
        dp = t7_draw_state['draw_points']
        ca = t7_draw_state['current_action']
        ti = t7_draw_state['temp_items']
        
        if tool == 'zigzag':
            for i in ti: t7_canvas.delete(i)
            ti.clear()
            if len(dp) > 1:
                zz = generate_zigzag(dp, step=sz*2 + 2, amplitude=sz*1.5)
                if zz: 
                    item = t7_canvas.create_line(*zz, fill="red", width=max(1, int(sz*0.8)), joinstyle=tk.MITER, tags="drawing")
                    t7_draw_state['history'].append([item])
            dp.clear()
            
        elif tool in ['line_red', 'needle']:
            for i in ti: t7_canvas.delete(i)
            ti.clear()
            if len(dp) > 0: 
                item = t7_canvas.create_line(dp[0][0], dp[0][1], x, y, fill="red", width=max(1, sz), tags="drawing")
                t7_draw_state['history'].append([item])
            dp.clear()
            
        elif tool == 'twinning_wisp':
            if ca:
                resampled = [dp[0]]
                for p in dp[1:]:
                    if math.hypot(p[0]-resampled[-1][0], p[1]-resampled[-1][1]) >= 10: 
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
                    item = t7_canvas.create_line(cx - nx*l, cy - ny*l, cx + nx*l, cy + ny*l, fill="red", width=max(1, sz), tags="drawing")
                    ca.append(item)
                t7_draw_state['history'].append(list(ca))
            ca.clear()
            dp.clear()
            
        elif tool in ['feather', 'cloud', 'pen_blue']:
            if ca: 
                t7_draw_state['history'].append(list(ca))
            ca.clear()
            dp.clear()
            
        t7_on_hover(event)

    def t7_undo_action(event=None):
        if notebook.index(notebook.select()) != 6: return 
        if isinstance(root.focus_get(), tk.Entry): return
        if t7_draw_state['history']:
            for item in t7_draw_state['history'].pop():
                try: t7_canvas.delete(item)
                except: pass

    root.bind("<Control-z>", t7_undo_action)
    root.bind("<Control-Z>", t7_undo_action)
    
    t7_canvas.bind("<Motion>", t7_on_hover)
    t7_canvas.bind("<Leave>", lambda e: t7_canvas.delete("preview"))
    t7_canvas.bind("<Button-1>", t7_on_click)
    t7_canvas.bind("<B1-Motion>", t7_on_drag)
    t7_canvas.bind("<ButtonRelease-1>", t7_on_release)

    # KHỞI TẠO CÁC NÚT BẤM DƯỚI CÙNG
    btn_nhap = tk.Button(bottom_frame, text="📥 NHẬP (Đẩy sang Tab 2)", command=push_to_tab2, font=("Arial", 10, "bold"), bg="#d4edda", width=25, height=2)
    btn_sua = tk.Button(bottom_frame, text="✏️ CHỈNH SỬA DÒNG", command=sua_dong_tab2, font=("Arial", 10, "bold"), bg="#fff3cd", height=2)
    btn_copy = tk.Button(bottom_frame, text="📋 COPY DỮ LIỆU", command=copy_data_tab2, font=("Arial", 10, "bold"), bg="#e2e3e5", height=2)
    btn_xoa = tk.Button(bottom_frame, text="🗑️ XÓA DÒNG ĐÃ CHỌN", command=xoa_dong_tab2, font=("Arial", 10, "bold"), bg="#f8d7da", height=2)
    btn_xuat = tk.Button(bottom_frame, text="📤 XUẤT FILE EXCEL", command=export_excel, font=("Arial", 10, "bold"), bg="#fff3cd", width=20, height=2)

    lbl_tong_sl = tk.Label(bottom_frame, text="Tổng SL tạm tính: 0", font=("Arial", 12, "bold"), bg="#005b52", fg="#f1d09f")
    
    def update_tong_sl():
        total = sum(float(r[3]) for r in app_data.get("tab2_rows", []) if str(r[3]).replace('.','',1).isdigit())
        total_float = float(total)
        lbl_tong_sl.config(text=f"Tổng SL tạm (Tab 2): {int(total_float) if total_float.is_integer() else total_float}")

    def on_tab_change(event):
        selected_tab = notebook.index(notebook.select())
        btn_nhap.pack_forget()
        btn_sua.pack_forget()
        btn_xoa.pack_forget()
        btn_copy.pack_forget()
        btn_xuat.pack_forget()
        lbl_tong_sl.pack_forget()

        if selected_tab == 0:   
            bottom_frame.config(bg="#005b52", bd=1, relief="ridge")
            btn_nhap.pack(side=tk.LEFT, padx=20, pady=10)
            lbl_tong_sl.pack(side=tk.RIGHT, padx=20, pady=10)
            entries["Batch"].focus() 
        elif selected_tab == 1: 
            bottom_frame.config(bg="#005b52", bd=1, relief="ridge")
            btn_sua.pack(side=tk.LEFT, padx=(10, 5), pady=10)
            btn_xoa.pack(side=tk.LEFT, padx=(5, 10), pady=10)
            btn_copy.pack(side=tk.LEFT, padx=10, pady=10)
            btn_xuat.pack(side=tk.RIGHT, padx=20, pady=10)
            lbl_tong_sl.pack(side=tk.RIGHT, padx=20, pady=10)
        else:
            bottom_frame.config(bg=root.cget('bg'), bd=0, relief="flat")

    update_tong_sl()
    notebook.bind("<<NotebookTabChanged>>", on_tab_change)
    on_tab_change(None)

    # --- KHAI BÁO CÁC PHÍM TẮT TOÀN CẦU (GLOBAL HOTKEYS) ---
    def global_save(event=None):
        current_tab = notebook.index(notebook.select())
        if current_tab == 1: 
            export_excel()
        elif current_tab == 6:
            t7_open_save_dialog()

    def global_new(event=None):
        if notebook.index(notebook.select()) == 0: 
            push_to_tab2()
            
    def global_open_map(event=None):
        if notebook.index(notebook.select()) == 0:
            call_tab7_mapping()

    root.bind('<Control-s>', global_save)
    root.bind('<Control-S>', global_save)
    root.bind('<Control-n>', global_new)
    root.bind('<Control-N>', global_new)
    
    root.bind('<Control-Next>', global_open_map)
    
    tree.bind('<Delete>', lambda e: xoa_dong_tab2())

    root.after(100, lambda: entries["Batch"].focus())
    root.mainloop()

if __name__ == "__main__":
    main()