import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import json
import os
import subprocess 
import threading
import re

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

# Kểm tra thư viện lịch
try:
    from tkcalendar import DateEntry
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False

# --- CẤU HÌNH ĐƯỜNG DẪN TỐI ƯU ---
IMAGE_DIR = r"D:\OneDrive-NamTL-CAF\OneDrive - pnj.com.vn\PNJ Onedrive -NamTL\product\hinh-anh\HINH ANH SAN PHAM CAO"
TRACKING_FILE = r"D:\nam.tl01\nam.tl01\product\CAF\SCBH-PH\data\tracking-SCBH.xlsx"
DEFAULT_EXPORT_DIR = r"D:\nam.tl01\nam.tl01\product\CAF\SCBH-PH\data"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(CURRENT_DIR, "form_data.json")
INDEX_FILE = os.path.join(CURRENT_DIR, "image_index.json") 

# --- DỮ LIỆU MẶC ĐỊNH ---
default_data = {
    "categories": {
        "Loại hàng": ["Hàng mua lại TKD", "Hàng trưng bày", "Hàng bảo hành", "Hàng phản hồi", "Hàng tách đá", "Hủy nấu"],
        "Loại lỗi": ["Đá trầy (cạnh, mặt)", "Đá nhăm (cạnh, mặt)", "Đá cấn (cạnh, mặt)", "Đá rớt", "Thiếu giấy kiểm định"],
        "Người kiểm": ["E0225574-TRẦN LÊ NAM"]
    },
    "tab2_rows": [],
    "last_nguoi_kiem": "E0225574-TRẦN LÊ NAM",
    "image_search_dir": "" 
}

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "last_nguoi_kiem" not in data: data["last_nguoi_kiem"] = default_data["last_nguoi_kiem"]
                if "tab2_rows" not in data: data["tab2_rows"] = []
                if "image_search_dir" not in data: data["image_search_dir"] = ""
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

def main():
    root = tk.Tk()
    root.title("FORM KIỂM TRA ĐÁ SCBH CAF - V47 (Liên kết Mapping App V28)")
    
    try:
        root.state('zoomed')
    except:
        root.attributes('-zoomed', True) 

    bottom_frame = tk.Frame(root, bg="#cce5ff", bd=1, relief="ridge")
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

    notebook = ttk.Notebook(root)
    notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

    tab1_frame = tk.Frame(notebook, bg="#f0f0f0")
    tab2_frame = tk.Frame(notebook, bg="#f0f0f0")
    tab3_frame = tk.Frame(notebook, bg="#f0f0f0")
    tab4_frame = tk.Frame(notebook, bg="#f0f0f0") 
    tab5_frame = tk.Frame(notebook, bg="#f0f0f0") 
    tab6_frame = tk.Frame(notebook, bg="#f0f0f0") 

    notebook.add(tab1_frame, text=" 📝 TAB 1: NHẬP LIỆU ")
    notebook.add(tab2_frame, text=" 🗂️ TAB 2: DỮ LIỆU TẠM ")
    notebook.add(tab3_frame, text=" ⚙️ TAB 3: DANH MỤC ")
    notebook.add(tab4_frame, text=" 🔍 TAB 4: TÌM LỊCH SỬ ")
    notebook.add(tab5_frame, text=" 📊 TAB 5: THỐNG KÊ ")
    notebook.add(tab6_frame, text=" 🖼️ TAB 6: TÌM ẢNH LỖI ")

    root.bind('<Alt-Key-1>', lambda e: notebook.select(0))
    root.bind('<Alt-Key-2>', lambda e: notebook.select(1))
    root.bind('<Alt-Key-3>', lambda e: notebook.select(2))
    root.bind('<Alt-Key-4>', lambda e: notebook.select(3))
    root.bind('<Alt-Key-5>', lambda e: notebook.select(4)) 
    root.bind('<Alt-Key-6>', lambda e: notebook.select(5)) 

    # ==========================================
    # TAB 1: THIẾT KẾ GIAO DIỆN NHẬP LIỆU
    # ==========================================
    col_container = tk.Frame(tab1_frame, bg="#f0f0f0")
    col_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    col_container.columnconfigure(0, weight=2)
    col_container.columnconfigure(1, weight=1)
    col_container.rowconfigure(0, weight=1)

    left_big_frame = tk.LabelFrame(col_container, text="Thông tin Sản phẩm", font=("Arial", 10, "bold"), bg="#f0f0f0")
    left_big_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    
    left_big_frame.columnconfigure(0, weight=1)
    left_big_frame.columnconfigure(1, weight=1)
    left_big_frame.rowconfigure(0, weight=1)

    input_frame = tk.Frame(left_big_frame, bg="#f0f0f0")
    input_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

    img_container = tk.Frame(left_big_frame, bg="#f0f0f0")
    img_container.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

    fields_left = [
        "Batch", "Material/Mã SP", "Plant", "Loại hàng", 
        "Loại lỗi", "Tình trạng đá", "Ghi chú", "Ghi chú tham chiếu", "Hóa đơn", "Số lượng", "Ngày",
        "Số TN", "Số GIA", "Người kiểm"
    ]
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

    # --- CẬP NHẬT TÊN FILE MAPPING_APP-V28.PY ---
    def call_mapping_app():
        try:
            mapping_script = os.path.join(CURRENT_DIR, "mapping_app-v31.py")
            if os.path.exists(mapping_script):
                subprocess.Popen(["python", mapping_script])
            else:
                messagebox.showwarning("Lỗi", f"Không tìm thấy file:\n{mapping_script}\n\nVui lòng đặt file này cùng thư mục với phần mềm hiện tại.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở ứng dụng: {e}")

    btn_mapping = tk.Button(img_container, text="🖌️ MỞ MAPPING APP (BẢN VẼ)", command=call_mapping_app, font=("Arial", 10, "bold"), bg="#cce5ff", height=2)
    btn_mapping.pack(fill=tk.X, padx=5, pady=(5, 10))

    right_frame = tk.LabelFrame(col_container, text="Thông tin lịch sử SP", font=("Arial", 10, "bold"), bg="#f0f0f0")
    right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
    
    txt_history = tk.Text(right_frame, font=("Arial", 10), relief="sunken", bd=1, wrap=tk.WORD, state=tk.DISABLED, bg="#f9f9f9")
    txt_history.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # ==========================================
    # CÁC HÀM XỬ LÝ TAB 1
    # ==========================================
    def load_product_image(event=None):
        ma_sp = entries["Material/Mã SP"].get().strip()
        if not ma_sp:
            lbl_img.config(image='', text="Chưa nhập Mã SP")
        else:
            img_path = os.path.join(IMAGE_DIR, f"{ma_sp}.jpg")
            if os.path.exists(img_path):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((350, 350), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    lbl_img.config(image=photo, text="")
                    lbl_img.image = photo 
                except Exception as e:
                    lbl_img.config(image='', text=f"Lỗi mở ảnh: {e}")
            else:
                lbl_img.config(image='', text=f"Không tìm thấy ảnh:\n{ma_sp}.jpg")

        if event and event.keysym == 'Return':
            event.widget.tk_focusNext().focus()
            return "break"

    def load_history(event=None):
        raw_batch = entries["Batch"].get().strip()
        if not raw_batch: return
        real_batch = raw_batch.split("/")[-1] if "/" in raw_batch else raw_batch
        
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
        except ImportError:
            txt_history.insert(tk.END, "❌ Thiếu thư viện pandas, openpyxl.")
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
    
    entries["Batch"].bind("<Return>", on_batch_return)
    entries["Batch"].bind("<FocusOut>", lambda e: load_history())

    def focus_next_widget(event):
        event.widget.tk_focusNext().focus()
        return "break"

    for key, widget in entries.items():
        if key not in ["Material/Mã SP", "Batch"] and isinstance(widget, (tk.Entry, ttk.Combobox)):
            widget.bind("<Return>", focus_next_widget)

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
    # TAB 2: THIẾT KẾ BẢNG DỮ LIỆU TẠM
    # ==========================================
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
            for child in tree.get_children():
                col_data.append(str(tree.item(child, 'values')[col_index]))
            root.clipboard_clear()
            root.clipboard_append("\n".join(col_data))
            
        popup_menu.add_command(label=f"📐 Copy toàn bộ cột '{col_name}'", command=copy_column)
        
        def copy_row():
            row_str = "\t".join(str(v) for v in row_values)
            root.clipboard_clear()
            root.clipboard_append(row_str)
            
        popup_menu.add_command(label=f"📃 Copy cả dòng này (Tất cả cột)", command=copy_row)
        
        def copy_partial_row():
            idx_ghichu = columns.index("Ghi chú") + 1 
            row_str = "\t".join(str(v) for v in row_values[:idx_ghichu])
            root.clipboard_clear()
            root.clipboard_append(row_str)
            
        popup_menu.add_command(label=f"✂️ Copy dòng (Từ Batch đến Ghi chú)", command=copy_partial_row)
        popup_menu.post(event.x_root, event.y_root)

    tree.bind("<Button-3>", show_popup)

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
            if not raw_val: return
            lines = raw_val.split('\n')
            added_any = False
            for line in lines:
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
    search_so_tn = tk.Entry(search_input_frame, width=25, font=("Arial", 10))
    search_so_tn.grid(row=1, column=1, padx=5, pady=10)

    tk.Label(search_input_frame, text="Số GIA:", bg="#f0f0f0", font=("Arial", 10)).grid(row=1, column=2, padx=10, pady=10, sticky="e")
    search_so_gia = tk.Entry(search_input_frame, width=25, font=("Arial", 10))
    search_so_gia.grid(row=1, column=3, padx=5, pady=10)

    btn_frame_tab4 = tk.Frame(search_input_frame, bg="#f0f0f0")
    btn_frame_tab4.grid(row=0, column=4, rowspan=2, padx=30, pady=10)

    btn_search_tab4 = tk.Button(btn_frame_tab4, text="🔍 TÌM KIẾM", font=("Arial", 10, "bold"), bg="#cce5ff", height=2, width=15)
    btn_search_tab4.pack(side=tk.TOP, pady=2)
    
    btn_clear_tab4 = tk.Button(btn_frame_tab4, text="🔄 LÀM MỚI", font=("Arial", 9), bg="#e2e3e5", width=17)
    btn_clear_tab4.pack(side=tk.TOP, pady=2)

    result_frame_tab4 = tk.LabelFrame(tab4_frame, text="Kết quả tìm kiếm", font=("Arial", 10, "bold"), bg="#f0f0f0")
    result_frame_tab4.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    txt_search_result = tk.Text(result_frame_tab4, font=("Arial", 11), relief="sunken", bd=1, wrap=tk.WORD, state=tk.DISABLED, bg="#f9f9f9")
    txt_search_result.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def clear_search_inputs():
        search_ma_sp.delete(0, tk.END)
        search_plant.delete(0, tk.END)
        search_so_tn.delete(0, tk.END)
        search_so_gia.delete(0, tk.END)
        txt_search_result.config(state=tk.NORMAL)
        txt_search_result.delete("1.0", tk.END)
        txt_search_result.config(state=tk.DISABLED)

    btn_clear_tab4.config(command=clear_search_inputs)

    def perform_search(event=None):
        val_ma_sp = search_ma_sp.get().strip()
        val_plant = search_plant.get().strip()
        val_so_tn = search_so_tn.get().strip()
        val_so_gia = search_so_gia.get().strip()

        if not any([val_ma_sp, val_plant, val_so_tn, val_so_gia]):
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập ít nhất 1 tiêu chí tìm kiếm!")
            return

        txt_search_result.config(state=tk.NORMAL)
        txt_search_result.delete("1.0", tk.END)
        root.update()

        if not os.path.exists(TRACKING_FILE):
            txt_search_result.insert(tk.END, f"❌ Không tìm thấy file Excel tracking tại:\n{TRACKING_FILE}")
            txt_search_result.config(state=tk.DISABLED)
            return

        try:
            import pandas as pd
            df = pd.read_excel(TRACKING_FILE)

            def search_any_col(dataframe, keyword):
                return dataframe.astype(str).apply(lambda x: x.str.contains(keyword, case=False, na=False)).any(axis=1)

            mask = pd.Series(True, index=df.index)
            if val_ma_sp: mask = mask & search_any_col(df, val_ma_sp)
            if val_plant: mask = mask & search_any_col(df, val_plant)
            if val_so_tn: mask = mask & search_any_col(df, val_so_tn)
            if val_so_gia: mask = mask & search_any_col(df, val_so_gia)

            results = df[mask]
            txt_search_result.delete("1.0", tk.END)
            if results.empty:
                txt_search_result.insert(tk.END, "Không có dữ liệu lịch sử nào khớp với các tiêu chí tìm kiếm trên.")
            else:
                txt_search_result.insert(tk.END, f"Đã tìm thấy {len(results)} kết quả phù hợp:\n")
                for idx, row in results.iterrows():
                    txt_search_result.insert(tk.END, "-" * 60 + "\n")
                    for col in df.columns:
                        val = row[col]
                        if pd.notna(val) and str(val).strip():
                            txt_search_result.insert(tk.END, f"• {col}: {val}\n")
                txt_search_result.insert(tk.END, "-" * 60 + "\n")

        except ImportError:
            txt_search_result.insert(tk.END, "❌ Thiếu thư viện pandas, openpyxl.")
        except Exception as e:
            txt_search_result.insert(tk.END, f"❌ Lỗi đọc file Excel:\n{e}")
        finally:
            txt_search_result.config(state=tk.DISABLED)

    btn_search_tab4.config(command=perform_search)
    for entry in [search_ma_sp, search_plant, search_so_tn, search_so_gia]:
        entry.bind("<Return>", perform_search)

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

    def clear_date_filter():
        entry_tu_ngay.delete(0, tk.END)
        entry_den_ngay.delete(0, tk.END)
        
    btn_clear_date = tk.Button(filter_frame, text="✖ XÓA LỌC", command=clear_date_filter, font=("Arial", 9, "bold"), bg="#ffcccc")
    btn_clear_date.grid(row=0, column=4, padx=5, pady=10)

    btn_load_dashboard = tk.Button(filter_frame, text="⚙️ XỬ LÝ DỮ LIỆU", font=("Arial", 10, "bold"), bg="#e2e3e5", height=1)
    btn_load_dashboard.grid(row=0, column=5, padx=20, pady=10)

    kpi_container = tk.Frame(tab5_frame, bg="#f0f0f0")
    kpi_container.pack(fill=tk.X, padx=10, pady=10)

    chart_btn_frame = tk.Frame(tab5_frame, bg="#f0f0f0")
    chart_btn_frame.pack(fill=tk.X, padx=10, pady=15)

    btn_hm = tk.Button(chart_btn_frame, text="🔥 MỞ BẢN ĐỒ NHIỆT", font=("Arial", 10, "bold"), bg="#ffcccc", height=2, width=25, state=tk.DISABLED)
    btn_line = tk.Button(chart_btn_frame, text="📈 MỞ ĐƯỜNG XU HƯỚNG", font=("Arial", 10, "bold"), bg="#cce5ff", height=2, width=28, state=tk.DISABLED)
    btn_pareto = tk.Button(chart_btn_frame, text="📊 MỞ PHÂN TÍCH PARETO", font=("Arial", 10, "bold"), bg="#d4edda", height=2, width=25, state=tk.DISABLED)

    btn_hm.pack(side=tk.LEFT, padx=10)
    btn_line.pack(side=tk.LEFT, padx=10)
    btn_pareto.pack(side=tk.LEFT, padx=10)

    lbl_instruction = tk.Label(tab5_frame, text="Vui lòng bấm nút '⚙️ XỬ LÝ DỮ LIỆU' trước.\nSau đó chọn các nút bên trên để mở từng biểu đồ phân tích chuyên sâu.", font=("Arial", 12, "italic"), fg="#666666", bg="#f0f0f0")
    lbl_instruction.pack(pady=40)

    filtered_data = {'df': None, 'cols': {}}

    def process_data_and_kpi():
        for widget in kpi_container.winfo_children(): widget.destroy()
        btn_hm.config(state=tk.DISABLED); btn_line.config(state=tk.DISABLED); btn_pareto.config(state=tk.DISABLED)

        try: import pandas as pd
        except ImportError: return messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt pandas")

        if not os.path.exists(TRACKING_FILE): return messagebox.showerror("Lỗi", "Chưa có file Excel tracking.")

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

            filtered_data['df'] = df; filtered_data['cols'] = cols

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

            btn_hm.config(state=tk.NORMAL); btn_line.config(state=tk.NORMAL); btn_pareto.config(state=tk.NORMAL)
            lbl_instruction.config(text="Đã xử lý xong dữ liệu! Vui lòng chọn mở biểu đồ bên trên.", fg="green")

        except Exception as e: messagebox.showerror("Lỗi", f"Lỗi xử lý dữ liệu:\n{e}")

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
            
            def update_scrollregion(event): scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
            inner_frame.bind("<Configure>", update_scrollregion)
            
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
                        for label in ax.get_xticklabels(): label.set_rotation(30); label.set_ha('right')

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
                        for label in ax.get_xticklabels(): label.set_rotation(35); label.set_ha('right')

                canvas_agg = FigureCanvasTkAgg(fig, master=inner_frame)
                canvas_agg.draw()
                canvas_agg.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                inner_frame.update_idletasks()
                scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

            def on_zoom_in():
                if chart_state['scale'] < 3.0: chart_state['scale'] += 0.3; draw_chart_on_canvas()
            def on_zoom_out():
                if chart_state['scale'] > 0.6: chart_state['scale'] -= 0.3; draw_chart_on_canvas()
            def on_save():
                if chart_state['fig']:
                    filepath = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png")])
                    if filepath: chart_state['fig'].savefig(filepath, bbox_inches='tight'); messagebox.showinfo("Thành công", "Đã lưu ảnh biểu đồ!")

            btn_zoom_in.config(command=on_zoom_in); btn_zoom_out.config(command=on_zoom_out); btn_save.config(command=on_save)
            draw_chart_on_canvas()

        except Exception as e: messagebox.showerror("Lỗi Vẽ Biểu Đồ", str(e)); top.destroy()

    btn_hm.config(command=lambda: open_chart_window('heatmap'))
    btn_line.config(command=lambda: open_chart_window('line'))
    btn_pareto.config(command=lambda: open_chart_window('pareto'))

    # ==========================================
    # TAB 6: PIM SYSTEM - TÌM KIẾM HÌNH ẢNH LỖI (V45 TÍCH HỢP LỌC NGÀY)
    # ==========================================
    t6_top_frame = tk.LabelFrame(tab6_frame, text="Cấu hình & Đồng bộ Danh bạ Kho Ảnh", font=("Arial", 10, "bold"), bg="#f0f0f0")
    t6_top_frame.pack(fill=tk.X, padx=10, pady=5)

    tk.Label(t6_top_frame, text="Thư mục Kho ảnh:", bg="#f0f0f0", font=("Arial", 9)).grid(row=0, column=0, padx=10, pady=10)
    
    ent_search_dir = tk.Entry(t6_top_frame, width=50, font=("Arial", 10))
    ent_search_dir.insert(0, app_data.get("image_search_dir", ""))
    ent_search_dir.grid(row=0, column=1, padx=5, pady=10)

    def choose_search_dir():
        d = filedialog.askdirectory(title="Chọn Thư Mục Mẹ Chứa Ảnh")
        if d:
            ent_search_dir.delete(0, tk.END)
            ent_search_dir.insert(0, d)
            app_data["image_search_dir"] = d
            save_data(app_data)
            
    btn_browse_dir = tk.Button(t6_top_frame, text="📁 CHỌN THƯ MỤC", command=choose_search_dir, font=("Arial", 9))
    btn_browse_dir.grid(row=0, column=2, padx=5, pady=10)

    lbl_sync_status = tk.Label(t6_top_frame, text=f"Tổng số ảnh trong danh bạ hiện tại: {len(image_index_list)}", bg="#f0f0f0", font=("Arial", 9, "italic"), fg="blue")
    lbl_sync_status.grid(row=0, column=4, padx=20, pady=10)

    def run_sync_images():
        folder_path = ent_search_dir.get().strip()
        if not os.path.exists(folder_path):
            messagebox.showerror("Lỗi", "Đường dẫn không tồn tại. Vui lòng chọn lại Thư mục Kho ảnh.")
            return
            
        btn_sync.config(state=tk.DISABLED, text="⏳ ĐANG QUÉT...")
        lbl_sync_status.config(text="Đang quét thư mục... Vui lòng đợi!", fg="red")
        
        def task():
            global image_index_list
            temp_list = []
            for root_dir, dirs, files in os.walk(folder_path):
                for f in files:
                    if f.lower().endswith(('.jpg', '.jpeg', '.png', '.jfif')):
                        temp_list.append(os.path.join(root_dir, f))
                        
            image_index_list = temp_list
            with open(INDEX_FILE, "w", encoding="utf-8") as f:
                json.dump(image_index_list, f, ensure_ascii=False)
                
            def update_ui():
                btn_sync.config(state=tk.NORMAL, text="🔄 ĐỒNG BỘ LẠI KHO ẢNH")
                lbl_sync_status.config(text=f"✅ Đồng bộ thành công! Tìm thấy {len(image_index_list)} ảnh.", fg="green")
            root.after(0, update_ui)
            
        threading.Thread(target=task, daemon=True).start()

    btn_sync = tk.Button(t6_top_frame, text="🔄 ĐỒNG BỘ LẠI KHO ẢNH", command=run_sync_images, font=("Arial", 9, "bold"), bg="#d4edda")
    btn_sync.grid(row=0, column=3, padx=10, pady=10)

    # --- KHUNG TÌM KIẾM ĐA NHIỆM (CÓ LỌC NGÀY VÀ Lệnh .pack() được cập nhật đầy đủ) ---
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
        
    def t6_clear_date():
        t6_ent_tu_ngay.delete(0, tk.END)
        t6_ent_den_ngay.delete(0, tk.END)
        
    tk.Button(t6_search_frame, text="✖ XÓA", command=t6_clear_date, font=("Arial", 8, "bold"), bg="#ffcccc").pack(side=tk.LEFT, padx=5)

    tk.Label(t6_search_frame, text="Từ khóa:", bg="#f0f0f0", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(15, 5))
    ent_img_search = tk.Entry(t6_search_frame, width=30, font=("Arial", 12))
    ent_img_search.pack(side=tk.LEFT, padx=5)
    
    instruction_text = "💡 HƯỚNG DẪN THAO TÁC: Đưa chuột vào ảnh tìm được\n 👉 Click ĐÚP (Chuột trái) = Phóng to soi lỗi HD | 🖱️ Click CHUỘT PHẢI = Mở thẳng Folder gốc chứa ảnh"
    lbl_img_instruction = tk.Label(tab6_frame, text=instruction_text, bg="#fff3cd", fg="#856404", font=("Arial", 10, "bold"), justify=tk.CENTER, bd=1, relief="solid")
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
    
    def g_update_scroll(event): gallery_canvas.configure(scrollregion=gallery_canvas.bbox("all"))
    gallery_inner.bind("<Configure>", g_update_scroll)

    thumbnail_refs = [] 

    def open_image_location(filepath):
        try:
            norm_path = os.path.normpath(filepath)
            subprocess.run(['explorer', '/select,', norm_path])
        except Exception as e: messagebox.showerror("Lỗi", f"Không thể mở thư mục:\n{e}")

    def open_hd_image(filepath):
        top = tk.Toplevel(root)
        top.title("🔍 Chế độ Soi Ảnh HD")
        try: top.state('zoomed')
        except: top.attributes('-zoomed', True)
        
        tb_frame = tk.Frame(top, bg="#e2e3e5", bd=1, relief="raised")
        tb_frame.pack(side=tk.TOP, fill=tk.X)
        btn_open_loc = tk.Button(tb_frame, text="📂 MỞ THƯ MỤC GỐC", command=lambda: open_image_location(filepath), font=("Arial", 10, "bold"), bg="#fff3cd")
        btn_open_loc.pack(side=tk.RIGHT, padx=10, pady=5)
        lbl_info = tk.Label(tb_frame, text=f"Đang xem: {os.path.basename(filepath)}", bg="#e2e3e5", font=("Arial", 10, "italic"))
        lbl_info.pack(side=tk.LEFT, padx=10, pady=5)

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
                new_size = (int(pil_img.width * img_state['scale']), int(pil_img.height * img_state['scale']))
                resized = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                img_state['tk_img'] = ImageTk.PhotoImage(resized)
                if img_state['id']: img_canvas.delete(img_state['id'])
                img_state['id'] = img_canvas.create_image(0, 0, anchor="nw", image=img_state['tk_img'])
                img_canvas.configure(scrollregion=img_canvas.bbox("all"))
            render_img()
            def mouse_wheel(event):
                if event.delta > 0: img_state['scale'] *= 1.1
                else: img_state['scale'] /= 1.1
                render_img()
            img_canvas.bind("<MouseWheel>", mouse_wheel)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file ảnh:\n{e}")
            top.destroy()

    def perform_image_search(event=None):
        tu_ngay_str = t6_ent_tu_ngay.get().strip()
        den_ngay_str = t6_ent_den_ngay.get().strip()
        
        start_date_int = 0
        end_date_int = 99999999
        
        if tu_ngay_str:
            try: start_date_int = int(datetime.strptime(tu_ngay_str, "%d/%m/%Y").strftime("%Y%m%d"))
            except: pass
        if den_ngay_str:
            try: end_date_int = int(datetime.strptime(den_ngay_str, "%d/%m/%Y").strftime("%Y%m%d"))
            except: pass

        raw_kw = ent_img_search.get().strip().lower()
        keywords = raw_kw.split() if raw_kw else []
        
        for widget in gallery_inner.winfo_children(): widget.destroy()
        thumbnail_refs.clear()
        
        results = []
        for p in image_index_list:
            fname = os.path.basename(p)
            
            # --- KIỂM TRA ĐIỀU KIỆN LỌC NGÀY (Từ chuỗi yyyymmdd ở đầu tên file) ---
            if start_date_int > 0 or end_date_int < 99999999:
                match = re.search(r'^(\d{8})', fname)
                if match:
                    file_date_int = int(match.group(1))
                    if file_date_int < start_date_int or file_date_int > end_date_int:
                        continue 
                else:
                    continue 

            # --- KIỂM TRA TỪ KHÓA ---
            if keywords and not all(kw in p.lower() for kw in keywords):
                continue
                
            results.append(p)
        
        results.sort(reverse=True) 
        
        MAX_RES = 50
        display_res = results[:MAX_RES]
        
        if not display_res:
            tk.Label(gallery_inner, text="Không tìm thấy ảnh nào khớp với các tiêu chí lọc.", font=("Arial", 12), bg="white").pack(pady=20)
            return
            
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
                short_name = fname if len(fname) < 25 else fname[:22] + "..."
                tk.Label(frm, text=short_name, bg="white", font=("Arial", 9)).pack(pady=(5,0))
                
                lbl_img.bind("<Double-Button-1>", lambda e, p=path: open_hd_image(p))
                def on_right_click(event, p=path):
                    ctx_menu = tk.Menu(root, tearoff=0, font=("Arial", 10))
                    ctx_menu.add_command(label="🔍 Xem ảnh HD (Phóng to)", command=lambda: open_hd_image(p))
                    ctx_menu.add_separator()
                    ctx_menu.add_command(label="📂 Mở thư mục chứa ảnh này", command=lambda: open_image_location(p))
                    ctx_menu.post(event.x_root, event.y_root)
                lbl_img.bind("<Button-3>", on_right_click)

                col_idx += 1
                if col_idx >= COLUMNS: col_idx = 0; row_idx += 1
            except: pass 

    btn_img_search = tk.Button(t6_search_frame, text="🔍 TÌM ẢNH", command=perform_image_search, font=("Arial", 10, "bold"), bg="#cce5ff", width=15)
    btn_img_search.pack(side=tk.LEFT, padx=10)
    ent_img_search.bind("<Return>", lambda e: perform_image_search())

    # ==========================================
    # LOGIC NÚT BẤM VÀ XUẤT EXCEL THỰC TẾ
    # ==========================================
    def push_to_tab2():
        ma_sp = entries["Material/Mã SP"].get().strip()
        if not ma_sp: 
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập Mã SP!")
            return
            
        raw_batch = entries["Batch"].get().strip()
        real_batch = raw_batch.split("/")[-1] if "/" in raw_batch else raw_batch
        if not real_batch: real_batch = "none"

        ghi_chu_text = entries["Ghi chú"].get("1.0", tk.END).strip()
        if not ghi_chu_text: ghi_chu_text = "none"

        hoa_don = entries["Hóa đơn"].get().strip()
        if not hoa_don: hoa_don = "0"

        so_luong = entries["Số lượng"].get().strip()
        if not so_luong: so_luong = "1"

        plant = entries["Plant"].get().strip() or "none"
        loai_hang = entries["Loại hàng"].get().strip() or "none"
        loai_loi = entries["Loại lỗi"].get().strip() or "none"
        ngay = entries["Ngày"].get().strip() or "none"
        tinh_trang = entries["Tình trạng đá"].get().strip() or "none"
        so_tn = entries["Số TN"].get().strip() or "none"
        so_gia = entries["Số GIA"].get().strip() or "none"
        nguoi_kiem = entries["Người kiểm"].get().strip() or "none"

        data_row = (
            real_batch, ma_sp, plant, so_luong,
            loai_hang, loai_loi, ngay, tinh_trang,
            so_tn, so_gia, ghi_chu_text, nguoi_kiem, hoa_don
        )
        
        tree.insert("", tk.END, values=data_row)
        app_data["tab2_rows"].append(data_row)
        app_data["last_nguoi_kiem"] = entries["Người kiểm"].get()
        save_data(app_data)
        
        fields_to_clear = ["Batch", "Material/Mã SP", "Plant", "Loại hàng", "Loại lỗi", "Số TN", "Số GIA", "Hóa đơn"]
        for field in fields_to_clear:
            entries[field].delete(0, tk.END)
            if isinstance(entries[field], ttk.Combobox):
                entries[field].set('')
                
        entries["Ghi chú"].delete("1.0", tk.END)
        entries["Số lượng"].delete(0, tk.END)
        entries["Số lượng"].insert(0, "1")
                
        lbl_img.config(image='', text="Chưa tải ảnh\n(Nhập Mã SP và nhấn Enter)")
        
        txt_history.config(state=tk.NORMAL)
        txt_history.delete("1.0", tk.END)
        txt_history.config(state=tk.DISABLED)
        
        entries["Batch"].focus() 

    def xoa_dong_tab2():
        selected = tree.selection()
        if not selected: return
        indices = [tree.index(sel) for sel in selected]
        indices.sort(reverse=True)
        for sel in selected: tree.delete(sel)
        for idx in indices:
            if idx < len(app_data["tab2_rows"]): app_data["tab2_rows"].pop(idx)
        save_data(app_data)

    def sua_dong_tab2():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 dòng để chỉnh sửa!")
            return
        if len(selected) > 1:
            messagebox.showwarning("Cảnh báo", "Chỉ được chọn 1 dòng để chỉnh sửa mỗi lần!")
            return
            
        item = selected[0]
        idx = tree.index(item)
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
            if idx < len(app_data["tab2_rows"]):
                app_data["tab2_rows"][idx] = new_vals
                save_data(app_data)
                
            edit_win.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật dữ liệu dòng được chọn!")

        tk.Button(edit_win, text="💾 LƯU THAY ĐỔI", command=save_edit, bg="#d4edda", font=("Arial", 10, "bold"), width=15).pack(pady=20)

    def copy_data_tab2(event=None):
        selected_items = tree.selection()
        
        if selected_items:
            items_to_copy = selected_items
            msg = f"Đã copy {len(selected_items)} dòng được chọn!"
        else:
            items_to_copy = tree.get_children()
            msg = "Đã copy TOÀN BỘ bảng!"
            
        if not items_to_copy:
            if not event: messagebox.showwarning("Trống", "Không có dữ liệu để copy!")
            return

        headers = "\t".join(columns)
        lines = [headers]

        for item in items_to_copy:
            vals = tree.item(item, "values")
            lines.append("\t".join(str(v) for v in vals))

        clipboard_text = "\n".join(lines)
        root.clipboard_clear()
        root.clipboard_append(clipboard_text)
        
        if not event:
            messagebox.showinfo("Thành công", f"{msg}\nBạn có thể nhấn Ctrl+V để dán trực tiếp vào file Excel.")

    tree.bind("<Control-c>", copy_data_tab2)
    tree.bind("<Control-C>", copy_data_tab2)

    def export_excel():
        if not app_data["tab2_rows"]:
            messagebox.showwarning("Trống", "Không có dữ liệu nào ở Tab 2 để xuất!")
            return
        target_file = filedialog.askopenfilename(
            title="Chọn file Excel để cập nhật dữ liệu",
            initialdir=DEFAULT_EXPORT_DIR,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not target_file: return 
            
        try:
            import openpyxl
            wb = openpyxl.load_workbook(target_file)
            if "source" not in wb.sheetnames:
                messagebox.showerror("Lỗi", f"Không tìm thấy sheet mang tên 'source' trong file:\n{os.path.basename(target_file)}")
                wb.close()
                return
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
                
                if not loai_loi.strip() or loai_loi.lower() == "none": loai_loi = "none"
                if "-" in nguoi_kiem and nguoi_kiem.lower() != "none": nguoi_kiem = nguoi_kiem.split("-")[0].strip()
                
                tinh_trang_upper = tinh_trang.strip().upper()
                tt_short = "Đ" if tinh_trang_upper == "ĐẠT" else ("KĐ" if tinh_trang_upper == "KHÔNG ĐẠT" else tinh_trang)
                sl_val = int(so_luong) if str(so_luong).isdigit() else so_luong

                hd_val = 0
                if str(hoa_don).isdigit():
                    hd_val = int(hoa_don)
                else:
                    try: hd_val = float(hoa_don)
                    except ValueError: hd_val = 0

                ngay_obj = ngay
                is_valid_date = False
                if ngay.lower() != "none":
                    try:
                        ngay_obj = datetime.strptime(ngay.strip(), "%d/%m/%Y").date()
                        is_valid_date = True
                    except ValueError: pass 
                
                row_excel = [last_stt, real_batch, ma_sp, plant, sl_val, loai_hang, loai_loi, ngay_obj, tt_short, ghi_chu, so_tn, so_gia, nguoi_kiem, hd_val]
                sheet.append(row_excel)
                
                if is_valid_date:
                    date_cell = sheet.cell(row=sheet.max_row, column=8)
                    date_cell.number_format = 'DD/MM/YYYY'
                
            wb.save(target_file)
            wb.close()
            for item in tree.get_children(): tree.delete(item)
            app_data["tab2_rows"].clear()
            save_data(app_data)
            messagebox.showinfo("Thành công", f"Đã ghi nối thành công toàn bộ dữ liệu vào file:\n{os.path.basename(target_file)}")
            
        except PermissionError:
            messagebox.showerror("Lỗi Truy Cập", "Không thể ghi dữ liệu vì file Excel đang được mở.\nVui lòng đóng file Excel trước khi xuất!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi ghi file:\n{e}\n\nVui lòng cài đặt openpyxl nếu chưa có (pip install openpyxl).")

    # KHỞI TẠO CÁC NÚT BẤM (ĐƯỢC GẮN VÀO BOTTOM_FRAME)
    btn_nhap = tk.Button(bottom_frame, text="📥 NHẬP (Đẩy sang Tab 2)", command=push_to_tab2, font=("Arial", 10, "bold"), bg="#d4edda", width=25, height=2)
    btn_sua = tk.Button(bottom_frame, text="✏️ CHỈNH SỬA DÒNG", command=sua_dong_tab2, font=("Arial", 10, "bold"), bg="#fff3cd", height=2)
    btn_copy = tk.Button(bottom_frame, text="📋 COPY DỮ LIỆU", command=copy_data_tab2, font=("Arial", 10, "bold"), bg="#e2e3e5", height=2)
    btn_xoa = tk.Button(bottom_frame, text="🗑️ XÓA DÒNG ĐÃ CHỌN", command=xoa_dong_tab2, font=("Arial", 10, "bold"), bg="#f8d7da", height=2)
    btn_xuat = tk.Button(bottom_frame, text="📤 XUẤT FILE EXCEL", command=export_excel, font=("Arial", 10, "bold"), bg="#fff3cd", width=20, height=2)

    def on_tab_change(event):
        selected_tab = notebook.index(notebook.select())
        
        btn_nhap.pack_forget()
        btn_sua.pack_forget()
        btn_xoa.pack_forget()
        btn_copy.pack_forget()
        btn_xuat.pack_forget()

        if selected_tab == 0:   
            bottom_frame.config(bg="#cce5ff", bd=1, relief="ridge")
            btn_nhap.pack(side=tk.LEFT, padx=20, pady=10)
            entries["Batch"].focus() 
        elif selected_tab == 1: 
            bottom_frame.config(bg="#cce5ff", bd=1, relief="ridge")
            btn_sua.pack(side=tk.LEFT, padx=(10, 5), pady=10)
            btn_xoa.pack(side=tk.LEFT, padx=(5, 10), pady=10)
            btn_copy.pack(side=tk.LEFT, padx=10, pady=10)
            btn_xuat.pack(side=tk.RIGHT, padx=20, pady=10)
        else:
            bottom_frame.config(bg=root.cget('bg'), bd=0, relief="flat")

    notebook.bind("<<NotebookTabChanged>>", on_tab_change)
    on_tab_change(None)

    root.after(100, lambda: entries["Batch"].focus())
    root.mainloop()

if __name__ == "__main__":
    main()