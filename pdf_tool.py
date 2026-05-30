import sys
import os
import threading
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pythoncom
import win32com.client
from PIL import Image, ImageTk
from fpdf import FPDF
from pypdf import PdfWriter
import tempfile
import time

# --- GET EMBEDDED PATH FOR ASSETS ---
def get_asset_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- RE-ENGINEERED COMPACT CUTE PROGRESS BAR ---
class MiniProgress:
    def __init__(self, message, total_steps=100):
        self.root = tk.Tk()
        self.root.title("Little's PDF")
        self.root.overrideredirect(True)
        self.root.geometry("420x95")
        self.root.eval('tk::PlaceWindow . center')
        
        self.bg_color = "#dbdbdb"
        self.root.configure(bg=self.bg_color, highlightthickness=1, highlightbackground="#b5b5b5")
        self.root.attributes("-topmost", True)
        
        self.root.bind("<Button-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.do_drag)
        
        icon_path = get_asset_path('icon.ico')
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except:
                pass
        
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=12)
        
        left_frame = tk.Frame(main_frame, bg=self.bg_color)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        
        if os.path.exists(icon_path):
            try:
                img = Image.open(icon_path)
                img_width, img_height = img.size
                target_height = 52  
                target_width = int((img_width / img_height) * target_height)
                
                img_resized = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
                self.icon_img = ImageTk.PhotoImage(img_resized)
                
                icon_lbl = tk.Label(left_frame, image=self.icon_img, bg=self.bg_color, borderwidth=0)
                icon_lbl.pack()
            except:
                pass
                
        brand_lbl = tk.Label(left_frame, text="Little's PDF", bg=self.bg_color, fg="#333333", 
                             font=("Segoe UI Semibold", 9, "bold"))
        brand_lbl.pack(pady=(2, 0))

        right_frame = tk.Frame(main_frame, bg=self.bg_color)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.lbl = tk.Label(right_frame, text=message, bg=self.bg_color, fg="#1a1a1a", 
                            font=("Segoe UI", 11), anchor="w", justify=tk.LEFT, wraplength=250)
        self.lbl.pack(fill=tk.X, pady=(2, 6))
        
        self.progress_width = 205
        self.progress_height = 6 
        self.total_steps = total_steps
        
        bar_row = tk.Frame(right_frame, bg=self.bg_color)
        bar_row.pack(fill=tk.X, expand=True, pady=(2, 0))
        
        self.canvas = tk.Canvas(bar_row, width=self.progress_width, height=self.progress_height, 
                               bg="#c0c0c0", bd=0, highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, pady=2)
        
        self.canvas.create_rectangle(0, 0, self.progress_width, self.progress_height, fill="#c0c0c0", width=0)
        self.fill_bar = self.canvas.create_rectangle(0, 0, 0, self.progress_height, fill="#ed1c24", width=0)
        
        self.percent_lbl = tk.Label(bar_row, text="0%", bg=self.bg_color, fg="#333333", 
                                    font=("Segoe UI", 11, "bold"), width=4, anchor="e")
        self.percent_lbl.pack(side=tk.RIGHT, padx=(5, 0))

    def start_drag(self, event):
        self.x = event.x
        self.y = event.y

    def do_drag(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def update_progress(self, current_value, text=None):
        pct = current_value / self.total_steps
        new_width = int(self.progress_width * pct)
        self.canvas.coords(self.fill_bar, 0, 0, new_width, self.progress_height)
        
        self.percent_lbl.config(text=f"{int(pct * 100)}%")
        if text:
            self.lbl.config(text=text)
        self.root.update()

    def show_success(self, text):
        self.root.after(0, lambda: self._apply_success_ui(text))

    def _apply_success_ui(self, text):
        self.lbl.config(text=text)
        self.percent_lbl.config(text="100%")
        self.canvas.coords(self.fill_bar, 0, 0, self.progress_width, self.progress_height)
        
        # 1.5 Second Cinematic Linger before safely self-destructing
        self.root.after(1500, self.close)

    def close(self):
        self.root.destroy()

# --- EXCEL WORKSHEET SELECTION INTERFACE ---
class SheetSelector:
    def __init__(self, sheet_names):
        self.result = None
        self.win = tk.Tk()
        self.win.title("Little's PDF")
        self.win.geometry("350x280")
        self.win.eval('tk::PlaceWindow . center')
        
        icon_path = get_asset_path('icon.ico')
        if os.path.exists(icon_path):
            self.win.iconbitmap(icon_path)
            
        tk.Label(self.win, text="Select the Worksheet to Convert:", font=("Segoe UI", 10, "bold")).pack(pady=10)
        
        frame = tk.Frame(self.win)
        frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        self.lb = tk.Listbox(frame, selectmode=tk.SINGLE, font=("Segoe UI", 9), relief="flat", highlightthickness=1, highlightbackground="#cccccc")
        self.lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.lb.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lb.config(yscrollcommand=sb.set)
        
        for name in sheet_names:
            self.lb.insert(tk.END, name)
        if sheet_names:
            self.lb.selection_set(0)
            
        tk.Button(self.win, text="Convert Selected Sheet", bg="#0078D7", fg="white", 
                  font=("Segoe UI", 9, "bold"), command=self.confirm, relief="flat", pady=5, cursor="hand2").pack(pady=15)
        self.win.mainloop()

    def confirm(self):
        sel = self.lb.curselection()
        if sel:
            self.result = self.lb.get(sel[0])
        self.win.destroy()

# --- ABSOLUTE SILENT CONVERSION ENGINES WITH ZOMBIE SWEEPER ---
def convert_word(input_path, output_path):
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0 
        word.ScreenUpdating = False 
        try:
            word.Options.PrintBackground = False 
            word.Application.Interactive = False 
        except: pass
            
        abs_in = os.path.abspath(input_path)
        abs_out = os.path.abspath(output_path)
        doc = word.Documents.Open(abs_in, ReadOnly=True, Visible=False)
        doc.SaveAs(abs_out, 17)
        doc.Close(False)
    finally:
        # ZOMBIE SWEEPER: Ensures ghost processes are destroyed even on hard crashes
        if word:
            try: word.Application.Interactive = True
            except: pass
            try: word.Quit()
            except: pass
            del word

def convert_excel(input_path, output_path, targeted_sheet=None):
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        try: excel.Interactive = False 
        except: pass
            
        abs_in = os.path.abspath(input_path)
        abs_out = os.path.abspath(output_path)
        wb = excel.Workbooks.Open(abs_in, ReadOnly=True)
        if targeted_sheet:
            ws = wb.Worksheets(targeted_sheet)
            ws.ExportAsFixedFormat(0, abs_out)
        else:
            wb.ExportAsFixedFormat(0, abs_out)
        wb.Close(False)
    finally:
        if excel:
            try: excel.Interactive = True
            except: pass
            try: excel.Quit()
            except: pass
            del excel

def convert_ppt(input_path, output_path):
    ppt = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.DisplayAlerts = 1 
        abs_in = os.path.abspath(input_path)
        abs_out = os.path.abspath(output_path)
        presentation = ppt.Presentations.Open(abs_in, WithWindow=False, ReadOnly=True)
        presentation.SaveAs(abs_out, 32)
        presentation.Close()
    finally:
        if ppt:
            try: ppt.Quit()
            except: pass
            del ppt

def convert_text(input_path, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=11)
    with open(input_path, 'r', encoding='utf-8', errors='replace') as f:
        text = f.read()
    pdf.multi_cell(0, 5, text)
    pdf.output(output_path)

def convert_image(input_path, output_path):
    image = Image.open(input_path)
    img_converted = image.convert('RGB')
    img_converted.save(output_path, "PDF", resolution=100.0, save_all=True)

def route_conversion(input_path, output_path, excel_sheet=None):
    ext = os.path.splitext(input_path)[1].lower()
    if ext in ['.doc', '.docx']: convert_word(input_path, output_path)
    elif ext in ['.xls', '.xlsx']: convert_excel(input_path, output_path, excel_sheet)
    elif ext in ['.ppt', '.pptx']: convert_ppt(input_path, output_path)
    elif ext == '.txt': convert_text(input_path, output_path)
    elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp']: convert_image(input_path, output_path)
    elif ext == '.pdf': pass
    else: raise ValueError(f"Unsupported format: {ext}")

def handle_path_collision(output_path):
    if not os.path.exists(output_path):
        return output_path
        
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    
    ans = messagebox.askyesnocancel(
        "File Collision Detected",
        f"The output file already exists:\n{os.path.basename(output_path)}\n\n"
        "Click 'YES' to OVERWRITE the existing file.\n"
        "Click 'NO' to AUTO-RENAME and save a copy safely.\n"
        "Click 'Cancel' to abort conversion entirely."
    )
    root.destroy()
    
    if ans is True:
        return output_path
    elif ans is False:
        base, ext = os.path.splitext(output_path)
        counter = 1
        new_path = f"{base}_{counter}{ext}"
        while os.path.exists(new_path):
            counter += 1
            new_path = f"{base}_{counter}{ext}"
        return new_path
    else:
        return None

# --- PROCESS RUNNERS ---
def process_single_file(input_path, ui, excel_sheet=None):
    pythoncom.CoInitialize()
    success = False
    try:
        ui.update_progress(15, "initializing engines...")
        raw_output_path = os.path.splitext(input_path)[0] + ".pdf"
        
        final_output_path = handle_path_collision(raw_output_path)
        if not final_output_path:
            return
            
        ui.update_progress(50, "parsing file layout...")
        route_conversion(input_path, final_output_path, excel_sheet)
        success = True
    except Exception as e:
        messagebox.showerror("Conversion Error", str(e))
    finally:
        pythoncom.CoUninitialize()
        if success:
            ui.show_success("conversion complete!")
        else:
            ui.close()

def execute_merge(files, save_path, ui):
    pythoncom.CoInitialize()
    temp_files = []
    total_files = len(files)
    success = False
    try:
        merger = PdfWriter()
        for index, file in enumerate(files):
            step_base = int((index / total_files) * 80)
            ui.update_progress(10 + step_base, f"converting file {index+1} of {total_files}...")
            
            ext = os.path.splitext(file)[1].lower()
            if ext == '.pdf':
                merger.append(file)
            else:
                fd, temp_path = tempfile.mkstemp(suffix=".pdf")
                os.close(fd)
                temp_files.append(temp_path)
                route_conversion(file, temp_path)
                merger.append(temp_path)
                
        ui.update_progress(90, "compiling documents...")
        merger.write(save_path)
        merger.close()
        success = True
    except Exception as e:
        messagebox.showerror("Merge Error", f"Failed to complete operations: {e}")
    finally:
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                try: os.remove(temp_file)
                except: pass
        pythoncom.CoUninitialize()
        if success:
            ui.show_success("All elements processed and unified!")
        else:
            ui.close()

# --- DRAG & DROP INTERACTIVE LISTBOX ---
class DragDropListbox(tk.Listbox):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.bind('<Button-1>', self.setCurrentIndex)
        self.bind('<B1-Motion>', self.shiftItems)
        self.curIndex = None

    def setCurrentIndex(self, event):
        self.curIndex = self.nearest(event.y)

    def shiftItems(self, event):
        i = self.nearest(event.y)
        if i < self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i+1, x)
            self.curIndex = i
        elif i > self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i-1, x)
            self.curIndex = i

# --- MULTI FILE MERGER WINDOW ---
def launch_merger_ui(files):
    window = tk.Tk()
    window.title("Little's PDF - Merger Tool")
    window.geometry("580x420")
    window.eval('tk::PlaceWindow . center')
    window.configure(bg="#f4f4f4")
    
    icon_path = get_asset_path('icon.ico')
    if os.path.exists(icon_path):
        window.iconbitmap(icon_path)

    tk.Label(window, text="Click and Drag items to Rearrange", 
             font=("Segoe UI", 10, "bold"), fg="#333333", bg="#f4f4f4").pack(pady=(15, 5))

    frame = tk.Frame(window, bg="#f4f4f4")
    frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    listbox = DragDropListbox(frame, selectmode=tk.SINGLE, font=("Segoe UI", 9), 
                              relief="flat", highlightthickness=1, highlightbackground="#cccccc")
    
    v_scroll = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=listbox.xview)
    h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
    
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    listbox.config(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

    file_map = {}
    for f in files:
        base_name = os.path.basename(f)
        original_base = base_name
        counter = 1
        while base_name in file_map and file_map[base_name] != f:
            base_name = f"({counter}) {original_base}"
            counter += 1
        file_map[base_name] = f
        listbox.insert(tk.END, base_name)

    def move_up():
        sel = listbox.curselection()
        if not sel or sel[0] == 0: return
        idx = sel[0]
        val = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx - 1, val)
        listbox.selection_set(idx - 1)

    def move_down():
        sel = listbox.curselection()
        if not sel or sel[0] == listbox.size() - 1: return
        idx = sel[0]
        val = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx + 1, val)
        listbox.selection_set(idx + 1)

    def start_merge():
        ordered_bases = listbox.get(0, tk.END)
        if len(ordered_bases) < 2:
            messagebox.showwarning("Warning", "Need at least 2 files to merge.")
            return

        ordered_files = [file_map[b] for b in ordered_bases]

        default_dir = os.path.dirname(ordered_files[0])
        save_path = filedialog.asksaveasfilename(
            initialdir=default_dir,
            initialfile="Merged_Document.pdf",
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")], 
            title="Save Merged PDF As"
        )
        if not save_path: return
        
        window.destroy()
        
        ui = MiniProgress("initializing merger pipeline...")
        threading.Thread(target=execute_merge, args=(ordered_files, save_path, ui), daemon=True).start()
        ui.root.mainloop()

    btn_frame = tk.Frame(window, bg="#f4f4f4")
    btn_frame.pack(pady=15)

    tk.Button(btn_frame, text="Move Up", width=12, relief="flat", bg="#e0e0e0", command=move_up).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Move Down", width=12, relief="flat", bg="#e0e0e0", command=move_down).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Merge into PDF", width=20, bg="#0078D7", fg="white", 
              font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", command=start_merge).pack(side=tk.LEFT, padx=15)

    window.mainloop()

# --- SYSTEM ROUTER ---
if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        sys.exit()

    if len(args) == 1 and args[0].lower().endswith('.pdf'):
        sys.exit()

    if len(args) == 1:
        target_file = args[0]
        file_extension = os.path.splitext(target_file)[1].lower()
        selected_sheet_target = None
        
        if file_extension in ['.xls', '.xlsx']:
            pythoncom.CoInitialize()
            ex_app = None
            try:
                ex_app = win32com.client.Dispatch("Excel.Application")
                ex_app.Visible = False
                ex_app.DisplayAlerts = False
                try: ex_app.Interactive = False
                except: pass
                
                ex_wb = ex_app.Workbooks.Open(os.path.abspath(target_file), ReadOnly=True)
                names = [sheet.Name for sheet in ex_wb.Worksheets]
                ex_wb.Close(False)
            except:
                pass
            finally:
                if ex_app:
                    try: ex_app.Interactive = True
                    except: pass
                    try: ex_app.Quit()
                    except: pass
                    del ex_app
                pythoncom.CoUninitialize()
                
            if 'names' in locals() and len(names) > 1:
                selector = SheetSelector(names)
                selected_sheet_target = selector.result
                if not selected_sheet_target:
                    sys.exit()
                
        ui = MiniProgress("parsing file layout...")
        threading.Thread(target=process_single_file, args=(target_file, ui, selected_sheet_target), daemon=True).start()
        ui.root.mainloop()
    else:
        clean_args = [f for f in args]
        if len(clean_args) >= 1:
            launch_merger_ui(clean_args)