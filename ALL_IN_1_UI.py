import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk
import threading
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
from datetime import datetime
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

# ==============================================================================
# PART 1: CONSTANTS & BACKEND LOGIC
# ==============================================================================

# Spectrometer Mapping: [Source_Cell, Dest_Cell]
# Copied from Code 2, with E10 fix applied
SPECTRO_MAPPINGS = [
    ['C2',  'E14'],  # Heat Code
    ['B10', 'E16'],  # Element 1
    ['C10', 'E17'],  # Element 2
    ['D10', 'E18'],  # Element 3
    ['F10', 'E19'],  # Element 4
    ['E10', 'E20'],  # Element 5 (Fixed)
    ['G10', 'E21'],  # Element 6
    ['L10', 'E22'],  # Element 7
    ['S10', 'E23']   # Element 8
]

class LogHelper:
    """Helper to safely write to the GUI log window from a thread."""
    def __init__(self, text_widget, root):
        self.widget = text_widget
        self.root = root

    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] [{level}] {message}"
        
        # Schedule the UI update on the main thread
        self.root.after(0, lambda: self._write(formatted_msg, level))

    def _write(self, message, level):
        self.widget.configure(state='normal')
        
        # Add color tags based on level
        tag = "info"
        if level == "ERROR": tag = "error"
        elif level == "SUCCESS": tag = "success"
        elif level == "DATA": tag = "data"
        
        self.widget.insert(tk.END, message + "\n", tag)
        self.widget.see(tk.END)
        self.widget.configure(state='disabled')

# --- Logic: Docx Extraction ---
def extract_micro_data(docx_path, logger):
    results = {}
    if not docx_path or not os.path.exists(docx_path):
        return results

    logger.log(f"Scanning Microstructure (DOCX): {os.path.basename(docx_path)}")
    
    try:
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
            tree = ET.fromstring(xml_content)
    except Exception as e:
        logger.log(f"Failed to read DOCX: {e}", "ERROR")
        return results

    all_text_chunks = []
    for elem in tree.iter():
        if elem.tag.endswith('}t') and elem.text and elem.text.strip():
            all_text_chunks.append(elem.text.strip())

    target_labels = {
        "Graphite Nodularity": "last", "Nodular Particles per mm²": "last",
        "Graphite Size": "last", "Graphite Form": "last",
        "Graphite Fraction": "last", "Ferrite / Pearlite Ratio": "first"
    }
    
    for label, preference in target_labels.items():
        target_index = -1
        
        # Find index
        if preference == "last":
            for i, chunk in enumerate(all_text_chunks):
                if label.lower() in chunk.lower(): target_index = i
        else: # first
            for i, chunk in enumerate(all_text_chunks):
                if label.lower() in chunk.lower(): 
                    target_index = i; break

        if target_index != -1:
            neighbors = all_text_chunks[target_index+1:target_index+6]
            found_value = None
            
            # Specific heuristic logic per label
            if label == "Graphite Fraction":
                for j, n in enumerate(neighbors):
                    if "%" in n and any(c.isdigit() for c in n): found_value = n; break
                    if any(c.isdigit() for c in n) and j+1 < len(neighbors) and neighbors[j+1] == "%":
                        found_value = f"{n}{neighbors[j+1]}"; break
            elif label == "Graphite Form":
                for n in neighbors:
                     if "(" in n and ")" in n: found_value = n; break
            elif label == "Ferrite / Pearlite Ratio":
                combined = "".join(neighbors[0:3])
                match = re.search(r"(\d+\.?\d*%\s*/\s*\d+\.?\d*%)", combined)
                if match: found_value = match.group(1)
            elif label == "Graphite Nodularity":
                for n in neighbors:
                    if "%" in n and len(n) > 1: found_value = n; break
            else:
                for n in neighbors:
                    if any(c.isdigit() for c in n) and not n.endswith('%'):
                        found_value = re.sub(r'[\s\.\,]+$', '', n); break
            
            if found_value:
                results[label] = found_value
                logger.log(f"  > Found {label}: {found_value}", "DATA")
    return results

# --- Logic: PDF Helpers ---
def find_value_neighbor(elements, label_text, required_keyword="Mpa"):
    label_bbox = None
    for element in elements:
        if label_text in element.get_text():
            label_bbox = element.bbox; break    
    if not label_bbox: return None

    lx0, ly0, lx1, ly1 = label_bbox
    closest_distance = 9999
    best_text = None
    
    for element in elements:
        text = element.get_text().strip()
        ex0, ey0, ex1, ey1 = element.bbox
        if label_text in text: continue
        # Simple proximity check
        if (ey0 < ly1 + 2) and (ey1 > ly0 - 2) and (ex0 >= lx0 - 5) and (required_keyword in text):
            distance = ex0 - lx1
            if distance < closest_distance:
                closest_distance = distance
                best_text = text
    
    if best_text:
        match = re.search(r"([\d\.]+)", best_text)
        return match.group(1) if match else best_text
    return None

def extract_pdf_elements(pdf_path, logger):
    elements = []
    if not pdf_path or not os.path.exists(pdf_path): return elements
    try:
        for page_layout in extract_pages(pdf_path, page_numbers=[0]):
            for element in page_layout:
                if isinstance(element, LTTextContainer): elements.append(element)
    except Exception as e:
        logger.log(f"Error reading PDF {os.path.basename(pdf_path)}: {e}", "ERROR")
    return elements

# --- Logic: Tensile PDF ---
def process_tensile(pdf_path, logger):
    if not pdf_path: return (None, None, None)
    logger.log(f"Scanning Tensile (PDF): {os.path.basename(pdf_path)}")
    elements = extract_pdf_elements(pdf_path, logger)
    if not elements: return (None, None, None)

    val_tensile = find_value_neighbor(elements, "Tensile Strength", "Mpa")
    val_yield = find_value_neighbor(elements, "Yield Strength", "Mpa")
    val_elong = find_value_neighbor(elements, "Elongation", "%")

    if val_tensile: logger.log(f"  > Tensile: {val_tensile}", "DATA")
    if val_yield: logger.log(f"  > Yield: {val_yield}", "DATA")
    if val_elong: logger.log(f"  > Elongation: {val_elong}", "DATA")

    return val_tensile, val_yield, val_elong

# --- Logic: Hardness PDF ---
def process_hardness(pdf_path, logger):
    if not pdf_path: return []
    logger.log(f"Scanning Hardness (PDF): {os.path.basename(pdf_path)}")
    elements = extract_pdf_elements(pdf_path, logger)
    if not elements: return []

    hardness_labels = [e for e in elements if "Hardness" in e.get_text()]
    # Sort top to bottom
    hardness_labels.sort(key=lambda x: x.bbox[3], reverse=True)
    
    results = []
    for label in hardness_labels:
        found = None
        # Try finding inside label text first
        match_inside = re.search(r"([\d\.]+)\s*HBW", label.get_text())
        if match_inside: 
            found = match_inside.group(1)
        else:
            # Look for neighbor to the right
            lx0, ly0, lx1, ly1 = label.bbox
            closest = 9999
            for element in elements:
                etext = element.get_text().strip()
                if "HBW" not in etext: continue
                ex0, ey0, ex1, ey1 = element.bbox
                if (ey0 < ly1 + 5) and (ey1 > ly0 - 5) and (ex0 > lx0):
                    dist = ex0 - lx1
                    if dist < closest:
                        nm = re.search(r"([\d\.]+)\s*HBW", etext)
                        if nm:
                            closest = dist
                            found = nm.group(1)
        if found:
            results.append(found)
            logger.log(f"  > Hardness Found: {found}", "DATA")
            if len(results) >= 2: break # Only need max 2 usually
            
    return results


# ==============================================================================
# PART 2: UI CLASS (Merged)
# ==============================================================================

class UnifiedMTCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal MTC Automation Tool")
        self.root.geometry("800x750")

        # --- Style Configuration ---
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # --- Variables ---
        self.path_mtc = tk.StringVar()       # Destination File
        
        # Mechanical Source Files
        self.path_micro = tk.StringVar()
        self.path_tensile = tk.StringVar()
        self.path_hardness = tk.StringVar()
        
        # Chemical Source File
        self.path_spectro = tk.StringVar()
        
        # Checkboxes for execution control
        self.do_mechanical = tk.BooleanVar(value=True)
        self.do_chemical = tk.BooleanVar(value=True)

        self._init_ui()

    def _init_ui(self):
        # 1. Main Title
        lbl_title = tk.Label(self.root, text="MTC & Report Automation Hub", font=("Segoe UI", 16, "bold"), fg="#333")
        lbl_title.pack(pady=10)

        # 2. Frame: Destination File (The Central File)
        frm_dest = tk.LabelFrame(self.root, text="DESTINATION: MTC Entry Sheet", font=("Segoe UI", 10, "bold"), fg="#D32F2F", padx=10, pady=10)
        frm_dest.pack(fill="x", padx=15, pady=5)
        self._build_file_selector(frm_dest, "MTC Excel File (.xlsx)*:", self.path_mtc)

        # 3. Inputs Container (Split logic visually)
        frm_inputs = tk.Frame(self.root)
        frm_inputs.pack(fill="both", padx=15, pady=5)

        # 3A. Left Side: Mechanical Sources
        frm_mech = tk.LabelFrame(frm_inputs, text="Mechanical Data Sources", font=("Segoe UI", 9, "bold"), padx=5, pady=5)
        frm_mech.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        tk.Checkbutton(frm_mech, text="Run Mechanical Extraction", variable=self.do_mechanical, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Separator(frm_mech, orient="horizontal").pack(fill="x", pady=5)
        
        self._build_file_selector(frm_mech, "Micro DOCX:", self.path_micro, is_compact=True, file_type="docx")
        self._build_file_selector(frm_mech, "Tensile PDF:", self.path_tensile, is_compact=True, file_type="pdf")
        self._build_file_selector(frm_mech, "Hardness PDF:", self.path_hardness, is_compact=True, file_type="pdf")

        # 3B. Right Side: Chemical Source
        frm_chem = tk.LabelFrame(frm_inputs, text="Chemical Data Sources", font=("Segoe UI", 9, "bold"), padx=5, pady=5)
        frm_chem.pack(side="right", fill="both", expand=True, padx=(5, 0))

        tk.Checkbutton(frm_chem, text="Run Chemical Transfer", variable=self.do_chemical, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Separator(frm_chem, orient="horizontal").pack(fill="x", pady=5)
        
        self._build_file_selector(frm_chem, "Spectro XLSX:", self.path_spectro, is_compact=True, file_type="xlsx")

        # 4. Logger Area
        frm_log = tk.LabelFrame(self.root, text="Process Log", padx=5, pady=5)
        frm_log.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.txt_log = scrolledtext.ScrolledText(frm_log, height=12, state='disabled', font=("Consolas", 9))
        self.txt_log.pack(fill="both", expand=True)
        # Config tags for colors
        self.txt_log.tag_config("error", foreground="red")
        self.txt_log.tag_config("success", foreground="green")
        self.txt_log.tag_config("data", foreground="blue")
        self.txt_log.tag_config("info", foreground="black")

        # 5. Progress Bar & Buttons
        self.lbl_status = tk.Label(self.root, text="Ready to start", anchor="w", fg="gray")
        self.lbl_status.pack(fill="x", padx=15)
        
        self.progress = ttk.Progressbar(self.root, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=15, pady=(0, 10))

        frm_btns = tk.Frame(self.root)
        frm_btns.pack(fill="x", padx=15, pady=10)
        
        tk.Button(frm_btns, text="START PROCESSING", command=self.start_processing, 
                  bg="#2E7D32", fg="white", font=("Segoe UI", 11, "bold"), height=2, width=25).pack(side="left")
        
        tk.Button(frm_btns, text="EXIT", command=self.close_app, 
                  bg="#C62828", fg="white", font=("Segoe UI", 11, "bold"), height=2, width=15).pack(side="right")

    def _build_file_selector(self, parent, label, variable, is_compact=False, file_type="xlsx"):
        frame = tk.Frame(parent)
        frame.pack(fill="x", pady=2)
        
        if is_compact:
            tk.Label(frame, text=label, width=12, anchor="w", font=("Segoe UI", 8)).pack(side="left")
            entry_width = 15
        else:
            tk.Label(frame, text=label, width=20, anchor="w").pack(side="left")
            entry_width = 40
            
        tk.Entry(frame, textvariable=variable).pack(side="left", fill="x", expand=True, padx=5)
        
        ftypes = [("All Files", "*.*")]
        if file_type == "xlsx": ftypes = [("Excel Files", "*.xlsx")]
        if file_type == "pdf": ftypes = [("PDF Files", "*.pdf")]
        if file_type == "docx": ftypes = [("Word Files", "*.docx")]
        
        tk.Button(frame, text="Choose File", width=10, 
                  command=lambda: self._browse(variable, ftypes)).pack(side="left")

    def _browse(self, variable, ftypes):
        f = filedialog.askopenfilename(filetypes=ftypes)
        if f: variable.set(f)

    def close_app(self):
        if messagebox.askokcancel("Exit", "Do you really want to quit?"):
            self.root.destroy()

    # --- Processing Logic ---

    def start_processing(self):
        # 1. Validations
        mtc_file = self.path_mtc.get()
        if not mtc_file:
            messagebox.showwarning("Missing Destination", "Please select the 'MTC Entry Sheet' (Destination) file.")
            return
        if not self.do_chemical.get() and not self.do_mechanical.get():
            messagebox.showwarning("Nothing Selected", "Please select at least one operation (Mechanical or Chemical).")
            return

        # 2. UI Lock
        self.progress['value'] = 0
        self.txt_log.configure(state='normal')
        self.txt_log.delete(1.0, tk.END)
        self.txt_log.configure(state='disabled')
        
        # 3. Threading
        threading.Thread(target=self._run_thread, args=(mtc_file,), daemon=True).start()

    def _run_thread(self, mtc_file):
        logger = LogHelper(self.txt_log, self.root)
        logger.log("--- Starting Process ---")

        try:
            # CHECK DESTINATION
            if not os.path.exists(mtc_file):
                logger.log(f"Destination file not found: {mtc_file}", "ERROR")
                return
            
            # OPEN DESTINATION (Do this once!)
            logger.log("Opening MTC Destination file...")
            self.update_prog(10, "Opening Workbook...")
            try:
                wb_dest = openpyxl.load_workbook(mtc_file)
                ws_dest = wb_dest.active
            except PermissionError:
                logger.log("Error: File is Open. Please close Excel and try again.", "ERROR")
                return

            # --- PART A: CHEMICAL (SPECTRO) ---
            if self.do_chemical.get():
                spectro_file = self.path_spectro.get()
                if spectro_file and os.path.exists(spectro_file):
                    logger.log(f"Starting Chemical Transfer from: {os.path.basename(spectro_file)}")
                    self.update_prog(20, "Processing Chemical Data...")
                    
                    try:
                        wb_src = openpyxl.load_workbook(spectro_file, data_only=True)
                        ws_src = wb_src.active
                        
                        count = 0
                        for src_cell, dst_cell in SPECTRO_MAPPINGS:
                            val = ws_src[src_cell].value
                            ws_dest[dst_cell].value = val
                            # Detailed debug in log
                            # logger.log(f"  Transfer: {src_cell} ({val}) -> {dst_cell}") 
                            count += 1
                        logger.log(f"Chemical Data Transferred: {count} cells.", "SUCCESS")
                    except Exception as e:
                        logger.log(f"Error in Chemical Transfer: {e}", "ERROR")
                else:
                    logger.log("Skipping Chemical: File missing or not selected.", "info")

            # --- PART B: MECHANICAL ---
            if self.do_mechanical.get():
                self.update_prog(40, "Processing Mechanical Inputs...")
                
                # 1. Micro structure
                if self.path_micro.get():
                    micro_res = extract_micro_data(self.path_micro.get(), logger)
                    # Writing to Destination
                    micro_map = {
                        "Graphite Nodularity": 'T36', "Nodular Particles per mm²": 'T37',
                        "Graphite Size": 'T38', "Graphite Form": 'T39',
                        "Graphite Fraction": 'T40', "Ferrite / Pearlite Ratio": 'T41'
                    }
                    for k, cell in micro_map.items():
                        if k in micro_res: ws_dest[cell] = micro_res[k]
                
                self.update_prog(60, "Processing Tensile/Hardness...")
                
                # 2. Tensile
                if self.path_tensile.get():
                    t_tens, t_yield, t_elong = process_tensile(self.path_tensile.get(), logger)
                    if t_tens: ws_dest['E26'] = t_tens
                    if t_yield: ws_dest['E27'] = t_yield
                    if t_elong: ws_dest['E28'] = t_elong

                # 3. Hardness
                if self.path_hardness.get():
                    hard_vals = process_hardness(self.path_hardness.get(), logger)
                    if len(hard_vals) > 0: ws_dest['E29'] = hard_vals[0]
                    if len(hard_vals) > 1: ws_dest['E30'] = hard_vals[1]

            # --- SAVE ---
            logger.log("Saving MTC Entry Sheet...")
            self.update_prog(90, "Saving file...")
            wb_dest.save(mtc_file)
            
            self.update_prog(100, "Done!")
            logger.log("PROCESS COMPLETED SUCCESSFULLY.", "SUCCESS")
            messagebox.showinfo("Done", "Data transfer and update complete!")

        except Exception as e:
            logger.log(f"CRITICAL ERROR: {str(e)}", "ERROR")
            messagebox.showerror("Error", str(e))
        finally:
             self.root.after(0, lambda: self.lbl_status.configure(text="Ready"))

    def update_prog(self, val, text):
        self.root.after(0, lambda: self._update_prog_ui(val, text))

    def _update_prog_ui(self, val, text):
        self.progress['value'] = val
        self.lbl_status.configure(text=text)

if __name__ == "__main__":
    root = tk.Tk()
    # High DPI Awareness for Windows (makes text look sharp)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
        
    app = UnifiedMTCApp(root)
    root.mainloop()
