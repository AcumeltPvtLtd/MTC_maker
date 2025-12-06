import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk  # Import ttk for the modern Progress Bar
import openpyxl
import os

# --- Configuration: Mapping ---
# [Source_Cell (Spectrometer), Destination_Cell (MTC Entry)]
CELL_MAPPINGS = [
    ['C2',  'E14'],  # Heat Code
    ['B10', 'E16'],  # Element 1
    ['C10', 'E17'],  # Element 2
    ['D10', 'E18'],  # Element 3
    ['F10', 'E19'],  # Element 4
    ['E10', 'E20'],  # Element 5 (FIXED: Changed P10 to E10)
    ['G10', 'E21'],  # Element 6
    ['L10', 'E22'],  # Element 7
    ['S10', 'E23']   # Element 8
]

class MTCTransferApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MTC Data Transfer Tool")
        self.root.geometry("500x480")  # Increased height for widgets

        # Variables to store file paths
        self.source_path = ""
        self.dest_path = ""

        # --- UI Layout ---
        
        # 1. Title
        title_label = tk.Label(root, text="Spectrometer to MTC Transfer", font=("Arial", 14, "bold"))
        title_label.pack(pady=15)

        # 2. Source File Section
        self.btn_source = tk.Button(root, text="1. Select Spectrometer Report (.xlsx)", 
                                    command=self.select_source_file, width=40, bg="#e1e1e1")
        self.btn_source.pack(pady=5)
        
        self.lbl_source = tk.Label(root, text="No file selected", fg="gray", wraplength=450)
        self.lbl_source.pack(pady=5)

        # 3. Destination File Section
        self.btn_dest = tk.Button(root, text="2. Select MTC Entry Sheet (.xlsx)", 
                                  command=self.select_dest_file, width=40, bg="#e1e1e1")
        self.btn_dest.pack(pady=5)
        
        self.lbl_dest = tk.Label(root, text="No file selected", fg="gray", wraplength=450)
        self.lbl_dest.pack(pady=5)

        # --- Progress Bar Section ---
        self.lbl_status = tk.Label(root, text="Ready", font=("Arial", 10, "italic"))
        self.lbl_status.pack(pady=(15, 0))

        # We use a frame to hold the progress bar
        self.progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=5)

        # 4. Action Buttons Frame
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=20)

        # Run Button
        self.btn_run = tk.Button(btn_frame, text="TRANSFER DATA", 
                                 command=self.run_transfer, 
                                 width=20, height=2, bg="green", fg="white", font=("Arial", 10, "bold"))
        self.btn_run.pack(side=tk.LEFT, padx=10)

        # --- Close Button ---
        self.btn_close = tk.Button(btn_frame, text="CLOSE", 
                                   command=self.close_app, 
                                   width=15, height=2, bg="#cc0000", fg="white", font=("Arial", 10, "bold"))
        self.btn_close.pack(side=tk.LEFT, padx=10)

    def select_source_file(self):
        filename = filedialog.askopenfilename(
            title="Select Spectrometer Report",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if filename:
            self.source_path = filename
            self.lbl_source.config(text=f"Source: {os.path.basename(filename)}", fg="black")

    def select_dest_file(self):
        filename = filedialog.askopenfilename(
            title="Select MTC Entry Sheet",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if filename:
            self.dest_path = filename
            self.lbl_dest.config(text=f"Dest: {os.path.basename(filename)}", fg="black")

    def update_progress(self, value, text):
        """Helper to update bar and label text immediately"""
        self.progress['value'] = value
        self.lbl_status.config(text=text)
        self.root.update_idletasks() # Force UI refresh

    def run_transfer(self):
        # Validation
        if not self.source_path or not self.dest_path:
            messagebox.showwarning("Missing Files", "Please select both files first.")
            return

        try:
            # Step 0: Loading
            self.update_progress(10, "Loading workbooks... Please wait.")
            
            # 1. Load Source (data_only=True gets the calculated value, not the formula)
            wb_source = openpyxl.load_workbook(self.source_path, data_only=True)
            ws_source = wb_source.active

            # 2. Load Destination
            wb_dest = openpyxl.load_workbook(self.dest_path)
            ws_dest = wb_dest.active

            self.update_progress(30, "Transferring Data...")

            # 3. Transfer Data Loop
            total_mappings = len(CELL_MAPPINGS)
            
            for i, (src_cell, dst_cell) in enumerate(CELL_MAPPINGS):
                # Read
                value = ws_source[src_cell].value
                # Write
                ws_dest[dst_cell].value = value
                
                # Update Progress Calculation (Mapping steps from 30% to 80%)
                percent = 30 + int(((i + 1) / total_mappings) * 50)
                self.update_progress(percent, f"{percent}% - Copying {src_cell} to {dst_cell}...")

            # 4. Save
            self.update_progress(85, "Saving destination file... (Don't close)")
            wb_dest.save(self.dest_path)
            
            # Done
            self.update_progress(100, "Transfer Complete!")
            messagebox.showinfo("Success", "Data transferred and saved successfully!")
            
            # Reset UI slightly after
            self.lbl_status.config(text="Ready")
            self.progress['value'] = 0

        except PermissionError:
            self.update_progress(0, "Error: File Open")
            messagebox.showerror("Error", "Permission Denied.\nPlease close the Excel files if they are open.")
        except Exception as e:
            self.update_progress(0, "Error")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    def close_app(self):
        if messagebox.askokcancel("Quit", "Do you want to close the program?"):
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MTCTransferApp(root)
    root.mainloop()
