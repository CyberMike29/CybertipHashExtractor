import os
import re
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import pandas as pd
import pdfplumber


def extract_file_info(pdf_path: str) -> list[dict]:
    extracted_records = []
    full_text = ""
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    md5_regex = re.compile(r"(?:MD5|Hash)[\s:]*([a-fA-F0-9]{32})", re.IGNORECASE)
    filename_regex = re.compile(r"(?:Filename|File\s*Name)[\s:]*(.*)", re.IGNORECASE)

    lines = full_text.splitlines()
    current_filename = None
    i = 0

    while i < len(lines):
        line_clean = lines[i].strip()
        
        # 1. Detect filename
        fn_match = filename_regex.search(line_clean)
        if fn_match:
            parts = [fn_match.group(1).strip()]
            i += 1
            
            # Consume continuation lines until reaching MD5, another header, or a blank line
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line:
                    break
                if md5_regex.search(next_line) or re.search(r"^[A-Za-z\s]+:", next_line):
                    break
                
                parts.append(next_line)
                i += 1
            
            # Join wrapped line fragments
            merged_filename = "".join(parts).strip()

            # Truncate immediately after the dot and 3-4 extension characters
            # e.g., "photo_123_n.jpg Created: 2023..." becomes "photo_123_n.jpg"
            clean_match = re.search(r"^(.*?\.[a-zA-Z0-9]{3,4})", merged_filename)
            if clean_match:
                current_filename = clean_match.group(1).strip()
            else:
                current_filename = merged_filename

            continue

        # 2. Detect MD5 hash
        md5_match = md5_regex.search(line_clean)
        if md5_match:
            md5_val = md5_match.group(1).strip()
            extracted_records.append({
                "Source PDF": Path(pdf_path).name,
                "Filename": current_filename if current_filename else "Not Found",
                "MD5": md5_val,
            })
            current_filename = None

        i += 1

    # Fallback if no structured records were found
    if not extracted_records:
        raw_md5s = re.findall(r"\b([a-fA-F0-9]{32})\b", full_text)
        if raw_md5s:
            for h in raw_md5s:
                extracted_records.append({
                    "Source PDF": Path(pdf_path).name,
                    "Filename": "Located by raw hash pattern",
                    "MD5": h,
                })

    return extracted_records


class PDFExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Hash & File Info Extractor")
        self.geometry("620x520")
        self.minsize(550, 450)

        self.selected_files = []
        self._build_ui()

    def _build_ui(self):
        # 1. PDF File Selection Section
        file_frame = ttk.LabelFrame(self, text="Step 1: Select Input PDFs", padding=10)
        file_frame.pack(fill="x", padx=12, pady=(10, 5))

        btn_select_files = ttk.Button(file_frame, text="Browse PDFs...", command=self.choose_files)
        btn_select_files.pack(side="left", padx=(0, 10))

        self.lbl_files = ttk.Label(file_frame, text="No PDF files selected", foreground="gray")
        self.lbl_files.pack(side="left", fill="x", expand=True)

        # 2. Output Location Section
        out_frame = ttk.LabelFrame(self, text="Step 2: Choose Destination Excel File", padding=10)
        out_frame.pack(fill="x", padx=12, pady=5)

        btn_select_dest = ttk.Button(out_frame, text="Save As...", command=self.choose_output_path)
        btn_select_dest.pack(side="left", padx=(0, 10))

        self.output_path_var = tk.StringVar(value="")
        entry_dest = ttk.Entry(out_frame, textvariable=self.output_path_var)
        entry_dest.pack(side="left", fill="x", expand=True)

        # 3. Action Section
        action_frame = ttk.Frame(self, padding=5)
        action_frame.pack(fill="x", padx=12, pady=5)

        self.btn_run = ttk.Button(action_frame, text="Start Extraction", command=self.start_processing_thread)
        self.btn_run.pack(fill="x")

        # 4. Progress and Logs
        log_frame = ttk.LabelFrame(self, text="Status & Logs", padding=10)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(5, 10))

        self.progress = ttk.Progressbar(log_frame, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))

        self.log_box = ScrolledText(log_frame, height=10, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)

    def choose_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            self.selected_files = list(files)
            count = len(self.selected_files)
            self.lbl_files.config(
                text=f"{count} file{'s' if count != 1 else ''} selected",
                foreground="black"
            )
            self.log(f"Selected {count} file(s).")

    def choose_output_path(self):
        output_file = filedialog.asksaveasfilename(
            title="Choose Excel Destination",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if output_file:
            self.output_path_var.set(output_file)

    def log(self, text: str):
        self.log_box.config(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def start_processing_thread(self):
        if not self.selected_files:
            messagebox.showwarning("Missing Files", "Please select at least one PDF file.")
            return

        out_path = self.output_path_var.get().strip()
        if not out_path:
            messagebox.showwarning("Missing Output Path", "Please specify where to save the Excel file.")
            return

        # Disable button and run extraction in a thread to keep GUI responsive
        self.btn_run.config(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self._process_worker, args=(out_path,), daemon=True).start()

    def _process_worker(self, output_path: str):
        all_records = []
        total = len(self.selected_files)

        self.log("\nStarting extraction...")

        for index, pdf_path in enumerate(self.selected_files, start=1):
            fname = Path(pdf_path).name
            self.log(f"[{index}/{total}] Reading: {fname}")
            try:
                records = extract_file_info(pdf_path)
                all_records.extend(records)
                self.log(f"  -> Found {len(records)} record(s)")
            except Exception as e:
                self.log(f"  -> Error reading {fname}: {e}")

            # Update progress bar
            self.progress["value"] = (index / total) * 100

        if all_records:
            try:
                # 1. Save to Excel
                df = pd.DataFrame(all_records)
                df.to_excel(output_path, index=False)
                self.log(f"\nDone! Successfully saved {len(df)} entries to:\n{output_path}")

                # 2. Save MD5 hashes to a separate text file
                txt_path = Path(output_path).with_suffix(".txt")
                
                # Option A: Unique hashes only (remove dict.fromkeys if duplicates are needed)
                hashes = list(dict.fromkeys(r["MD5"] for r in all_records if r.get("MD5")))
                
                with open(txt_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(hashes) + "\n")
                
                self.log(f"MD5 hashes saved to:\n{txt_path}")

                messagebox.showinfo(
                    "Success",
                    f"Extracted {len(df)} entries to Excel and saved {len(hashes)} MD5 hashes to text file!"
                )
            except Exception as e:
                self.log(f"\nError saving files: {e}")
                messagebox.showerror("Export Failed", f"Could not write outputs: {e}")


if __name__ == "__main__":
    app = PDFExtractorApp()
    app.mainloop()
