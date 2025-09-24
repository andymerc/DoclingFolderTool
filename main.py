
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docling.document_converter import DocumentConverter

def convert_folder(input_folder, output_folder, progress_bar, status_label):
    converter = DocumentConverter()

    # Count total files first
    all_files = []
    for root, _, files in os.walk(input_folder):
        for filename in files:
            all_files.append(Path(root) / filename)

    total_files = len(all_files)
    if total_files == 0:
        messagebox.showwarning("No Files", "No files found in the input folder.")
        return

    progress_bar["maximum"] = total_files
    progress_bar["value"] = 0

    for idx, file_path in enumerate(all_files, start=1):
        try:
            # Convert file
            doc = converter.convert(str(file_path)).document
            markdown_text = doc.export_to_markdown()

            # Save with same folder structure
            relative_path = file_path.relative_to(input_folder).with_suffix(".md")
            save_path = Path(output_folder) / relative_path
            save_path.parent.mkdir(parents=True, exist_ok=True)

            with open(save_path, "w", encoding="utf-8") as f:
                f.write(markdown_text)

            status_label.config(text=f"Processed {idx}/{total_files}: {file_path.name}")

        except Exception as e:
            status_label.config(text=f"Skipped {file_path.name} (Error)")

        # Update progress bar
        progress_bar["value"] = idx
        progress_bar.update()

    messagebox.showinfo("Done", f"Processed {total_files} files.")

def select_input_folder():
    folder = filedialog.askdirectory(title="Select Input Folder")
    input_var.set(folder)

def select_output_folder():
    folder = filedialog.askdirectory(title="Select Output Folder")
    output_var.set(folder)

def run_conversion():
    input_folder = input_var.get()
    output_folder = output_var.get()

    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders.")
        return

    convert_folder(input_folder, output_folder, progress_bar, status_label)

# GUI setup
root = tk.Tk()
root.title("Docling Folder Converter")
root.geometry("550x250")

input_var = tk.StringVar()
output_var = tk.StringVar()

# Input selection
tk.Label(root, text="Input Folder:").pack(anchor="w", padx=10, pady=5)
frame_in = tk.Frame(root)
frame_in.pack(fill="x", padx=10)
tk.Entry(frame_in, textvariable=input_var, width=50).pack(side="left", expand=True, fill="x")
tk.Button(frame_in, text="Browse", command=select_input_folder).pack(side="left", padx=5)

# Output selection
tk.Label(root, text="Output Folder:").pack(anchor="w", padx=10, pady=5)
frame_out = tk.Frame(root)
frame_out.pack(fill="x", padx=10)
tk.Entry(frame_out, textvariable=output_var, width=50).pack(side="left", expand=True, fill="x")
tk.Button(frame_out, text="Browse", command=select_output_folder).pack(side="left", padx=5)

# Progress bar + status
progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.pack(pady=15, padx=10)

status_label = tk.Label(root, text="Waiting to start...", anchor="w")
status_label.pack(fill="x", padx=10)

# Run button
tk.Button(root, text="Run Conversion", command=run_conversion).pack(pady=10)

root.mainloop()
