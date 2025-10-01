# Docling Folder Converter
# A GUI tool to convert all documents in a folder to Markdown using Docling.
# Cleaning after conversion to remove unwanted sections. Making DOCX weekly reports LLM-Ready
# Andy Mercado, 9/24/2025

import os
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docling.document_converter import DocumentConverter


# ---------- Text Cleaning ----------
def cut_intro_sections(text: str) -> str:
    """
    1) Cut everything before the first real bold section header (**Name**) that is NOT
       'To', 'From', 'Date', or 'Subject'.
    2) Remove unwanted blocks (Personnel, Meetings, Safety, Kudos, Training, Goals, Standby).
    """
    # Normalize newlines
    t = text.replace("\r\n", "\n").replace("\r", "\n")

    # --- STEP 1: find first real section header ---
    memo = {"to", "from", "date", "subject"}
    header_pat = re.compile(r"\*\*\s*([^\n*]+?)\s*:?\s*\*\*", re.IGNORECASE)

    first_real_start = None
    for m in header_pat.finditer(t):
        title = m.group(1).strip().lower()
        if title not in memo:
            first_real_start = m.start()
            break

    if first_real_start is not None:
        t = t[first_real_start:].lstrip()

    # --- STEP 2: remove unwanted sections ---
    unwanted_keywords = [
        "weekly personnel",
        "personnel",
        "meeting",
        "training",
        "safety",
        "compliance",
        "kudos",
        "webinar",
        "standby",
        "automation standby",
        "automation overtime",
        "six months goals"
    ]

    matches = list(header_pat.finditer(t))
    if not matches:
        return t.lstrip()

    out_parts = []
    for i, m in enumerate(matches):
        # normalize header
        title_norm = re.sub(r"[^a-z0-9]+", " ", m.group(1).strip().lower()).strip()
        block_start = m.start()
        block_end = matches[i + 1].start() if i + 1 < len(matches) else len(t)

        block = t[block_start:block_end]
        # skip if header matches any unwanted keyword
        if any(kw in title_norm for kw in unwanted_keywords):
            continue
        out_parts.append(block)

    cleaned = "".join(out_parts).lstrip()
    return cleaned


def cut_off_after_standby(markdown_text: str) -> str:
    """Remove everything from Standby section onward (Automation Standby, Standby, etc)."""
    cutoff_pat = re.compile(r"\*\*\s*(automation\s+)?standby\s*:?[\s*]*", re.IGNORECASE)
    m = cutoff_pat.search(markdown_text)
    if m:
        return markdown_text[:m.start()].rstrip()
    return markdown_text


def split_into_chunks(markdown_text: str):
    """Split into chunks by bold headers (colon optional)."""
    parts = re.split(r"(?=\*\*[^*]+?\*\*)", markdown_text)
    return [p.strip() for p in parts if p.strip()]


def clean_text_pipeline(markdown_text: str) -> str:
    """Apply all cleaning steps in sequence."""
    t = cut_intro_sections(markdown_text)
    t = cut_off_after_standby(t)
    chunks = split_into_chunks(t)
    return "\n\n".join(chunks)



# ---------- Conversion ----------
def convert_folder(input_folder, output_folder, progress_bar, status_label):
    converter = DocumentConverter()

    # Count total files first
    all_files = []
    for root, _, files in os.walk(input_folder):
        for filename in files:
            all_files.append(Path(root) / filename)

    print("Total files found:", len(all_files))
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

            # Apply cleaning pipeline
            cleaned_text = clean_text_pipeline(markdown_text)

            # Save with same folder structure
            relative_path = file_path.relative_to(input_folder).with_suffix(".md")
            save_path = Path(output_folder) / relative_path
            save_path.parent.mkdir(parents=True, exist_ok=True)

            with open(save_path, "w", encoding="utf-8") as f:
                f.write(cleaned_text)

            status_label.config(text=f"Processed {idx}/{total_files}: {file_path.name}")

        except Exception as e:
            status_label.config(text=f"Skipped {file_path.name} (Error: {e})")

        # Update progress bar
        progress_bar["value"] = idx
        progress_bar.update()

    messagebox.showinfo("Done", f"Processed {total_files} files.")


# ---------- GUI ----------
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