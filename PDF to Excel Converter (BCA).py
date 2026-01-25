import sys
import re
from pathlib import Path
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

#!/usr/bin/env python3
"""
pdf_to_bank_csv.py

Convert "Rekening Giro BCA" PDF statements into clean CSV/Excel.
Usage:
    python pdf_to_bank_csv.py input.pdf output.csv
    python pdf_to_bank_csv.py input.pdf output.xlsx

Dependencies:
    pip install pdfplumber pandas openpyxl
"""


HEADERS = ["TANGGAL", "KETERANGAN", "CBG", "MUTASI", "SALDO"]


def detect_bank(path):
    """Look at the first page text to guess the bank/statement format.
    Returns one of: 'BNI', 'BNI_USD', 'MCM', 'BRI', or 'GENERIC'."""
    try:
        with pdfplumber.open(path) as pdf:
            if len(pdf.pages) == 0:
                return "GENERIC"
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return "GENERIC"
    t = text.upper()
    # BNI (IDR) / BNI USD
    if "BNI" in t:
        # if USD present and BNI mentioned, treat as BNI_USD
        if "USD" in t or "DOLLAR" in t or "US DOLLAR" in t:
            return "BNI_USD"
        return "BNI"
    # MCM (some statements contain MCM or MANDIRI CORPORATE MARKETS-like markers)
    if "MCM" in t or "MANDIRI" in t and "MCM" in t:
        return "MCM"
    # BRI
    if "BANK RAKYAT INDONESIA" in t or "BRI" in t:
        return "BRI"
    return "GENERIC"

def normalize_whitespace(s):
    return re.sub(r"\s+", " ", s).strip()


def normalize_numeric_token(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    s = s.replace('\xa0', '')
    # If comma used as decimal and no dots -> convert
    if ',' in s and '.' not in s:
        if re.search(r',[0-9]{1,2}$', s):
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    # If multiple dots, keep only last as decimal separator
    if s.count('.') > 1:
        last = s.rfind('.')
        intpart = s[:last].replace('.', '').replace(',', '')
        frac = s[last+1:]
        s = intpart + '.' + frac
    s = s.replace(',', '')
    s = re.sub(r'[^0-9\.\-]', '', s)
    s = re.sub(r'\.{2,}', '.', s)
    return s


def normalize_date_str(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    # try pandas parse first
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    # fallback: try some common formats
    from datetime import datetime
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%d/%m/%y', '%d-%m-%y'):
        try:
            return datetime.strptime(s.split()[0], fmt).strftime('%Y-%m-%d')
        except Exception:
            continue
    return s


def clean_description(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    # remove long numeric sequences
    s = re.sub(r'\b\d{6,}\b', '', s)
    # remove tokens like TO:000000000
    s = re.sub(r'\bT?O?:?0{3,}\d+\b', '', s)
    # remove long alphanumeric journal codes
    s = re.sub(r'\b[A-Z0-9]{8,}\b', '', s)
    return normalize_whitespace(s)

def clean_number(s):
    if s is None: return ""
    s = s.strip()
    # remove non-number except dot, comma, minus; then unify comma thousands -> remove commas
    s = re.sub(r"[^\d\.,\-]", "", s)
    # If both comma and dot present, assume comma thousands -> remove commas
    if "," in s and "." in s:
        s = s.replace(",", "")
    # If only commas present and they look like thousands (3-digit groups), remove them
    elif "," in s and re.search(r"\d+,\d{3}(?:,\d{3})*", s):
        s = s.replace(",", "")
    # collapse multiple dots
    s = re.sub(r"\.{2,}", ".", s)
    return s


def format_amount_str(s):
    """Return a cleaned amount string without trailing .00; keep decimals if present."""
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    # unify decimal separator to dot (should already be dot from clean_number)
    # try parse as float
    try:
        v = float(s)
    except Exception:
        # fallback: remove trailing .00 if present
        if s.endswith('.00'):
            return s[:-3]
        if s.endswith(',00'):
            return s[:-3]
        return s
    # if integer-valued, return as integer string
    if abs(v - int(v)) < 1e-9:
        return str(int(v))
    # otherwise return normalized float without unnecessary trailing zeros
    out = ('%f' % v).rstrip('0').rstrip('.')
    return out

def detect_columns_from_headers(words, page_width, headers=None):
    # words: list of dicts from pdfplumber.extract_words()
    # returns (boundaries, header_index_map) where header_index_map maps header->col_index
    if headers is None:
        headers = HEADERS
    header_positions = {}
    for h in headers:
        # match header token anywhere in word text (upper)
        matches = [w for w in words if w["text"].strip().upper() == h.upper()]
        if matches:
            header_positions[h.upper()] = matches[0]["x0"]
    # fallback: approximate equal-width columns if not enough headers found
    if len(header_positions) < 2:
        cols = []
        for i in range(6):
            cols.append(i * page_width / 5.0)
        boundaries = cols
        header_index_map = {}
        return boundaries, header_index_map
    # build sorted x positions for headers
    items = sorted(header_positions.items(), key=lambda kv: kv[1])
    xs = [pos for _, pos in items]
    left = 0.0
    right = page_width
    boundaries = [left]
    for i in range(len(xs)-1):
        boundaries.append((xs[i] + xs[i+1]) / 2.0)
    boundaries.append(right)
    # map header token -> column index (based on order)
    header_index_map = {}
    for idx, (h, _) in enumerate(items):
        header_index_map[h] = idx
    return boundaries, header_index_map

def assign_word_to_col(x, boundaries):
    # boundaries are edges [x0, x1, x2, ...] length = ncols+1
    for i in range(len(boundaries)-1):
        if boundaries[i] <= x < boundaries[i+1]:
            return i
    return len(boundaries)-2

def parse_pdf(path, bank=None):
    """Generic parsing pipeline with support for bank-specific header tokens and column mapping.
    If a bank is supplied, choose a header token set most likely matching that bank's statement.
    """
    records = []
    # Choose header tokens per bank
    bank_headers = None
    if bank:
        bk = bank.upper()
        if bk == "BNI":
            bank_headers = ["POSTING DATE", "EFFECTIVE DATE", "BRANCH", "JOURNAL", "TRANSACTION DESCRIPTION", "AMOUNT", "DB/CR", "BALANCE"]
        elif bk == "BNI_USD":
            bank_headers = ["POSTING DATE", "EFFECTIVE DATE", "BRANCH", "JOURNAL", "TRANSACTION DESCRIPTION", "AMOUNT", "DB/CR", "BALANCE"]
        elif bk == "MCM":
            bank_headers = ["POSTING DATE", "REMARK", "REFERENCE NO", "DEBIT", "CREDIT", "BALANCE"]
        elif bk == "BRI":
            bank_headers = ["POSTING DATE", "TIME", "REMARK", "DEBET", "CREDIT", "TELLER ID"]
    if bank_headers is None:
        bank_headers = HEADERS

    with pdfplumber.open(path) as pdf:
        pages_words = [(page, page.extract_words()) for page in pdf.pages]

        # Build header candidates from top lines to detect repeated page headers
        header_counts = {}
        for page, words in pages_words:
            if not words:
                continue
            lines = {}
            for w in words:
                y = round(w["top"])
                lines.setdefault(y, []).append(w)
            top_ys = sorted(lines.keys())[:6]
            for y in top_ys:
                row_words = sorted(lines[y], key=lambda w: w["x0"])
                text = " ".join(w["text"] for w in row_words)
                text_norm = normalize_whitespace(text).upper()
                if not text_norm:
                    continue
                header_counts[text_norm] = header_counts.get(text_norm, 0) + 1

        header_texts = {t for t, c in header_counts.items() if c >= 2}

        # Add bank-specific skip keywords
        known_skip_keywords = ["TANGGAL", "SALDO AWAL", "CATATAN", "NO. REKENING", "REKENING GIRO"]
        if bank:
            bk = bank.upper()
            if bk in ("BNI", "BNI_USD"):
                known_skip_keywords += ["ACCOUNT STATEMENT", "ACCOUNT NO", "ACCOUNT NUMBER", "POSTING DATE", "AMOUNT", "BALANCE"]
            elif bk == "MCM":
                known_skip_keywords += ["REMARK", "DEBIT", "CREDIT", "BALANCE", "POSTING DATE"]
            elif bk == "BRI":
                known_skip_keywords += ["ACCOUNT STATEMENT", "OPENING BALANCE", "CLOSING BALANCE", "DEBET", "CREDIT"]

        col_boundaries = None
        header_map = {}
        prev_record = None
        for page, words in pages_words:
            if not words:
                continue
            if col_boundaries is None:
                col_boundaries, header_map = detect_columns_from_headers(words, page.width, headers=bank_headers)
            lines = {}
            for w in words:
                y = round(w["top"])
                lines.setdefault(y, []).append(w)
            for y in sorted(lines.keys()):
                row_words = sorted(lines[y], key=lambda w: w["x0"])
                cols = [""] * (len(col_boundaries)-1)
                for w in row_words:
                    col_idx = assign_word_to_col(w["x0"], col_boundaries)
                    if cols[col_idx]:
                        cols[col_idx] += " " + w["text"]
                    else:
                        cols[col_idx] = w["text"]
                cols = [normalize_whitespace(c) for c in cols]
                upper_line = " ".join(cols).upper()

                # skip known footers/headers
                if any(h in upper_line for h in known_skip_keywords):
                    continue
                # skip repeated top-of-page headers detected earlier
                skip_flag = False
                for ht in header_texts:
                    if not ht:
                        continue
                    if ht in upper_line or upper_line in ht:
                        skip_flag = True
                        break
                if skip_flag:
                    continue

                # heuristics: find date, amount, balance, description columns using header_map
                date_col = 0
                desc_col = 1
                amount_col = len(cols)-2 if len(cols) >= 2 else len(cols)-1
                balance_col = len(cols)-1
                if header_map:
                    keys = {k.upper(): v for k, v in header_map.items()}
                    for candidate in ("TRANSACTION DESCRIPTION", "REMARK", "REMARKS", "REMARKS/DETAILS", "KETERANGAN"):
                        if candidate in keys:
                            desc_col = keys[candidate]
                            break
                    for candidate in ("POSTING DATE", "TANGGAL", "TIME", "EFFECTIVE DATE"):
                        if candidate in keys:
                            date_col = keys[candidate]
                            break
                    for candidate in ("AMOUNT", "DEBIT", "CREDIT", "MUTASI"):
                        if candidate in keys:
                            amount_col = keys[candidate]
                            break
                    for candidate in ("BALANCE", "SALDO"):
                        if candidate in keys:
                            balance_col = keys[candidate]
                            break

                def safe_get(cols, idx):
                    return cols[idx] if 0 <= idx < len(cols) else ""

                tanggal = safe_get(cols, date_col)
                keterangan = safe_get(cols, desc_col)
                # For robustness, detect numeric tokens in the whole row and pick last tokens as amounts
                row_text = " ".join(cols)
                num_tokens = re.findall(r'[-+]?(?:\d{1,3}(?:[.,]\d{3})+|\d+)(?:[.,]\d+)?', row_text)
                mutasi = safe_get(cols, amount_col)
                saldo = safe_get(cols, balance_col)
                if num_tokens:
                    if len(num_tokens) >= 1:
                        saldo = normalize_numeric_token(num_tokens[-1])
                    if len(num_tokens) >= 2:
                        mutasi = normalize_numeric_token(num_tokens[-2])
                    else:
                        mutasi = ""
                # clean description from noisy numeric/journal tokens
                keterangan = clean_description(keterangan)

                if (not tanggal) and (not clean_number(mutasi)) and (not clean_number(saldo)):
                    if prev_record is not None:
                        prev_record["KETERANGAN"] = (prev_record["KETERANGAN"] + " " + keterangan).strip()
                    continue

                rec = {
                    "TANGGAL": tanggal,
                    "KETERANGAN": keterangan,
                    "CBG": safe_get(cols, 2) if len(cols) > 2 else "",
                    "MUTASI": clean_number(mutasi),
                    "SALDO": clean_number(saldo),
                }
                for k in rec:
                    rec[k] = normalize_whitespace(rec[k])
                records.append(rec)
                prev_record = records[-1]
    return records


def run_conversion(input_path, output_path):
    inp = Path(input_path)
    out = Path(output_path)
    if not inp.exists():
        raise FileNotFoundError(f"Input file not found: {inp}")
    bank = detect_bank(str(inp))
    recs = parse_pdf(str(inp), bank=bank)
    if not recs:
        raise RuntimeError("No records parsed.")
    df = pd.DataFrame(recs, columns=["TANGGAL", "KETERANGAN", "CBG", "MUTASI", "SALDO"])

    # Clean amounts and add numeric columns
    df_clean = df.copy()
    for col in ("MUTASI", "SALDO"):
        df_clean[col] = df_clean[col].fillna("").astype(str).apply(format_amount_str)

    def to_number(x):
        s = clean_number(str(x))
        s = s.replace(',', '')
        if s == "":
            return float('nan')
        try:
            return float(s)
        except Exception:
            return float('nan')

    df_clean['MUTASI_NUM'] = df_clean['MUTASI'].apply(to_number)
    df_clean['SALDO_NUM'] = df_clean['SALDO'].apply(to_number)

    # Normalize date to ISO where possible
    df_clean['DATE_ISO'] = df_clean['TANGGAL'].apply(normalize_date_str)

    # Preserve currency from bank detection
    currency = 'USD' if bank == 'BNI_USD' else 'IDR'
    df_clean['CURRENCY'] = currency

    # Create Debit/Credit columns heuristically
    df_clean['DEBIT'] = df_clean['MUTASI_NUM'].where(df_clean['MUTASI_NUM'] > 0)
    df_clean['CREDIT'] = df_clean['MUTASI_NUM'].where(df_clean['MUTASI_NUM'] < 0).abs()

    # Bank-specific post-processing hints
    if bank == 'BNI_USD':
        # For USD statements, keep two decimal places
        df_clean['MUTASI'] = df_clean['MUTASI_NUM'].map(lambda v: f"{v:.2f}" if pd.notna(v) else "")
        df_clean['SALDO'] = df_clean['SALDO_NUM'].map(lambda v: f"{v:.2f}" if pd.notna(v) else "")

    # Write output
    if out.suffix.lower() in [".xls", ".xlsx"]:
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Raw", index=False)
            df_clean.to_excel(writer, sheet_name="Clean", index=False)
    else:
        df_clean.to_csv(out, index=False, encoding="utf-8-sig")
    return len(df)


def gui_main():
    root = tk.Tk()
    root.title("PDF to Excel Converter (BCA)")
    root.geometry("760x200")

    inp_var = tk.StringVar()
    out_var = tk.StringVar()
    status_var = tk.StringVar()

    # store selected input paths (list of strings)
    inp_list = []

    def set_input_list(lst):
        nonlocal inp_list
        inp_list = [str(Path(p)) for p in lst]
        if not inp_list:
            inp_var.set("")
        elif len(inp_list) == 1:
            inp_var.set(inp_list[0])
        else:
            inp_var.set(f"{len(inp_list)} files selected")

    def browse_files():
        sel = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files","*.pdf")])
        if sel:
            set_input_list(sel)
            status_var.set(f"Selected {len(sel)} files")

    def browse_folder():
        d = filedialog.askdirectory(title="Select folder containing PDFs")
        if d:
            p = Path(d)
            pdfs = sorted([str(x) for x in p.glob("*.pdf")])
            if not pdfs:
                messagebox.showwarning("No PDFs", "No PDF files found in selected folder.")
                return
            set_input_list(pdfs)
            status_var.set(f"Selected folder: {d} ({len(pdfs)} files)")

    def add_folders():
        # allow adding multiple folders by repeated folder selection
        added = 0
        while True:
            d = filedialog.askdirectory(title="Select a folder containing PDFs (Cancel to stop)")
            if not d:
                break
            p = Path(d)
            pdfs = sorted([str(x) for x in p.glob("*.pdf")])
            if not pdfs:
                messagebox.showwarning("No PDFs", f"No PDF files found in selected folder: {d}")
            else:
                # merge into existing list
                new_list = list(inp_list) + pdfs
                # remove duplicates while preserving order
                seen = set()
                merged = []
                for it in new_list:
                    if it not in seen:
                        seen.add(it)
                        merged.append(it)
                set_input_list(merged)
                added += len(pdfs)
                status_var.set(f"Added {len(pdfs)} files from {d} (total {len(inp_list)})")
            # ask if user wants to add another folder
            again = messagebox.askyesno("Add another?", "Add another folder?")
            if not again:
                break

    def clear_selection():
        set_input_list([])
        status_var.set("Selection cleared")

    def browse_out():
        # If multiple inputs selected we ask for a directory, otherwise allow save-as file
        if len(inp_list) > 1:
            d = filedialog.askdirectory(title="Select output folder")
            if d:
                out_var.set(d)
        else:
            p = filedialog.asksaveasfilename(title="Save as", defaultextension=".xlsx",
                                             filetypes=[("Excel","*.xlsx"), ("CSV","*.csv")])
            if p:
                out_var.set(p)

    def do_convert():
        out_path = out_var.get().strip()
        if not inp_list:
            messagebox.showwarning("Missing", "Please select one or more input PDFs (Files or Folder).")
            return
        if not out_path:
            messagebox.showwarning("Missing", "Please select an output file or folder.")
            return

        # Determine batch or single
        try:
            status_var.set("Processing...")
            root.update_idletasks()
            total = 0
            errors = []
            if len(inp_list) > 1:
                out_dir = Path(out_path)
                if not out_dir.exists() or not out_dir.is_dir():
                    messagebox.showwarning("Output", "For multiple input files please select an output folder.")
                    status_var.set("Cancelled")
                    return
                for i, inp in enumerate(inp_list, 1):
                    try:
                        out_file = out_dir / (Path(inp).stem + ".xlsx")
                        n = run_conversion(inp, str(out_file))
                        total += n
                        status_var.set(f"Processed {i}/{len(inp_list)}: {out_file.name}")
                        root.update_idletasks()
                    except Exception as e:
                        errors.append(f"{inp}: {e}")
                msg = f"Processed {len(inp_list)} files, total records {total}"
                if errors:
                    msg += f"\n{len(errors)} errors."
                    messagebox.showwarning("Completed with errors", msg + "\nSee console for details.")
                else:
                    messagebox.showinfo("Done", msg)
            else:
                inp = inp_list[0]
                out = Path(out_path)
                # if out is directory (user selected folder), build filename from input
                if out.exists() and out.is_dir():
                    out_file = out / (Path(inp).stem + ".xlsx")
                else:
                    out_file = out
                n = run_conversion(inp, str(out_file))
                total += n
                messagebox.showinfo("Done", f"Wrote {n} records to {out_file}")
            status_var.set(f"Done: {total} records")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            status_var.set("Error")

    frm = tk.Frame(root, padx=10, pady=10)
    frm.pack(fill=tk.BOTH, expand=True)

    tk.Label(frm, text="Input (Files or Folder):").grid(row=0, column=0, sticky=tk.W)
    tk.Entry(frm, textvariable=inp_var, width=70).grid(row=0, column=1, padx=6)
    tk.Button(frm, text="Browse files...", command=browse_files).grid(row=0, column=2, padx=4)
    tk.Button(frm, text="Browse folder...", command=browse_folder).grid(row=0, column=3)
    tk.Button(frm, text="Add folder(s)...", command=add_folders).grid(row=0, column=4, padx=4)
    tk.Button(frm, text="Clear", command=clear_selection).grid(row=0, column=5)

    tk.Label(frm, text="Output (file or folder):").grid(row=1, column=0, sticky=tk.W, pady=8)
    tk.Entry(frm, textvariable=out_var, width=70).grid(row=1, column=1, padx=6)
    tk.Button(frm, text="Browse...", command=browse_out).grid(row=1, column=2)

    tk.Button(frm, text="Convert", command=do_convert, width=14).grid(row=2, column=1, pady=12)
    tk.Label(frm, textvariable=status_var).grid(row=3, column=0, columnspan=4, sticky=tk.W)

    root.mainloop()


if __name__ == "__main__":
    def main():
        args = sys.argv[1:]
        if not args:
            gui_main()
            return

        # treat last arg as output (file or directory), others as inputs
        if len(args) < 2:
            print("Usage: python PDF to Excel Converter (BCA).py <input1.pdf> [<input2.pdf> ...] <output.xlsx|out_dir>")
            sys.exit(1)
        *in_args, out_arg = args
        inputs = [Path(p) for p in in_args]
        out = Path(out_arg)

        # if single input and it's a folder -> process all PDFs inside
        if len(inputs) == 1 and inputs[0].is_dir():
            inputs = sorted(list(inputs[0].glob("*.pdf")))

        if not inputs:
            print("No input files specified.")
            sys.exit(1)

        # decide whether out is directory or a single output file
        out_is_dir = False
        if out.exists() and out.is_dir():
            out_is_dir = True
            out_dir = out
        else:
            # if multiple inputs, require out be a directory
            if len(inputs) > 1:
                # try to create directory
                try:
                    out.mkdir(parents=True, exist_ok=True)
                    out_is_dir = True
                    out_dir = out
                except Exception:
                    print("When specifying multiple input files, the last argument must be an output directory.")
                    sys.exit(1)
            else:
                out_is_dir = False
                out_dir = out.parent

        total = 0
        for inp in inputs:
            if not Path(inp).exists():
                print("Input file not found:", inp)
                continue
            if out_is_dir:
                out_file = Path(out_dir) / (Path(inp).stem + ".xlsx")
            else:
                out_file = out
            try:
                n = run_conversion(str(inp), str(out_file))
                print(f"Wrote {n} records to {out_file}")
                total += n
            except Exception as e:
                print(f"Error processing {inp}: {e}")
        print(f"Processed {len(inputs)} files, total records {total}")

    main()