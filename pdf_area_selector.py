#!/usr/bin/env python
"""PDF Area Selector with zoom, scrollbars, and fit-to-width.

Features:
  - Load PDF (PyMuPDF)
  - Page navigation (Prev / Next)
  - Draw multiple rectangular selection areas (tables) per page
  - Zoom In / Zoom Out (rectangles scale accordingly)
  - Fit page width button
  - Scrollbars for large pages
  - Save / Load selections to JSON
  - Extract only selected areas to a single Excel file (Camelot)

Install deps (Windows cmd):
  pip install PyMuPDF camelot-py pandas openpyxl pillow

Run:
  python pdf_area_selector.py
"""

import os
import re
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


EXCEL_ILLEGAL_CHAR_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
PT_DECIMAL_RE = re.compile(r"^\s*([+-]?)\s*(\d{1,3}(?:[ .\u00A0\u202F]\d{3})*|\d+),(\d+)\s*$")

try:
    import fitz  # PyMuPDF
except ImportError as e:
    raise SystemExit("Missing dependency PyMuPDF. Install with: pip install PyMuPDF") from e

try:
    import camelot
except ImportError as e:
    raise SystemExit("Missing dependency camelot-py. Install with: pip install camelot-py[base]") from e

try:
    import pandas as pd
except ImportError as e:
    raise SystemExit("Missing dependency pandas. Install with: pip install pandas") from e

try:
    from PIL import Image, ImageTk
except ImportError as e:
    raise SystemExit("Missing dependency Pillow. Install with: pip install pillow") from e


class PDFAreaSelector:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Area Selector - Camelot")

        # State
        self.pdf_path = None
        self.doc = None  # fitz.Document
        self.page_index = 0
        self.zoom = 2.0  # rendering scale (1.0 = 72 dpi)
        self.page_photo = None  # keep reference
        self.rect_start = None
        self.current_rect_id = None
        # selections: page_index -> list[(x0,y0,x1,y1)] in current zoom pixel coords
        self.selections = {}

        self._build_ui()
        # Bind after canvas exists
        self._bind_canvas()

    # ---------------- UI ----------------
    def _build_ui(self):
        top = tk.Frame(self.root)
        top.pack(fill=tk.X, padx=6, pady=4)

        self.entry_pdf = tk.Entry(top, width=55)
        self.entry_pdf.pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(top, text="Browse", command=self.browse_pdf).pack(side=tk.LEFT, padx=2)
        tk.Button(top, text="Load", command=self.load_pdf).pack(side=tk.LEFT, padx=2)

        self.flavor_var = tk.StringVar(value="stream")
        ttk.Radiobutton(top, text="Stream", variable=self.flavor_var, value="stream").pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(top, text="Lattice", variable=self.flavor_var, value="lattice").pack(side=tk.LEFT, padx=2)

        tk.Button(top, text="Zoom +", command=self.zoom_in).pack(side=tk.LEFT, padx=2)
        tk.Button(top, text="Zoom -", command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        tk.Button(top, text="Fit", command=self.fit_page_width).pack(side=tk.LEFT, padx=2)

        tk.Button(top, text="Prev", command=self.prev_page).pack(side=tk.LEFT, padx=6)
        tk.Button(top, text="Next", command=self.next_page).pack(side=tk.LEFT, padx=2)
        tk.Button(top, text="Extract", command=self.extract).pack(side=tk.LEFT, padx=8)
        tk.Button(top, text="Quit", command=self.root.quit).pack(side=tk.RIGHT, padx=2)

        body = tk.Frame(self.root)
        body.pack(fill=tk.BOTH, expand=True)

        # Canvas + scrollbars
        canvas_frame = tk.Frame(body)
        canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.v_scroll = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas = tk.Canvas(
            canvas_frame,
            bg="white",
            cursor="tcross",
            xscrollcommand=self.h_scroll.set,
            yscrollcommand=self.v_scroll.set,
        )
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.v_scroll.config(command=self.canvas.yview)
        self.h_scroll.config(command=self.canvas.xview)

        # Side panel
        side = tk.Frame(body)
        side.pack(side=tk.RIGHT, fill=tk.Y, padx=6)
        tk.Label(side, text="Selections (current page)").pack(anchor="w")
        self.listbox = tk.Listbox(side, width=35, height=18)
        self.listbox.pack(fill=tk.Y, pady=2)
        tk.Button(side, text="Remove Selected", command=self.remove_selected_rect).pack(fill=tk.X, pady=2)
        tk.Button(side, text="Clear Page", command=self.clear_page_rects).pack(fill=tk.X, pady=2)
        tk.Button(side, text="Save Areas JSON", command=self.save_selections_json).pack(fill=tk.X, pady=4)
        tk.Button(side, text="Load Areas JSON", command=self.load_selections_json).pack(fill=tk.X, pady=2)
        self.page_label = tk.Label(side, text="Page: -/-")
        self.page_label.pack(pady=6)
        self.status_var = tk.StringVar(value="Idle")
        tk.Label(side, textvariable=self.status_var, wraplength=200, justify=tk.LEFT, fg="#333").pack(anchor="w", pady=6)

    def _bind_canvas(self):
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)

    # ---------------- PDF LOADING ----------------
    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[["PDF files", "*.pdf"], ["All files", "*"]])
        if path:
            self.entry_pdf.delete(0, tk.END)
            self.entry_pdf.insert(0, path)

    def load_pdf(self):
        path = self.entry_pdf.get().strip()
        if not path:
            messagebox.showwarning("No path", "Please specify a PDF path")
            return
        if not os.path.isfile(path):
            messagebox.showerror("Not found", f"File not found:\n{path}")
            return
        try:
            self.doc = fitz.open(path)
            self.pdf_path = path
            self.page_index = 0
            self.selections.setdefault(self.page_index, [])
            self.update_page()
            self.status("Loaded PDF")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")

    # ---------------- PAGE RENDERING ----------------
    def update_page(self):
        if not self.doc:
            return
        page = self.doc[self.page_index]
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        mode = "RGB" if pix.n < 4 else "RGBA"
        img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
        self.page_photo = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.canvas.create_image(0, 0, image=self.page_photo, anchor=tk.NW)
        self.redraw_rectangles()
        self.refresh_listbox()
        self.page_label.config(text=f"Page: {self.page_index + 1}/{len(self.doc)}")
        self.status(f"Rendered page {self.page_index+1} @ {int(self.zoom*72)} dpi (zoom {self.zoom:.2f}x)")

    def next_page(self):
        if self.doc and self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self.selections.setdefault(self.page_index, [])
            self.update_page()

    def prev_page(self):
        if self.doc and self.page_index > 0:
            self.page_index -= 1
            self.update_page()

    # ---------------- RECTANGLES ----------------
    def on_canvas_press(self, event):
        if not self.doc:
            return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        self.rect_start = (cx, cy)
        self.current_rect_id = None

    def on_canvas_drag(self, event):
        if not self.rect_start:
            return
        x0, y0 = self.rect_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        if self.current_rect_id:
            self.canvas.delete(self.current_rect_id)
        self.current_rect_id = self.canvas.create_rectangle(
            x0, y0, x1, y1, outline="lime", width=2, dash=(3, 2)
        )

    def on_canvas_release(self, event):
        if not self.rect_start:
            return
        x0, y0 = self.rect_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.rect_start = None
        x_min, x_max = sorted([x0, x1])
        y_min, y_max = sorted([y0, y1])
        if abs(x_max - x_min) < 5 or abs(y_max - y_min) < 5:
            if self.current_rect_id:
                self.canvas.delete(self.current_rect_id)
            self.current_rect_id = None
            return
        if self.current_rect_id:
            self.canvas.delete(self.current_rect_id)
        self.canvas.create_rectangle(x_min, y_min, x_max, y_max, outline="red", width=2)
        self.selections.setdefault(self.page_index, []).append((x_min, y_min, x_max, y_max))
        self.refresh_listbox()
        self.status(f"Added rect {x_min},{y_min},{x_max},{y_max}")
        self.current_rect_id = None

    def redraw_rectangles(self):
        for (x0, y0, x1, y1) in self.selections.get(self.page_index, []):
            self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", width=2)

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for i, (x0, y0, x1, y1) in enumerate(self.selections.get(self.page_index, [])):
            self.listbox.insert(tk.END, f"#{i+1} ({int(x0)},{int(y0)}) -> ({int(x1)},{int(y1)})")

    def remove_selected_rect(self):
        sel = self.listbox.curselection()
        if not sel:
            self.status("No selection chosen")
            return
        idx = sel[0]
        rects = self.selections.get(self.page_index, [])
        if idx < len(rects):
            removed = rects.pop(idx)
            self.update_page()
            self.status(f"Removed rectangle {removed}")

    def clear_page_rects(self):
        if self.page_index in self.selections:
            self.selections[self.page_index] = []
            self.update_page()
            self.status("Cleared rectangles for page")

    # ---------------- SAVE / LOAD SELECTIONS ----------------
    def save_selections_json(self):
        if not any(self.selections.values()):
            messagebox.showinfo("Nothing", "No selections to save")
            return
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[["JSON", "*.json"]])
        if not path:
            return
        data = {str(k): v for k, v in self.selections.items() if v}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        self.status(f"Saved areas -> {path}")

    def load_selections_json(self):
        if not self.doc:
            messagebox.showwarning("Load PDF", "Load a PDF first")
            return
        path = filedialog.askopenfilename(filetypes=[["JSON", "*.json"], ["All", "*"]])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in data.items():
                self.selections[int(k)] = [tuple(r) for r in v]
            self.update_page()
            self.status("Loaded saved areas")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load JSON: {e}")

    # ---------------- EXTRACTION ----------------
    def extract(self):
        if not self.doc:
            messagebox.showwarning("No PDF", "Load a PDF first")
            return
        if not any(v for v in self.selections.values()):
            messagebox.showinfo("No selections", "No rectangles selected")
            return
        flavor = self.flavor_var.get()
        default_out = os.path.splitext(self.pdf_path)[0] + ".xlsx"
        out_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=os.path.basename(default_out),
            filetypes=[["Excel", "*.xlsx"]],
        )
        if not out_path:
            return
        self.status("Starting extraction ...")
        threading.Thread(target=self._do_extract, args=(flavor, out_path), daemon=True).start()

    def _do_extract(self, flavor: str, out_path: str):
        try:
            frames = []
            for p_idx, rects in sorted(self.selections.items(), key=lambda kv: kv[0]):
                if not rects:
                    continue
                page = self.doc[p_idx]
                mat = fitz.Matrix(self.zoom, self.zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                width_px, height_px = pix.width, pix.height
                for (x0, y0, x1, y1) in rects:
                    area = self._rect_to_camelot_area(x0, y0, x1, y1, width_px, height_px, page)
                    try:
                        tables = camelot.read_pdf(
                            self.pdf_path,
                            pages=str(p_idx + 1),
                            flavor=flavor,
                            table_areas=[area],
                            split_text=True,
                        )
                    except Exception as e:
                        self.status(f"Error p{p_idx+1} area {area}: {e}")
                        continue
                    if len(tables) == 0:
                        self.status(f"No table p{p_idx+1} area {area}")
                        continue
                    best = max(tables, key=lambda t: t.df.shape[0] * t.df.shape[1])
                    df = best.df.copy()
                    df.insert(0, "Page", p_idx + 1)
                    frames.append(df)
                    self.status(f"Added p{p_idx+1} area {area} rows={df.shape[0]}")
            if not frames:
                messagebox.showinfo("Result", "No tables extracted.")
                self.status("No tables extracted.")
                return
            combined = pd.concat(frames, ignore_index=True)
            combined = combined.map(self._sanitize_excel_value)
            combined.to_excel(out_path, index=False, header=False)
            self.status(f"Saved {len(combined)} rows -> {out_path}")
            messagebox.showinfo("Done", f"Extraction complete. Rows: {len(combined)}\nSaved: {out_path}")
        except Exception as e:
            self.status(f"Extraction error: {e}")
            messagebox.showerror("Error", f"Extraction failed: {e}")

    def _rect_to_camelot_area(self, x0, y0, x1, y1, width_px, height_px, page):
        """Convert canvas rectangle (top-left origin) to Camelot area string.

        Camelot expects coordinates with origin at bottom-left: x1,y1,x2,y2 where
        (x1,y1) is lower-left and (x2,y2) upper-right. We captured rectangle in
        top-left origin pixel space at the current zoom. Add a small padding so
        bottom rows aren't clipped.
        """
        pdf_rect = page.rect  # PyMuPDF uses top-left origin
        pdf_w = pdf_rect.width
        pdf_h = pdf_rect.height
        scale_x = pdf_w / width_px
        scale_y = pdf_h / height_px
        # rectangle in pdf top-origin coordinates
        left = x0 * scale_x
        right = x1 * scale_x
        top = y0 * scale_y
        bottom = y1 * scale_y
        # padding (points)
        pad = 4
        top = max(0, top - pad)
        bottom = min(pdf_h, bottom + pad)
        # convert to bottom-left origin:
        lower_y = pdf_h - bottom  # y1 (lower)
        upper_y = pdf_h - top     # y2 (upper)
        # clamp
        lower_y = max(0, min(pdf_h, lower_y))
        upper_y = max(0, min(pdf_h, upper_y))
        return f"{left:.2f},{lower_y:.2f},{right:.2f},{upper_y:.2f}"

    def _sanitize_excel_value(self, value):
        if not isinstance(value, str):
            return value
        cleaned = EXCEL_ILLEGAL_CHAR_RE.sub("", value).strip()
        match = PT_DECIMAL_RE.match(cleaned)
        if not match:
            return cleaned

        sign, integer_part, decimal_part = match.groups()
        integer_part = re.sub(r"[ .\u00A0\u202F]", "", integer_part)
        return float(f"{sign}{integer_part}.{decimal_part}")

    # ---------------- ZOOM ----------------
    def zoom_in(self):
        if self.zoom >= 4.0:
            self.status("Max zoom reached")
            return
        old = self.zoom
        self.zoom *= 1.25
        self.rescale_rects(old)
        self.update_page()

    def zoom_out(self):
        if self.zoom <= 0.5:
            self.status("Min zoom reached")
            return
        old = self.zoom
        self.zoom /= 1.25
        self.rescale_rects(old)
        self.update_page()

    def fit_page_width(self):
        if not self.doc:
            return
        self.root.update_idletasks()
        total_w = self.root.winfo_width() or 1000
        side_panel_w = 340  # approx side panel width
        avail = max(200, total_w - side_panel_w)
        page_w_pts = self.doc[self.page_index].rect.width
        if page_w_pts > 0:
            old = self.zoom
            self.zoom = max(0.4, min(5.0, avail / page_w_pts))
            self.rescale_rects(old)
            self.update_page()
            self.status(f"Fit width -> zoom {self.zoom:.2f}x")

    def rescale_rects(self, old_zoom):
        if old_zoom == 0 or not self.selections:
            return
        factor = self.zoom / old_zoom
        for p, rects in list(self.selections.items()):
            self.selections[p] = [
                (x0 * factor, y0 * factor, x1 * factor, y1 * factor) for (x0, y0, x1, y1) in rects
            ]

    # ---------------- UTIL ----------------
    def status(self, msg: str):
        self.status_var.set(msg)
        print(msg)


def main():
    root = tk.Tk()
    app = PDFAreaSelector(root)
    root.mainloop()


if __name__ == "__main__":
    main()

    def on_canvas_release(self, event):
        if not self.rect_start:
            return
        x0, y0 = self.rect_start
        x1 = self.canvas.canvasx(event.x)
        y1 = self.canvas.canvasy(event.y)
        self.rect_start = None
        # Normalize
        x_min, x_max = sorted([x0, x1])
        y_min, y_max = sorted([y0, y1])
        if abs(x_max - x_min) < 5 or abs(y_max - y_min) < 5:
            # too small
            if self.current_rect_id:
                self.canvas.delete(self.current_rect_id)
            self.current_rect_id = None
            return
        # finalize drawn rectangle (replace preview)
        if self.current_rect_id:
            self.canvas.delete(self.current_rect_id)
        rect_id = self.canvas.create_rectangle(x_min, y_min, x_max, y_max, outline='red', width=2)
        # store selection
        self.selections.setdefault(self.page_index, []).append((x_min, y_min, x_max, y_max))
        self.refresh_listbox()
        self.status(f"Added rect {x_min},{y_min},{x_max},{y_max}")
        self.current_rect_id = None

    def redraw_rectangles(self):
        rects = self.selections.get(self.page_index, [])
        for (x0,y0,x1,y1) in rects:
            self.canvas.create_rectangle(x0,y0,x1,y1, outline='red', width=2)

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for i, (x0,y0,x1,y1) in enumerate(self.selections.get(self.page_index, [])):
            self.listbox.insert(tk.END, f"#{i+1} ({x0},{y0}) -> ({x1},{y1})")

    def remove_selected_rect(self):
        sel = self.listbox.curselection()
        if not sel:
            self.status("No selection chosen in list")
            return
        idx = sel[0]
        rects = self.selections.get(self.page_index, [])
        if idx < len(rects):
            removed = rects.pop(idx)
            if not rects:
                # optionally remove key
                pass
            self.update_page()
            self.status(f"Removed rectangle {removed}")

    def clear_page_rects(self):
        if self.page_index in self.selections:
            self.selections[self.page_index] = []
            self.update_page()
            self.status("Cleared rectangles for page")

    # ---------------- SAVE / LOAD AREAS ----------------
    def save_selections_json(self):
        if not self.selections:
            messagebox.showinfo("Nothing", "No selections to save")
            return
        save_path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[["JSON","*.json"]])
        if not save_path:
            return
        # store as dict of page->list of rects
        data = {str(k): v for k,v in self.selections.items() if v}
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        self.status(f"Saved areas to {save_path}")

    def load_selections_json(self):
        if not self.doc:
            messagebox.showwarning("Load PDF first", "Load a PDF before loading areas")
            return
        path = filedialog.askopenfilename(filetypes=[["JSON","*.json"],["All","*"]])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for k, v in data.items():
                p = int(k)
                self.selections[p] = [tuple(r) for r in v]
            self.update_page()
            self.status("Loaded saved areas")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load JSON: {e}")

    # ---------------- EXTRACTION ----------------
    def extract(self):
        if not self.doc:
            messagebox.showwarning("No PDF", "Load a PDF first")
            return
        if not any(v for v in self.selections.values()):
            messagebox.showinfo("No selections", "No rectangles selected")
            return
        flavor = self.flavor_var.get()
        out_default = os.path.splitext(self.pdf_path)[0] + '.xlsx'
        out_path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=os.path.basename(out_default), filetypes=[["Excel","*.xlsx"]])
        if not out_path:
            return
        self.status("Starting extraction ...")
        thread = threading.Thread(target=self._do_extract, args=(flavor, out_path), daemon=True)
        thread.start()

    def _do_extract(self, flavor, out_path):
        try:
            frames = []
            for p_idx, rects in sorted(self.selections.items(), key=lambda kv: kv[0]):
                if not rects:
                    continue
                page = self.doc[p_idx]
                # Render once to get pixel dimensions used for scaling
                mat = fitz.Matrix(self.zoom, self.zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                width_px, height_px = pix.width, pix.height
                for r in rects:
                    x0,y0,x1,y1 = r
                    area = self._rect_to_camelot_area(x0,y0,x1,y1,width_px,height_px,page)
                    try:
                        tables = camelot.read_pdf(self.pdf_path, pages=str(p_idx+1), flavor=flavor, table_areas=[area], split_text=True)
                    except Exception as e:
                        self.status(f"Error p{p_idx+1} area {area}: {e}")
                        continue
                    if len(tables)==0:
                        self.status(f"No table p{p_idx+1} area {area}")
                        continue
                    # pick largest table
                    best = max(tables, key=lambda t: t.df.shape[0]*t.df.shape[1])
                    df_best = best.df.copy()
                    df_best.insert(0, 'Page', p_idx+1)
                    frames.append(df_best)
                    self.status(f"Added p{p_idx+1} area {area} rows={df_best.shape[0]}")
            if not frames:
                self.status("No tables extracted.")
                messagebox.showinfo("Result", "No tables extracted.")
                return
            combined = pd.concat(frames, ignore_index=True)
            combined.to_excel(out_path, index=False, header=False)
            self.status(f"Saved {len(combined)} rows -> {out_path}")
            messagebox.showinfo("Done", f"Extraction complete. Rows: {len(combined)}\nSaved: {out_path}")
        except Exception as e:
            self.status(f"Extraction error: {e}")
            messagebox.showerror("Error", f"Extraction failed: {e}")

    def _rect_to_camelot_area(self, x0,y0,x1,y1,width_px,height_px,page):
        # Camelot expects l,t,r,b with origin at top-left in PDF coordinate system (points)
        pdf_rect = page.rect
        scale_x = pdf_rect.width / width_px
        scale_y = pdf_rect.height / height_px
        l = x0 * scale_x
        r = x1 * scale_x
        t = y0 * scale_y
        b = y1 * scale_y
        # Clip
        t = max(0, min(pdf_rect.height, t))
        b = max(0, min(pdf_rect.height, b))
        return f"{l:.2f},{t:.2f},{r:.2f},{b:.2f}"

    # ---------------- UTIL ----------------
    def status(self, msg):
        self.status_var.set(msg)
        # Also print to console for logging
        print(msg)

    # ---------------- ZOOM ----------------
    def zoom_in(self):
        if self.zoom >= 4.0:
            self.status("Max zoom reached")
            return
        old = self.zoom
        self.zoom *= 1.25
        self.rescale_rects(old)
        self.update_page()

    def zoom_out(self):
        if self.zoom <= 0.5:
            self.status("Min zoom reached")
            return
        old = self.zoom
        self.zoom /= 1.25
        self.rescale_rects(old)
        self.update_page()

    def rescale_rects(self, old_zoom):
        if not self.selections or old_zoom == 0:
            return
        factor = self.zoom / old_zoom
        for p_idx, rects in list(self.selections.items()):
            new_rects = []
            for (x0,y0,x1,y1) in rects:
                new_rects.append((x0*factor, y0*factor, x1*factor, y1*factor))
            self.selections[p_idx] = new_rects

    def fit_page(self):
        if not self.doc:
            return
        # Determine available width in canvas frame (approx root width minus side panel ~ 320px)
        self.root.update_idletasks()
        total_w = self.root.winfo_width()
        side_w = 320  # approximate side panel width
        avail = max(200, total_w - side_w)
        page = self.doc[self.page_index]
        page_w = page.rect.width  # points (72 dpi)
        # current zoom gives page_w * zoom * (1 point ~= 1 pixel at 72 dpi)
        # We want: page_w * zoom = avail  -> zoom = avail / page_w
        if page_w > 0:
            old = self.zoom
            self.zoom = max(0.4, min(5.0, avail / page_w))
            self.rescale_rects(old)
            self.update_page()
            self.status(f"Fit page width -> zoom {self.zoom:.2f}x")


def main():
    root = tk.Tk()
    app = PDFAreaSelector(root)
    root.mainloop()

if __name__ == '__main__':
    main()
