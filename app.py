import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json


class DataReconTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Recon Tool")
        self.root.geometry("960x750")
        self.root.minsize(800, 600)

        self.df1 = None
        self.df2 = None
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.mapping_rows = []  # [(col1_name, tk.StringVar)]

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        self._build_file_panel()
        self._build_mapping_panel()
        self._build_results_panel()
        self._build_status_bar()

    def _build_file_panel(self):
        frame = ttk.LabelFrame(self.root, text="Files", padding=10)
        frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        frame.columnconfigure(1, weight=1)

        for row, (label, var, cmd) in enumerate([
            ("File 1:", self.file1_path, lambda: self._browse(self.file1_path)),
            ("File 2:", self.file2_path, lambda: self._browse(self.file2_path)),
        ]):
            ttk.Label(frame, text=label, width=7).grid(row=row, column=0, sticky="w")
            ttk.Entry(frame, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6, pady=2)
            ttk.Button(frame, text="Browse…", command=cmd).grid(row=row, column=2)

        btn_row = ttk.Frame(frame)
        btn_row.grid(row=2, column=0, columnspan=3, pady=(8, 0))
        ttk.Button(btn_row, text="Load Files & Map Columns", command=self._load_files).pack()

    def _build_mapping_panel(self):
        outer = ttk.LabelFrame(self.root, text="Column Mapping", padding=10)
        outer.grid(row=1, column=0, sticky="ew", padx=10, pady=4)
        outer.columnconfigure(0, weight=1)

        # Header labels
        hdr = ttk.Frame(outer)
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(hdr, text="File 1 Column", width=30, font=("", 9, "bold")).pack(side="left", padx=(4, 0))
        ttk.Label(hdr, text="File 2 Column (mapped to)", width=34, font=("", 9, "bold")).pack(side="left")

        # Scrollable canvas for mapping rows
        canvas_frame = ttk.Frame(outer)
        canvas_frame.grid(row=1, column=0, sticky="ew")

        self.map_canvas = tk.Canvas(canvas_frame, height=160, highlightthickness=0)
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.map_canvas.yview)
        self.map_canvas.configure(yscrollcommand=vsb.set)
        self.map_canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.map_inner = ttk.Frame(self.map_canvas)
        self.map_canvas_window = self.map_canvas.create_window((0, 0), window=self.map_inner, anchor="nw")
        self.map_inner.bind("<Configure>", self._on_map_resize)
        self.map_canvas.bind("<Configure>", self._on_canvas_resize)

        # Mapping action buttons
        btn_row = ttk.Frame(outer)
        btn_row.grid(row=2, column=0, pady=(6, 0))
        ttk.Button(btn_row, text="Save Mapping…", command=self._save_mapping).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Load Mapping…", command=self._load_mapping).pack(side="left", padx=4)

        # Compare button
        ttk.Button(self.root, text="▶  Compare", command=self._compare).grid(
            row=1, column=0, sticky="e", padx=18, pady=4)

    def _build_results_panel(self):
        frame = ttk.LabelFrame(self.root, text="Results", padding=10)
        frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=4)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = ("ID", "Field (File 1 name)", "File 1 Value", "File 2 Value")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        widths = (140, 180, 250, 250)
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=80)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure("diff", background="#ffe0e0")
        self.tree.tag_configure("missing", background="#fff3cd")

        ttk.Button(frame, text="Export Results…", command=self._export_results).grid(
            row=2, column=0, sticky="e", pady=(6, 0))

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="Load two Excel files to begin.")
        bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        bar.grid(row=3, column=0, sticky="ew", padx=0, pady=0)

    # ------------------------------------------------------------------
    # Canvas / resize helpers
    # ------------------------------------------------------------------

    def _on_map_resize(self, _event):
        self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all"))

    def _on_canvas_resize(self, event):
        self.map_canvas.itemconfig(self.map_canvas_window, width=event.width)

    # ------------------------------------------------------------------
    # File loading
    # ------------------------------------------------------------------

    def _browse(self, var):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            var.set(path)

    def _load_files(self):
        p1, p2 = self.file1_path.get().strip(), self.file2_path.get().strip()
        if not p1 or not p2:
            messagebox.showwarning("Missing Files", "Select both files before loading.")
            return
        try:
            self.df1 = pd.read_excel(p1, dtype=str).fillna("")
            self.df2 = pd.read_excel(p2, dtype=str).fillna("")
        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            return

        self._set_status(
            f"File 1: {len(self.df1)} rows, {len(self.df1.columns)} columns  |  "
            f"File 2: {len(self.df2)} rows, {len(self.df2.columns)} columns"
        )
        self._build_mapping_rows()

    # ------------------------------------------------------------------
    # Column mapping UI
    # ------------------------------------------------------------------

    def _build_mapping_rows(self, existing_map=None):
        for widget in self.map_inner.winfo_children():
            widget.destroy()
        self.mapping_rows.clear()

        cols1 = list(self.df1.columns)
        cols2 = list(self.df2.columns)
        options = ["-- ignore --"] + cols2

        id_col = cols1[0]
        ttk.Label(self.map_inner, text=f"ID column (key): {id_col}",
                  foreground="#555").grid(row=0, column=0, columnspan=2,
                                          sticky="w", padx=6, pady=(0, 4))

        for i, col1 in enumerate(cols1[1:], start=1):
            ttk.Label(self.map_inner, text=col1, width=28, anchor="w").grid(
                row=i, column=0, padx=(6, 2), pady=1, sticky="w")

            var = tk.StringVar()
            if existing_map and col1 in existing_map:
                mapped = existing_map[col1]
                var.set(mapped if mapped in options else "-- ignore --")
            elif col1 in cols2:
                var.set(col1)
            else:
                var.set("-- ignore --")

            cb = ttk.Combobox(self.map_inner, textvariable=var,
                              values=options, width=30, state="readonly")
            cb.grid(row=i, column=1, padx=(2, 6), pady=1)
            self.mapping_rows.append((col1, var))

    def _get_mapping(self):
        """Return {file1_col: file2_col} for all non-ignored rows."""
        return {col1: var.get()
                for col1, var in self.mapping_rows
                if var.get() != "-- ignore --"}

    def _save_mapping(self):
        if not self.mapping_rows:
            messagebox.showwarning("No Mapping", "Load files and build a mapping first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON mapping", "*.json"), ("All files", "*.*")])
        if not path:
            return
        with open(path, "w") as f:
            json.dump(self._get_mapping(), f, indent=2)
        self._set_status(f"Mapping saved to {path}")

    def _load_mapping(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showwarning("No Files", "Load both files before applying a mapping.")
            return
        path = filedialog.askopenfilename(
            filetypes=[("JSON mapping", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path) as f:
                mapping = json.load(f)
        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            return
        self._build_mapping_rows(existing_map=mapping)
        self._set_status(f"Mapping loaded from {path}")

    # ------------------------------------------------------------------
    # Comparison
    # ------------------------------------------------------------------

    def _compare(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showwarning("No Files", "Load both files first.")
            return

        mapping = self._get_mapping()
        if not mapping:
            messagebox.showwarning("No Mapping", "Map at least one column pair before comparing.")
            return

        id1 = self.df1.columns[0]
        id2 = self.df2.columns[0]

        # Deduplicate on ID (keep first occurrence, flag warning)
        dup1 = self.df1[id1].duplicated().any()
        dup2 = self.df2[id2].duplicated().any()
        if dup1 or dup2:
            which = " and ".join(filter(None, [
                "File 1" if dup1 else "",
                "File 2" if dup2 else ""
            ]))
            messagebox.showwarning(
                "Duplicate IDs",
                f"{which} contain duplicate ID values. Only the first occurrence of each ID will be used.")

        df1 = self.df1.drop_duplicates(subset=id1).set_index(id1)
        df2 = self.df2.drop_duplicates(subset=id2).set_index(id2)

        for row in self.tree.get_children():
            self.tree.delete(row)

        ids_only_in_1 = set(df1.index) - set(df2.index)
        ids_only_in_2 = set(df2.index) - set(df1.index)
        common_ids = set(df1.index) & set(df2.index)

        diff_count = 0

        for record_id in sorted(ids_only_in_1, key=str):
            self.tree.insert("", "end",
                             values=(record_id, "(missing in File 2)", "", ""),
                             tags=("missing",))
            diff_count += 1

        for record_id in sorted(ids_only_in_2, key=str):
            self.tree.insert("", "end",
                             values=(record_id, "(missing in File 1)", "", ""),
                             tags=("missing",))
            diff_count += 1

        for record_id in sorted(common_ids, key=str):
            row1 = df1.loc[record_id]
            row2 = df2.loc[record_id]
            for col1, col2 in mapping.items():
                val1 = str(row1.get(col1, "")).strip()
                val2 = str(row2.get(col2, "")).strip()
                if val1 != val2:
                    self.tree.insert("", "end",
                                     values=(record_id, col1, val1, val2),
                                     tags=("diff",))
                    diff_count += 1

        if diff_count == 0:
            self._set_status("Compare complete — no differences found.")
            messagebox.showinfo("Results", "Files match on all mapped fields.")
        else:
            self._set_status(f"Compare complete — {diff_count} difference(s) found.")

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def _export_results(self):
        rows = [self.tree.item(r)["values"] for r in self.tree.get_children()]
        if not rows:
            messagebox.showinfo("Nothing to Export", "Run a comparison first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx"), ("All files", "*.*")])
        if not path:
            return
        cols = ["ID", "Field (File 1 name)", "File 1 Value", "File 2 Value"]
        pd.DataFrame(rows, columns=cols).to_excel(path, index=False)
        self._set_status(f"Results exported to {path}")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _set_status(self, msg):
        self.status_var.set(msg)


if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.25)
    except Exception:
        pass
    DataReconTool(root)
    root.mainloop()
