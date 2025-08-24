import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
import sys
import subprocess
import traceback
from pymkv import MKVFile
import threading

class Pymkv2MergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch MKV Merger - J0nan")
        self.root.geometry("750x360")

        self.mkvtoolnix_path = tk.StringVar()
        self.folder1_path = tk.StringVar()
        self.folder2_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.track_selections = {}

        self.metadata_title = tk.StringVar()

        self.include_chapters_file1 = tk.BooleanVar(value=False)
        self.include_chapters_file2 = tk.BooleanVar(value=False)
        self.include_global_tags_file1 = tk.BooleanVar(value=False)
        self.include_global_tags_file2 = tk.BooleanVar(value=False)
        self.include_attachments_file1 = tk.BooleanVar(value=False)
        self.include_attachments_file2 = tk.BooleanVar(value=False)

        self.file_jsons = {1: None, 2: None}
        self.sample_paths = {1: None, 2: None}

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)

        self.create_path_selection_widgets(main_frame)
        ttk.Button(main_frame, text="Analyze Files & Select Tracks", command=self.setup_track_selection).pack(pady=12)

        self.progress_window = None
        self.merge_progressbar = None
        self.merge_progress_label = None
        self.accept_button = None

        self.start_merge_button = None
        self.merge_thread = None

    def create_path_selection_widgets(self, parent):
        path_frame = ttk.LabelFrame(parent, text="File and Folder Paths")
        path_frame.pack(fill="x", expand=True, padx=5, pady=5)

        ttk.Label(path_frame, text="MKVToolNix Path:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.mkvtoolnix_path, width=68).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse...", command=self.browse_mkvtoolnix).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(path_frame, text="Input Folder 1:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.folder1_path, width=68).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse...", command=lambda: self.browse_folder(self.folder1_path)).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(path_frame, text="Input Folder 2:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.folder2_path, width=68).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse...", command=lambda: self.browse_folder(self.folder2_path)).grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(path_frame, text="Output Folder:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.output_folder_path, width=68).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(path_frame, text="Browse...", command=lambda: self.browse_folder(self.output_folder_path)).grid(row=3, column=2, padx=5, pady=5)

    def browse_mkvtoolnix(self):
        path = filedialog.askdirectory(title="Select MKVToolNix Installation Folder")
        if path:
            self.mkvtoolnix_path.set(path)
            mkvmerge_exe = "mkvmerge.exe" if sys.platform == "win32" else "mkvmerge"
            MKVFile.mkvmerge_path = os.path.join(path, mkvmerge_exe)
            print(f"[INFO] mkvmerge path set to: {MKVFile.mkvmerge_path}")

    def browse_folder(self, path_var):
        path = filedialog.askdirectory()
        if path:
            path_var.set(path)
            print(f"[INFO] Folder path set to: {path}")

    def validate_paths(self):
        if not self.mkvtoolnix_path.get() or not os.path.isdir(self.mkvtoolnix_path.get()):
            messagebox.showerror("Error", "MKVToolNix path is not set or invalid.")
            print(f"[ERROR] MKVToolNix path is not set or invalid")
            return False
        mkvmerge_exe = "mkvmerge.exe" if sys.platform == "win32" else "mkvmerge"
        if not os.path.exists(os.path.join(self.mkvtoolnix_path.get(), mkvmerge_exe)):
            messagebox.showerror("Error", f"mkvmerge not found at: {self.mkvtoolnix_path.get()}")
            print(f"[ERROR] mkvmerge not found at: {self.mkvtoolnix_path.get()}")
            return False
        if not self.folder1_path.get() or not self.folder2_path.get() or not self.output_folder_path.get():
            messagebox.showerror("Error", "All input and output folders must be selected.")
            print(f"[ERROR] All input and output folders must be selected")
            return False
        return True

    def find_matching_files(self):
        try:
            files1 = set(os.listdir(self.folder1_path.get()))
            files2 = set(os.listdir(self.folder2_path.get()))
            return sorted(list(files1.intersection(files2)))
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"Folder not found: {e.filename}")
            print(f"[ERROR] Folder not found: {e.filename}")
            return []

    def setup_track_selection(self):
        if not self.validate_paths(): return
        matching_files = self.find_matching_files()
        if not matching_files:
            messagebox.showinfo("Information", "No matching files found.")
            print(f"[INFO] No matching files found")
            return

        sample_filename = matching_files[0]
        path1 = os.path.join(self.folder1_path.get(), sample_filename)
        path2 = os.path.join(self.folder2_path.get(), sample_filename)

        try:
            mkv1 = MKVFile(path1)
            mkv2 = MKVFile(path2)
        except Exception as e:
            messagebox.showerror("Error Analyzing File", f"Could not analyze files.\n\nError: {e}")
            print(f"[ERROR] Could not analyze files")
            traceback.print_exc()
            return

        self.sample_paths[1] = path1
        self.sample_paths[2] = path2
        self.file_jsons[1] = self.parse_mkvmerge_json(path1)
        self.file_jsons[2] = self.parse_mkvmerge_json(path2)

        has_chapters1 = bool(self.file_jsons[1] and self.file_jsons[1].get("chapters"))
        has_chapters2 = bool(self.file_jsons[2] and self.file_jsons[2].get("chapters"))

        # tags detection - mkvmerge may store tags under 'tags' or 'global_tags'
        has_tags1 = bool(self.file_jsons[1] and (self.file_jsons[1].get("tags") or self.file_jsons[1].get("global_tags")))
        has_tags2 = bool(self.file_jsons[2] and (self.file_jsons[2].get("tags") or self.file_jsons[2].get("global_tags")))

        has_attachments1 = bool(self.file_jsons[1] and self.file_jsons[1].get("attachments"))
        has_attachments2 = bool(self.file_jsons[2] and self.file_jsons[2].get("attachments"))

        self.track_window = tk.Toplevel(self.root)
        self.track_window.title("Track Selection and Customization")

        folder1_name = os.path.basename(os.path.normpath(self.folder1_path.get()))
        folder2_name = os.path.basename(os.path.normpath(self.folder2_path.get()))
        self.create_track_widgets(self.track_window, f"File 1: {sample_filename} (from '{folder1_name}')", mkv1, 1)
        self.create_track_widgets(self.track_window, f"File 2: {sample_filename} (from '{folder2_name}')", mkv2, 2)

        self.create_global_properties_widgets(self.track_window, mkv1, sample_filename,
                                             has_chapters1=has_chapters1, has_chapters2=has_chapters2,
                                             has_tags1=has_tags1, has_tags2=has_tags2,
                                             has_attachments1=has_attachments1, has_attachments2=has_attachments2)

        button_frame = ttk.Frame(self.track_window)
        button_frame.pack(pady=10, fill="x")
        ttk.Button(button_frame, text="Save Preset", command=self.save_preset).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Load Preset", command=self.load_preset).pack(side="left", padx=5)

        self.start_merge_button = ttk.Button(button_frame, text="Start Merging", command=self.start_merging)
        self.start_merge_button.pack(side="right", padx=10)

    def create_global_properties_widgets(self, parent, mkv1, filename,
                                         has_chapters1=False, has_chapters2=False,
                                         has_tags1=False, has_tags2=False,
                                         has_attachments1=False, has_attachments2=False):
        frame = ttk.LabelFrame(parent, text="Global Properties")
        frame.pack(padx=10, pady=5, fill="x", expand=True)

        ttk.Label(frame, text="Metadata Title:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        default_title = mkv1.title or os.path.splitext(filename)[0]
        self.metadata_title.set(default_title)
        ttk.Entry(frame, textvariable=self.metadata_title, width=50).grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="we")

        chapter_frame = ttk.Frame(frame)
        chapter_frame.grid(row=1, column=0, columnspan=4, padx=5, pady=(2,4), sticky="w")

        row1 = ttk.Frame(chapter_frame)
        row1.pack(anchor="w", pady=2, fill="x")
        cb1 = ttk.Checkbutton(row1, text="Include chapters from File 1", variable=self.include_chapters_file1)
        cb1.pack(side="left")
        if not has_chapters1:
            cb1.state(["disabled"])
            ttk.Label(row1, text="(No chapters found or mkvmerge unavailable for File 1)", foreground="gray").pack(side="left", padx=8)

        row2 = ttk.Frame(chapter_frame)
        row2.pack(anchor="w", pady=2, fill="x")
        cb2 = ttk.Checkbutton(row2, text="Include chapters from File 2", variable=self.include_chapters_file2)
        cb2.pack(side="left")
        if not has_chapters2:
            cb2.state(["disabled"])
            ttk.Label(row2, text="(No chapters found or mkvmerge unavailable for File 2)", foreground="gray").pack(side="left", padx=8)

        tags_frame = ttk.Frame(frame)
        tags_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=(2,4), sticky="w")

        tags_row1 = ttk.Frame(tags_frame)
        tags_row1.pack(anchor="w", pady=2, fill="x")
        tcb1 = ttk.Checkbutton(tags_row1, text="Include global tags from File 1", variable=self.include_global_tags_file1)
        tcb1.pack(side="left")
        view_tags_btn1 = ttk.Button(tags_row1, text="View tags", command=lambda idx=1: self.show_global_tags_window(idx))
        view_tags_btn1.pack(side="left", padx=8)
        if not has_tags1:
            tcb1.state(["disabled"])
            view_tags_btn1.state(["disabled"])
            ttk.Label(tags_row1, text="(No global tags found or mkvmerge unavailable for File 1)", foreground="gray").pack(side="left", padx=8)

        tags_row2 = ttk.Frame(tags_frame)
        tags_row2.pack(anchor="w", pady=2, fill="x")
        tcb2 = ttk.Checkbutton(tags_row2, text="Include global tags from File 2", variable=self.include_global_tags_file2)
        tcb2.pack(side="left")
        view_tags_btn2 = ttk.Button(tags_row2, text="View tags", command=lambda idx=2: self.show_global_tags_window(idx))
        view_tags_btn2.pack(side="left", padx=8)
        if not has_tags2:
            tcb2.state(["disabled"])
            view_tags_btn2.state(["disabled"])
            ttk.Label(tags_row2, text="(No global tags found or mkvmerge unavailable for File 2)", foreground="gray").pack(side="left", padx=8)

        attach_frame = ttk.Frame(frame)
        attach_frame.grid(row=3, column=0, columnspan=4, padx=5, pady=(2,8), sticky="w")

        attach_row1 = ttk.Frame(attach_frame)
        attach_row1.pack(anchor="w", pady=2, fill="x")
        acb1 = ttk.Checkbutton(attach_row1, text="Include attachments from File 1", variable=self.include_attachments_file1)
        acb1.pack(side="left")
        view_attach_btn1 = ttk.Button(attach_row1, text="View attachments...", command=lambda idx=1: self.show_attachments_window(idx))
        view_attach_btn1.pack(side="left", padx=8)
        if not has_attachments1:
            acb1.state(["disabled"])
            view_attach_btn1.state(["disabled"])
            ttk.Label(attach_row1, text="(No attachments found or mkvmerge unavailable for File 1)", foreground="gray").pack(side="left", padx=8)

        attach_row2 = ttk.Frame(attach_frame)
        attach_row2.pack(anchor="w", pady=2, fill="x")
        acb2 = ttk.Checkbutton(attach_row2, text="Include attachments from File 2", variable=self.include_attachments_file2)
        acb2.pack(side="left")
        view_attach_btn2 = ttk.Button(attach_row2, text="View attachments...", command=lambda idx=2: self.show_attachments_window(idx))
        view_attach_btn2.pack(side="left", padx=8)
        if not has_attachments2:
            acb2.state(["disabled"])
            view_attach_btn2.state(["disabled"])
            ttk.Label(attach_row2, text="(No attachments found or mkvmerge unavailable for File 2)", foreground="gray").pack(side="left", padx=8)

    def parse_mkvmerge_json(self, filepath):
        mkvmerge_exe = "mkvmerge.exe" if sys.platform == "win32" else "mkvmerge"
        mkvmerge_path = os.path.join(self.mkvtoolnix_path.get(), mkvmerge_exe)

        if not os.path.exists(mkvmerge_path):
            print(f"[INFO] mkvmerge not found at: {mkvmerge_path}. JSON parsing disabled for {filepath}")
            return None

        try:
            process = subprocess.run([mkvmerge_path, "-J", filepath], capture_output=True, text=True, check=True)
            out = process.stdout
            if not out:
                print(f"[INFO] mkvmerge returned no output for file: {filepath}")
                return None
            data = json.loads(out)
            return data
        except subprocess.CalledProcessError as cpe:
            print(f"[ERROR] mkvmerge identified error for file {filepath}: returncode={cpe.returncode}")
            if cpe.stderr:
                print(cpe.stderr)
            traceback.print_exc()
            return None
        except Exception as e:
            print(f"[ERROR] Exception while parsing mkvmerge JSON for {filepath}: {e}")
            traceback.print_exc()
            return None

    def show_attachments_window(self, file_index):
        data = self.file_jsons.get(file_index)
        if not data:
            messagebox.showinfo("No data", "No mkvmerge JSON info is available for this file.")
            return
        attachments = data.get("attachments") or []
        if not attachments:
            messagebox.showinfo("No attachments", "No attachments found for this file.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Attachments - File {file_index}")
        win.geometry("700x300")

        cols = ("id", "file_name", "mime_type", "size", "description")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c.replace("_", " ").title())
            tree.column(c, width=120, anchor="w")
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        for a in attachments:
            aid = a.get("id") or a.get("attachment_id") or ""
            fname = a.get("file_name") or a.get("name") or ""
            mtype = a.get("mime_type") or a.get("content_type") or ""
            size = a.get("size") or ""
            desc = a.get("description") or ""
            tree.insert("", "end", values=(str(aid), str(fname), str(mtype), str(size), str(desc)))

        def on_double_click(event):
            sel = tree.selection()
            if not sel: return
            item = tree.item(sel[0])
            values = item.get("values", [])
            file_name = values[1] if len(values) > 1 else None
            if not file_name:
                messagebox.showinfo("Info", "No filename available for this attachment.")
                return
            path = self.sample_paths.get(file_index)
            if not path:
                messagebox.showerror("Error", "Original file path not available.")
                return
            save_to = filedialog.asksaveasfilename(initialfile=file_name)
            if not save_to:
                return
            mkvextract_exe = "mkvextract.exe" if sys.platform == "win32" else "mkvextract"
            mkvextract_path = os.path.join(self.mkvtoolnix_path.get(), mkvextract_exe)
            if os.path.exists(mkvextract_path):
                try:
                    aid = values[0]
                    subprocess.run([mkvextract_path, "attachments", "extract", path, f"{aid}:{save_to}"], check=True)
                    messagebox.showinfo("Success", f"Attachment saved to:\n{save_to}")
                except subprocess.CalledProcessError as e:
                    messagebox.showerror("Error", f"mkvextract failed: {e}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not extract attachment: {e}")
            else:
                messagebox.showinfo("mkvextract not found",
                    "mkvextract not found in MKVToolNix path.\n"
                    "Install / point to mkvextract or extract attachments using your usual tool.")

        tree.bind("<Double-1>", on_double_click)

        ttk.Label(win, text="Double-click an attachment to try extracting it (requires mkvextract).").pack(padx=8, pady=(0,8), anchor="w")

    def show_global_tags_window(self, file_index):
        data = self.file_jsons.get(file_index)
        if not data:
            messagebox.showinfo("No data", "No mkvmerge JSON info is available for this file.")
            return
        tags = data.get("tags") or data.get("global_tags") or {}
        if not tags:
            messagebox.showinfo("No global tags", "No global tags found for this file.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Global tags - File {file_index}")
        win.geometry("700x420")

        text = tk.Text(win, wrap="none")
        text.pack(fill="both", expand=True, padx=6, pady=6)

        try:
            pretty = json.dumps(tags, indent=2, ensure_ascii=False)
        except Exception:
            pretty = str(tags)
        text.insert("1.0", pretty)
        text.config(state="disabled")

        vscroll = ttk.Scrollbar(win, orient="vertical", command=text.yview)
        hscroll = ttk.Scrollbar(win, orient="horizontal", command=text.xview)
        text.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.pack(side="right", fill="y")
        hscroll.pack(side="bottom", fill="x")

    def create_track_widgets(self, parent, title, mkv_file, file_index):
        frame = ttk.LabelFrame(parent, text=title)
        frame.pack(padx=10, pady=10, fill="x", expand=True)
        self.track_selections[file_index] = []

        headers = ["Include", "ID", "Type", "Codec", "Language", "Name", "Default", "Forced"]
        for i, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=i, padx=5, sticky="w")

        for i, track in enumerate(mkv_file.tracks):
            include_var = tk.BooleanVar(value=True)
            lang_var = tk.StringVar(value=track.language or "und")
            name_var = tk.StringVar(value=track.track_name or "")
            default_var = tk.BooleanVar(value=track.default_track)
            forced_var = tk.BooleanVar(value=track.forced_track)
            self.track_selections[file_index].append({
                "track_obj": track, "include": include_var, "language": lang_var,
                "name": name_var, "default": default_var, "forced": forced_var
            })
            row = i + 1
            ttk.Checkbutton(frame, variable=include_var).grid(row=row, column=0)
            ttk.Label(frame, text=str(track.track_id)).grid(row=row, column=1, sticky="w")
            ttk.Label(frame, text=track.track_type.capitalize()).grid(row=row, column=2, sticky="w")
            codec_str = getattr(track, 'track_codec', 'N/A')
            ttk.Label(frame, text=codec_str).grid(row=row, column=3, sticky="w")
            ttk.Entry(frame, textvariable=lang_var, width=6).grid(row=row, column=4, sticky="w")
            ttk.Entry(frame, textvariable=name_var, width=28).grid(row=row, column=5, sticky="w")
            ttk.Checkbutton(frame, variable=default_var).grid(row=row, column=6)
            ttk.Checkbutton(frame, variable=forced_var).grid(row=row, column=7)

    def show_progress_window(self, maximum):
        if self.progress_window and tk.Toplevel.winfo_exists(self.progress_window):
            try:
                self.progress_window.destroy()
            except Exception:
                pass

        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Merging Progress")
        self.progress_window.geometry("450x140")
        self.progress_window.resizable(False, False)
        self.progress_window.transient(self.root)
        self.progress_window.protocol("WM_DELETE_WINDOW", lambda: None)

        container = ttk.Frame(self.progress_window, padding=12)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Batch merge in progress...", font=(None, 11, 'bold')).pack(anchor="w")

        self.merge_progressbar = ttk.Progressbar(container, orient="horizontal", length=380, mode="determinate")
        self.merge_progressbar.pack(pady=(12, 6))
        self.merge_progressbar["maximum"] = maximum
        self.merge_progress_label = ttk.Label(container, text="Preparing...")
        self.merge_progress_label.pack()

        self.accept_button = None
        self.progress_window.lift()
        self.root.update_idletasks()

    def finish_progress_window(self, final_text, success=True):
        def _finish():
            if not (self.progress_window and tk.Toplevel.winfo_exists(self.progress_window)):
                return
            self.progress_window.protocol("WM_DELETE_WINDOW", self.progress_window.destroy)
            self.merge_progress_label.config(text=final_text)

            if not self.accept_button:
                try:
                    geom = self.progress_window.geometry().split('+')[0]
                    w, h = geom.split('x')
                    new_h = max(int(h), 200)
                    self.progress_window.geometry(f"{w}x{new_h}")
                except Exception:
                    self.progress_window.geometry("450x200")

                self.accept_button = ttk.Button(self.progress_window, text="Accept", command=self.progress_window.destroy, width=14)
                self.accept_button.pack(side="bottom", pady=14, ipadx=8, ipady=6)
                self.progress_window.update_idletasks()

            if success and self.merge_progressbar:
                try:
                    self.merge_progressbar["value"] = self.merge_progressbar["maximum"]
                except Exception:
                    pass

            if self.start_merge_button:
                try:
                    if getattr(self.start_merge_button, "winfo_exists", lambda: 0)():
                        self.start_merge_button.state(["!disabled"])
                except tk.TclError:
                    pass

        self.root.after(0, _finish)

    def start_merging(self):
        if self.start_merge_button:
            try:
                if getattr(self.start_merge_button, "winfo_exists", lambda: 0)():
                    self.start_merge_button.state(["disabled"])
            except tk.TclError:
                pass

        try:
            if hasattr(self, 'track_window') and self.track_window and tk.Toplevel.winfo_exists(self.track_window):
                self.track_window.destroy()
        except Exception:
            pass

        matching_files = self.find_matching_files()
        if not matching_files:
            messagebox.showinfo("Information", "No matching files found.")
            print(f"[ERROR] No matching files found")
            if self.start_merge_button:
                try:
                    if getattr(self.start_merge_button, "winfo_exists", lambda: 0)():
                        self.start_merge_button.state(["!disabled"])
                except tk.TclError:
                    pass
            return

        self.show_progress_window(len(matching_files))

        self.merge_thread = threading.Thread(target=self._merge_worker, args=(matching_files,), daemon=True)
        self.merge_thread.start()

    def _merge_worker(self, matching_files):
        error_occurred = False

        for i, filename in enumerate(matching_files):
            def ui_update(idx=i, fname=filename):
                try:
                    if self.merge_progressbar:
                        self.merge_progressbar["value"] = idx + 1
                    if self.merge_progress_label:
                        self.merge_progress_label.config(text=f"Processing: {fname} ({idx+1}/{len(matching_files)})")
                except Exception:
                    pass
            self.root.after(0, ui_update)

            try:
                print(f"\n[INFO] Starting merge for {filename} (file {i+1}/{len(matching_files)})")
                src1 = os.path.join(self.folder1_path.get(), filename)
                src2 = os.path.join(self.folder2_path.get(), filename)
                final_output = os.path.join(self.output_folder_path.get(), filename)

                if not os.path.exists(src1) or not os.path.exists(src2):
                    print(f"[ERROR] Source files missing for {filename}")
                    raise FileNotFoundError(f"Source files missing for {filename}")

                src_mkv1 = MKVFile(src1)
                src_mkv2 = MKVFile(src2)

                new_mkv = MKVFile()
                new_mkv.title = self.metadata_title.get()

                for selection in self.track_selections.get(1, []):
                    if selection["include"].get():
                        matching_track = next((t for t in src_mkv1.tracks if t.track_id == selection["track_obj"].track_id), None)
                        if matching_track:
                            self.apply_track_settings(matching_track, selection)
                            new_mkv.add_track(matching_track)

                for selection in self.track_selections.get(2, []):
                    if selection["include"].get():
                        matching_track = next((t for t in src_mkv2.tracks if t.track_id == selection["track_obj"].track_id), None)
                        if matching_track:
                            self.apply_track_settings(matching_track, selection)
                            new_mkv.add_track(matching_track)

                if not self.include_chapters_file1.get():
                    print(f"[INFO] Chapters NOT requested: removing chapters from File 1: {src1}")
                    try:
                        src_mkv1.no_chapters()
                    except AttributeError:
                        print("[WARN] MKVFile.no_chapters() not available in this pymkv2 version.")
                if not self.include_chapters_file2.get():
                    print(f"[INFO] Chapters NOT requested: removing chapters from File 2: {src2}")
                    try:
                        src_mkv2.no_chapters()
                    except AttributeError:
                        print("[WARN] MKVFile.no_chapters() not available in this pymkv2 version.")

                if not self.include_global_tags_file1.get():
                    print(f"[INFO] Global tags NOT requested: removing global tags from File 1: {src1}")
                    try:
                        src_mkv1.no_global_tags()
                    except AttributeError:
                        print("[WARN] MKVFile no_global_tags method not available in this pymkv2 version.")
                if not self.include_global_tags_file2.get():
                    print(f"[INFO] Global tags NOT requested: removing global tags from File 2: {src2}")
                    try:
                        src_mkv2.no_global_tags()
                    except AttributeError:
                        print("[WARN] MKVFile no_global_tags/no_tags method not available in this pymkv2 version.")

                if not self.include_attachments_file1.get():
                    print(f"[INFO] Attachments NOT requested: removing attachments from File 1: {src1}")
                    try:
                        src_mkv1.no_attachments()
                    except AttributeError:
                        print("[WARN] MKVFile.no_attachments() not available in this pymkv2 version.")
                if not self.include_attachments_file2.get():
                    print(f"[INFO] Attachments NOT requested: removing attachments from File 2: {src2}")
                    try:
                        src_mkv2.no_attachments()
                    except AttributeError:
                        print("[WARN] MKVFile.no_attachments() not available in this pymkv2 version.")

                try:
                    new_mkv.mux(final_output, silent=True)
                except Exception as e:
                    print(f"[ERROR] pymkv2 mux failed for {filename}: {e}")
                    traceback.print_exc()
                    raise

            except Exception as e:
                error_occurred = True
                err_text = f"Error merging {filename}: {e}"
                print(f"[ERROR] {err_text}")
                traceback.print_exc()
                self.root.after(0, lambda: self.finish_progress_window(err_text, success=False))
                break

        if not error_occurred:
            self.root.after(0, lambda: self.finish_progress_window("All files merged successfully!", success=True))

    def apply_track_settings(self, track, selection):
        track.language = selection["language"].get()
        track.track_name = selection["name"].get()
        track.default_track = selection["default"].get()
        track.forced_track = selection["forced"].get()

    def save_preset(self):
        preset_data = {
            "global_properties": {
                "title": self.metadata_title.get(),
                "include_chapters_file1": self.include_chapters_file1.get(),
                "include_chapters_file2": self.include_chapters_file2.get(),
                "include_global_tags_file1": self.include_global_tags_file1.get(),
                "include_global_tags_file2": self.include_global_tags_file2.get(),
                "include_attachments_file1": self.include_attachments_file1.get(),
                "include_attachments_file2": self.include_attachments_file2.get()
            },
            "tracks": {}
        }
        for file_index, selections in self.track_selections.items():
            preset_data["tracks"][file_index] = [{
                "track_id": s["track_obj"].track_id, "include": s["include"].get(),
                "language": s["language"].get(), "name": s["name"].get(),
                "default": s["default"].get(), "forced": s["forced"].get()
            } for s in selections]

        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if filepath:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, indent=4)
            messagebox.showinfo("Success", "Preset saved successfully.")

    def load_preset(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filepath:
            with open(filepath, 'r', encoding='utf-8') as f:
                preset_data = json.load(f)

            global_props = preset_data.get("global_properties", {})
            self.metadata_title.set(global_props.get("title", ""))
            self.include_chapters_file1.set(global_props.get("include_chapters_file1", False))
            self.include_chapters_file2.set(global_props.get("include_chapters_file2", False))
            self.include_global_tags_file1.set(global_props.get("include_global_tags_file1", False))
            self.include_global_tags_file2.set(global_props.get("include_global_tags_file2", False))
            self.include_attachments_file1.set(global_props.get("include_attachments_file1", False))
            self.include_attachments_file2.set(global_props.get("include_attachments_file2", False))

            track_data = preset_data.get("tracks", {})
            for file_index_str, tracks in track_data.items():
                file_index = int(file_index_str)
                for track_settings in tracks:
                    for ui_track in self.track_selections.get(file_index, []):
                        if ui_track["track_obj"].track_id == track_settings["track_id"]:
                            ui_track["include"].set(track_settings["include"])
                            ui_track["language"].set(track_settings["language"])
                            ui_track["name"].set(track_settings["name"])
                            ui_track["default"].set(track_settings["default"])
                            ui_track["forced"].set(track_settings["forced"])
                            break
            messagebox.showinfo("Success", "Preset loaded successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = Pymkv2MergerApp(root)
    root.mainloop()
