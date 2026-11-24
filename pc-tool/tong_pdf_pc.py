# -*- coding: utf-8 -*-
"""
통PDF (TongPDF-PC) - 교무실 Windows PC용 무설치 PDF 일괄 변환기
지원: .hwp, .hwpx (한글), .show (한쇼), .ppt, .pptx (파워포인트), .doc, .docx (워드), 이미지
"""

import os
import sys
import time
import queue
import threading
import subprocess
import traceback
import winreg
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SUPPORTED_EXTS = {
    ".hwp", ".hwpx",
    ".show",
    ".ppt", ".pptx",
    ".doc", ".docx",
    ".jpg", ".jpeg", ".png", ".bmp", ".webp",
}

HWP_EXTS = {".hwp", ".hwpx"}
SHOW_EXTS = {".show"}
PPT_EXTS = {".ppt", ".pptx"}
WORD_EXTS = {".doc", ".docx"}
IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


def get_hwp_security_module_name():
    """레지스트리에 등록된 HwpAutomation 보안 모듈 이름이 있으면 반환"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\HNC\HwpAutomation\Modules",
            0,
            winreg.KEY_READ,
        )
        try:
            i = 0
            while True:
                name, val, _ = winreg.EnumValue(key, i)
                if val and os.path.exists(str(val)):
                    return name
                i += 1
        except OSError:
            pass
        finally:
            winreg.CloseKey(key)
    except OSError:
        pass
    return None


def try_register_hwp_security_module():
    """PC에 FilePathCheckerModule DLL이 있으면 레지스트리에 자동 등록"""
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "FilePathCheckerModuleExample.dll"),
        r"C:\Program Files (x86)\Hnc\Office 2024\HOffice130\Bin\FilePathCheckerModuleExample.dll",
        r"C:\Program Files (x86)\Hnc\Office 2022\HOffice120\Bin\FilePathCheckerModuleExample.dll",
        r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\FilePathCheckerModuleExample.dll",
        r"C:\Program Files (x86)\Hnc\Office 2018\HOffice100\Bin\FilePathCheckerModuleExample.dll",
    ]
    found = next((p for p in candidates if os.path.exists(p)), None)
    if not found:
        return False, (
            "보안 모듈 DLL(FilePathCheckerModuleExample.dll)을 찾지 못했습니다.\n\n"
            "걱정하지 마세요! DLL이 없어도 변환은 정상 작동하며,\n"
            "한글이 처음 열릴 때 뜨는 접근 허용 창에서 [모두 허용]을 한 번만 누르시면 됩니다."
        )
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\HNC\HwpAutomation\Modules")
        winreg.SetValueEx(key, "FilePathCheckerModuleExample", 0, winreg.REG_SZ, found)
        winreg.CloseKey(key)
        return True, f"한글 보안 모듈이 자동 등록되었습니다.\n경로: {found}"
    except Exception as e:
        return False, f"레지스트리 등록 실패: {e}"


class OfficeConverter:
    """Windows COM 자동화를 이용한 로컬 PDF 변환 엔진"""

    def __init__(self):
        self.hwp = None
        self.ppt = None
        self.hshow = None
        self.word = None

    def _get_hwp(self):
        if self.hwp is None:
            import win32com.client as wc
            hwp = wc.Dispatch("HWPFrame.HwpObject")
            mod_name = get_hwp_security_module_name()
            if mod_name:
                try:
                    hwp.RegisterModule("FilePathCheckDLL", mod_name)
                except Exception:
                    pass
            try:
                hwp.SetMessageBoxMode(0x00010000 | 0x00001000 | 0x00000100 | 0x00000010 | 0x00000001)
            except Exception:
                pass
            self.hwp = hwp
        return self.hwp

    def _get_ppt(self):
        if self.ppt is None:
            import win32com.client as wc
            ppt = wc.Dispatch("PowerPoint.Application")
            self.ppt = ppt
        return self.ppt

    def _get_hshow(self):
        if self.hshow is None:
            import win32com.client as wc
            hshow = wc.Dispatch("HShow.Application")
            self.hshow = hshow
        return self.hshow

    def _get_word(self):
        if self.word is None:
            import win32com.client as wc
            word = wc.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            self.word = word
        return self.word

    def convert_hwp(self, src, dst):
        hwp = self._get_hwp()
        opened = hwp.Open(src, "", "forceopen:true;versionwarning:false;suspendpassword:true")
        if not opened:
            raise RuntimeError("한글 파일을 열지 못했습니다. (암호화되었거나 손상된 파일)")
        try:
            ok = hwp.SaveAs(dst, "PDF", "")
            if not ok or not os.path.exists(dst):
                # Fallback via HAction FileSaveAsPdf
                act = hwp.CreateAction("FileSaveAsPdf")
                pset = act.CreateSet()
                act.GetDefault(pset)
                pset.SetItem("FileName", dst)
                pset.SetItem("Format", "PDF")
                act.Execute(pset)
            if not os.path.exists(dst):
                raise RuntimeError("한글 PDF 저장에 실패했습니다.")
        finally:
            try:
                hwp.Clear(1)
            except Exception:
                pass

    def convert_ppt(self, src, dst):
        try:
            ppt = self._get_ppt()
            pres = ppt.Presentations.Open(src, ReadOnly=True, Untitled=False, WithWindow=False)
            try:
                # 32 = ppSaveAsPDF
                pres.SaveAs(dst, 32)
            finally:
                pres.Close()
            if os.path.exists(dst):
                return
        except Exception as ppt_err:
            # Fallback to Hancom Show if MS PowerPoint is not installed
            try:
                self.convert_hshow(src, dst)
                return
            except Exception:
                raise RuntimeError(f"파워포인트 변환 실패: {ppt_err}")

    def convert_hshow(self, src, dst):
        hshow = self._get_hshow()
        pres = None
        try:
            # Try Presentations.Open first, then app.Open
            try:
                pres = hshow.Presentations.Open(src)
            except Exception:
                pres = hshow.Open(src)

            saved = False
            # Try SaveAs / ExportAsFixedFormat variants supported across Hancom Show versions
            for attempt in (
                lambda: pres.SaveAs(dst, 32) if pres else None,
                lambda: pres.SaveAs(dst, "PDF") if pres else None,
                lambda: pres.ExportAsFixedFormat(dst, 2) if pres else None,
                lambda: hshow.ActivePresentation.SaveAs(dst, 32),
                lambda: hshow.ActivePresentation.SaveAs(dst, "PDF"),
            ):
                try:
                    attempt()
                    if os.path.exists(dst) and os.path.getsize(dst) > 0:
                        saved = True
                        break
                except Exception:
                    continue

            if not saved:
                raise RuntimeError("한쇼 PDF 내보내기 호출이 지원되지 않는 버전입니다.")
        finally:
            if pres is not None:
                try:
                    pres.Close()
                except Exception:
                    pass

    def convert_word(self, src, dst):
        try:
            word = self._get_word()
            doc = word.Documents.Open(src, ReadOnly=True)
            try:
                # 17 = wdFormatPDF
                doc.SaveAs2(dst, FileFormat=17)
            finally:
                doc.Close(False)
            if os.path.exists(dst):
                return
        except Exception as word_err:
            # Fallback to HWP (Hancom Office opens .docx/.doc well)
            try:
                self.convert_hwp(src, dst)
                return
            except Exception:
                raise RuntimeError(f"워드 변환 실패: {word_err}")

    def convert_image(self, src, dst):
        from PIL import Image
        with Image.open(src) as im:
            rgb = im.convert("RGB")
            rgb.save(dst, "PDF", resolution=150.0)

    def convert(self, src, out_dir=None):
        src_abs = os.path.abspath(src)
        ext = os.path.splitext(src_abs)[1].lower()
        base = os.path.splitext(os.path.basename(src_abs))[0]
        target_dir = os.path.abspath(out_dir) if out_dir else os.path.dirname(src_abs)
        os.makedirs(target_dir, exist_ok=True)
        dst = os.path.join(target_dir, base + ".pdf")

        # Avoid overwriting an existing PDF with identical name silently if needed, or overwrite cleanly
        if os.path.exists(dst):
            try:
                os.remove(dst)
            except Exception:
                dst = os.path.join(target_dir, f"{base}_{int(time.time())}.pdf")

        if ext in HWP_EXTS:
            self.convert_hwp(src_abs, dst)
        elif ext in SHOW_EXTS:
            self.convert_hshow(src_abs, dst)
        elif ext in PPT_EXTS:
            self.convert_ppt(src_abs, dst)
        elif ext in WORD_EXTS:
            self.convert_word(src_abs, dst)
        elif ext in IMG_EXTS:
            self.convert_image(src_abs, dst)
        else:
            raise RuntimeError(f"지원하지 않는 파일 형식입니다: {ext}")

        return dst

    def close(self):
        for name, obj in (
            ("hwp", self.hwp),
            ("ppt", self.ppt),
            ("hshow", self.hshow),
            ("word", self.word),
        ):
            if obj is not None:
                try:
                    obj.Quit()
                except Exception:
                    pass
        self.hwp = self.ppt = self.hshow = self.word = None


class TongPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("통PDF (TongPDF-PC) · 한글·한쇼·PPT 원본 100% PDF 일괄 변환기")
        self.root.geometry("820x620")
        self.root.minsize(720, 520)
        self.root.configure(bg="#f1f5f9")

        self.files = []  # list of abs paths
        self.out_dir = tk.StringVar(value="")
        self.is_running = False
        self.msg_queue = queue.Queue()

        self._build_ui()
        self._hook_drag_and_drop()
        self.root.after(120, self._poll_queue)

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass

        # Header
        header = tk.Frame(self.root, bg="#1e3a8a", padx=18, pady=14)
        header.pack(fill="x")
        tk.Label(
            header,
            text="통PDF (TongPDF-PC) · 교무실 무설치 PDF 일괄 변환기",
            font=("Malgun Gothic", 14, "bold"),
            fg="#ffffff",
            bg="#1e3a8a",
        ).pack(anchor="w")
        tk.Label(
            header,
            text="한글(.hwp/.hwpx) · 한쇼(.show) · 파워포인트(.ppt/.pptx) · 워드(.doc/.docx) · 이미지 → 원본 100% PDF",
            font=("Malgun Gothic", 9),
            fg="#dbeafe",
            bg="#1e3a8a",
        ).pack(anchor="w", pady=(3, 0))

        # Main container
        body = tk.Frame(self.root, bg="#f1f5f9", padx=16, pady=12)
        body.pack(fill="both", expand=True)

        # Drop banner
        drop_box = tk.Frame(body, bg="#eff6ff", highlightbackground="#3b82f6", highlightthickness=2, padx=14, pady=12)
        drop_box.pack(fill="x", pady=(0, 10))
        tk.Label(
            drop_box,
            text="📂 변환할 파일이나 폴더를 이 창으로 끌어다 놓으세요 (Drag & Drop)",
            font=("Malgun Gothic", 11, "bold"),
            fg="#1e3a8a",
            bg="#eff6ff",
        ).pack()
        tk.Label(
            drop_box,
            text="내 PC에 설치된 한컴오피스·파워포인트를 자동 구동합니다 · 외부 서버 전송 없음 (보안 100%)",
            font=("Malgun Gothic", 8),
            fg="#475569",
            bg="#eff6ff",
        ).pack(pady=(2, 0))

        # Toolbar
        toolbar = tk.Frame(body, bg="#f1f5f9")
        toolbar.pack(fill="x", pady=(0, 8))

        tk.Button(toolbar, text="+ 파일 추가", command=self.add_files, font=("Malgun Gothic", 9, "bold"), bg="#ffffff", padx=10, pady=5).pack(side="left", padx=(0, 6))
        tk.Button(toolbar, text="+ 폴더 통째로 추가", command=self.add_folder, font=("Malgun Gothic", 9, "bold"), bg="#ffffff", padx=10, pady=5).pack(side="left", padx=(0, 6))
        tk.Button(toolbar, text="선택 삭제", command=self.remove_selected, font=("Malgun Gothic", 9), bg="#ffffff", padx=8, pady=5).pack(side="left", padx=(0, 6))
        tk.Button(toolbar, text="목록 비우기", command=self.clear_all, font=("Malgun Gothic", 9), bg="#ffffff", padx=8, pady=5).pack(side="left", padx=(0, 6))
        tk.Button(toolbar, text="한글 보안 팝업 해제 설정", command=self.setup_hwp_security, font=("Malgun Gothic", 9), bg="#fef3c7", fg="#92400e", padx=10, pady=5).pack(side="right")

        # Output folder row
        out_row = tk.Frame(body, bg="#f1f5f9")
        out_row.pack(fill="x", pady=(0, 8))
        tk.Label(out_row, text="저장 폴더:", font=("Malgun Gothic", 9, "bold"), bg="#f1f5f9", fg="#334155").pack(side="left")
        tk.Entry(out_row, textvariable=self.out_dir, font=("Malgun Gothic", 9)).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(out_row, text="폴더 선택", command=self.pick_out_dir, font=("Malgun Gothic", 9), bg="#ffffff", padx=8).pack(side="left", padx=(0, 4))
        tk.Button(out_row, text="원본과 같은 폴더", command=lambda: self.out_dir.set(""), font=("Malgun Gothic", 8), bg="#ffffff", padx=6).pack(side="left")

        # Treeview list
        list_frame = tk.Frame(body, bg="#ffffff", bd=1, relief="solid")
        list_frame.pack(fill="both", expand=True)

        cols = ("name", "ext", "status")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("name", text="파일명")
        self.tree.heading("ext", text="형식")
        self.tree.heading("status", text="상태")
        self.tree.column("name", width=430, anchor="w")
        self.tree.column("ext", width=70, anchor="center")
        self.tree.column("status", width=230, anchor="w")

        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Bottom controls & progress
        bottom = tk.Frame(body, bg="#f1f5f9", pady=10)
        bottom.pack(fill="x")

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))

        btn_row = tk.Frame(bottom, bg="#f1f5f9")
        btn_row.pack(fill="x")

        self.status_lbl = tk.Label(
            btn_row,
            text="파일을 추가한 뒤 [일괄 PDF 변환 시작]을 눌러주세요. (저장 폴더 비우면 원본과 같은 폴더에 저장)",
            font=("Malgun Gothic", 9),
            fg="#334155",
            bg="#f1f5f9",
        )
        self.status_lbl.pack(side="left")

        self.open_out_btn = tk.Button(
            btn_row,
            text="결과 폴더 열기",
            command=self.open_output_folder,
            font=("Malgun Gothic", 9, "bold"),
            bg="#ffffff",
            padx=12,
            pady=7,
        )
        self.open_out_btn.pack(side="right", padx=(6, 0))

        self.start_btn = tk.Button(
            btn_row,
            text="⚡ 일괄 PDF 변환 시작",
            command=self.start_convert,
            font=("Malgun Gothic", 10, "bold"),
            bg="#1d4ed8",
            fg="#ffffff",
            activebackground="#1e40af",
            activeforeground="#ffffff",
            padx=18,
            pady=7,
        )
        self.start_btn.pack(side="right")

    def _hook_drag_and_drop(self):
        try:
            import windnd
            windnd.hook_dropfiles(self.root, func=self._on_drop_files)
        except Exception:
            pass

    def _on_drop_files(self, raw_files):
        paths = []
        for item in raw_files:
            if isinstance(item, bytes):
                for enc in ("mbcs", "utf-8", "cp949"):
                    try:
                        paths.append(item.decode(enc))
                        break
                    except Exception:
                        continue
            else:
                paths.append(str(item))
        self.root.after(0, lambda: self._enqueue_paths(paths))

    def _enqueue_paths(self, paths):
        added = 0
        existing = set(self.files)
        for p in paths:
            if os.path.isdir(p):
                for r, _, fnames in os.walk(p):
                    for fn in fnames:
                        full = os.path.abspath(os.path.join(r, fn))
                        ext = os.path.splitext(full)[1].lower()
                        if ext in SUPPORTED_EXTS and full not in existing:
                            self.files.append(full)
                            existing.add(full)
                            self.tree.insert("", "end", iid=full, values=(os.path.basename(full), ext.upper(), "대기 중"))
                            added += 1
            elif os.path.isfile(p):
                full = os.path.abspath(p)
                ext = os.path.splitext(full)[1].lower()
                if ext in SUPPORTED_EXTS and full not in existing:
                    self.files.append(full)
                    existing.add(full)
                    self.tree.insert("", "end", iid=full, values=(os.path.basename(full), ext.upper(), "대기 중"))
                    added += 1
        self.status_lbl.config(text=f"총 {len(self.files)}개 파일 대기 중 (방금 {added}개 추가됨)")

    def add_files(self):
        selected = filedialog.askopenfilenames(
            title="변환할 문서 선택",
            filetypes=[
                ("지원 문서 전체", "*.hwp *.hwpx *.show *.ppt *.pptx *.doc *.docx *.jpg *.jpeg *.png *.bmp *.webp"),
                ("한글 파일", "*.hwp *.hwpx"),
                ("한쇼·파워포인트", "*.show *.ppt *.pptx"),
                ("워드 파일", "*.doc *.docx"),
                ("이미지 파일", "*.jpg *.jpeg *.png *.bmp *.webp"),
                ("모든 파일", "*.*"),
            ],
        )
        if selected:
            self._enqueue_paths(list(selected))

    def add_folder(self):
        folder = filedialog.askdirectory(title="변환할 파일이 들어있는 폴더 선택")
        if folder:
            self._enqueue_paths([folder])

    def remove_selected(self):
        if self.is_running:
            return
        for iid in self.tree.selection():
            if iid in self.files:
                self.files.remove(iid)
            self.tree.delete(iid)
        self.status_lbl.config(text=f"총 {len(self.files)}개 파일 대기 중")

    def clear_all(self):
        if self.is_running:
            return
        self.files.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.progress["value"] = 0
        self.status_lbl.config(text="목록을 비웠습니다.")

    def pick_out_dir(self):
        folder = filedialog.askdirectory(title="PDF를 저장할 폴더 선택")
        if folder:
            self.out_dir.set(folder)

    def setup_hwp_security(self):
        ok, msg = try_register_hwp_security_module()
        if ok:
            messagebox.showinfo("한글 보안 모듈 설정 완료", msg)
        else:
            messagebox.showinfo("안내", msg)

    def open_output_folder(self):
        target = self.out_dir.get().strip()
        if not target and self.files:
            target = os.path.dirname(self.files[0])
        if target and os.path.isdir(target):
            subprocess.Popen(["explorer", target])
        else:
            messagebox.showinfo("안내", "먼저 파일을 추가하거나 저장 폴더를 선택해 주세요.")

    def start_convert(self):
        if self.is_running:
            return
        if not self.files:
            messagebox.showwarning("알림", "변환할 파일을 먼저 추가해 주세요.")
            return

        self.is_running = True
        self.start_btn.config(state="disabled", text="변환 진행 중...")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0

        files_copy = list(self.files)
        out_dir = self.out_dir.get().strip() or None

        t = threading.Thread(target=self._worker, args=(files_copy, out_dir), daemon=True)
        t.start()

    def _worker(self, files, out_dir):
        import pythoncom
        pythoncom.CoInitialize()
        converter = OfficeConverter()
        ok_count = 0
        fail_count = 0
        try:
            for idx, fpath in enumerate(files, start=1):
                self.msg_queue.put(("row", fpath, "변환 중..."))
                self.msg_queue.put(("status", f"({idx}/{len(files)}) {os.path.basename(fpath)} 변환 중..."))
                try:
                    dst = converter.convert(fpath, out_dir=out_dir)
                    ok_count += 1
                    self.msg_queue.put(("row", fpath, f"완료 → {os.path.basename(dst)}"))
                except Exception as e:
                    fail_count += 1
                    err_short = str(e).splitlines()[0][:80]
                    self.msg_queue.put(("row", fpath, f"실패: {err_short}"))
                self.msg_queue.put(("progress", idx))
        finally:
            converter.close()
            pythoncom.CoUninitialize()
            self.msg_queue.put(("done", ok_count, fail_count))

    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                kind = msg[0]
                if kind == "row":
                    _, iid, status = msg
                    if self.tree.exists(iid):
                        vals = list(self.tree.item(iid, "values"))
                        vals[2] = status
                        self.tree.item(iid, values=vals)
                elif kind == "status":
                    self.status_lbl.config(text=msg[1])
                elif kind == "progress":
                    self.progress["value"] = msg[1]
                elif kind == "done":
                    _, ok_c, fail_c = msg
                    self.is_running = False
                    self.start_btn.config(state="normal", text="⚡ 일괄 PDF 변환 시작")
                    self.status_lbl.config(text=f"변환 완료! 성공 {ok_c}개 / 실패 {fail_c}개")
                    messagebox.showinfo("변환 완료", f"일괄 변환이 완료되었습니다.\n\n성공: {ok_c}개\n실패: {fail_c}개")
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)


def run_cli(paths):
    import pythoncom
    pythoncom.CoInitialize()
    conv = OfficeConverter()
    try:
        for p in paths:
            dst = conv.convert(p)
            print(f"[OK] {p} -> {dst}")
    finally:
        conv.close()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--cli":
        run_cli(sys.argv[2:])
    else:
        root = tk.Tk()
        app = TongPdfApp(root)
        root.mainloop()
