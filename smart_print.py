import os
import threading
import fitz  
import win32print
import win32ui
import win32con
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageWin, ImageDraw, ImageFont

ctk.set_appearance_mode("dark")  # "light", "dark", or "system"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

class SmartDuplexManager:

    def __init__(self, root):
        self.root = root
        self.root.title("Smart Print Manager")
        self.root.geometry("600x800")
        self.file_path = None
        self.doc = None

        self.build_ui()
        self.load_printers()


    def build_ui(self):
        header_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        header_frame.pack(pady=(20, 10), fill="x", padx=20)
        
        ctk.CTkLabel(header_frame, text="Smart Duplex Manager", font=ctk.CTkFont(size=24, weight="bold")).pack()
        ctk.CTkLabel(header_frame, text="Configure and manage your print jobs.", text_color="gray", font=ctk.CTkFont(size=13)).pack()

        file_card = ctk.CTkFrame(self.root, corner_radius=10)
        file_card.pack(pady=10, padx=20, fill="x")
        
        self.btn_select = ctk.CTkButton(file_card, text="Select PDF Document", command=self.select_pdf, font=ctk.CTkFont(weight="bold"))
        self.btn_select.pack(pady=(15, 5))
        
        self.file_label = ctk.CTkLabel(file_card, text="No file selected", text_color="gray", font=ctk.CTkFont(size=12))
        self.file_label.pack(pady=(0, 15))

        printer_card = ctk.CTkFrame(self.root, corner_radius=10)
        printer_card.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(printer_card, text="Target Printer", font=ctk.CTkFont(weight="bold")).pack(pady=(15, 5), padx=15, anchor="w")
        self.printer_combo = ctk.CTkOptionMenu(printer_card, dynamic_resizing=False)
        self.printer_combo.pack(pady=(0, 15), padx=15, fill="x")

        config_card = ctk.CTkFrame(self.root, corner_radius=10)
        config_card.pack(pady=10, padx=20, fill="x")

        row1 = ctk.CTkFrame(config_card, fg_color="transparent")
        row1.pack(fill="x", pady=10, padx=15)
        
        ctk.CTkLabel(row1, text="Pages (e.g., 1-5):").pack(side="left")
        self.page_range_var = ctk.StringVar(value="all")
        ctk.CTkEntry(row1, textvariable=self.page_range_var, width=120).pack(side="left", padx=10)
        
        ctk.CTkLabel(row1, text="Step:").pack(side="left", padx=(10, 0))
        self.step_var = ctk.StringVar(value="1")
        ctk.CTkEntry(row1, textvariable=self.step_var, width=50).pack(side="left", padx=10)

        row2 = ctk.CTkFrame(config_card, fg_color="transparent")
        row2.pack(fill="x", pady=5, padx=15)
        ctk.CTkLabel(row2, text="Orientation:").pack(side="left")
        self.orient_var = ctk.StringVar(value="auto")
        ctk.CTkRadioButton(row2, text="Auto-Rotate", variable=self.orient_var, value="auto").pack(side="left", padx=15)
        ctk.CTkRadioButton(row2, text="Original", variable=self.orient_var, value="original").pack(side="left")

        row3 = ctk.CTkFrame(config_card, fg_color="transparent")
        row3.pack(fill="x", pady=(5, 15), padx=15)
        ctk.CTkLabel(row3, text="Flip Edge:").pack(side="left")
        self.flip_mode = ctk.StringVar(value="long")
        ctk.CTkRadioButton(row3, text="Long Edge", variable=self.flip_mode, value="long").pack(side="left", padx=15)
        ctk.CTkRadioButton(row3, text="Short Edge", variable=self.flip_mode, value="short").pack(side="left")

        toggle_card = ctk.CTkFrame(self.root, corner_radius=10)
        toggle_card.pack(pady=10, padx=20, fill="x")

        self.duplex_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(toggle_card, text="Hardware Duplex (Configure in OS first)", variable=self.duplex_var).pack(pady=(15, 5), padx=15, anchor="w")

        self.blank_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(toggle_card, text="Auto-add blank page for odd counts", variable=self.blank_var).pack(pady=5, padx=15, anchor="w")

        self.print_pagenum_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(toggle_card, text="Stamp PDF page number on output", variable=self.print_pagenum_var).pack(pady=(5, 15), padx=15, anchor="w")

        self.status_label = ctk.CTkLabel(self.root, text="Ready", text_color="gray", font=ctk.CTkFont(size=12))
        self.status_label.pack(pady=(10, 0))

        self.progress = ctk.CTkProgressBar(self.root, width=400)
        self.progress.pack(pady=10)
        self.progress.set(0)

        self.print_button = ctk.CTkButton(self.root, text="PRINT DOCUMENT", command=self.start_print_thread, height=45, font=ctk.CTkFont(size=15, weight="bold"))
        self.print_button.pack(pady=10)

    def load_printers(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        names = [p[2] for p in printers]
        if names:
            self.printer_combo.configure(values=names)
            self.printer_combo.set(names[0])

    def select_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            try:
                if self.doc: self.doc.close()
                self.doc = fitz.open(file_path)
                self.file_path = file_path
                total_pages = len(self.doc)
                filename = os.path.basename(file_path)
                self.file_label.configure(text=f"{filename}  |  {total_pages} Pages", text_color="white")
                self.status_label.configure(text=f"Loaded {total_pages} pages.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open PDF:\n{e}")
                self.file_path = None
                self.doc = None

    def parse_page_range(self, range_str, max_pages):
        range_str = range_str.strip().lower()
        if not range_str or range_str == "all": return list(range(max_pages))
        pages = set()
        for part in range_str.split(','):
            part = part.strip()
            if not part: continue
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    start = max(1, start)
                    end = min(max_pages, end)
                    if start <= end: pages.update(range(start - 1, end))
                except ValueError: raise ValueError(f"Invalid range format: '{part}'")
            else:
                try:
                    p = int(part)
                    if 1 <= p <= max_pages: pages.add(p - 1)
                except ValueError: raise ValueError(f"Invalid page number: '{part}'")
        return sorted(list(pages))

    def start_print_thread(self):
        if not self.file_path or not self.doc:
            messagebox.showerror("Error", "Select a valid PDF first.")
            return
        threading.Thread(target=self.print_document).start()

    def print_document(self):
        printer_name = self.printer_combo.get()
        orig_pages = len(self.doc)

        try:
            job_pages = self.parse_page_range(self.page_range_var.get(), orig_pages)
            step_size = max(1, int(self.step_var.get()))
            job_pages = job_pages[::step_size]
        except ValueError as e:
            self.root.after(0, lambda: messagebox.showerror("Invalid Range", str(e)))
            return

        if not job_pages:
            self.root.after(0, lambda: messagebox.showinfo("Info", "No valid pages selected."))
            return

        self.root.after(0, lambda: self.print_button.configure(state="disabled"))

        if self.duplex_var.get():
            try:
                self.print_job(printer_name, "Hardware Duplex Job", job_pages)
                self.root.after(0, lambda: messagebox.showinfo("Done", "Job sent to printer."))
                self.root.after(0, lambda: self.status_label.configure(text="Print Job Complete!"))
            except Exception as ex:
                self.root.after(0, lambda: messagebox.showerror("Print Error", str(ex)))
            finally:
                self.root.after(0, lambda: self.print_button.configure(state="normal"))
            return

        odd_pages = job_pages[0::2]
        even_pages = job_pages[1::2]

        if len(odd_pages) > len(even_pages) and self.blank_var.get():
            even_pages.append(None)

        try:
            self.print_job(printer_name, "Manual Duplex - Odd Pass", odd_pages, is_even_pass=False)
            self.root.after(0, self.prompt_flip, printer_name, even_pages)
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showerror("Print Error", str(ex)))
            self.root.after(0, lambda: self.print_button.configure(state="normal"))

    def prompt_flip(self, printer_name, even_pages):
        if not even_pages or all(p is None for p in even_pages):
            messagebox.showinfo("Done", "Print complete. (No even pages to flip).")
            self.print_button.configure(state="normal")
            self.progress.set(0)
            self.status_label.configure(text="Print Job Complete!")
            return

        self.status_label.configure(text="Waiting for user to flip paper...")
        resp = messagebox.askokcancel("Flip Paper", "Odd pages printed.\n\nFlip the stack and reinsert it. Press OK to print the back sides.")
        
        if resp:
            threading.Thread(target=self.print_even_pass, args=(printer_name, even_pages)).start()
        else:
            self.print_button.configure(state="normal")
            self.progress.set(0)
            self.status_label.configure(text="Print cancelled.")

    def print_even_pass(self, printer_name, even_pages):
        try:
            reversed_even = list(reversed(even_pages))
            self.print_job(printer_name, "Manual Duplex - Even Pass", reversed_even, is_even_pass=True)
            self.root.after(0, lambda: messagebox.showinfo("Done", "Printing Complete."))
            self.root.after(0, lambda: self.status_label.configure(text="Print Job Complete!"))
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showerror("Print Error", str(ex)))
        finally:
            self.root.after(0, lambda: self.print_button.configure(state="normal"))
            self.root.after(0, lambda: self.progress.set(0))

    def print_job(self, printer_name, job_name, page_indices, is_even_pass=False):
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            hdc.StartDoc(job_name)

            HORZRES = hdc.GetDeviceCaps(win32con.HORZRES)
            VERTRES = hdc.GetDeviceCaps(win32con.VERTRES)
            canvas_is_wide = HORZRES > VERTRES
            total_passes = len(page_indices)

            for i, page_index in enumerate(page_indices):
                hdc.StartPage()

                if page_index is not None:
                    actual_page_num = page_index + 1
                    status_text = f"Printing Page {actual_page_num} ... ({i+1} of {total_passes})"
                    self.root.after(0, lambda t=status_text: self.status_label.configure(text=t))

                    page = self.doc.load_page(page_index)
                    pix = page.get_pixmap(dpi=300)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    if self.print_pagenum_var.get():
                        draw = ImageDraw.Draw(img)
                        try:
                            font = ImageFont.truetype("arial.ttf", 50)
                        except IOError:
                            font = ImageFont.load_default()
                        text = f"Page {actual_page_num}"
                        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
                        text_w, text_h = right - left, bottom - top
                        margin = 60
                        x, y = img.width - text_w - margin, img.height - text_h - margin
                        draw.rectangle((x-10, y-10, x+text_w+10, y+text_h+10), fill="white")
                        draw.text((x, y), text, font=font, fill="black")

                    if self.orient_var.get() == "auto":
                        if (img.width > img.height) != canvas_is_wide:
                            img = img.rotate(90, expand=True)

                    if is_even_pass and self.flip_mode.get() == "short":
                        img = img.rotate(180, expand=True)

                    img_ratio = img.width / img.height
                    page_ratio = HORZRES / VERTRES

                    if img_ratio > page_ratio:
                        draw_w, draw_h = HORZRES, int(HORZRES / img_ratio)
                    else:
                        draw_w, draw_h = int(VERTRES * img_ratio), VERTRES

                    offset_x, offset_y = (HORZRES - draw_w) // 2, (VERTRES - draw_h) // 2

                    dib = ImageWin.Dib(img)
                    dib.draw(hdc.GetHandleOutput(), (offset_x, offset_y, offset_x + draw_w, offset_y + draw_h))

                hdc.EndPage()
                self.root.after(0, lambda current=i+1, total=total_passes: self.progress.set(current/total))

            hdc.EndDoc()
            hdc.DeleteDC()
        finally:
            win32print.ClosePrinter(hprinter)

if __name__ == "__main__":
    app_root = ctk.CTk()
    app = SmartDuplexManager(app_root)
    app_root.mainloop()