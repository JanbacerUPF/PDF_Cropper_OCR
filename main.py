import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Canvas, Scrollbar
import os
import io

try:
    import fitz  # PyMuPDF for PDF rendering and editing
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from PIL import Image, ImageTk
import win32com.client as win32  # For Word automation (PDF → DOCX)


class PDFCropper:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Blanker with Preview")
        self.root.geometry("900x700")

        self.pdf_path = None
        self.pdf_document = None
        self.current_page = 0
        self.preview_image = None
        self.blank_rectangles = []
        self.pdf_width = 0
        self.pdf_height = 0

        if not HAS_PYMUPDF:
            messagebox.showwarning(
                "Warning",
                "PyMuPDF not installed. Preview feature will be limited.\n"
                "Install with: pip install PyMuPDF"
            )

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left panel
        control_frame = ttk.Frame(main_frame, width=300)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        control_frame.pack_propagate(False)

        # File selection
        file_frame = ttk.LabelFrame(control_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(file_frame, text="Select PDF File", command=self.select_file).pack(pady=5)
        self.file_label = ttk.Label(file_frame, text="No file selected", wraplength=250)
        self.file_label.pack(pady=5)

        # Page navigation (hidden until PDF is loaded)
        self.nav_frame = ttk.LabelFrame(control_frame, text="Page Navigation", padding="10")
        nav_buttons = ttk.Frame(self.nav_frame)
        nav_buttons.pack()
        ttk.Button(nav_buttons, text="◀", command=self.prev_page, width=3).pack(side=tk.LEFT, padx=2)
        self.page_label = ttk.Label(nav_buttons, text="Page 0/0")
        self.page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(nav_buttons, text="▶", command=self.next_page, width=3).pack(side=tk.LEFT, padx=2)
        self.nav_frame.pack_forget()

        # Margin sliders
        blank_frame = ttk.LabelFrame(control_frame, text="Blank Areas (points)", padding="10")
        blank_frame.pack(fill=tk.X, pady=(0, 10))
        params = [("Left:", "left_var"), ("Top:", "top_var"), ("Right:", "right_var"), ("Bottom:", "bottom_var")]
        MAX_MARGIN_POINTS = 200

        for i, (label, var_name) in enumerate(params):
            ttk.Label(blank_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            var = tk.DoubleVar(value=0)
            setattr(self, var_name, var)
            display_var = tk.StringVar(value="0.00")

            def on_slider_change(val, v=var, dv=display_var):
                dv.set(f"{float(val):.2f}")
                self.update_preview()

            slider = ttk.Scale(blank_frame, from_=0, to=MAX_MARGIN_POINTS,
                               orient="horizontal", variable=var, command=on_slider_change)
            slider.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            ttk.Label(blank_frame, textvariable=display_var).grid(row=i, column=2, padx=5)

        ttk.Label(blank_frame, text="72 points = 1 inch", font=("Arial", 8)).grid(row=4, column=0, columnspan=3, pady=5)

        # Preview controls
        preview_controls = ttk.LabelFrame(control_frame, text="Preview Controls", padding="10")
        preview_controls.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(preview_controls, text="Update Preview", command=self.update_preview).pack(pady=2)
        ttk.Button(preview_controls, text="Reset Areas", command=self.reset_blank).pack(pady=2)

        # Action buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(button_frame, text="Blank Areas PDF", command=self.blank_pdf).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Blank Areas DOCX (OCR)", command=self.blank_and_convert_to_docx).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(fill=tk.X, pady=2)
        ttk.Label(button_frame, text="Liverbula S.L. ©", font=("Arial", 9, "italic"), foreground="gray").pack(fill=tk.X, pady=(10, 0))

        # Right panel (Preview)
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        canvas_frame = ttk.Frame(preview_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = Canvas(canvas_frame, bg="white")
        v_scrollbar = Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.create_text(200, 200, text="Select a PDF file to preview", font=("Arial", 16), fill="gray")

    def select_file(self):
        self.pdf_path = filedialog.askopenfilename(
            title="Select PDF file", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if self.pdf_path:
            self.file_label.config(text=f"Selected: {os.path.basename(self.pdf_path)}")
            self.load_pdf()

    def load_pdf(self):
        if not HAS_PYMUPDF:
            messagebox.showinfo("Info", "Preview requires PyMuPDF. Install with: pip install PyMuPDF")
            return
        try:
            self.pdf_document = fitz.open(self.pdf_path)
            self.current_page = 0
            self.nav_frame.pack(fill=tk.X, pady=(0, 10))
            self.update_page_info()
            self.render_page()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {e}")

    def update_page_info(self):
        if self.pdf_document:
            total_pages = len(self.pdf_document)
            self.page_label.config(text=f"Page {self.current_page + 1}/{total_pages}")
            rect = self.pdf_document[self.current_page].rect
            self.pdf_width, self.pdf_height = rect.width, rect.height

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_page_info()
            self.render_page()

    def next_page(self):
        if self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.update_page_info()
            self.render_page()

    def render_page(self):
        if not self.pdf_document:
            return
        try:
            page = self.pdf_document[self.current_page]
            self.canvas.update_idletasks()
            scale_factor = min(
                self.canvas.winfo_width() / self.pdf_width,
                self.canvas.winfo_height() / self.pdf_height
            ) * 0.95
            pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor))
            pil_image = Image.open(io.BytesIO(pix.tobytes("ppm")))
            self.image_width, self.image_height = pil_image.width, pil_image.height
            self.preview_image = ImageTk.PhotoImage(pil_image)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.preview_image)
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to render page: {e}")

    def update_preview(self):
        if not self.preview_image:
            return
        for rect in self.blank_rectangles:
            self.canvas.delete(rect)
        self.blank_rectangles.clear()
        px_x = self.image_width / self.pdf_width
        px_y = self.image_height / self.pdf_height
        for side, coords in [
            (self.left_var.get(), (0, 0, self.left_var.get() * px_x, self.image_height)),
            (self.top_var.get(), (0, 0, self.image_width, self.top_var.get() * px_y)),
            (self.right_var.get(), (self.image_width - self.right_var.get() * px_x, 0, self.image_width, self.image_height)),
            (self.bottom_var.get(), (0, self.image_height - self.bottom_var.get() * px_y, self.image_width, self.image_height))
        ]:
            if side > 0:
                self.blank_rectangles.append(
                    self.canvas.create_rectangle(*coords, fill="white", outline="red", width=2, stipple="gray50")
                )

    def reset_blank(self):
        for var in [self.left_var, self.top_var, self.right_var, self.bottom_var]:
            var.set(0)
        self.update_preview()

    def apply_blank_to_pdf(self, pdf_doc, left, top, right, bottom):
        for page in pdf_doc:
            w, h = page.rect.width, page.rect.height
            if left: page.draw_rect(fitz.Rect(0, 0, left, h), color=(1, 1, 1), fill=(1, 1, 1))
            if top: page.draw_rect(fitz.Rect(0, 0, w, top), color=(1, 1, 1), fill=(1, 1, 1))
            if right: page.draw_rect(fitz.Rect(w - right, 0, w, h), color=(1, 1, 1), fill=(1, 1, 1))
            if bottom: page.draw_rect(fitz.Rect(0, h - bottom, w, h), color=(1, 1, 1), fill=(1, 1, 1))

    def blank_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return
        if not HAS_PYMUPDF:
            messagebox.showerror("Error", "PyMuPDF is required for blanking functionality.")
            return
        try:
            left, top, right, bottom = map(float, [
                self.left_var.get(), self.top_var.get(),
                self.right_var.get(), self.bottom_var.get()
            ])
            pdf_doc = fitz.open(self.pdf_path)
            self.apply_blank_to_pdf(pdf_doc, left, top, right, bottom)
            output_path = f"{os.path.splitext(self.pdf_path)[0]}_blanked.pdf"
            pdf_doc.save(output_path)
            pdf_doc.close()
            messagebox.showinfo("Success", f"Blanked PDF saved as:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def blank_and_convert_to_docx(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return
        if not HAS_PYMUPDF:
            messagebox.showerror("Error", "PyMuPDF is required for blanking functionality.")
            return
        try:
            left, top, right, bottom = map(float, [
                self.left_var.get(), self.top_var.get(),
                self.right_var.get(), self.bottom_var.get()
            ])
            pdf_doc = fitz.open(self.pdf_path)
            self.apply_blank_to_pdf(pdf_doc, left, top, right, bottom)
            base_name = os.path.splitext(self.pdf_path)[0]
            output_pdf_path = os.path.abspath(f"{base_name}_blanked.pdf")
            pdf_doc.save(output_pdf_path)
            pdf_doc.close()
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(output_pdf_path)
            output_docx_path = os.path.abspath(f"{base_name}_blanked.docx")
            doc.SaveAs2(output_docx_path, FileFormat=16)
            doc.Close()
            word.Quit()
            messagebox.showinfo("Success", f"Blanked and converted file saved as:\n{output_docx_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


def main():
    root = tk.Tk()
    PDFCropper(root)
    root.mainloop()


if __name__ == "__main__":
    main()
