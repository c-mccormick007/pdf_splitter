import sys
import os
import PyPDF2
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk


class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter")
        self.root.geometry("1500x1000")
        #self.root.state('zoomed')

        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.pdf_path = ""
        self.output_dir = r'T:\Accounts Payable\- AP 2024  -  DIGITAL FILING CABINET\COGS - DROPSHIPS\1 - SORT - RENAME'

        self.pages_to_split = []
        self.ranges_to_split = []
        self.range_list = []

        self.checkboxes = []
        self.page_checkboxes = []
        self.thumbnails = []  # Store thumbnails to prevent garbage collection

        self.range_counter = 0
        self.create_widgets()

        self.root.bind('<Return>', self.add_range)
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop)

    def create_widgets(self):
        tk.Label(self.frame, text="Drag and drop a PDF file here or use the button to select one:",
                 wraplength=300).pack(pady=10)
        select_button = tk.Button(self.frame, text="Select PDF", command=self.select_file)
        select_button.pack(pady=5)
        self.add_rangebtn = tk.Button(self.frame, text="Add Range", command=self.add_range)
        self.add_rangebtn.pack(pady=10)
        self.split_button = tk.Button(self.frame, text="Split Selected Pages", command=self.split_selected)
        self.split_button.pack(pady=10)
        tk.Label(self.frame, text="Select pages to split:").pack(pady=5)

        # Scrollable area for thumbnails with fixed width and height
        self.canvas = tk.Canvas(self.frame, height=400, width=600)  # Fixed height and width
        self.h_scrollbar = tk.Scrollbar(self.frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        self.scrollable_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.pack(fill=tk.X)  # Keep fixed width, expand horizontally only if needed
        self.h_scrollbar.pack(fill=tk.X)

    def select_file(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            self.load_pdf_pages()

    def drop(self, event):
        self.pdf_path = event.data.strip('{}')
        if self.pdf_path.endswith(".pdf"):
            self.clear_thumbnails()
            self.load_pdf_pages()
        else:
            messagebox.showerror("Invalid File", "Please drop a valid PDF file.")

    def clear_thumbnails(self):
        # Clear all widgets inside the scrollable_frame (removing all thumbnails)
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # Clear the thumbnails list
        self.thumbnails.clear()
    def load_pdf_pages(self):
        try:
            self.reader = PyPDF2.PdfReader(open(self.pdf_path, 'rb'))
            num_pages = len(self.reader.pages)

            for widget in self.checkboxes:
                widget.destroy()
            self.checkboxes.clear()
            self.page_checkboxes.clear()
            self.thumbnails.clear()

            for i in range(num_pages):
                var = tk.IntVar()
                checkbox = tk.Checkbutton(self.scrollable_frame, text=f"Page {i + 1}", variable=var)
                checkbox.pack(side=tk.LEFT, padx=5)
                self.checkboxes.append(checkbox)
                self.page_checkboxes.append((var, i))


                # Generate and display thumbnails
                thumbnail = self.generate_thumbnail(self.pdf_path, i)
                if thumbnail:
                    label = tk.Label(self.scrollable_frame, image=thumbnail)
                    label.pack(side=tk.LEFT, padx=5)
                    self.thumbnails.append(thumbnail)

            self.scrollable_frame.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def generate_thumbnail(self, pdf_path, page_num):
        try:
            doc = fitz.open(pdf_path)
            page = doc.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((400, 450))  # Larger thumbnail size

            thumbnail = ImageTk.PhotoImage(img)
            return thumbnail
        except Exception as e:
            messagebox.showerror("Error", f"Could not generate thumbnail: {str(e)}")
            return None

    def add_range(self, event=None):
        selected_indices = [index for var, index in self.page_checkboxes if var.get() == 1]
        selected_indices.sort()

        ranges = []
        if selected_indices:
            start = selected_indices[0]
            end = start
            for i in range(1, len(selected_indices)):
                if selected_indices[i] == end + 1:
                    end = selected_indices[i]
                else:
                    ranges.append([start, end])
                    start = selected_indices[i]
                    end = start
            ranges.append([start, end])

        for start, end in ranges:
            self.range_counter += 1
            if start == end:
                range_list_entry = tk.Label(self.frame, text=f"range {self.range_counter}: Page: {start + 1}")
            else:
                range_list_entry = tk.Label(self.frame,
                                            text=f"range {self.range_counter}: Pages: {start + 1} - {end + 1}")
            range_list_entry.pack(pady=10)
            self.range_list.append(range_list_entry)
            self.ranges_to_split.append([start, end])

        for var, index in self.page_checkboxes:
            if var.get() == 1:
                var.set(0)

        # Force geometry update and resize the window
        self.root.update_idletasks()
        self.root.geometry(f'{self.root.winfo_width()}x{self.root.winfo_height()}')

    def split_selected(self):
        ranges = self.ranges_to_split
        if ranges:
            self.split_pdf(ranges)
            self.cleanup()

    def split_pdf(self, ranges):
        try:
            for i, (start, end) in enumerate(ranges):
                writer = PyPDF2.PdfWriter()
                for j in range(start, end + 1):
                    writer.add_page(self.reader.pages[j])

                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = os.path.join(self.output_dir, f"split_{timestamp}_{i + 1}.pdf")

                with open(output_filename, 'wb') as outfile:
                    writer.write(outfile)
            messagebox.showinfo("Success", "PDF split successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def cleanup(self):
        try:
            delete_original = messagebox.askyesno("Delete Original PDF",
                                                  "Delete the original PDF file after splitting?")
            if delete_original:
                all_pages_set = set(range(len(self.reader.pages)))
                selected_pages_set = set()
                for start, end in self.ranges_to_split:
                    selected_pages_set.update(range(start, end + 1))

                if selected_pages_set == all_pages_set:
                    self.reader.stream.close()
                    os.remove(self.pdf_path)
                else:
                    messagebox.showwarning("Warning",
                                           "Not all pages have been selected for splitting. Original PDF not deleted.")
            self.pdf_path = ""

            self.pages_to_split = []
            self.ranges_to_split = []

            for widget in self.checkboxes:
                widget.destroy()

            self.checkboxes = []
            self.page_checkboxes = []
            self.range_list = []

        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()
