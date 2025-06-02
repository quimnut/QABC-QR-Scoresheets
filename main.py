import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import csv
import qrcode
import numpy as np
# apt install python3-pil.imagetk mac needs to use python compiled against homebrews tk
from PIL import Image, ImageOps, ImageDraw, ImageFont, ImageTk
import fitz
import io
import operator
import cv2
from pyzbar.pyzbar import decode

def parseCSV(bcoemCSV):
    with open(bcoemCSV, encoding='latin') as csvfile:
        readCSV = csv.DictReader(csvfile, delimiter=',')
        entries = []
        for row in readCSV:
            entry = row
            entry['SubCategory'] = entry['Subcategory']
            entry['EntryNumber'] = entry['Judging Number']
            entry['SubCategoryName'] = entry['Style']
            entry['SpecialIngredients'] = entry['Required Info'].replace("\n", "; ").strip()
            entries.append(entry)
    return entries

def makeEntriesPreview(entries):
    condensed_list = []
    condensed_list_headers = ['Judging Number', 'Style', 'Required Info', 'Location', 'Table', 'Flight', 'Round']
    for row in entries:
        cl = [row['Judging Number'], row['Style'], row['SpecialIngredients'], row['Location'], row['Table'], row['Flight'], row['Round']]
        condensed_list.append(cl)
    return condensed_list, condensed_list_headers

def makeQR(entry,desired_size):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=3,
    )
    qr.add_data(entry['Judging Number'])
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
    old_size = img.size # resize
    ratio = float(desired_size)/max(old_size)
    new_size = tuple([int(x*ratio) for x in old_size])
    qr2 = img.resize(new_size, Image.LANCZOS)
    qr3 = Image.new("RGB", (desired_size, desired_size))
    qr3.paste(qr2, ((desired_size-new_size[0])//2, (desired_size-new_size[1])//2))
    qr4 = Image.new(qr3.mode, (desired_size, desired_size), (255,255,255)) 
    qr4.paste(qr3, (0, 0))
    return qr4 

def recommendTextSize(text): # this tries to fit in required info into its text box
    special_font_size = 7
    if (len(text) <= 50):
        special_font_size = 6
    elif (len(text) >= 70):
        special_font_size = 5
    elif (len(text) >= 90):
        special_font_size = 4
    elif (len(text) >= 110):
        special_font_size = 3
    elif (len(text) >= 135):
        special_font_size = 2
    return special_font_size

def dummyEntry():
    entry = {}
    entry['Judging Number'] = 'SAMPLE'
    entry['Location'] = 'Sample Day 1'
    entry['Table'] = '09: Table 9: Stoots'
    entry['Flight'] = '1'
    entry['Round'] = '1'
    entry['Style'] = 'Imperial Stout [BJCP 20C]'
    entry['SpecialIngredients'] = 'Sample Special Ingredients'
    entry['SubCategory'] = '04'
    entry['Category'] = '10'
    entry['EntryNumber'] = entry['Judging Number']
    entry['SubCategoryName'] = entry['Style']
    return entry

def previewPDF(filename):
    entry = dummyEntry()
    pdf = fitz.open(filename)
    page = pdf[0]
    qr = makeQR(entry,360)
    st = io.BytesIO()
    qr.save(st, 'PNG')
    rect = fitz.Rect(205,720,275,790)
    page.insert_image(rect, stream=st)
    special_font_size = recommendTextSize(entry['SpecialIngredients'])
    entry['FooterText'] = 'Location : ' + entry['Location'] + ' | Table : ' + entry['Table'] + ' | Flight : ' + entry['Flight']
    for widg in page.widgets():
        if widg.field_name in entry.keys():
            widg.field_value = entry[widg.field_name]
            widg.update()
        if (widg.field_name == 'SpecialIngredients'):
            widg.text_fontsize = special_font_size
            widg.update()
    pix = page.get_pixmap()
    img = pix.tobytes("ppm")
    pdf.close()
    return img

def generatePDF(entries, sheet, category, outputdir, number_of_pages):
    entries.sort(key=operator.itemgetter('Flight'))
    entries.sort(key=operator.itemgetter('Table'))
    entries.sort(key=operator.itemgetter('Round'))
    entries.sort(key=operator.itemgetter('Location'))
    outfile = fitz.Document() 
    pdf = fitz.open(sheet)
    for entry in entries:
        if not int(entry['Category']) == int(category):
            #print(str(entry['Category']) + ' ' + str(category))
            continue
        #print('Generating ' + entry['Judging Number'])
        qr = makeQR(entry,360)
        st = io.BytesIO()
        qr.save(st, 'PNG')
        rect = fitz.Rect(205,720,275,790)
        page = pdf[0]
        page.insert_image(rect, stream=st)
        special_font_size = recommendTextSize(entry['SpecialIngredients'])
        entry['FooterText'] = 'Location : ' + entry['Location'] + ' | Table : ' + entry['Table'] + ' | Flight : ' + entry['Flight']
        for widg in page.widgets():
            if widg.field_name in entry.keys():
                widg.field_value = entry[widg.field_name]
                widg.update()
            if (widg.field_name == 'SpecialIngredients'):
                widg.text_fontsize = special_font_size
                widg.update()
        mat = fitz.Matrix(4,4)  
        pix = page.get_pixmap(matrix = mat)
        png = pix.tobytes("png")
        for n in range(int(number_of_pages)):
            #print('Adding page ' + entry['Judging Number'] + ' ' + str(n))
            rect = fitz.paper_rect("a4")
            outpage = outfile.new_page(width=rect.width, height=rect.height)
            outpage.insert_image(rect, stream=png)
    o_filename = os.path.join(outputdir, f'category_{category}.pdf')
    outfile.save(o_filename, garbage=3, deflate=True)

def insertOrCreateScoresheet(fn, pageno, doc):
    snew = False
    if os.path.exists(fn):
        ss = fitz.open(fn)
    else:
        ss = fitz.Document()
        snew = True
    ss.insert_pdf(doc,from_page=pageno,to_page=pageno,start_at=-1)
    if snew:
        ss.save(fn, garbage=3, deflate=True)
    else:
        ss.save(fn,incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP, deflate=True)
    ss.close()

def rotate_image(image, angle):
  image_center = tuple(np.array(image.shape[1::-1]) / 2)
  rot_mat = cv2.getRotationMatrix2D(image_center, angle, 1.0)
  result = cv2.warpAffine(image, rot_mat, image.shape[1::-1], flags=cv2.INTER_LINEAR)
  return result

def sortScannedScoresheets(f_scanned, f_outputdir):
    doc = fitz.open(f_scanned)
    for page in range(doc.page_count):
        dox = doc[page]
        pix = dox.get_pixmap()
        pix = cv2.imdecode(np.frombuffer(pix.tobytes(), dtype=np.uint8), -1)
        oh, ow, oc = pix.shape
        h1 = ( oh / 5 ) * 4
        h1 = int(h1)
        h2 = int(oh)
        #h2 = int(h2)
        w1 = ( ow / 5 ) 
        w1 = int(w1)
        w2 = w1 * 3
        w2 = int(w2)
        img_gray = cv2.cvtColor(pix, cv2.COLOR_BGR2GRAY)
        for once in range(1):
            code = decode(img_gray)
            if code:
                judging_number = code[0].data
                continue
            img_crop = img_gray[h1:h2,w1:w2]
            code = decode(img_crop)
            if code:
                judging_number = code[0].data
                continue
            img_bw = cv2.adaptiveThreshold(img_gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY,11,2)
            code = decode(img_bw)
            if code:
                judging_number = code[0].data
                continue
            code = decode(pix)
            if code:
                judging_number = code[0].data
                continue
            count = 0
            while count != 359: # rotate it until we get it, works on a few stubborn cases
                count = count + 1
                img_new = rotate_image(img_crop, count)
                img_new = cv2.resize(img_new, (300,200))
                img_new = cv2.adaptiveThreshold(img_new,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY,11,2)
                code = decode(img_new)
                if code:
                    judging_number = code[0].data
                    continue
        try:
            judging_number
        except NameError:
            fn = os.path.join(f_outputdir, 'rejects.pdf')
            insertOrCreateScoresheet(fn, page, doc)
        else:
            fn = judging_number.decode("utf-8-sig") + '.pdf'
            out = os.path.join(f_outputdir, fn)
            insertOrCreateScoresheet(out, page, doc)
    doc.close()

def getNumberofPages(filename):
    pdf = fitz.open(filename)
    count = pdf.page_count
    pdf.close()
    return count

def deletePage(filename, page):
    pdf = fitz.open(filename)
    pdf.delete_page(page)
    if pdf.page_count == 0:
        pdf.close()
        os.remove(filename)
    else:
        pdf.save(filename, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP, deflate=True)
        pdf.close()

# window functions
def process_reject_tk(f_outputdir: str, page_number: int, prev_geometry: str = None):
    rejects_path = os.path.join(f_outputdir, 'rejects.pdf')
    if not os.path.isfile(rejects_path):
        messagebox.showerror("Error", f"No rejects.pdf found in:\n{f_outputdir}")
        return
    
    try:
        doc = fitz.open(rejects_path)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to open rejects.pdf:\n{e}")
        return

    real_idx = page_number - 1
    if real_idx < 0 or real_idx >= doc.page_count:
        messagebox.showerror("Error", f"Page {page_number} is out of range (1–{doc.page_count}).")
        doc.close()
        return

    page = doc[real_idx]
    pix = page.get_pixmap()
    ppm_data = pix.tobytes("ppm")
    doc.close()

    top = tk.Toplevel()
    top.title("Process Reject")
    top.resizable(False, False)

    button_frame = ttk.Frame(top, padding=(10, 10))
    button_frame.pack(side='top', fill='x')

    def on_delete():
        deletePage(rejects_path, real_idx)
        old_geom = top.geometry()
        top.destroy()
        update_rejected_controls()

        try:
            new_total = getNumberofPages(rejects_path)
        except Exception:
            new_total = 0

        if page_number <= new_total:
            process_reject_tk(f_outputdir, page_number, prev_geometry=old_geom)


    delete_btn = ttk.Button(button_frame, text="Delete This Page", command=on_delete)
    delete_btn.grid(row=0, column=0, padx=(0, 5))

    ttk.Label(button_frame, text="or").grid(row=0, column=1, padx=5)

    assigned_var = tk.StringVar()
    assign_entry = ttk.Entry(button_frame, width=6, textvariable=assigned_var)
    assign_entry.grid(row=0, column=2, padx=(0, 5))

    def on_correct():
        s = assigned_var.get().strip()
        if not s:
            messagebox.showerror("Error", "Please enter a 6-digit judging number.")
            return
        if len(s) < 6:
            s = s.zfill(6)

        new_filename = s + ".pdf"
        save_path = os.path.join(f_outputdir, new_filename)

        try:
            doc_here = fitz.open(rejects_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to re-open rejects.pdf:\n{e}")
            return

        try:
            insertOrCreateScoresheet(save_path, real_idx, doc_here)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write corrected file:\n{e}")
            doc_here.close()
            return

        doc_here.close()
        deletePage(rejects_path, real_idx)

        old_geom = top.geometry()
        top.destroy()
        update_rejected_controls()

        try:
            new_total = getNumberofPages(rejects_path)
        except Exception:
            new_total = 0

        if page_number <= new_total:
            process_reject_tk(f_outputdir, page_number, prev_geometry=old_geom)
        # else: no more pages

    correct_btn = ttk.Button(
        button_frame,
        text="Assign Judging Number",
        command=on_correct
    )
    correct_btn.grid(row=0, column=3, padx=(0, 5))

    def on_close():
        top.destroy()
        update_rejected_controls()

    close_btn = ttk.Button(button_frame, text="Close", command=on_close)
    close_btn.grid(row=0, column=4, padx=5)
    button_frame.columnconfigure(5, weight=1)

    img_frame = ttk.Frame(top, padding=(10, 0, 10, 10))
    img_frame.pack(side='top', fill='both', expand=True)

    ttk.Label(img_frame, text="Rejected Scoresheet:").pack(anchor='w')

    # Convert raw PPM bytes to a PhotoImage
    try:
        photo = tk.PhotoImage(data=ppm_data)
    except tk.TclError:
        # Fallback: write to a temp PPM file and reopen
        temp_ppm = os.path.join(f_outputdir, "_temp_reject.ppm")
        with open(temp_ppm, "wb") as tmpf:
            tmpf.write(ppm_data)
        photo = tk.PhotoImage(file=temp_ppm)
        os.remove(temp_ppm)

    img_label = ttk.Label(img_frame, image=photo)
    img_label.image = photo  
    img_label.pack(pady=(5, 0))

    top.update_idletasks()   

    if prev_geometry:
        top.geometry(prev_geometry)
    else:
        # Otherwise, center on screen
        w = top.winfo_width()
        h = top.winfo_height()
        sw = top.winfo_screenwidth()
        sh = top.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        top.geometry(f"{w}x{h}+{x}+{y}")

    top.transient()
    top.grab_set()
    top.wait_window()  


# Initialize main window
root = tk.Tk()
root.title("Scoresheet Management Application")
root.geometry("800x250")
root.configure(bg='#2a2a2a')

# Create notebook
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Style customization for dark theme
# style = ttk.Style()
# style.configure('TLabel', background='#2a2a2a', foreground='white')
# style.configure('TButton', background='#444', foreground='white')
# style.configure('TEntry', fieldbackground='#444', foreground='white')
# style.configure('TSpinbox', fieldbackground='#444', foreground='white')

# --- Scoresheet Generation ---
scoresheet_tab = ttk.Frame(notebook)
notebook.add(scoresheet_tab, text='Scoresheet Generation')

def create_input_row(parent, label_text):
    frame = ttk.Frame(parent)
    ttk.Label(frame, text=label_text).grid(row=0, column=0, sticky='w', padx=5)
    entry = ttk.Entry(frame)
    entry.grid(row=0, column=1, sticky='we', padx=5)

    def browse():
        file_path = filedialog.askopenfilename(title=f'Select {label_text}')
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def preview():
        file_path = entry.get()
        if not file_path:
            messagebox.showerror("Error", f"Please select a file for {label_text}.")
            return
        ext = os.path.splitext(file_path)[1].lower()

        if label_text.lower().startswith("entries paid"):
            try:
                raw_data = parseCSV(file_path)
                if not raw_data:
                    messagebox.showinfo("Empty CSV", "The CSV appears to have no rows.")
                    return

                columns = list(raw_data[0].keys())
                original_rows = [
                    [row[col] for col in columns] for row in raw_data
                ]

                popup = tk.Toplevel(parent)
                popup.title("CSV Preview: " + os.path.basename(file_path))
                popup.geometry("800x600")  # initial size; user can resize

                search_frame = ttk.Frame(popup, padding=(10, 5))
                search_frame.pack(fill="x")
                ttk.Label(search_frame, text="Search:").pack(side="left")
                search_var = tk.StringVar()
                search_entry = ttk.Entry(search_frame, textvariable=search_var)
                search_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

                table_frame = ttk.Frame(popup)
                table_frame.pack(fill="both", expand=True, padx=10, pady=(0,10))

                vsb = ttk.Scrollbar(table_frame, orient="vertical")
                hsb = ttk.Scrollbar(table_frame, orient="horizontal")

                tree = ttk.Treeview(
                    table_frame,
                    columns=columns,
                    show="headings",
                    yscrollcommand=vsb.set,
                    xscrollcommand=hsb.set,
                )
                vsb.config(command=tree.yview)
                hsb.config(command=tree.xview)

                vsb.pack(side="right", fill="y")
                hsb.pack(side="bottom", fill="x")
                tree.pack(side="left", fill="both", expand=True)

                def sort_column(col, reverse):
                    l = [(tree.set(k, col), k) for k in tree.get_children("")]
                    try:
                        l.sort(key=lambda t: float(t[0]) if t[0] not in ("", None) else float("-inf"), reverse=reverse)
                    except ValueError:
                        l.sort(key=lambda t: t[0].lower(), reverse=reverse)

                    for index, (_, k) in enumerate(l):
                        tree.move(k, "", index)

                    tree.heading(col, command=lambda: sort_column(col, not reverse))

                for col in columns:
                    tree.heading(col, text=col, command=lambda _col=col: sort_column(_col, False))
                    tree.column(col, width=100, anchor="w")

                for row_values in original_rows:
                    tree.insert("", "end", values=row_values)

                def filter_rows(*args):
                    search_text = search_var.get().strip().lower()
                    for item in tree.get_children():
                        tree.delete(item)

                    if search_text == "":
                        for rv in original_rows:
                            tree.insert("", "end", values=rv)
                        return

                    for rv in original_rows:
                        if any(search_text in (str(cell).lower()) for cell in rv):
                            tree.insert("", "end", values=rv)

                search_var.trace_add("write", filter_rows)

                popup.bind("<Escape>", lambda e: popup.destroy())
                return

            except Exception as e:
                messagebox.showerror("Error", f"Cannot read CSV file:\n{e}")
                return

        else:
            try:
                raw = previewPDF(file_path)
                pil_img = Image.open(io.BytesIO(raw))
                tk_img  = ImageTk.PhotoImage(pil_img)          

                popup = tk.Toplevel() 
                popup.title("Scoresheet Preview")

                img_label = ttk.Label(popup, image=tk_img)
                img_label.pack(padx=10, pady=10)
                img_label.image = tk_img   
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read file:\n{e}")

    ttk.Button(frame, text='Browse', command=browse).grid(row=0, column=2, padx=5)
    ttk.Button(frame, text='Preview', command=preview).grid(row=0, column=3, padx=5)

    frame.columnconfigure(1, weight=1)
    return entry, frame

inputs = {}
labels = ["Beer Scoresheet", "Cider Scoresheet", "Mead Scoresheet", "Entries Paid CSV w/ all flights assigned"]
for label in labels:
    entry, frame = create_input_row(scoresheet_tab, label)
    frame.pack(fill='x', padx=10, pady=3)
    inputs[label] = entry

output_frame = ttk.Frame(scoresheet_tab)
output_frame.pack(fill='x', padx=10, pady=10)

ttk.Label(output_frame, text='Output Folder').grid(row=0, column=0, sticky='w', padx=5)
output_entry = ttk.Entry(output_frame)
output_entry.grid(row=0, column=1, sticky='we', padx=5)

def browse_output_folder():
    folder_path = filedialog.askdirectory(title='Select Output Folder')
    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)

ttk.Button(output_frame, text='Browse', command=browse_output_folder).grid(row=0, column=2, padx=5)

ttk.Label(output_frame, text='Copies').grid(row=0, column=3, sticky='e', padx=5)
copies_spinbox = ttk.Spinbox(output_frame, from_=1, to=100, width=5)
copies_spinbox.set(1)
copies_spinbox.grid(row=0, column=4, padx=5)

def generate_pdf():
    f_entries    = inputs["Entries Paid CSV w/ all flights assigned"].get().strip()
    f_beersheet  = inputs["Beer Scoresheet"].get().strip()
    f_meadsheet  = inputs["Mead Scoresheet"].get().strip()
    f_cidersheet = inputs["Cider Scoresheet"].get().strip()
    f_outputdir = output_entry.get().strip()
    copies_text = copies_spinbox.get().strip()

    if not f_entries or not f_entries.lower().endswith(".csv"):
        messagebox.showerror("Missing/Invalid CSV", "Please select a valid Entries CSV file.")
        return

    for path, label in [
        (f_beersheet,  "Beer sheet PDF"),
        (f_meadsheet,  "Mead sheet PDF"),
        (f_cidersheet, "Cider sheet PDF"),
    ]:
        if not path or not path.lower().endswith(".pdf"):
            messagebox.showerror("Missing/Invalid Template", f"Please select a valid {label}.")
            return
        if not os.path.isfile(path):
            messagebox.showerror("File Not Found", f"{label} not found:\n{path}")
            return

    if not f_outputdir or not os.path.isdir(f_outputdir):
        messagebox.showerror("Missing/Invalid Output Dir", "Please select a valid output directory.")
        return

    try:
        number_of_sheets = int(copies_text)
        if number_of_sheets < 1:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid Copies", "Please enter a positive integer for Number of Sheets.")
        return

    try:
        entries = parseCSV(f_entries)
        if not entries:
            messagebox.showinfo("No Rows", "CSV parsed successfully but contains no rows.")
            return
    except Exception as e:
        messagebox.showerror("CSV Parse Error", f"Could not parse CSV:\n{e}")
        return

    try:
        for i in range(1, 21):
            if   i == 19:
                chosen_sheet = f_meadsheet
            elif i == 20:
                chosen_sheet = f_cidersheet
            elif i == 69: # custom category for aabc 2025 cider styles
                chosen_sheet = f_cidersheet
            else:
                chosen_sheet = f_beersheet

            generatePDF(entries, chosen_sheet, i, f_outputdir, number_of_sheets)

    except Exception as e:
        messagebox.showerror("Generation Error", f"Failed while generating PDFs:\n{e}")
        return

    messagebox.showinfo("Success", "All PDFs generated successfully.")

ttk.Button(output_frame, text='Generate PDF', command=generate_pdf).grid(row=0, column=5, padx=5)

output_frame.columnconfigure(1, weight=1)

# --- Scoresheet Processing ---
processing_tab = ttk.Frame(notebook)
notebook.add(processing_tab, text='Scoresheet Processing')

ttk.Label(processing_tab, text='Scanned Scoresheets').grid(row=0, column=0, sticky='w', padx=5, pady=(0,5))
scanned_entry = ttk.Entry(processing_tab)
scanned_entry.grid(row=0, column=1, sticky='we', padx=5, pady=(0,5))

def browse_scanned():
    file_path = filedialog.askopenfilename(title='Select Scanned Scoresheets')
    if file_path:
        scanned_entry.delete(0, tk.END)
        scanned_entry.insert(0, file_path)
        update_rejected_controls()

ttk.Button(processing_tab, text='Browse…', command=browse_scanned).grid(row=0, column=2, padx=5, pady=(0,5))

ttk.Label(processing_tab, text='Sorted PDF Folder').grid(row=1, column=0, sticky='w', padx=5, pady=(0,5))
sorted_var = tk.StringVar()
sorted_entry = ttk.Entry(processing_tab, textvariable=sorted_var)
sorted_entry.grid(row=1, column=1, sticky='we', padx=5, pady=(0,5))

def browse_sorted():
    folder = filedialog.askdirectory(title='Select Sorted PDF Folder')
    if folder:
        sorted_var.set(folder)
        update_rejected_controls()

ttk.Button(processing_tab, text='Browse…', command=browse_sorted).grid(row=1, column=2, padx=5, pady=(0,5))

# Row 2: “Process Scoresheets” button
def process_scanned(folder_path=None):
    if folder_path is None:
        folder_path = sorted_var.get().strip()
    scanned_path = scanned_entry.get().strip()
    if not scanned_path or not os.path.isfile(scanned_path):
        messagebox.showerror("Error", "Please select a valid scanned‐scoresheets PDF.")
        return
    if not folder_path or not os.path.isdir(folder_path):
        messagebox.showerror("Error", "Please select a valid folder to output sorted PDFs.")
        return

    sortScannedScoresheets(scanned_path, folder_path)
    update_rejected_controls()

ttk.Button(processing_tab, text='Process Scoresheets', command=lambda: process_scanned(sorted_var.get().strip())).grid(row=3, column=2, padx=5, pady=(10,0))

ttk.Label(processing_tab, text='Review rejected scans:').grid(row=4, column=0, sticky='w', padx=5, pady=(10,0))
page_var = tk.StringVar()
page_combobox = ttk.Combobox(processing_tab, textvariable=page_var, state='disabled', width=5)
page_combobox.grid(row=4, column=1, padx=5, pady=(10,0), sticky='w')

def review_page():
    folder = sorted_var.get().strip()
    if not folder or not os.path.isdir(folder):
        messagebox.showerror("Error", "Please select a valid sorted-PDF folder first.")
        return
    rejects_pdf = os.path.join(folder, 'rejects.pdf')
    if not os.path.isfile(rejects_pdf):
        messagebox.showerror("Error", "'rejects.pdf' not found in:\n" + folder)
        return
    try:
        page_num = int(page_var.get())
    except (ValueError, TypeError):
        messagebox.showerror("Error", "Please select a valid page number.")
        return

    process_reject_tk(folder, page_num)

review_btn = ttk.Button(processing_tab, text='Review', command=review_page, state='disabled')
review_btn.grid(row=4, column=2, padx=5, pady=(10,0))

processing_tab.columnconfigure(1, weight=1)

def update_rejected_controls():
    folder = sorted_var.get().strip()
    rejects_pdf = os.path.join(folder, 'rejects.pdf')
    if not folder or not os.path.isdir(folder) or not os.path.isfile(rejects_pdf):
        page_combobox['values'] = []
        page_combobox.config(state='disabled')
        review_btn.config(state='disabled')
        return

    try:
        total_pages = getNumberofPages(rejects_pdf)
    except Exception:
        total_pages = 0

    if total_pages <= 0:
        page_combobox['values'] = []
        page_combobox.config(state='disabled')
        review_btn.config(state='disabled')
    else:
        pages = [str(i) for i in range(1, total_pages+1)]
        page_combobox.config(state='readonly', values=pages)
        page_combobox.current(0)
        review_btn.config(state='normal')


update_rejected_controls()
root.mainloop()
