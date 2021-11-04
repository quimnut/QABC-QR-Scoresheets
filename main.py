import PySimpleGUI as sg
import os
import csv
import qrcode
import numpy as np
from PIL import Image, ImageOps, ImageDraw, ImageFont
import fitz
import io
import operator
import cv2
from pyzbar.pyzbar import decode


# old function that reads entries from a CSV file
def parseCSV(bcoemCSV):
    with open(bcoemCSV, encoding='utf-8-sig') as csvfile:
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
    qr2 = img.resize(new_size, Image.ANTIALIAS)
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
     


main_layout = [[
    sg.Frame(layout=[[
        sg.Text("Beer Scoresheet"),
        sg.In(size=(25, 1), enable_events=True, key="-BEERSHEET-"),
        sg.FileBrowse(),
        sg.Button("Preview", enable_events=True, key="-PREVIEWBEERSHEET-")
    ],
    [
        sg.Text("Cider Scoresheet"),
        sg.In(size=(25, 1), enable_events=True, key="-CIDERSHEET-"),
        sg.FileBrowse(),
        sg.Button("Preview", enable_events=True, key="-PREVIEWCIDERSHEET-")
    ],
    [
        sg.Text("Mead Scoresheet"),
        sg.In(size=(25, 1), enable_events=True, key="-MEADSHEET-"),
        sg.FileBrowse(),
        sg.Button("Preview", enable_events=True, key="-PREVIEWMEADSHEET-")
    ],
    [
        sg.Text("Entries Paid CSV w/ all flights assigned"),
        sg.In(size=(25, 1), enable_events=True, key="-ENTRIES-"),
        sg.FileBrowse(),
        sg.Button("Preview", enable_events=True, key="-PREVIEWCSV-")
    ]], title='Inputs:',element_justification='right', pad=(0,0)) 
    ],
    [
    sg.Frame(layout=[[
        sg.Text("Output Folder"),
        sg.In(size=(25, 1), enable_events=True, key="-OUTPUTDIR-"),
        sg.FolderBrowse(),
        sg.Text('Copies'),sg.Combo([1, 2, 3], default_value=1, key='-COPIES-'),
        sg.Button("Generate PDF", enable_events=True, key="-OUTPUTBEERSHEETS-")
    ]], title='Output:',element_justification='right', pad=(0,0))
    ],
    [
    sg.Frame(layout=[[
        sg.Text("Scanned Scoresheets"),
        sg.In(size=(25, 1), enable_events=True, key="-SCANFILE-"),
        sg.FileBrowse()
    ],[
        sg.Text("Sorted PDF Folder"),
        sg.In(size=(25, 1), enable_events=True, key="-SORTEDOUTPUTDIR-"),
        sg.FolderBrowse() 
    ],[
        sg.Button("Go", enable_events=True, key="-SORTSCORESHEETS-")
    ]], title='Process Scanned Scoresheets:',element_justification='right', pad=(0,0))
    ]
    ]


if __name__ == '__main__':
    sg.theme('DarkBlue14')
    window = sg.Window(title="ScoreSheet Wizard", layout=main_layout, margins=(100, 50))

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED: 
            break
        elif event == "-PREVIEWBEERSHEET-":
            filename = values['-BEERSHEET-']
            if os.path.isfile(filename):
                img = previewPDF(filename)
                sg.popup_no_buttons(image=img)
            else:
                filename = sg.popup_get_file("", no_window=True)
                if filename == '':
                    break
                window['-BEERSHEET-'].update(filename)
        elif event == "-PREVIEWMEADSHEET-":
            filename = values['-MEADSHEET-']
            if os.path.isfile(filename):
                img = previewPDF(filename)
                sg.popup_no_buttons(image=img)
            else:
                filename = sg.popup_get_file("", no_window=True)
                if filename == '':
                    break
                window['-MEADSHEET-'].update(filename)
        elif event == "-PREVIEWCIDERSHEET-":
            filename = values['-CIDERSHEET-']
            if os.path.isfile(filename):
                img = previewPDF(filename)
                sg.popup_no_buttons(image=img)
            else:
                filename = sg.popup_get_file("", no_window=True)
                if filename == '':
                    break
                window['-CIDERSHEET-'].update(filename)                
        elif event == "-PREVIEWCSV-":
            filename = values['-ENTRIES-']
            if os.path.isfile(filename):
                entries = parseCSV(filename)
                condensed_list, condensed_list_headers = makeEntriesPreview(entries)
                layout = [[sg.Table(values=condensed_list,
                    headings=condensed_list_headers,
                    max_col_width=50,
                    auto_size_columns=True,
                    justification='left',
                    # alternating_row_color='lightblue',
                    num_rows=min(len(condensed_list), 20))]]
                sg.Window('Preview', layout, keep_on_top=True).read()
        elif event == "-OUTPUTBEERSHEETS-":
            f_entries = values['-ENTRIES-']
            if not os.path.isfile(f_entries):
                sg.popup_error("Entries CSV not found")
                break
            f_beersheet = values['-BEERSHEET-']
            if not os.path.isfile(f_beersheet):
                sg.popup_error("Beer Scoresheet not found")
                break
            f_meadsheet = values['-MEADSHEET-']
            if not os.path.isfile(f_meadsheet):
                sg.popup_error("Mead Scoresheet not found")
                break
            f_cidersheet = values['-CIDERSHEET-']
            if not os.path.isfile(f_cidersheet):
                sg.popup_error("Cider Scoresheet not found")
                break
            f_outputdir = values['-OUTPUTDIR-']
            if not os.path.isdir(f_outputdir):
                sg.popup_error("Output folder not found")
                break
            number_of_sheets = int(values['-COPIES-'])
            entries = parseCSV(f_entries)
            for i in range(1,21):
                if i == 19:
                    chosen_sheet = f_meadsheet
                elif i == 20:
                    chosen_sheet = f_cidersheet
                else:
                    chosen_sheet = f_beersheet
                generatePDF(entries, chosen_sheet, i, f_outputdir, number_of_sheets)
        elif event == "-SORTSCORESHEETS-":
            f_scanned = values['-SCANFILE-']
            if not os.path.isfile(f_scanned):
                sg.popup_error("Scanned Scoresheets not found")
                break
            f_outputdir = values['-SORTEDOUTPUTDIR-']
            if not os.path.isdir(f_outputdir):
                sg.popup_error("Sorted PDF folder not found")
                break
            sortScannedScoresheets(f_scanned, f_outputdir)
            
    window.close()
