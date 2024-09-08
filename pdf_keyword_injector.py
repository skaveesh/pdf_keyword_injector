import os
import pymupdf
from tkinter import Tk, Label, Entry, Checkbutton, IntVar, Button, OptionMenu, StringVar, messagebox, Frame

FONT_SIZE = 2
FONT_TIMES_ROMAN= "Times-Roman"

color_dict = {
    "white": {
        "RGB": (1, 1, 1),
        "HEX": 16777215 # RGB HEX value for white
    },
    "red": {
        "RGB": (1, 0, 0), 
        "HEX": 16711680 # RGB HEX value for red
    },
    "black": {
        "RGB": (0, 0, 0),
        "HEX": 0 # RGB HEX value for black
    },
}

dropdown_color_options = [ 
    "white", 
    "red", 
    "black",
] 

def main():
    # create root window
    root = Tk()

    # root window title and dimension
    root.title("PDF Keyword Injector v1.0 by @skaveesh")
    # Set geometry(width x height)
    root.geometry('640x420')

    # create a frame with padding
    frame = Frame(root, padx=20, pady=20)
    frame.grid(column=0, row=0)

    # adding a label to the root window
    lbl = Label(frame, text = "Enter the PDF path?")
    lbl.grid(column =0, row =0)

    # adding Entry Field
    txt = Entry(frame, width=100)
    txt.grid(column =0, row =1)

    input_pdf_file_path = get_first_pdf_file()

    if input_pdf_file_path is not None:
        lbl2 = Label(frame, text = "Default PDF file: " + input_pdf_file_path)
        lbl2.grid(column=0, row=2)

    Label(frame).grid(column=0, row=3, pady=1)

    lb3 = Label(frame, text = "Enter the output PDF path?")
    lb3.grid(column=0, row=4)

    txt2 = Entry(frame, width=100)
    txt2.grid(column =0, row =5)
    txt2.configure(state='disabled')

    Label(frame).grid(column=0, row=6, pady=1)

    # adding Checkbutton for overwrite PDF
    overwrite_pdf_check= IntVar()
    chk = Checkbutton(frame, text="Overwrite input PDF", variable=overwrite_pdf_check, onvalue=1, offvalue=0, command=lambda: toggle_output_file_path_text(overwrite_pdf_check, txt2))
    chk.select()  # Set default value to checked
    chk.grid(column=0, row=7)

    # adding Checkbutton for remove existing keywords
    remove_existing_keywords= IntVar()
    chk2 = Checkbutton(frame, text="Remove existing keywords", variable=remove_existing_keywords, onvalue=1, offvalue=0)
    chk2.select()  # Set default value to checked
    chk2.grid(column=0, row=8)

    Label(frame).grid(column=0, row=9, pady=1)

    lb4 = Label(frame, text = "Keyword color:")
    lb4.grid(column=0, row=10)

    dropdown_selected_color_option = StringVar() 
    dropdown_selected_color_option.set(dropdown_color_options[0])
    drop = OptionMenu(frame , dropdown_selected_color_option , *dropdown_color_options) 
    drop.grid(column=0, row=11)

    Label(frame).grid(column=0, row=12, pady=1)

    lb5 = Label(frame, text = "Enter the text to write?")
    lb5.grid(column=0, row=13)

    txt3 = Entry(frame, width=100)
    txt3.grid(column =0, row =14)

    Label(frame).grid(column=0, row=15, pady=1)

    # button widget with red color text inside
    btn = Button(frame, text = "Modify PDF" , fg = "red", command=lambda: validate_and_modify(root, txt, txt2, txt3, input_pdf_file_path, overwrite_pdf_check, dropdown_selected_color_option, remove_existing_keywords))
    # Set Button Grid
    btn.grid(column=0, row=16)

    # Execute Tkinter
    root.mainloop()


def toggle_output_file_path_text(overwrite_pdf_check, txt2):
    if overwrite_pdf_check.get() == 0:
        txt2.configure(state='normal')
    else:
        txt2.configure(state='disabled')


def validate_and_modify(root, txt, txt2, txt3, input_pdf_file_path, overwrite_pdf_check, dropdown_selected_color_option, remove_existing_keywords):
    
    if input_pdf_file_path is None and txt.get() == "":
        root.withdraw()
        messagebox.showerror("Error", "No PDF file found in the current directory and the PDF path is empty.")
        root.deiconify()
        return
    
    if not txt.get() == "":
        validate_file_path(root, txt.get(), "Invalid input file path")
        input_pdf_file_path = txt.get()

    if overwrite_pdf_check.get() == 0:
        output_pdf_file_path = txt2.get()
    else:
        output_pdf_file_path = input_pdf_file_path

    try:
        modify_pdf_file(input_pdf_file_path, output_pdf_file_path, txt3.get(), dropdown_selected_color_option, remove_existing_keywords)
        messagebox.showinfo("Success", "The file was saved successfully.")
    except Exception as e:
        print(e)
        root.withdraw()
        messagebox.showerror("Error occurred while saving the file", str(e))
        root.deiconify()


def modify_pdf_file(input_pdf_file_path, output_pdf_file_path, text_to_write, dropdown_selected_color_option, remove_existing_keywords_var):
    # Open the input PDF file
    pdf_document = pymupdf.open(input_pdf_file_path)

    # Remove existing keywords
    if remove_existing_keywords_var.get() == 1:
        remove_existing_keywords(pdf_document)
    
    # Get the first page
    first_page = pdf_document[0]
    
    # Define the position (bottom of the page)
    page_width = first_page.rect.width
    page_height = first_page.rect.height
    x_position = 10  # 10 points from the left
    y_position = page_height - 10  # 10 points from the bottom
    
    # Add the text to the first page
    first_page.insert_text((x_position, y_position), text_to_write, fontname=FONT_TIMES_ROMAN, fontsize=FONT_SIZE, color=color_dict[dropdown_selected_color_option.get()]["RGB"])

    incremental_var = False
    if input_pdf_file_path == output_pdf_file_path:
        incremental_var = True

    # Save the modified PDF to the output file path
    pdf_document.save(output_pdf_file_path, incremental=incremental_var, encryption=pymupdf.PDF_ENCRYPT_KEEP)
    pdf_document.close()


def remove_existing_keywords(pdf_document):
    first_page = pdf_document.load_page(0)

    # Extract text blocks
    text_blocks = first_page.get_text("dict")["blocks"]
    
    # Iterate through each block
    for block in text_blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    if span["size"] == FONT_SIZE and span["font"] == FONT_TIMES_ROMAN:
                        for color_option in dropdown_color_options:
                            if span["color"] == color_dict[color_option]["HEX"]:
                                print(f"\nremoved text::: {span['text']}")
                                rect = pymupdf.Rect(span["bbox"])
                                first_page.add_redact_annot(rect)

    first_page.apply_redactions()


def get_first_pdf_file(): 
    pdf_files = [file for file in os.listdir() if file.endswith(".pdf")] 
    if pdf_files: 
        return pdf_files[0] 
    else: 
        return None


def validate_file_path(root, file_path, err_msg):
    if not os.path.isfile(file_path):
        root.withdraw()
        messagebox.showerror("Error", err_msg)
        root.deiconify()
        raise FileNotFoundError(f"{file_path} does not exist.")
        

if __name__ == "__main__":
    main()