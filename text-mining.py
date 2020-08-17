#!/usr/bin/env python3
import pandas as pd
import textract
import spacy
from tkinter import Tk, filedialog, messagebox
from tkinter import ttk

root = Tk()
root.title('Simple Text Mining App')
root.resizable(False, False)

style = ttk.Style()
style.configure("TFrame", background="ivory")
style.configure("TLabel", foreground="darkslategrey", background="ivory",
                font="serif 14")
style.configure("I.TLabel", foreground="darkslategrey", background="ivory",
                font="serif 13 normal italic ")
style.configure("TButton", foreground="slategrey", background="aquamarine",
                font="serif 12")


extensions = ['.pdf', '.csv', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.odt',
              '.json', '.htm', '.html', '.tsv',  '.pptx', '.epub', '.log',
              '.rtf', '.jpeg', '.jpg', '.gif', '.ogg',  '.png', '.msg', '.wav',
              '.eml', '.mp3', '.ps', '.psv', '.tff', '.tif', '.tiff'
              ]


def open_file():
    filename = filedialog.askopenfilename(initialdir='.',
                                          title="Please select document",
                                          filetypes=(('pdf', '*.pdf'),
                                                     ("plain text", "*.txt"),
                                                     ('word', '*.docx'),
                                                     ("all files", "*.*"))
                                          )
    return filename


def get_text(progress):
    # select file and obtain it's contents
    progress.place(relx=0.15, rely=0.7)
    progress['value'] = 5
    file_name = open_file()

    if not file_name:  # if no file is selected
        messagebox.showinfo(message="Please select a file to proceed")
        progress.destroy()

    try:
        text = textract.process(file_name).decode()
        return text
    except UnicodeDecodeError:
        messagebox.showerror(message='Unable to parse file')


def process_text():
    """
    Makes predictions about named entities using spaCy's small core model,
    and saves the results as an excel file.
    """
    # initialise progress bar
    progress_bar = ttk.Progressbar(root, orient='horizontal', length=400,
                                   mode='determinate', value=5)
    progress_bar.place(relx=0.15, rely=0.7)
    progress_bar['value'] = 5

    text = get_text(progress_bar)
    if not text:
        messagebox.showwarning(message="Couldn't find text to process.")
        progress_bar.destroy()

    progress_bar['value'] = 15
    # loading the model
    nlp = spacy.load("en_core_web_sm")
    progress_bar['value'] = 35

    # Getting named entity information
    doc = nlp(text)
    ent_text = []
    ent_labels = []
    ent_sentences = []

    for entity in doc.ents:
        ent_text.append(entity.text)  # the entity's text
        ent_labels.append(entity.label_)  # the type of entity
        # Getting some context on the obtained named-entity
        ent_sentences.append(entity.sent.text.replace('\n', ' '))

    progress_bar['value'] = 75

    extracted_info = pd.DataFrame({'Entity': ent_text, 'Type': ent_labels,
                                   'Context': ent_sentences})
    labels = extracted_info.Type.unique()
    ent_descriptions = [spacy.explain(_label) for _label in labels]
    Descriptions = pd.DataFrame({"Type": labels,
                                 "Description": ent_descriptions})

    progress_bar['value'] = 90

    # getting save-as name
    output_file = filedialog.asksaveasfilename(initialdir='.',
                                               initialfile="text_results.xlsx",
                                               filetypes=[("Excel", '.xlsx')])

    # Saving the extracted_info as an excel file (a sheet for each type)
    with pd.ExcelWriter(output_file) as writer:
        Descriptions.to_excel(writer, sheet_name="DEFINITIONS", index=False)

        for ent_name, data in extracted_info.groupby('Type'):
            data[["Entity", "Context"]].to_excel(writer, sheet_name=ent_name,
                                                 index=False)

    messagebox.showinfo(message=f'Done! Results saved as {output_file!r}')
    progress_bar.destroy()  # remove progress_bar after completion


frame = ttk.Frame(root, width=600, height=350)
frame['padding'] = (5, 10)
frame['borderwidth'] = 2

intro_text = "This is a simple application useful for extracting " +\
             "information from text files in a variety of formats." +\
             " Supported extensions include:"
intro = ttk.Label(frame, text=intro_text,  wraplength=480)
intro.place(relx=0.07, rely=0.1, relwidth=0.86, relheight=0.25)

supported_extensions = ttk.Label(frame, text='  '.join(extensions),
                                 wraplength=350, style='I.TLabel')
supported_extensions.place(relx=0.2, rely=0.35)

file_select_button = ttk.Button(frame, text='Select file', width=25,
                                command=process_text)
file_select_button.place(relx=0.3, rely=0.8)


frame.pack()
root.bind('<Return>', process_text)
root.mainloop()
