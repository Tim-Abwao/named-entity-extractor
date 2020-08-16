#!/usr/bin/env python3
import pandas as pd
import textract
import spacy
from tkinter import Tk, filedialog, messagebox
from tkinter import ttk

root = Tk()
root.title('Simple Text Mining App')


def open_file():
    filename = filedialog.askopenfilename(initialdir='.',
                                          title="Please select document",
                                          filetypes=(("all files", "*.*"),
                                                     ("plain text", "*.txt"),
                                                     ('pdf', '*.pdf'),
                                                     ('word', '*.docx')))
    return filename


def process_text():
    """
    Makes predictions about named entities using spaCy's small core model,
    and saves the results as an excel file.
    """
    # initialise progress bar
    progress_bar = ttk.Progressbar(root, orient='horizontal', length=400,
                                   mode='determinate', value=5)
    progress_bar.place(relx=0.15, rely=0.7)

    # select file and obtain it's contents
    file_name = open_file()
    text = textract.process(file_name).decode()
    progress_bar['value'] = 15

    # loading the model
    nlp = spacy.load("en_core_web_sm")

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

    extracted_info = pd.DataFrame({'Entity': ent_text, 'Type': ent_labels,
                                   'Context': ent_sentences})
    labels = extracted_info.Type.unique()
    ent_descriptions = [spacy.explain(_label) for _label in labels]
    Descriptions = pd.DataFrame({"Type": labels,
                                 "Description": ent_descriptions})

    # getting destination
    output_file = filedialog.asksaveasfilename()

    # Saving the extracted_info as an excel file (a sheet for each type)
    with pd.ExcelWriter(output_file) as writer:
        Descriptions.to_excel(writer, sheet_name="DEFINITIONS", index=False)

        for ent_name, data in extracted_info.groupby('Type'):
            data[["Entity", "Context"]].to_excel(writer, sheet_name=ent_name,
                                                 index=False)

    messagebox.showinfo(message=f'Done! Results saved as {output_file!r}')


frame = ttk.Frame(root, width=600, height=400)
frame['padding'] = (5, 10)
frame['borderwidth'] = 2

intro_text = "This is a simple application useful for extracting text from " +\
             "files in various formats."
intro = ttk.Label(frame, text=intro_text,  wraplength=480)
intro.place(relx=0.07, rely=0.1, relwidth=0.86, relheight=0.2)

file_select_button = ttk.Button(frame, text='Select file', width=30,
                                command=process_text)
file_select_button.place(relx=0.3, rely=0.4)

# file_save_button = ttk.Button(frame, text='Select file', width=30,
#                                 command=process_text)
# file_save_button.place(relx=0.3, rely=0.3)

frame.pack()


root.mainloop()
