#!/usr/bin/env python3
from tkinter import Tk, filedialog
import os
import pandas as pd
import spacy
from tika import unpack


def select_file():
    current_directory = os.getcwd()
    # Tk GUI to select document to process
    root = Tk()
    root.filename = filedialog.askopenfilename(
            initialdir=current_directory,
            title="Please select document",
            filetypes=(("all files", "*.*"),
                       ("plain text", "*.txt"),
                       ('pdf', '*.pdf'),
                       ('word', '*.docx'))
        )
    return root.filename


def custom_print(text):
    print(f'\n\t*** {text} ***')


def extract_text():
    text_file = select_file()
    # Extracting text
    custom_print('Extracting text using Apache Tika...')
    parsed = unpack.from_file(text_file, 'http://localhost:9998/')
    text = parsed["content"]
    custom_print('Successfully extracted text')
    return text


def process_text():
    text = extract_text()
    # Processing the text using spaCy
    custom_print('Processing Text With SpaCy...')
    nlp = spacy.load("en_core_web_sm")
    custom_print('SpaCy model loaded. Obtaining Named Entities...')
    doc = nlp(text)

    # Getting named entity information
    ent_text = []
    ent_labels = []
    ent_sentences = []
    for entity in doc.ents:
        ent_text.append(entity.text)  # the entity's text
        ent_labels.append(entity.label_)  # the type of entity
        # Getting some context on the obtained named-entity
        ent_sentences.append(entity.sent.text.replace('\n', ' '))

    custom_print(f'Extracted {len(doc.ents)} entities')
    custom_print('Saving results to excel file...')

    text_data = pd.DataFrame({'Entity': ent_text, 'Type': ent_labels,
                              'Context': ent_sentences})
    labels = text_data.Type.unique()
    ent_descriptions = [spacy.explain(_label) for _label in labels]
    Descriptions = pd.DataFrame({"Type": labels,
                                 "Description": ent_descriptions})

    # Tk GUI to select destination directory
    save_dir = filedialog.askdirectory(
        title='Please select destination folder'
    )
    output_file = save_dir + '/text-mining.xlsx'

    # Exporting the entity-info to an excel file (a sheet for each type)
    with pd.ExcelWriter('text-mining.xlsx') as writer:
        Descriptions.to_excel(writer, sheet_name="DEFINITIONS", index=False)
        for ent_name, data in text_data.groupby('Type'):
            data[["Entity", "Context"]].to_excel(writer, sheet_name=ent_name,
                                                 index=False)

    custom_print(f'Done! Saved file: {output_file!r}')


if __name__ == '__main__':
    process_text()
