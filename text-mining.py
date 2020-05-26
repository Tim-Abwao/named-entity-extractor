#!/usr/bin/env python3
from tkinter import Tk, filedialog
import os
import pandas as pd
import spacy
from tika import unpack

cwd = os.getcwd() 
# Tk GUI to select document to process
root = Tk()
root.filename = filedialog.askopenfilename(initialdir=cwd,
                                           title="Please select document",
                                           filetypes=(("all files", "*.*"),
                                                      ("plain text", "*.txt"),
                                                      ('pdf', '*.pdf'),
                                                      ('word', '*.docx')))

# Extracting text
print('-'*50, '\nExtracting text using Apache Tika...\n', '-'*50, sep='')
parsed = unpack.from_file(root.filename, 'http://localhost:9998/')
text = parsed["content"]
print('\n\n\t', '*'*3, ' Successfully extracted text ', '*'*3, '\n\n', sep='')

# Processing the text using spaCy
print('-'*50, '\nProcessing Text With SpaCy...\n', '-'*50, sep='')
nlp = spacy.load("en_core_web_sm")
print('\nSpaCy model loaded.\n\nObtaining Named Entities...')
doc = nlp(text)  

# Getting named-entity information
ent_text = []
ent_labels = []
ent_sentences = []
for entity in doc.ents:
    ent_text.append(entity.text)  # the entity's text
    ent_labels.append(entity.label_)  # the type of entity (PERSON, ORG, etc)
    # Getting some context on the obtained named-entity
    ent_sentences.append(entity.sent.text.replace('\n', ' '))

print('\n\t', '*** Extracted {} entities ***\n'.format(len(doc.ents)), sep='')
print('Saving results to excel file...\n\n')


text_data = pd.DataFrame({'Entity': ent_text, 'Type': ent_labels, 
                          'Context': ent_sentences})
labels = text_data.Type.unique()
ent_descriptions = [spacy.explain(_label) for _label in labels]
Descriptions = pd.DataFrame({"Type": labels, "Description": ent_descriptions})

# Tk GUI to select destination directory
save_dir = filedialog.askdirectory(title='Please select destination folder')
output_file = save_dir + '/text-mining.xlsx'

# Exporting the entity-info to an excel file (a sheet for each type)
with pd.ExcelWriter('text-mining.xlsx') as writer:
    Descriptions.to_excel(writer, sheet_name="DEFINITIONS", index=False)
    for ent_name, data in text_data.groupby('Type'):
        data[["Entity","Context"]].to_excel(writer, sheet_name=ent_name, 
                                            index=False)

print('\t', '***', ' Done! ', '***', '\n\nSaved file: ' + output_file, sep='')
