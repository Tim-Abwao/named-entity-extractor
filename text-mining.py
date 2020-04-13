#!/usr/bin/env python3
from tkinter import Tk, filedialog
import os
import pandas as pd
import spacy
from tika import unpack

cwd = os.getcwd()  # getting current working directory
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
doc = nlp(text)  # applying sPacy NLP pipeline

# Getting named-entity information
ent_text = []
ent_label = []
ent_sentence = []
for entity in doc.ents:
    ent_text.append(entity.text)  # the entity's text
    ent_label.append(entity.label_)  # the type of entity (PERSON, ORG, etc)
    # Getting some context on the obtained named-entity
    ent_sentence.append(entity.sent.text.replace('\n', ' '))
# Getting brief descriptions of each entity type
ent_description = [spacy.explain(_label) for _label in ent_label]

print('\n\t', '*** Extracted {} entities ***\n'.format(len(doc.ents)), sep='')
print('Saving results to excel file...\n\n')

# Combining obtained information
df = pd.DataFrame({'Entity': ent_text, 'Type': ent_label,
                   'Description': ent_description, 'Context': ent_sentence})

# Tk GUI to select destination directory
save_dir = filedialog.askdirectory(title='Please select destination folder')
output_file = save_dir + '/text-mining.xlsx'

# Exporting the entity-info into an excel file, with a sheet for each type:
entities = list(df.groupby('Type'))  # a list of (type, dataframe) tuples.
with pd.ExcelWriter(output_file) as writer:
    for ent_name, data in entities:
        data.to_excel(writer, sheet_name=ent_name, index=False)

print('\t', '*'*3, ' Done! ', '*'*3, '\n\nSaved file: ' + output_file, sep='')
