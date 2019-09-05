#using Tkinter graphical user interface to select text file
from tkinter import filedialog
from tkinter import *
import os
cwd = os.getcwd() # get current working directory
root = Tk()
root.filename = filedialog.askopenfilename(initialdir = cwd, title = "Select file", filetypes = (("all files","*.*"),("plain text","*.txt"),('pdf','*.pdf')))

#obtaining text from file using Apache Tika
print('--------------------------------------\nExtracting Text With Tika...\n--------------------------------------')
import tika
from tika import unpack
parsed = unpack.from_file(root.filename)
text=parsed["content"]
print('--------------------------------------\nSuccessfully Extracted Text\n--------------------------------------\n\n')
print('--------------------------------------\nProcessing Text With SpaCy..\n--------------------------------------')

#processing the text using spaCy 
import spacy
nlp = spacy.load("en_core_web_sm")
print('SpaCy model imported.\n\nObtaining Named Entities...')
doc = nlp(text)

#obtaining entity information
ent_text=[]
ent_label=[]
ent_sentence=[]
for entity in doc.ents:
    ent_text.append(entity.text) #the entity's text
    ent_label.append(entity.label_) # the type of entity, a string e.g. PERSON,GPE, etc
    ent_sentence.append(entity.sent.text.replace('\n',' ')) # a sentence with the entity, as one line with newline - \n - removed

ent_description=[spacy.explain(_label) for _label in ent_label] # a brief description of the entity type

print('\n*** Extracted {} entities ***\n'.format(len(doc.ents)))
print('Saving entity info to file...')

#Saving the information as a dataframe
import pandas as pd
df=pd.DataFrame({'Entity': ent_text,'Type': ent_label,'Description': ent_description,'Context':ent_sentence})

#Separating info for each entity label. This gives a list of (label, label_dataframe) tuples.
entities=list(df.groupby('Type'))

#Exporting the entity-info into an excel file, with a sheet for each type:
save_dir = filedialog.askdirectory(title='Select destination directory') # a GUI to select destination directory
saved_file=save_dir+'/text-info.xlsx'
with pd.ExcelWriter(saved_file) as writer:
    for x in entities:
        x[1].to_excel(writer, sheet_name=x[0], index=False)
        
print('--------------------------------------\nDone! File saved at '+saved_file+'\n--------------------------------------')

