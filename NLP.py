#using Tkinter GUI to select the document
from tkinter import filedialog
from tkinter import *
import os
cwd = os.getcwd() # getting current working directory
root = Tk()
root.filename = filedialog.askopenfilename(initialdir = cwd, title = "Please navigate to & select the document", filetypes = (("all files","*.*"),("plain text","*.txt"),('pdf','*.pdf'),('word','*.docx')))

#obtaining text from file using Apache Tika
print('-'*50,'\nExtracting text using Apache Tika...\n','-'*50, sep='')
print("Connecting to the Tika REST Server. This could take a minute, but is only necessary once, after which the tika 'jar file' will be temporarily available at '/tmp/tika-server.jar'")
import tika
from tika import unpack
parsed = unpack.from_file(root.filename)
text=parsed["content"] #getting text content
print('\n\n\t','*'*3,' Successfully extracted text ', '*'*3,'\n\n', sep='')
print('-'*50,'\nProcessing Text With SpaCy...\n','-'*50, sep='')

#processing the text using spaCy 
import spacy
nlp = spacy.load("en_core_web_sm")
print('\nSpaCy model loaded.\n\nObtaining Named Entities...')
doc = nlp(text) #applying the sPacy pipeline

#obtaining entity information
ent_text=[]
ent_label=[]
ent_sentence=[]
for entity in doc.ents:
    ent_text.append(entity.text) # the entity's text
    ent_label.append(entity.label_) # the type of entity e.g. PERSON,GPE, etc
    ent_sentence.append(entity.sent.text.replace('\n',' ')) # some context on the obtained entity

ent_description=[spacy.explain(_label) for _label in ent_label] # a brief description of the entity type

print('\n\t','*** Extracted {} entities ***\n'.format(len(doc.ents)), sep='')
print('Saving results to excel file...\n\n')

#Saving the information as a dataframe
import pandas as pd
df=pd.DataFrame({'Entity': ent_text,'Type': ent_label,'Description': ent_description,'Context':ent_sentence})

#Separating info for each entity label. This gives a list of (label, label_dataframe) tuples.
entities=list(df.groupby('Type'))

#Exporting the entity-info into an excel file, with a sheet for each type:
save_dir = filedialog.askdirectory(title='Please select destination folder') # a GUI to select destination directory
saved_file=save_dir+'/text-mining.xlsx'
with pd.ExcelWriter(saved_file) as writer:
    for x in entities:
        x[1].to_excel(writer, sheet_name=x[0], index=False)
        
print('\t','*'*3,' Done! ','*'*3, '\n\nFile saved at: '+saved_file, sep='')

