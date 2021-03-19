#!/usr/bin/env python3
from tkinter.filedialog import askopenfilename, asksaveasfilename
import pandas as pd
from pathlib import Path
import spacy
import textract


def select_file():
    """Create a file dialog to help navigate to and select a file.

    Returns:
        `pathlib.Path`: The path to the file.
    """
    # select file
    file = askopenfilename(
        initialdir='.', title='Please select a document',
        filetypes=(('pdf', '*.pdf'),
                   ('text', '*.txt'),
                   ('word', '*.docx'),
                   ('all files', '*.*'))
        )
    if not file:  # If file dialog is closed with no file selected
        return None
    else:
        return Path(file)


def get_text_from_file():
    """Read and parse the contents of a file as a string.

    Returns:
        str : The file's contents as one continuous string
    """
    filepath = select_file()

    try:
        text = textract.process(filepath)
    except TypeError:
        return None

    return text.decode()


def extract_entity_info(text):
    """Get information on the named entities present in the text.

    Args:
        text (str): The text to process

    Returns:
        `pandas.DataFrame`: A dataframe with 3 columns - 'Entity', 'Type' and
        'Context'.
    """
    # Load the pretrained spaCy model
    nlp = spacy.load('en_core_web_md')

    # Get entity predictions
    doc = nlp(text)
    entity_info = ((entity.text, entity.label_, entity.sent.text)
                   for entity in doc.ents)

    entity_df = pd.DataFrame(entity_info,
                             columns=['Entity', 'Type', 'Context'])

    # Remove 'newline' to get continuous text
    return entity_df.applymap(lambda x: x.replace('\n', ' '))


def get_entity_descriptions(entity_df):
    """Get a dataframe of entity descriptions.

    Args:
        entity_df (`pandas.DataFrame`): A dataframe of entity information with
        3 columns - 'Entity', 'Type' and 'Context'.

    Returns:
        `pandas.DataFrame`: A pandas dataframe with entity types and their
        respective descriptions.
    """

    entity_types = entity_df['Type'].unique()
    return pd.DataFrame(
        {'Type': entity_types,
         'Description': map(spacy.explain, entity_types)})


def select_file_destination():
    """Create a file dialog to help select a destination filename.

    Returns:
        `pathlib.Path`: The path to a file.
    """
    # Get desired output-file-name
    output_file = asksaveasfilename(
        initialdir='.', initialfile='text_results.xlsx',
        filetypes=[('Excel', '.xlsx')]
    )
    if not output_file:  # If file dialog is closed with no file selected
        return Path('entity-info.xlsx')
    else:
        return Path(output_file)


def save_results_to_excel(entity_df):
    """Save extracted information as an excel file, with a sheet for each
    entity type.

    Args:
        entity_df (`pandas.DataFrame`): A dataframe of entity information with
        3 columns - 'Entity', 'Type' and 'Context'.
    Returns:
        str : The path to the created file.
    """
    # Get destination filename
    output_file = select_file_destination()

    with pd.ExcelWriter(output_file) as writer:
        # Get entity descriptions & write them as the first sheet
        entity_descriptions = get_entity_descriptions(entity_df)
        entity_descriptions.to_excel(writer, index=False,
                                     sheet_name='DEFINITIONS')
        # Create a sheet for each entity type
        for ent_name, ent_data in entity_df.groupby('Type'):
            (ent_data[['Entity', 'Context']]
                .dropna()
                .to_excel(writer, index=False, sheet_name=ent_name))
    return output_file.name
