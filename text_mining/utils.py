#!/usr/bin/env python3
import logging
from pathlib import Path
from tkinter.filedialog import askopenfilename, asksaveasfilename

import pandas as pd
import spacy
import textract

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s"
)


def get_input_file():
    """Create a file-dialog to navigate to and select a file.

    Returns
    -------
    The path to the file.
    """
    file = askopenfilename(
        initialdir=".",
        title="Please select a document",
        filetypes=(
            ("pdf", "*.pdf"),
            ("text", "*.txt"),
            ("word", "*.docx"),
            ("all files", "*.*"),
        ),
    )
    if not file:  # If file dialog is closed with no file selected
        return None
    else:
        logging.info(f"Found file {file!r}")
        return Path(file)


def read_text_from_file():
    """Get the contents of a file as a string.

    Returns
    -------
    The file's contents as one continuous string.
    """
    filepath = get_input_file()

    if filepath is None:
        return None
    else:
        try:
            # Get the file's contents as bytes
            text = textract.process(filepath)
        except UnicodeDecodeError:
            return None

    return text.decode()  # Convert bytes to string


def extract_entity_info(text):
    """Get the named entities present in the text.

    Parameters
    ----------
    text : str
        The text to process.

    Returns
    -------
    A dictionary of entity predictions and descriptions.
    """
    # Load the pre-trained spaCy model
    logging.info("Analysing extracted text...")
    nlp = spacy.load("en_core_web_md")

    # Get entity predictions
    doc = nlp(text)
    entity_info = (
        (entity.text, entity.label_, entity.sent.text) for entity in doc.ents
    )
    entity_predictions = (
        pd.DataFrame(entity_info, columns=["Entity", "Type", "Context"])
        # Remove 'newline' to get continuous text
        .applymap(lambda x: x.replace("\n", " "))
    )

    logging.info(f"Done. {len(entity_predictions):,} named entities found.")

    # Get entity-type descriptions
    entity_types = entity_predictions["Type"].unique()
    entity_descriptions = pd.DataFrame(
        {"Type": entity_types, "Description": map(spacy.explain, entity_types)}
    )
    return {
        "predictions": entity_predictions,
        "descriptions": entity_descriptions,
    }


def set_output_file():
    """Create a file dialog to select a destination for the results.

    Returns
    -------
    The path to the output-file.
    """
    output_file = asksaveasfilename(
        initialdir=".",
        initialfile="text_results.xlsx",
        filetypes=[("Excel", ".xlsx")],
    )
    if not output_file:  # If file dialog is closed with no file selected
        return Path("entity-info.xlsx")
    else:
        return Path(output_file)


def save_results_to_excel(entity_info):
    """Save extracted information as an excel file, with a sheet for each
    entity type.

    Parameters
    ----------
    entity_info : dict
        A dictionary with entity predictions and descriptions.

    Returns
    -------
    The path to the output file.
    """
    # Get destination filename
    output_file = set_output_file()

    entity_predictions = entity_info["predictions"]
    entity_descriptions = entity_info["descriptions"]

    # Write entity info to an excel file
    with pd.ExcelWriter(output_file) as writer:
        # Write entity descriptions in the first sheet
        entity_descriptions.to_excel(
            writer, index=False, sheet_name="DEFINITIONS"
        )
        # Create a sheet for each entity predicted type
        for ent_name, ent_data in entity_predictions.groupby("Type"):
            (
                ent_data[["Entity", "Context"]]
                .dropna()
                .to_excel(writer, index=False, sheet_name=ent_name)
            )

    logging.info(f"Results saved as '{output_file}'")
    return output_file
