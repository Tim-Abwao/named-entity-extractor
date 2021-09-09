#!/usr/bin/env python3
import logging
from pathlib import Path
from tkinter.filedialog import askopenfilename, asksaveasfilename
from typing import Dict, Optional

import pandas as pd
import spacy
import textract
from spacy.lang import en

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s"
)
CHUNK_SIZE = 100000


def get_input_file() -> Optional[Path]:
    """Create a file-dialog to navigate to and select a file.

    Returns
    -------
    Path
        Input file.
    """
    filename = askopenfilename(
        initialdir=".",
        title="Please select a document",
        filetypes=(
            ("pdf", "*.pdf"),
            ("text", "*.txt"),
            ("word", "*.docx"),
            ("all files", "*.*"),
        ),
    )

    return Path(filename) if filename else None


def read_text_from_file() -> Optional[str]:
    """Get the contents of a file as a string.

    Returns
    -------
    Optional[str]
        A file's contents as one continuous string.
    """
    filepath = get_input_file()

    if filepath is None:
        return None
    else:
        logging.info(f"Found '{filepath}'. Extracting text...")
        return textract.process(filepath).decode()


def process_text_chunk(nlp: en.English, text_chunk: str) -> pd.DataFrame:
    """Get named entity info in a portion of the text not more than
    `CHUNK_SIZE` characters long. Chunking is necessary to avoid memory
    allocation errors.

    Parameters
    ----------
    nlp : spacy.lang.en.English
        Pretrained model/pipeline
    text_chunk : str
        The text to process

    Returns
    -------
    pd.DataFrame
        Entity name, type and context as columns.
    """
    doc = nlp(text_chunk)
    entity_info = (
        (entity.text, entity.label_, entity.sent.text) for entity in doc.ents
    )
    return pd.DataFrame(
        entity_info, columns=["Entity", "Type", "Context"]
    ).drop_duplicates()


def extract_entity_info(text: str) -> Dict[str, pd.DataFrame]:
    """Get the named entities present in the text.

    Parameters
    ----------
    text : str
        The text to process

    Returns
    -------
    Dict[str, pd.DataFrame]
        Named entity info and named entity type descriptions.
    """
    logging.info("Predicting named entities...")

    nlp = spacy.load("en_core_web_md")
    text_chunks = (
        text[idx: idx + CHUNK_SIZE] for idx in range(0, len(text), CHUNK_SIZE)
    )
    entity_data_list = [
        process_text_chunk(nlp, chunk) for chunk in text_chunks
    ]
    entity_df = (
        pd.concat(entity_data_list, ignore_index=True)
        .drop_duplicates()
        .applymap(lambda txt: txt.replace("\n", " "))
    )

    logging.info(f"Done. {len(entity_df):,} named entities found.")

    entity_types = entity_df["Type"].unique()
    entity_descriptions = pd.DataFrame(
        {"Type": entity_types, "Description": map(spacy.explain, entity_types)}
    )
    return {
        "predictions": entity_df,
        "descriptions": entity_descriptions,
    }


def set_output_file() -> Path:
    """Create a file dialog to select a destination for the results.

    Returns
    -------
    Path
        The desired output file's path.
    """
    output_file = asksaveasfilename(
        initialdir=".",
        initialfile="text_results.xlsx",
        filetypes=[("Excel", ".xlsx")],
    )

    return Path(output_file) if output_file else Path("entity-info.xlsx")


def save_results_to_excel(entity_info: Dict[str, pd.DataFrame]) -> Path:
    """Save extracted information as an excel file, with a sheet for each
    entity type.

    Parameters
    ----------
    entity_info : Dict[str, pd.DataFrame]
        Named entity predictions and type descriptions.

    Returns
    -------
    The path to the output file.
    """
    output_file = set_output_file()

    entity_predictions = entity_info["predictions"]
    entity_descriptions = entity_info["descriptions"]

    with pd.ExcelWriter(output_file) as writer:
        # Write entity descriptions in the first sheet
        entity_descriptions.to_excel(
            writer, index=False, sheet_name="DEFINITIONS"
        )
        # Create a sheet for each entity type
        for ent_name, ent_data in entity_predictions.groupby("Type"):
            (
                ent_data[["Entity", "Context"]]
                .dropna()
                .sort_values(by="Entity")
                .to_excel(writer, index=False, sheet_name=ent_name)
            )

    logging.info(f"Results saved as '{output_file}'\n")
    return output_file
