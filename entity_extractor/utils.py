import logging
from pathlib import Path
from typing import Dict

import pandas as pd
import spacy
from spacy.lang import en

CHUNK_SIZE = 100_000


def process_text_chunk(nlp: en.English, text_chunk: str) -> pd.DataFrame:
    """Get named entity predictions for a portion of the text not more than
    `CHUNK_SIZE` characters long. Chunking is necessary to avoid memory
    allocation errors.

    Args:
        nlp (spacy.lang.en.English): Pretrained model/pipeline.
        text_chunk (str): The text to process.

    Returns:
        pd.DataFrame: Entity name, type and context as columns.
    """
    doc = nlp(text_chunk)
    return pd.DataFrame(
        (
            (entity.text, entity.label_, entity.sent.text)
            for entity in doc.ents
        ),
        columns=["Entity", "Type", "Context"],
    ).drop_duplicates()


def extract_entity_info(text: str) -> Dict[str, pd.DataFrame]:
    """Predict the named entities present in the supplied `text`.

    Args:
        text (str): The text to process.

    Returns:
        Dict[str, pd.DataFrame]: Named entity info & type descriptions.
    """
    logging.info("Predicting named entities...")
    nlp = spacy.load("en_core_web_md")
    chunked_entity_data = (
        process_text_chunk(nlp, text_chunk=text[idx: idx + CHUNK_SIZE])
        for idx in range(0, len(text), CHUNK_SIZE)
    )
    entity_data = (
        pd.concat(chunked_entity_data, ignore_index=True)
        .drop_duplicates()
        .applymap(lambda txt: txt.replace("\n", " "))
    )
    logging.info(f"Done. {len(entity_data):,} named entities found.")

    entity_types = entity_data["Type"].unique()
    entity_descriptions = pd.DataFrame(
        {"Type": entity_types, "Description": map(spacy.explain, entity_types)}
    ).sort_values(by="Type")

    return {
        "predictions": entity_data,
        "descriptions": entity_descriptions,
    }


def save_results_to_excel(
    entity_info: Dict[str, pd.DataFrame], output_file: Path
):
    """Save extracted information as an excel file, with a sheet for each
    entity type.

    Args:
        entity_info (Dict[str, pandas.DataFrame]): "predictions"=Named entity
            data. "descriptions"=named entity types.
        output_file (Path): Where to save the results.
    """
    with pd.ExcelWriter(output_file) as writer:
        # Write entity descriptions in the first sheet
        entity_info["descriptions"].to_excel(
            writer, index=False, sheet_name="DEFINITIONS"
        )
        # Create a sheet for each entity type
        for ent_name, ent_data in entity_info["predictions"].groupby("Type"):
            (
                ent_data[["Entity", "Context"]]
                .dropna()
                .sort_values(by="Entity")
                .to_excel(writer, index=False, sheet_name=ent_name)
            )
    logging.info(f"Results saved as '{output_file}'\n")
