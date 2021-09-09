#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk
from typing import Optional

from entity_extractor import utils

supported_formats = [
    ".csv,",
    ".doc,",
    ".docx,",
    ".eml,",
    ".epub,",
    ".gif,",
    ".htm,",
    ".html,",
    ".jpeg,",
    ".jpg,",
    ".json,",
    ".log,",
    ".mp3,",
    ".msg,",
    ".odt,",
    ".ogg,",
    ".pdf,",
    ".png,",
    ".pptx,",
    ".ps,",
    ".psv,",
    ".rtf,",
    ".tab,",
    ".tff,",
    ".tif,",
    ".tiff,",
    ".tsv,",
    ".txt,",
    ".wav,",
    ".xls,",
    ".xlsx",
]


class EntityExtractor(tk.Frame):
    def __init__(self, master: Optional[tk.Tk]) -> None:
        """Extract text from files and predict the named entities present

        Parameters
        ----------
        master : Optional[Tk]
            The root (top-level) window.
        """
        super().__init__(master)
        self.master = master
        self.master.title("Named Entity Extractor")
        self.master.geometry("600x350")
        self.master.resizable(False, False)
        self.master.wm_iconphoto(
            True, tk.PhotoImage(file="entity_extractor/icon.png")
        )
        self.pack()
        self._create_widgets()

    intro_text = (
        "Quickly collect entity information (people, places, companies, ...)"
        " from text files in a variety of formats:"
    )
    file_extensions = "  ".join(supported_formats)

    def _create_widgets(self) -> None:
        """Add widgets to the frame/window."""

        self.canvas = tk.Canvas(self, width=600, height=350)

        # Add background image
        self.background_image = tk.PhotoImage(file="entity_extractor/bg.png")
        self.canvas.create_image(
            0, 0, image=self.background_image, anchor="nw"
        )

        # Add title
        self.canvas.create_text(
            (70, 40),
            anchor="nw",
            font=("Courier", 25, "bold"),
            text="named-entity-extractor",
        )

        # Add description
        self.canvas.create_text(
            (45, 90),
            anchor="nw",
            font=("Times", 13),
            text=self.intro_text,
            width=520,
        )

        # Add supported file extensions
        self.canvas.create_text(
            (100, 150),
            anchor="nw",
            font=("Courier", 11),
            text=self.file_extensions,
            width=400,
        )

        # Add button to select an input file
        self.button = tk.Button(
            self,
            bg="#777",
            command=self._process_text,
            default="active",
            fg="white",
            font=("Courier", 12, "bold"),
            relief="flat",
            text="Select file",
            width=25,
        )
        self.master.bind("<Return>", lambda _: self.button.invoke())
        self.canvas.create_window(
            (170, 260), anchor="nw", height=40, width=250, window=self.button
        )

        self.canvas.pack()

    def _process_text(self) -> None:
        """Obtain named-entity information from a file, and save the results
        in an excel file.
        """
        self.percent_complete = tk.IntVar(self, value=10)
        self.progress = ttk.Progressbar(
            self,
            length=400,
            mode="determinate",
            orient="horizontal",
            variable=self.percent_complete,
        )
        self.progress.place(relx=0.15, rely=0.9)

        self.text = utils.read_text_from_file()
        self.percent_complete.set(75)

        if self.text is None:
            tk.messagebox.showinfo(message="Please select a file to proceed")
            self.progress.destroy()
        else:
            entity_info = utils.extract_entity_info(self.text)
            self.percent_complete.set(95)
            output_filename = utils.save_results_to_excel(entity_info)

            self.progress.stop()
            self.progress.destroy()
            tk.messagebox.showinfo(
                message=f'Done! File saved as "{output_filename}"'
            )

            del self.text


def run_app() -> None:
    """Launch the graphical user interface."""
    root = tk.Tk()
    app = EntityExtractor(master=root)
    app.mainloop()
