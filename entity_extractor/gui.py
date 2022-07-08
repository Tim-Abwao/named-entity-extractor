#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk
from typing import Optional

from entity_extractor import utils

logging.basicConfig(
    format="[%(levelname)s %(asctime)s.%(msecs)03d] %(message)s",
    level=logging.INFO,
    datefmt="%H:%M:%S",
)

with open("entity_extractor/supported_formats.txt") as file:
    supported_formats = file.read().split()


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
        self.master.geometry("540x300")
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
    file_extensions = ",  ".join(supported_formats)

    def _create_widgets(self) -> None:
        """Add widgets to the frame/window."""
        self.canvas = tk.Canvas(self, width=540, height=320)
        self.background_image = tk.PhotoImage(file="entity_extractor/bg.png")
        self.canvas.create_image(
            0, 0, image=self.background_image, anchor="nw"
        )
        # Title
        self.canvas.create_text(
            (60, 20),
            anchor="nw",
            font=("Courier", 24, "bold"),
            text="named-entity-extractor",
        )
        # Summary
        self.canvas.create_text(
            (45, 80),
            anchor="nw",
            font=("Times", 13),
            text=self.intro_text,
            width=480,
        )
        # Supported file extensions
        self.canvas.create_text(
            (70, 125),
            anchor="nw",
            font=("Courier", 11),
            text=self.file_extensions,
            width=400,
        )
        # Button to select input file
        self.button = tk.Button(
            self,
            bg="#777",
            command=self._process_text,
            default="active",
            fg="white",
            font=("Courier", 12, "bold"),
            relief="flat",
            text="Select file",
        )
        self.canvas.create_window(
            (150, 235), anchor="nw", height=40, width=240, window=self.button
        )
        # Interpret 'Enter' key press as button press
        self.master.bind("<Return>", lambda _: self.button.invoke())

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
