import logging
import tkinter as tk
from pathlib import Path
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror
from typing import Optional

import textract

from entity_extractor import utils

logging.basicConfig(
    format="[%(levelname)s %(asctime)s.%(msecs)03d] %(message)s",
    level=logging.INFO,
    datefmt="%H:%M:%S",
)

with open("entity_extractor/assets/supported_formats.txt") as file:
    supported_formats = file.read().split()


class EntityExtractor(tk.Frame):
    """Extract text from files and predict the named entities present.

    Args:
        master (Optional[tkinter.Tk]): The root (top-level) window.
    """

    def __init__(self, master: Optional[tk.Tk]) -> None:
        super().__init__(master)
        self.master = master
        self.master.title("Named Entity Extractor")
        self.master.geometry("540x300")
        self.master.resizable(False, False)
        self.master.wm_iconphoto(
            True, tk.PhotoImage(file="entity_extractor/assets/icon.png")
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
        self.background_image = tk.PhotoImage(file="entity_extractor/assets/bg.png")
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
            (150, 220), anchor="nw", height=40, width=240, window=self.button
        )
        # Interpret 'Enter' key press as button press
        self.master.bind("<Return>", lambda _: self.button.invoke())

        self.canvas.pack()

    def _get_input_file(self) -> None:
        """Create a file-dialog to navigate to, and select a file."""
        filename = askopenfilename(
            initialdir=".",
            title="Please select a document",
            filetypes=(
                ("pdf", "*.pdf"),
                ("text", "*.txt"),
                ("word", "*.docx"),
                ("all files", "*.*"),
            ),
            parent=self,
        )
        if filename:
            self.input_file = Path(filename)
        else:
            showerror(
                title="File Error",
                message="Please select a valid file",
            )
            self.input_file = None

    def _extract_text(self) -> None:
        """Get the contents of the input file as a string."""
        logging.info(f"Found '{self.input_file}'. Extracting text...")
        try:
            self.text = textract.process(self.input_file).decode()
        except Exception as error:
            showerror(title="Error reading file", message=error)
            self.text = None

    def _get_output_filename(self) -> Path:
        """Create a file dialog to select a destination for the results."""
        output_file = asksaveasfilename(
            initialdir=".",
            initialfile="entity-info.xlsx",
            filetypes=[("Excel", ".xlsx")],
            parent=self,
        )
        self.output_file = Path(output_file or "entity-info.xlsx")

    def _process_text(self) -> None:
        """Obtain named-entity information from a file, and save the results
        in an excel file.
        """
        self.percent_complete = tk.IntVar(self, value=5)
        self.progress = ttk.Progressbar(
            self,
            length=400,
            mode="determinate",
            orient="horizontal",
            variable=self.percent_complete,
        )
        self.progress.place(relx=0.15, rely=0.9)

        self._get_input_file()
        self.percent_complete.set(15)

        if self.input_file:
            self._extract_text()
            self.percent_complete.set(45)

            if self.text:
                entity_info = utils.extract_entity_info(self.text)
                self.percent_complete.set(95)
                self._get_output_filename()
                utils.save_results_to_excel(
                    entity_info, output_file=self.output_file
                )
                self.progress.destroy()
                tk.messagebox.showinfo(
                    message=f'Results saved as "{self.output_file}"'
                )
                del self.text
            else:
                showerror(message="No text to process!")
        else:
            self.progress.destroy()


def run_app() -> None:
    """Launch the graphical user interface."""
    root = tk.Tk()
    app = EntityExtractor(master=root)
    app.mainloop()
