#!/usr/bin/env python3
from tkinter import Tk, messagebox, ttk, PhotoImage

from entity_extractor import utils


intro_text = """This is a simple app useful for extracting entity information \
from text files in a variety of formats. Supported file extensions include:"""

all_formats = [
    ".pdf",
    ".csv",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".txt",
    ".odt",
    ".json",
    ".htm",
    ".html",
    ".tsv",
    ".pptx",
    ".epub",
    ".log",
    ".rtf",
    ".jpeg",
    ".jpg",
    ".gif",
    ".ogg",
    ".png",
    ".msg",
    ".wav",
    ".eml",
    ".mp3",
    ".ps",
    ".psv",
    ".tff",
    ".tif",
    ".tiff",
]


class EntityExtractor(ttk.Frame):
    """Extract text from files and predict the named entities present."""

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Named Entity Extractor")
        self.master.resizable(False, False)
        self.configure(width=600, height=350)
        self.style = self._set_style()
        self._create_widgets()
        self.pack()

    _intro_text = intro_text

    _file_extensions = "  ".join(all_formats)

    def _create_widgets(self):
        """Add widgets to the frame/window."""
        # Introduction
        self.intro = ttk.Label(self, text=self._intro_text, wraplength=540)
        self.intro.place(relx=0.07, rely=0.12, relheight=0.25)

        # List of supported file extensions
        self.file_extensions = ttk.Label(
            self, wraplength=360, style="I.TLabel", text=self._file_extensions
        )
        self.file_extensions.place(relx=0.2, rely=0.4)

        # Button to select an input file
        self.select_input_file = ttk.Button(
            self, text="Select file", width=25, command=self._process_text
        )
        self.select_input_file.place(relx=0.3, rely=0.7, relheight=0.12)

    def _set_style(self):
        """Set style attributes for the widgets."""
        style = ttk.Style()
        # Frame style
        style.configure("TFrame", background="ivory")
        # Label style
        style.configure(
            "TLabel",
            foreground="darkslategrey",
            background="ivory",
            font="Times 15",
        )
        # Label style with bold, italic font
        style.configure(
            "I.TLabel",
            foreground="darkslategrey",
            background="ivory",
            font="Times 14 bold italic",
        )
        # Button style
        style.configure(
            "TButton",
            foreground="teal",
            background="aquamarine",
            font="serif 12",
        )

    def _process_text(self):
        """Obtain named-entity information from a text file, and save the
        output in an an excel file.
        """
        # Initialise the progress bar
        self.progress = ttk.Progressbar(
            self, orient="horizontal", length=400, mode="determinate", value=5
        )
        self.progress.place(relx=0.15, rely=0.9)

        # Acquire text from the input file
        self.text = utils.read_text_from_file()
        self.progress["value"] = 25

        if self.text is None:  # If the text content is empty
            messagebox.showinfo(message="Please select a file to proceed")
            self.progress.destroy()

        else:
            # Predict entities present in the text
            entity_info = utils.extract_entity_info(self.text)
            self.progress["value"] = 80

            # Save the results
            output_filename = utils.save_results_to_excel(entity_info)
            messagebox.showinfo(
                message=f'Done! File saved as "{output_filename}"'
            )

            # Remove progressbar after completion
            self.progress.destroy()


def run_app():
    """Start the graphical user interface."""
    root = Tk()
    root.wm_iconphoto(True, PhotoImage(file='entity_extractor/icon.png'))
    app = EntityExtractor(master=root)
    app.mainloop()
