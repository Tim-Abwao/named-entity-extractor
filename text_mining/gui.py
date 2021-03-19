#!/usr/bin/env python3
from tkinter import Tk, messagebox, ttk
from text_mining import utils


class TextMiningApp(ttk.Frame):
    """Extract text from files and predict the named entities present."""

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title('Simple Text Mining App')
        self.master.resizable(False, False)
        self.configure(width=600, height=350)
        self.style = self._set_style()
        self._create_widgets()
        self.pack()

    _intro_text = 'This is a simple app useful for extracting information '\
                  + 'from text files in a variety of formats. Supported file '\
                  + 'extensions include:'

    _file_extensions = '  '.join([
        '.pdf', '.csv', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.odt',
        '.json', '.htm', '.html', '.tsv',  '.pptx', '.epub', '.log', '.rtf',
        '.jpeg', '.jpg', '.gif', '.ogg',  '.png', '.msg', '.wav', '.eml',
        '.mp3', '.ps', '.psv', '.tff', '.tif', '.tiff'])

    def _create_widgets(self):
        """Add descriptive labels, and a button for selecting files."""
        self.intro = ttk.Label(self, text=self._intro_text,  wraplength=540)
        self.intro.place(relx=0.07, rely=0.1, relheight=0.25)

        self.file_ext_list = ttk.Label(self, wraplength=360, style='I.TLabel',
                                       text=self._file_extensions)
        self.file_ext_list.place(relx=0.2, rely=0.35)

        self.file_select = ttk.Button(self, text='Select file', width=25,
                                      command=self._process_text)
        self.file_select.place(relx=0.3, rely=0.7, relheight=0.12)

    def _set_style(self):
        """Set style attributes for the widgets."""
        style = ttk.Style()
        style.configure('TFrame', background='ivory')
        style.configure('TLabel', foreground='darkslategrey',
                        background='ivory', font='serif 14')
        style.configure('I.TLabel', foreground='darkslategrey',
                        background='ivory', font='serif 12 bold italic')
        style.configure('TButton', foreground='teal',
                        background='aquamarine', font='serif 12')

    def _process_text(self):
        """
        Obtain named-entity information from a text file, and save results as
        an excel file.
        """
        # initialise progress bar
        self.progress = ttk.Progressbar(self, orient='horizontal', length=400,
                                        mode='determinate', value=5)
        self.progress.place(relx=0.15, rely=0.9)

        # Acquire text from a file
        self.text = utils.get_text_from_file()
        self.progress['value'] = 25

        if self.text is None:  # If the text content is empty
            messagebox.showinfo(message='Please select a file to proceed')
            self.progress.destroy()

        else:
            # Get entity information from text
            entity_df = utils.extract_entity_info(self.text)
            self.progress['value'] = 80

            # Save the results
            output_filename = utils.save_results_to_excel(entity_df)
            messagebox.showinfo(
                message=f'Done! File saved as {output_filename!r}')

            # Remove progressbar after completion
            self.progress.destroy()


def launch_text_mining_app():
    """Start the graphical user interface."""
    app = TextMiningApp(master=Tk())
    app.mainloop()
