#!/usr/bin/env python3
from tkinter import Tk, filedialog, messagebox, ttk
import pandas as pd
import spacy
import textract


class TextMiningApp(ttk.Frame):
    """Extract text from files and predict entities present."""

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title('Simple Text Mining App')
        self.master.resizable(False, False)
        self.configure(width=600, height=350)
        self.style = self._set_style()
        self._create_widgets()
        self.pack()

    _intro_text = "This is a simple app useful for extracting information "\
                  + "from text files in a variety of formats. Supported file "\
                  + "extensions include:"

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
        style.configure("TFrame", background="ivory")
        style.configure("TLabel", foreground="darkslategrey",
                        background="ivory", font="serif 14")
        style.configure("I.TLabel", foreground="darkslategrey",
                        background="ivory", font="serif 12 bold italic")
        style.configure("TButton", foreground="teal",
                        background="aquamarine", font="serif 12")

    def _get_text(self):
        """Get text from selected document using textract."""
        # select file
        file = filedialog.askopenfilename(initialdir=".",
                                          title="Please select a document",
                                          filetypes=(("pdf", "*.pdf"),
                                                     ("text", "*.txt"),
                                                     ("word", "*.docx"),
                                                     ("all files", "*.*")))
        self.progress['value'] = 15

        if not file:  # if no file is selected
            messagebox.showinfo(message="Please select a file to proceed")
            self.progress.destroy()
            return None

        try:
            text = textract.process(file).decode()
            self.progress['value'] = 25
            return text
        except UnicodeDecodeError:
            messagebox.showerror(message='Unable to parse file')
            self.progress.destroy()
            return None

    def _extract_entities(self):
        """
        Make predictions on entities present in the text using one of spaCy's
        pretrained core models (en_core_web_sm currently).
        """
        # load pretrained spaCy model
        nlp = spacy.load("en_core_web_md")
        self.progress['value'] = 35

        # Predict entities
        doc = nlp(self.text)
        self.progress['value'] = 50

        entity_info = ((entity.text, entity.label_, entity.sent.text)
                       for entity in doc.ents)
        self.progress['value'] = 70

        # Create dataframe of extracted entities
        entity_df = pd.DataFrame(entity_info,
                                 columns=['Entity', 'Type', 'Context'])

        # remove unwanted 'newline' to get continuous text
        self.entity_df = entity_df.applymap(lambda x: x.replace('\n', ' '))

        # Create dataframe of entity descriptions
        entity_types = self.entity_df['Type'].unique()
        self.entity_descriptions = pd.DataFrame(
            {"Type": entity_types,
             "Description": map(spacy.explain, entity_types)})
        self.progress['value'] = 90

    def _save_results(self):
        """
        Save extracted information as an excel file, with a sheet for each
        entity type.
        """
        # get desired output file-name from user
        output_file = filedialog.asksaveasfilename(
            initialdir='.', initialfile="text_results.xlsx",
            filetypes=[("Excel", '.xlsx')]
        )
        if not output_file:  # if name not provided, or saving is cancelled
            self.progress.destroy()
        else:
            with pd.ExcelWriter(output_file) as writer:
                # a sheet for entity-type descriptions
                self.entity_descriptions.to_excel(writer,
                                                  sheet_name="DEFINITIONS",
                                                  index=False)
                # a sheet for each entity type
                for ent_name, ent_data in self.entity_df.groupby('Type'):
                    (ent_data[["Entity", "Context"]]
                        .dropna()
                        .to_excel(writer, index=False, sheet_name=ent_name))

            self.progress['value'] = 100
            messagebox.showinfo(
                message=f'Done! Results saved as {output_file!r}'
            )
            self.progress.destroy()  # remove progress_bar after completion

    def _process_text(self):
        """
        Obtain named-entity information from a text file, and save results as
        an excel file.
        """
        # initialise progress bar
        self.progress = ttk.Progressbar(self, orient='horizontal', length=400,
                                        mode='determinate', value=5)
        self.progress.place(relx=0.15, rely=0.9)

        # acquire text
        self.text = self._get_text()

        if self.text:
            self._extract_entities()
            self._save_results()


if __name__ == '__main__':
    app = TextMiningApp(master=Tk())
    app.mainloop()
