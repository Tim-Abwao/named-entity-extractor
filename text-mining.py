#!/usr/bin/env python3
from tkinter import Tk, filedialog, messagebox, ttk
import pandas as pd
import spacy
import textract


class TextMiningApp(ttk.Frame):
    """Extract and predict entities in text from files."""

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title('Simple Text Mining App')
        self.master.resizable(False, False)
        self.configure(width=600, height=350)
        self.style = self._set_style()
        self._create_widgets()
        self.pack()

    _intro_text = "This is a simple application useful for extracting " +\
                  "information from text files in a variety of formats." +\
                  " Supported extensions include:"

    _extensions = ('.pdf', '.csv', '.doc', '.docx', '.xls', '.xlsx', '.txt',
                   '.odt', '.json', '.htm', '.html', '.tsv',  '.pptx', '.epub',
                   '.log', '.rtf', '.jpeg', '.jpg', '.gif', '.ogg',  '.png',
                   '.msg', '.wav', '.eml', '.mp3', '.ps', '.psv', '.tff',
                   '.tif', '.tiff')

    def _create_widgets(self):
        """Add the descriptive labels, and a button for selecting files."""
        self.intro = ttk.Label(self, text=self._intro_text,  wraplength=480)
        self.intro.place(relx=0.07, rely=0.1, relwidth=0.86, relheight=0.25)
        self.extensions = ttk.Label(self, wraplength=350, style='I.TLabel',
                                    text='  '.join(self._extensions))
        self.extensions.place(relx=0.2, rely=0.35)
        self.file_select = ttk.Button(self, text='Select file', width=25,
                                      command=self._process_text)
        self.file_select.place(relx=0.3, rely=0.8)

    def _set_style(self):
        """Set the style attributes of the widgets."""
        style = ttk.Style()
        style.configure("TFrame", background="ivory")
        style.configure("TLabel", foreground="darkslategrey",
                        background="ivory", font="serif 14")
        style.configure("I.TLabel", foreground="darkslategrey",
                        background="ivory",
                        font="serif 13 normal italic ")
        style.configure("TButton", foreground="slategrey",
                        background="aquamarine", font="serif 12")

    def _get_text(self):
        """
        Obtain text from selected documents using textract, and return it as
        a string object.
        """
        # select file and obtain it's contents
        file_name = filedialog.askopenfilename(
            initialdir='.', title="Please select document",
            filetypes=(('pdf', '*.pdf'),
                       ("plain text", "*.txt"),
                       ('word', '*.docx'),
                       ("all files", "*.*"))
        )
        self.progress['value'] = 15
        if not file_name:  # if no file is selected
            messagebox.showinfo(message="Please select a file to proceed")
            self.progress.destroy()

        try:
            text = textract.process(file_name).decode()
            self.progress['value'] = 25
            return text
        except UnicodeDecodeError:
            messagebox.showerror(message='Unable to parse file')
            self.progress.destroy()

    def _extract_entities(self):
        """
        Make predictions on entities present in the text using one of spaCy's
        pretrained core models (en_core_web_sm currently).
        """
        # loading the pretrained model
        nlp = spacy.load("en_core_web_sm")
        self.progress['value'] = 35
        # Getting named entity information
        doc = nlp(self.text)
        self.progress['value'] = 50
        ent_info = [(entity.text, entity.label_, entity.sent.text)
                    for entity in doc.ents]
        self.progress['value'] = 70
        self.ent_df = pd.DataFrame(ent_info,
                                   columns=['Entity', 'Type', 'Context'])\
                        .applymap(lambda x: x.replace('\n', ' '))
        self.label_info = pd.DataFrame(
            {"Type": self.ent_df['Type'].unique(),
             "Description": map(spacy.explain, self.ent_df['Type'].unique())})
        self.progress['value'] = 90

    def _save_results(self):
        """
        Save extracted information as an excel file, with a sheet for every
        entity type.
        """
        # getting 'save-as' name
        output_file = filedialog.asksaveasfilename(
            initialdir='.', initialfile="text_results.xlsx",
            filetypes=[("Excel", '.xlsx')]
        )
        if not output_file:  # if name not provided, or saving is cancelled
            self.progress.destroy()

        with pd.ExcelWriter(output_file) as writer:
            # creating a sheet for entity type descriptions
            self.label_info.to_excel(writer, sheet_name="DEFINITIONS",
                                     index=False)
            # creating a sheet for each entity type
            for ent_name, data in self.ent_df.groupby('Type'):
                data[["Entity", "Context"]].dropna()\
                    .to_excel(writer, index=False, sheet_name=ent_name)

        self.progress['value'] = 100
        messagebox.showinfo(message=f'Done! Results saved as {output_file!r}')
        self.progress.destroy()  # remove progress_bar after completion

    def _process_text(self):
        """
        Apply all the methods to extract text from documents, obtain named
        entity information and export the results to an excel file.
        """
        # initialise progress bar
        self.progress = ttk.Progressbar(self, orient='horizontal', length=400,
                                        mode='determinate', value=5)
        self.progress.place(relx=0.15, rely=0.7)
        # acquire the text
        self.text = self._get_text()
        if not self.text:
            messagebox.showwarning(message="Couldn't find text to process.")
            self.progress.destroy()
        # obtain entity info
        self._extract_entities()
        # save results
        self._save_results()


if __name__ == '__main__':
    app = TextMiningApp(master=Tk())
    app.mainloop()
