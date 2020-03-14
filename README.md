# A simple program for text mining using spaCy NLP
[SpaCy](https://spacy.io/) is a powerful, user-friendly, open-source Natural Language Processing package in Python made by [Explosions AI](https://explosion.ai/).

Text to be processed will be derived from documents using `tika-python`. [Apache Tika](http://tika.apache.org) can extract text and metadata from hundreds of file types.

[tika-python](https://github.com/chrismattmann/tika-python) connects over the internet to the Tika REST server, but you can optionally install Apache Tika locally on your system.  
 
## Prerequisites
- Basic knowledge of **Python** and **jupyter notebooks**.
- **Apache Tika** is a **Java** module: you'll need to have Java installed.

## Getting Started
Download the files, and set up a virtual environment with:
``` bash
git clone https://github.com/Tim-Abwao/spaCy-Text-App.git
cd spaCy-Text-App
python3 -m venv venv
source venv/bin/activate
```

Optionally update pip, then install the required modules:
``` bash
pip install -U pip
pip install -r requirements.txt
```

Be sure to install the [*spaCy* model](https://spacy.io/models). For the small model:
``` bash
python -m spacy download en_core_web_sm
```
You can then select between the jupyter notebook `NLP.ipynb` and the Python script `NLP.py`.

## (a). Using the Jupyter Notebook
This lets you run the code interactively - cell by cell -  so that you can easily see what each piece of code does, and make any desired changes.

The following command should open up a browser window (or tab, if your default browser is already open) and display the `NLP.ipynb` notebook .
``` bash
jupyter notebook NLP.ipynb
```

You could also just enter
``` bash
jupyter notebook
```
then manually find and open the **NLP.ipynb** notebook.

## (b). Using the Python Script
Run the `NLP.py` script in your terminal with:
``` bash
python NLP.py
```

A [Tkinter](https://docs.python.org/3/library/tkinter.html#module-tkinter)-based GUI will then pop up and prompt you to navigate to, and select the document to process. The script should then continue running, and print what's happening at several stages. Afterwards, you'll be prompted to select a *destination directory*, where the extracted information will be saved as an excel file called **text-mining.xlsx**. 

That's all. Enjoy.
