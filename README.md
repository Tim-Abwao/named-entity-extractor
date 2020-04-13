# Basic text mining using spaCy NLP
[SpaCy](https://spacy.io/) is a powerful, user-friendly, open-source Natural Language Processing package in Python, made by [Explosions AI](https://explosion.ai/). Text to be processed will be extracted from documents using [tika-python](https://github.com/chrismattmann/tika-python). 

[Apache Tika](http://tika.apache.org) can extract text and metadata from "*over a thousand different file types*". The `tika-python` package connects over the internet to the Tika REST server; but you can optionally run the tika server locally (preferred here).
 
## Prerequisites
- Basic knowledge of **Python** and **jupyter notebooks**.
- **Apache Tika** is a [Java]('https://www.java.com/en/')-based program: you'll need to have Java installed.

## Getting Started
1. Download the files, and set up a virtual environment:

``` bash
git clone https://github.com/Tim-Abwao/text-mining-spacy.git
cd text-mining-spacy
python3 -m venv venv
source venv/bin/activate
```

2. Install the required packages:

``` bash
pip install -U pip
pip install -r requirements.txt
python -m spacy download en_core_web_sm
```
3. Download the tika-server jar file, and verify its integrity as advised. You can get it [here](https://www.apache.org/dyn/closer.cgi/tika/tika-server-1.24.jar). Afterwards, open a new terminal where it is, and start the tika server:

```bash
java -jar tika-server-1.24.jar
```


Navigate back to the terminal with the active virtual environment. You can then choose to use either the Jupyter notebook or the Python script.

## (a). Using the Jupyter Notebook
This lets you run the code interactively - cell by cell -  so that you can easily see what each piece of code does, and make any desired changes.

The following command should start the notebook server, and display the `text-mining.ipynb` notebook in your browser.

``` bash
jupyter notebook text-mining.ipynb

```

## (b). Using the Python Script
Run the following:
``` bash
python text-mining.py

```

A [Tkinter](https://docs.python.org/3/library/tkinter.html#module-tkinter) GUI will then pop up to help you navigate to, and select the document to process. The script should then continue running. Afterwards, you'll be prompted to select a **destination directory**, where the extracted information will be saved as an excel file called **text-mining.xlsx**. 

That's all. Enjoy.
