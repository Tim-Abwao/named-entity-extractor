# A Simple Program for Text Mining Using SpaCy NLP
[SpaCy](https://spacy.io/) is a powerful, open source Natural Language Processing package in Python, made by [Explosions AI](https://explosion.ai/).

The text to be processed will be extracted from text files using [tika-python](https://github.com/chrismattmann/tika-python). `tika-python` connects over the internet to the Tika REST server, but you can optionally install [Apache Tika](http://tika.apache.org) locally on your system. Apache Tika can extract text and metadata from hundreds of file types.
 
## Prerequisites
- Basic knowledge of **Python**, and experience with **jupyter notebooks**, **venv** and **pip**. 
- The **Tika** module requires you to have **Java** installed.

## Getting Started
- Set up a *virtual environment* in your project directory, .i.e. in the **SpaCy-Text-App** folder.  In your terminal/shell/command line, type the following:
```
cd ..path/to/spaCy-Text-App
python3 -m venv venv
```
- Activate the *virtual environment*, and install the required modules:
```
source venv/bin/activate
pip install -r requirements.txt
```
- For **Windows**, do this instead:
```
.\venv\Scripts\activate
pip install -r requirements.txt
```
Be sure to install the *spaCy* model. To add the small model:
```
python -m spacy download en_core_web_sm
```
Once your *virtual environment* is up and running, and you're done installing the required modules, you can choose to use either the jupyter notebook **NLP.ipynb**, or the Python script **NLP.py**.

## (a). Using the Jupyter Notebook
This lets you run the code piece by piece, so you can easily see what each piece of code (cell) does, and make changes.

In your terminal, at your project directory, with the virtual environment active, run the following:
```
python3 -m pip install jupyter

jupyter notebook NLP.ipynb
```
It should open up a browser window (or tab, if your default browser is already open) with the notebook.
You could also just type `jupyter notebook`, then navigate to the **NLP.ipynb** notebook.

## (b). Using the Python Script
In your terminal, at your project directory, with the virtual environment active, run the following:
```
python NLP.py
```

You'll first be prompted to select the text file to process. A *Select File* window should pop up, to help navigate and choose a text file. The script should then continue running just fine, and print what's happening at several stages. Afterwards, you'll be prompted to select a *destination directory*, where the extracted information will be saved as an excel file called **text-info.xlsx**. 

That's all. Enjoy!.
