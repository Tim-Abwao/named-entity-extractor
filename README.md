# Basic text mining using spaCy NLP

[spaCy][1] is a powerful, user-friendly, open-source Natural Language Processing package in Python, made by [Explosions][3].

Text to be processed will be extracted from documents using [tika-python][4]. [Apache Tika][5] can extract text and metadata from "*over a thousand different file types*".

## Prerequisites

- Basic knowledge of *Python* and *Jupyter notebooks*.
- *Apache Tika* is a *Java* - based program.

## Getting Started

1. Download the files, and set up a virtual environment:

    ```bash
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

3. Download the tika-server *jar* file, and verify its integrity as advised. You can get it [here][6].

    Afterwards, open a new terminal window or tab, navigate to where it is, and start the tika server using the command:

    ```bash
    java -jar tika-server-1.24.jar
    ```

4. Navigate back to the terminal with the active virtual environment.

## I. Using the Jupyter Notebook

This lets you run the code interactively so that you can easily see what each piece of code does, and make any desired changes. It also covers more of *spaCy*'s features.

To start the notebook server and display the `text-mining.ipynb` file in your browser, use the command:

``` bash
jupyter notebook text-mining.ipynb
```

## II. Using the Python Script

The script is programmed to extract [named entities][7] plus some context.

To run it:

```bash
python text-mining.py
```

A [Tkinter][8] [GUI][9] should pop up to help you navigate to, and select the document to process. Afterwards, you'll be prompted to select a *destination directory*, where the extracted information will be saved as an excel file named *text-mining.xlsx*.

That's all. Enjoy.

[1]: https://spacy.io/
[2]: https://en.wikipedia.org/wiki/Natural_language_processing
[3]: https://explosion.ai/
[4]: https://github.com/chrismattmann/tika-python
[5]: http://tika.apache.org
[6]: https://www.apache.org/dyn/closer.cgi/tika/tika-server-1.24.jar
[7]: https://en.wikipedia.org/wiki/Named_entity
[8]: https://docs.python.org/3/library/tkinter.html#module-tkinter
[9]: https://en.wikipedia.org/wiki/Graphical_user_interface
