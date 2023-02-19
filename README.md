# docx-live-reload

This is a command line tool that is useful for experimenting with Word documents and learning about WordprocessingML structures.

## Requirements

It works on Windows only and MS Word must be installed.

## Usage

Running: `docx-live-reload example.docx`

Will create extract the content of the document to a `test.docx__extracted` folder, making changes to `example.docx__extracted\word\document.xml` or `..styles.xml` will cause the document to reload in MS Word and show the changes.

Also making changes to `example.docx` using `python-docx` for example will also cause the document to reload in MS Word and show the changes.

After running the tool you can type `r` to manaually reload the document in MS Word and type `q` to quit.

## Installation

`pip install docx-live-reload`
