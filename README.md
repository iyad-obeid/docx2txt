doc2txt
===========

Converts Word Docx file to Text, including headers and footers

This uses the docx.py engine written by mikemaccana to create a script that converts word docx files to text.
The script is simple to use:

python docx2txt.py foo.docx foo.txt

The default is that it will extract the document body as well as the headers and footers. If you don't want the headers and footers you can add -noheader and/or -nofooter to the command line. There is also a -v (verbose) and -h (help).
