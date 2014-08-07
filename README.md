python-docx
===========

Converts Word Docx file to Text, including headers and footers

This uses the docx.py engine written by mikemaccana to create a script that converts word docx files to text.
The script is simple to use:

python docx2text.py foo.docx foo.txt

The default is that it will just extract the document body. If you want the header and footer as well you can add -header 
and/or -footer to the command line. There is also a -v (verbose) and -h (help).
