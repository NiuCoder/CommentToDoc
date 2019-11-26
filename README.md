# CommentToDoc
A python-implemented tool which can transfer code comments(.h file) to microsoft word document.

# How to work
put the [CommentToDoc.py](https://github.com/NiuCoder/CommentToDoc/blob/master/CommentToDoc.py) and the head file, for example [sample_api.h](https://github.com/NiuCoder/CommentToDoc/blob/master/sample_api.h) into the sample fold. Run the .py file and the doc will be generated, it will cost lot of time \
if there are two many files or file size is too large. 

# Note
- Only code comments in .h file are supported, except you change the source code
- The generated document can be reformatted if you like, but it depend on python-docx lib
- The sample_api.h show the format of the comment, if you want to change the comment style, \
the code may not work, then you should also change the soure code
- The core algorithm of the code is regex, so you can build your own code with another language with the regex
