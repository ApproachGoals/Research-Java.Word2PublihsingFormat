# Research.Java.Word2PublihsingFormat

This project is developed for academic research.
#
The User can through this tool convert the academic document among IEEE, ACM, Spring and Elsevier formats.
The tool can reduce the edit task before the thesis or paper's publification.
I know that there are some problems in the tool. I will improve it continuously.
In later, I will convert the project into a Web or webservice, in order to more academic researcher to get this benefit.
#
Technique
#
The main body of the OOXML file format is the XML data. They can be easily manipulated.
The project uses the Docx4j library to manipulate the OOXML format (Word) file. Each OOXML node in the main part of the Word file will be scanned and operated, the style, the context structure and so on. The size of some special nodes such like image and table are modified to make suitable to the column width.

