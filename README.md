# Getting Started

The document is organized as follows:
  - Requirements
  - Quick start
  - Installation
  - Code Documentation
  - Usage
  - Launch
  - Reporting bugs
  - License	
  
# Requirements

  - You need Python 2.7 run this code.
  - Below are the list of Packages required for this code:
  -     os, sys, re, time, pyPdf, xlsxwriter, subprocess, docx2html, 

> The design goal is to read either Docx or PFD files 
> and parse the data and use the key words passsed to
> the config and extract the Obligation points in the 
> input documents and create the Obgligation 
> Excel file as the result set. 
> Note : This is only driver , and this has to be appended with 
> some frame work in Python , Django etc.

### Quick start

We have used a number of open source projects to work properly:
* [AngularJS] - HTML enhanced for web apps - Not yet fetured in this version!
* [Ace Editor] - awesome web-based text editor - Not yet fetured in this version!
* [markdown-it] - Markdown parser done right. Fast and easy to extend.
* [node.js] - evented I/O for the backend. - Not yet fetured in this version!
* [Gulp] - the streaming build system 
* [jQuery] - FOr UI end. - Not yet fetured in this version!
* [Libreoffice] - Soffice opensorce package to be installed.

### Installation

We need the following to run :
1. soffice
2. docx2html
3. pYpdf
4. HTMLParser
5. TextStringObject
6. ContentStream
7. xlsxwriter

Install the dependencies and devDependencies and start the server.

```sh
$ cd Installed_Python_Directory
$ pip install pyPdf
$ pip install xlsxwriter
$ pip install subprocess
$ pip install docx2html
$ pip install HTMLParser
$ pip install re
$ sudo apt-get install gnome-web-print
$ sudo apt-get install xkhtmltopdf
$ sudo add-apt-repository ppa:libreoffice/ppa
$ sudo apt-get install libreoffice
$ sudo apt-get update
```

### Code Documentation

Below are the Plugins that can be derived from this module

| Plugin | README |
| ------ | ------ |
| passdoc(filename.html) | This takes HTML file as input and gives the Object List output |
| CreateExcel(DataObject, List) | Generate the Excel file with the input list of obligation and their respective clause and other flags. |
| extract_text() | This will etract the object from string. |
| scrape_text() | This is a AI opject to identify the best module to use for the input |
| create_index_html(File) | This gives the encoding for the input using AI trained module |
| Patternmatch(DataSet) | This will get the Obligation from the complete parsed data set list input |
| Method1(HTML_Input, Caluse ) | This will apply Method 1 of parsing data using the inbulit module |
| Method2(HTML_Input, Caluse ) | This will apply Method 2 of parsing data using the AI regular expression module |
| DataComputer(HTML_Input) | This will extract the File Header defination to exclude the header and footer in compete set. |
| DocxToPdf(File) | This will convert any Docx to PDF file. This will convert any chart ond graph set available in the code in convertion using the complete set module. |

### Usage

To convert any DOCX to PDF :
we can use the below command, Output file will be generated in the same path
```sh
$ python Driver_Utils.py Input.docx
```

To extract the Oblication from file using method 1 :
```sh
$ python  Driver_Utils.py input_file.pdf 1
or
$ python  Driver_Utils.py input_file.docx 1
```
To extract the Oblication from file using method 2 :
```sh
$ python  Driver_Utils.py input_file.pdf 2
or
$ python  Driver_Utils.py input_file.docx 2
```


### Launch
```sh
run -d -p 8000:8080 --restart="always" <file>/data:${package.json.version}
```

### Reporting Bug

Please  report for any bug.

License
----
NA
**@rkoots Development...**
