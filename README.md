# Extract emails from rtf,txt,text,doc,docx and PDF file

* Install Python3 and Pip3
* pip3 install -r requirements.txt
* python extract_emails.py --help
Note:
If your file has `doc` extension then you must have
 * On windows you must install pypiwin32
 * On Linux or Mac Install Libre Office

pypiwin32 is Windows python module so ignore install error on linux based os.


Options
* Use dir option to provide the directory/folder absolute path, default is current folder
* file option to scan only one file
* ext option to restrict the scanning of file extensions, default all supported extensions
* dst option to set the out file name, by default it will print on the console