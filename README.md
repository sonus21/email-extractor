# Extract emails from rtf, txt, text, doc,docx, and PDF file

* Install Python3 and Pip3
* pip3 install -r requirements.txt
* python extract_emails.py --help

Note:

If your file has a `doc` extension then you must have
 * On Windows you must install pypiwin32
 * On Linux or Mac Install Libre Office

pypiwin32 is a Windows python module so ignore the install error on Linux-based os.


Options
* `--dir` option to provide the directory/folder **absolute path**, default is current folder
* `--file` option to scan only one file
* `--ext` option to restrict the scanning of file extensions, default all supported extensions
* `--dst` option to set the output file name, by default it will print on the console

**NOTE**: Change the output file for each run otherwise it will overwrite the existing results.


#### Usage

Extract emails from a specific file  xyz.pdf

```shell script
python extract_emails.py --file=xyz.pdf --dst=emails.txt
```

Extract emails from all files from a folder/directory XYZ
```
python extract_emails.py --dir=XYZ --dst=emails.txt
```


While scanning a folder/directory you can specify file extensions as well, for example, it should only scan pdf files and then do


```shell script
python extract_emails.py --dir=XYZ --dst=emails.txt --ext pdf
```

Scan directory but only parse doc and pdf files

 ```shell script
python extract_emails.py --dir=XYZ --dst=emails.txt --ext pdf doc
```
