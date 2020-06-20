import codecs
import os
import re
import sys
import traceback
from docx import Document
import PyPDF2
import argparse
from .striprtf import striprtf


def get_file_ext(file_name):
    tokens = file_name.split(".")
    ext = None
    if len(tokens) > 0:
        ext = tokens[-1].lower()
    return ext


def get_text(file_name):
    ext = get_file_ext(file_name)
    if ext is None or ext in ["txt", "rst", "text", "adoc"]:
        try:
            with codecs.open(file_name, "r", "utf-8") as f:
                return f.read()
        except Exception:
            print("File could not be read ", file_name)
            traceback.print_exc()
    elif ext == "rtf":
        try:
            with codecs.open(file_name, "r", "utf-8") as f:
                return striprtf(f.read())
        except Exception:
            print("File could not be read ", file_name)
            traceback.print_exc()
    elif ext in ["pdf"]:
        full_text = []
        with open(file_name, 'rb') as f:
            reader = PyPDF2.PdfFileReader(f)
            for pageNumber in range(reader.numPages):
                page = reader.getPage(pageNumber)
                try:
                    txt = page.extractText()
                    full_text.append(txt)
                except Exception:
                    print("Error PDF reader ", file_name, pageNumber)
                    traceback.print_exc()
        return "\n".join(full_text)
    elif ext in ["docx"]:
        doc = Document(file_name)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    elif ext in ["doc"]:
        if os.name == 'nt':
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False
            wb = word.Documents.Open(file_name)
            doc = wb.ActiveDocument
            return doc.Range().Text


        os.system("/Applications/LibreOffice.app/Contents/MacOS/soffice  --headless --convert-to txt:Text " + file_name)
        fileX = os.path.split(file_name)[1].split(".") +".txt"
        try:
            with codecs.open(fileX, "r", "utf-8") as f:
                return f.read()
        except Exception:
            print("File could not be read ", fileX)
            traceback.print_exc()

    else:
        print("Unknown file extension", file_name)
    return ""


def list_files(path):
    files = []
    # r = root, d = directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            files.append(os.path.join(r, file))

    lst = [file for file in files]
    return lst


def get_emails(fileName):
    txt = get_text(fileName)
    regex = re.compile(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)")
    return regex.findall(txt)


def get_files(dir, extensions):
    if extensions is None:
        return list_files(dir)
    filtered_files = []
    for file in list_files(dir):
        ext = get_file_ext(file)
        if ext in extensions:
            filtered_files.append(file)
    return filtered_files


def main(args):
    if args.file is not None:
        files = [args.file]
    else:
        files = get_files(args.dir, args.ext)
    out = sys.stdout
    if args.dst is not None:
        out = open(args.dst, mode="w")
    for file in files:
        emails = get_emails(file)
        for email in emails:
            out.write(email)
            out.write("\n")
    if args.dst is not None:
        out.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Extract emails from file')
    parser.add_argument("--dir", type=str, help="Directory/Folder Name", default=".")
    parser.add_argument("--file", type=str, help="File")
    parser.add_argument("--ext", type=str, nargs='*', help="File extensions")
    parser.add_argument("--dst", type=str, help="Destination file name")
    args = parser.parse_args()
    main(args)
