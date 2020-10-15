import glob
import win32com.client
import os


def pdf_to_docx(pdfs_path):
    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0
    for i, doc in enumerate(glob.iglob(pdfs_path)):
        print(doc)
        filename = doc.split('\\')[-1]
        filepath = doc.replace(filename, "")
        print(filepath)
        in_file = os.path.abspath(doc)
        print(in_file)
        wb = word.Documents.Open(in_file)
        out_file = os.path.abspath(filepath + filename[0:-4] + ".docx".format(i))
        print("outfile\n", out_file)
        wb.SaveAs2(out_file, FileFormat=16)  # file format for docx
        print("success...")
        wb.Close()
    word.Quit()


pdf_to_docx("D:\\parse-api\\file\\a.pdf")
