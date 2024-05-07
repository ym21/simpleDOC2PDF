"""
SimplePDF

This project is licensed under the MIT License, see the LICENSE file for details
"""
import pathlib
import re
import tempfile
from tkinter import filedialog, messagebox
import win32com.client
import pypdf


WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_DOCUMENT_WITH_MARKUP = 7

XL_TYPE_PDF = 0

RESULT_FILE = "まとめ"


def get_folder(folder_path):
    """
    フォルダからdoc,docx,xls,xlsx,pdfを抜き出しリストに変換
    """
    files = [p for p in folder_path.glob('*.*') if re.search(fr'(?i)(^(?!~\$).+\.(doc|docx|xls|xlsx)$)|(.*(?<!{RESULT_FILE})\.pdf$)', str(p.name))]
    return files

def save_path(filepath, dirpath):
    """
    変換したPDFを保存するパス
    """
    return dirpath / filepath.with_suffix('.pdf').name


def doc2pdf(files, dir_name):
    """
    wordをPDFに変換
    """
    app = win32com.client.Dispatch('Word.Application')

    try:
        for file in files:
            print(file.name, '処理中')

            inputfile = str(file.resolve())
            outputfile = str(save_path(file, dir_name).resolve())

            doc = app.Documents.Open(inputfile)

            doc.ExportAsFixedFormat(OutputFileName=outputfile, ExportFormat=WD_EXPORT_FORMAT_PDF, Item=WD_EXPORT_DOCUMENT_CONTENT)

            doc.Close()
    except Exception as e:
        print('======ERROR======')
        print(str(e))
    finally:
        app.Quit()


def wb2pdf(files, dir_name):
    """
    excelをPDFに変換
    """
    app = win32com.client.Dispatch('Excel.Application')

    try:
        for file in files:
            print(file.name, '処理中')

            inputfile = str(file.resolve())
            outputfile = str(save_path(file, dir_name).resolve())

            wb = app.Workbooks.Open(inputfile)

            wb.ActiveSheet.ExportAsFixedFormat(Type=XL_TYPE_PDF, Filename=outputfile)
            wb.Close()
    except Exception as e:
        print('======ERROR======')
        print(str(e))
    finally:
        app.Quit()



def main(dir_name, ibp):
    """
    メイン
    """
    folder = pathlib.Path(dir_name)
    insert_blank_page = ibp

    files = get_folder(folder)

    word_files = [f for f in files if re.search(r'(?i)\.(doc|docx)', str(f.suffix))]

    excel_files = [f for f in files if re.search(r'(?i)\.(xls|xlsx)', str(f.suffix))]
    

    with tempfile.TemporaryDirectory() as temp_dir:
        tmp = pathlib.Path(temp_dir)

        if len(word_files) > 0:
            doc2pdf(word_files, tmp)
        
        if len(excel_files) > 0:
            wb2pdf(excel_files, tmp)

        pdf_files = [f if re.search(r'(?i)\.(pdf)', str(f.suffix)) else save_path(f, tmp) for f in files]

        if len(pdf_files) > 0:

            print('PDF結合中')
            pdf = pypdf.PdfWriter()

            try:
                for f in pdf_files:
                    pdf.append(f)

                    #奇数ページ数の後に空白ページを追加
                    if insert_blank_page:
                        render = pypdf.PdfReader(f)
                        page_number = len(render.pages)
                        if page_number % 2 == 1:
                            pdf.add_blank_page()

                
                pdf.write(folder / f'{RESULT_FILE}.pdf')
                print(pathlib.Path(folder,f'{RESULT_FILE}.pdf').resolve(),"にPDFが保存されました")
            except Exception as e:
                print('======ERROR======')
                print(str(e))
            finally:
                pdf.close()

if __name__ == '__main__':
    select_dir = filedialog.askdirectory(title='PDFに変換するディレクトリを選択してください')
    if select_dir:
        is_insert_blank_page = messagebox.askyesno('空白ページを追加しますか？','奇数ページのファイルの後ろに空白のページを追加しますか？')
        print("PDFへの変換を開始")
        main(select_dir, is_insert_blank_page)
    else:
        print("フォルダが選択されていません")

    messagebox.showinfo('確認','処理が終了しました。')
