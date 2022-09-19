from pathlib import Path
from docxtpl import DocxTemplate
from mailmerge import MailMerge
import win32com.client as win32
import os

def merge(path, context=None, **kwargs):
    template_path = str(Path().resolve())+path
    template = MailMerge(template_path)
    template.merge(**kwargs)

    mergedoc_path = path.split('\\')
    mergedoc_path[len(mergedoc_path)-1] = 'mergedoc.docx'
    mergedoc_path = '\\'.join(mergedoc_path)

    template.write(mergedoc_path)

    if context:
        mergedoc = DocxTemplate(mergedoc_path)
        mergedoc.render(context)
        mergedoc.save(mergedoc_path)

    word = win32.DispatchEx('Word.Application')
    mergepdf = word.Documents.Open(mergedoc_path)
    mergepdf_path = mergedoc_path.split('.')[0] + '.pdf'
    mergepdf.SaveAs(mergepdf_path, 17)
    mergepdf.Close()
    os.remove(mergedoc_path)
    word.Quit()

    return mergepdf_path()