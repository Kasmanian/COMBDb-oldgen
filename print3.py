import win32com.client as win32
# Open MS Word
word = win32.gencache.EnsureDispatch('Word.Application')

word_file = r"C:\Users\simmsk\Desktop\preliminary_culture_results_template.docx"
doc = word.Documents.Open(word_file)
# change to a .html
txt_path = word_file.split('.')[0] + '.html'

# wdFormatFilteredHTML has value 10
# saves the doc as an html
doc.SaveAs(txt_path, 10)

doc.Close()
# noinspection PyBroadException
try:
    word.ActiveDocument()
except Exception:
    word.Quit()