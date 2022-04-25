import win32com.client

a = win32com.client.Dispatch("Access.Application")
# a.visible = 1  
filename = r'C:\Users\simmsk\Desktop\COPY.CDSNew.mdb'
db = a.OpenCurrentDatabase(filename)

report_name = 'Culture Worksheet II'
a.DoCmd.OpenReport(report_name)
a.Quit()