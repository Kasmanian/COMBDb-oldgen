from docxtpl import DocxTemplate
from pathlib import Path

#Just change these according to your needs

#This is done to obtain the absolute paths to the input and output documents,
#because it is more reliable than using the relative path

template = DocxTemplate(str(Path().resolve())+r'\COMBDb\templates\test.docx')

#Specify all your headers in the headers column
context = {
'headers' : ['Component', 'Component Version', 'Server FQDN', 'Application port', 'DB SID', 'DB Port', 'Infos'],
'servers': []
}

#Fictious appserver list
appserver = ['a','b']

#Add data to servers 1 and 2 using a list and not a dict, remember to add
#an empty string for the Infos, as well, otherwise the border won't be drawn
for i in appserver: 
    server_1= ["Tomcat",7,i,5000," ",200,""]
    server_2= ["Apache",2.4,i," "," ",200,""]
    context['servers'].append(server_1)
    context['servers'].append(server_2)

template.render(context)
template.save(str(Path().resolve())+r'\COMBDb\templates\output.docx')