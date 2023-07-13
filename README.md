# COMBDb
The Clinical Oral Microbiology Database (COMBDb) is a desktop application designed to store and query lab reports for clients and clinicians of the Dental School. Using Python and MongoDB in concurrence with a local database located within the school, this application will streamline routine tasks such as generating printable quality assurance reports, invoices, and patient culture worksheets. Administrators on COMBDb will have full, cautionary access to database operations, including the ability to alter, delete, and add fields, while the application itself will be as automated as possible in these operations to marginalize errors. Students on COMBDb will have very limited access to database operations, keeping patient information confidential and allowing read-only queries. COMBDb is designed to reduce workload and make a workday in the lab as time-efficient as possible.

DOWNLOADING THE APPLICATION
There is a file called COMBDb containing the executable, its necessary directories, and a COMBDb.accdb file that was used for testing. A new accdb file can be created and import the necessary tables from the test file. Downloading the zip will give you access to COMBDb.

Packages:
- pip install PyQt5
- pip install pywin32
- pip install docx-mailmerge
- pip install docxtpl
- pip install PyQtWebEngine
- pip install bcrypt
- pip install pyodbc
