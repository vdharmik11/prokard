# exam-python-project-prokard
exam-python-project-prokard created by GitHub Classroom

Created By:  

(SEM - 5 Big Data and Analytics)  


Name | Enrollment No.
-----|---------------
Raj Rao|16012121023
Herik Taily|16012121033  
Ninad Thaker|16012121034
Dharmik Vyas|16012121038  
  
  
The PROKARD is a python based Progress Report Card Generating and Mailing System. As simple as it sounds, it takes a predefined
fromatted excel file containing various types of data of students along with their mail address, a sender with a mail address 
(less secure enabled in case of Gmail) and password can generate progress card and mail it with just a click !  
  
<b>Required Libraries:</b>  
* pandas  
* openpyxl  
* xlrd  
* cx_Freeze
  
<b>Usage:</b>  
  
A predefined excel file with sample data is available with this repository, download it and make all necessary modification required. 
Run the PROKARD.py file and click on file, locate and select the Excel file, then provided Sender's Email address and password and press
submit button. Once the mailing process will be completed, the sender will be notified that the mails have been sent successfully !
  
<b>Standalone executable file for Windows:</b>  
  
To generate standalone .exe file, run setup.py - This will include all the above mentioned required libraries in the executable package and can be ported to any Windows PC.   
  
For further details or to see the graphical workflow of the PROKARD, check the PDF provided in the repository.

<b>Note:</b> Format regarding Excel File which has to be followed strictly, otherwise mismatching/empty cell value will lead to errors 
or incorrect data. (The Excel, however, can be modified along with modification of code to maintain the correct form).

Feel free to download and modify the source code of this project.
