##########################################################
##########################################################
##                                                      ##
##          PROJECT: PROKARD (PROGRESS REPORT)          ##
##                                                      ##
##                        BY                            ##
##                                                      ##
##                                                      ##
##             RAJ     RAO    (16012121023)             ##
##             HERIK   TAILY  (16012121033)             ##
##             NINAD   THAKER (16012121034)             ##
##             DHARMIK VYAS   (16012121038)             ##
##                                                      ##
##                                                      ##
##                                                      ##
##########################################################
##########################################################

##########################################################
#                    IMPORTING MODULES                   #
##########################################################

# IMPORTING MODULES FOR EXCEL
from pandas import *
import pandas as pd
from openpyxl import load_workbook
import xlrd

# IMPORTING MODULES FOR GUI
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
from tkinter.filedialog import askopenfilename

# IMPORTING MODULES FOR MAIL
import smtplib
import email.utils
from email.mime.multipart import *
from email.mime.text import *

# IMPORTING MODULE FOR PDF READING
from pathlib import Path
import glob
import os
from time import *
import webbrowser

class Main:
    
    ##########################################################
    #                    GLOABL VARIABLES                    #
    ##########################################################

    toaddr=""
    mailbody=""
    subject=""
    file=""

    ##########################################################
    #                    PROGRAM CLOSING                     #
    ##########################################################

    def Close():
        window.destroy()

    ##########################################################
    #                     FILE SELECTION                     #
    ##########################################################

    def File():
        filename=askopenfilename(title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))    
        if(filename==""):
             tkinter.messagebox.showinfo(message="Please Select a file")   
        else:
            global file
            file=filename
 
    ##########################################################
    #                    LOGIN FUNCTION                      #
    ##########################################################
    def Login():
        # GETTING USERNAME
        fromaddr=e1.get()
        # GETTING PASSWORD
        password=e2.get()
        if(fromaddr!="" and password!=""):
            
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.ehlo()
            
            try:
                server.starttls()
                server.login(fromaddr, password)
            except:
                tkinter.messagebox.showinfo(message="Either Email ID or Password is Wrong!")
                e1.focus()
                e2.delete(0,'end')
            
            book=xlrd.open_workbook(file)
            sheet2=book.sheet_by_index(1)
            rowx=sheet2.nrows
            rowx-=1
            def Pbar(byte,step):
                maxbyte=1
                byte+=step
                progress["value"]=byte
                if byte < maxbyte:
                    pbar.after(100)
                    pbar.update()
            def Pbar_init():
                byte=0
                step=(1/(rowx-1))
                progress["value"]=0
                progress["maximum"]=1
                Pbar(byte,step)
            Pbar_init()

        ##########################################################
        #                    MAIN FUNCTION                       #
        ##########################################################
            
            book=xlrd.open_workbook(file)
            sheet2=book.sheet_by_index(1)
            ntime=sheet2.nrows
            ntime-=1
            s2=pd.read_excel(file,sheet_name="Sheet2",header=None,index=False)

            # N TIME MAIL
            while(ntime!=0):
                try:
                    # COUNTING OF COLS AND ROWS
                    book=xlrd.open_workbook(file)
                    sheet3=book.sheet_by_index(2)
                    cols=sheet3.ncols
                    rows=sheet3.nrows

                    # REGULAR SHEETS
                    s1=pd.read_excel(file,sheet_name="Sheet1",header=None,index=False)
                    s2=pd.read_excel(file,sheet_name="Sheet2",header=None,index=False)
                    s3=pd.read_excel(file,sheet_name="Sheet3",header=None,index=False)
                    
                    # LOADING COUNTER
                    cd=s1.iloc[6][1]

                    # LOADING EMAIL
                    for i in range(cd+1,cd+2):
                        global toaddr
                        toaddr=str(s2.iloc[i][3])

                    # CREATING FILE WITH ENROLLMENT NO.
                    for i in range(cd+1,cd+2):
                        name=str(s2.iloc[i][1])+".html"
                    f=open(name,"w")

                    # START OF HTML
                    f.write("<html><head><style>th{background-color:#FFCE72}</style></head><body><img src=\"https://drive.google.com/uc?id=1yjCtsM1zT4OgnCSiTEewCOPLPUsYfa-T\" width=300px height=100px style=\"display:block;margin:auto\"><table align=\"center\">")

                    # REPORT HEADING
                    f.write("<br><br><h3 align=\"center\">"+str(s1.iloc[7][0])+"</h3><br>")

                    # DETAILS OF INSTITUTE, PROGRAM AND BRANCH
                    f.write("<tr><td><b>INSTITUTE:</b>"+str(s1.iloc[0][0])+"</td></tr>")
                    f.write("<tr><td><b>PROGRAM:</b>"+str(s1.iloc[1][0])+"</td></tr>")
                    f.write("<tr><td><b>BRANCH:</b>"+str(s1.iloc[2][0])+"<b>SPECIALISATION:</b>"+str(s1.iloc[3][0])+"</td></tr>")

                    # SEMESTER AND MONTH & YEAR OF EXAM 
                    f.write("<table border=1 align=\"center\" rules=all width=720><th width=\"50%\">SEMESTER</th><th width=\"50%\">MONTH & YEAR OF EXAM</th><tr>")
                    for i in range(cd+1,cd+2):
                        global subject
                        subject=str(s1.iloc[5][0])
                        f.write("<tr><td align=center>"+str(s1.iloc[4][1])+"</td><td align=center>"+str(s1.iloc[5][0])+"</td></tr></table><br/><br/>")

                    # ENROLLMENT NO AND NAME
                    f.write("<table border=1 align=\"center\" rules=all width=720><tr><th width=\"50%\">NAME</th><th width=\"50%\">ENROLLMENT NO.</th></tr>")
                    for i in range(cd+1,cd+2):
                        f.write("<tr><td align=center>"+str(s2.iloc[i][2])+"</td><td align=center>"+str(s2.iloc[i][1])+"</td></tr></table><br><br>")

                    # COURSE CODE , COURSE TITLE , MAX MARK , EARNED MARKS
                    f.write("<table border=1 align=\"center\" rules=all width=720><tr><th>Course Code</th><th>Course Title</th><th>Maximum Marks</th><th>Marks Obtained</th></tr>")
                    for i in range(1):
                        for j in range(cols-1):
                            f.write("<tr><td>"+str(s3.iloc[i][j+1])+"</td><td>"+str(s3.iloc[i+1][j+1])+"</td><td align=\"center\">"+str(s3.iloc[i+2][j+1])+"</td><td align=\"center\">"+str(s3.iloc[cd+4][j+1])+"</td></tr>")
                    add=0
                    tot=0
                    # STUDENT MARK
                    for i in range(cd+1,cd+2):
                        for j in range(1,cols):
                            try:
                                d=int(str((s3.iloc[cd+4][j])))
                            except:
                                d=0
                            add=add+d
                    # SUBJECT TOTAL MARK
                    for i in range(1,cols-1):
                        tot=tot+int(s3.iloc[2][i])
                    f.write("<tr><th align=\"center\" colspan=2>TOTAL</th><th align=\"center\">"+str(tot)+"</th><th align=\"center\">"+str(add)+"</th></tr>")

                    f.write("</table>")

                    # END OF HTML
                    f.write("</table></body></html>")

                    # CLOSING FILE
                    f.close()

                    # MAIL TO RECEIVER BODY
                    mb=open(name,'r')
                    global mailbody
                    mailbody=mb.read()

                    # SENDING MAIL
                    msg = MIMEMultipart()
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = subject
                    body = mailbody
                    msg.attach(MIMEText(body, 'html'))
                    text = msg.as_string()
                    server.sendmail(fromaddr, toaddr, text)
                    
                    # INCREMENTING COUNTER
                    cd+=1
                    if(cd==rowx):
                        cd=0
                    writer=pd.ExcelWriter(file,engine="openpyxl")
                    book=load_workbook(file)
                    writer.book=book
                    writer.sheets=dict((ws.title,ws) for ws in book.worksheets)
                    df=pd.read_excel(file,header=None)
                    df.to_excel(writer,sheet_name="Sheet1",header=None,index=False)
                    cd1=pd.DataFrame({'Data':[cd]})
                    cd1.to_excel(writer,sheet_name="Sheet1",header=None,index=False,startcol=1,startrow=6)
                    writer.save()
                    ntime-=1
                    if(cd!=0):
                        Pbar((cd/(rowx-1)),(1/(rowx-1)))
                    
                except IndexError:
                    break
                except NameError:
                    tkinter.messagebox.showinfo(message="Please select a file first !!")
            server.quit()
            tkinter.messagebox.showinfo(title="Info",message="All Mails have been sent Successfully !")
            
        elif(fromaddr==""):
            tkinter.messagebox.showinfo(title="Info",message="Please Enter Email ID !")
            e1.focus()
        elif(password==""):
            tkinter.messagebox.showinfo(title="Info",message="Please Enter Password !")
            e2.focus()

    ##########################################################
    #                         GUIDE                          #
    ##########################################################
    def Guide():

        webbrowser.open_new(r'PROKARD_USER_GUIDE.pdf')
    
        
    ##########################################################
    #                       SEARCHING                        #
    ##########################################################

    def Search():
        try:
            file=askopenfilename(title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
            # COUNTING OF COLS
            book=xlrd.open_workbook(file)
            sheet3=book.sheet_by_index(2)
            cols=sheet3.ncols
            rows=sheet3.nrows
            # REGULAR SHEETS
            s2=pd.read_excel(file,sheet_name="Sheet2",header=None,index=False)
            s3=pd.read_excel(file,sheet_name="Sheet3",header=None,index=False)

            search=Tk()
            search.title("Search")
            search.geometry("250x250")
            search.configure(background="#FFCE72")
            sframe=Frame(search)
            sframe.configure(background="#FFCE72")
            sf1=Label(search,text="Enrollment No :",background="#FFCE72")
            sf1.pack(padx=50,pady=30)
            sf2=Entry(search,bd=5,width=20)
            sf2.pack(padx=5,pady=5)
            sf2.focus()

            def Submit():
                en=sf2.get()
                count=0
                if(en==""):
                    tkinter.messagebox.showinfo(title="Info",message="Please Enter Enrollment Number !")
                else:
                    for i in range(4,rows):
                        ss=str(s3.iloc[i][0])
                        if(en!=ss):
                            i+=1
                            count+=1
                        else:
                            search.destroy()
                            result=Tk()
                            result.title(en)
                            rframe=Frame(result)
                            rframe.configure(background="#FFCE72")
                            def Clo():
                                result.destroy()
                            rc=[]                            
                            for i in range(1):
                                for j in range(0,cols-1):
                                    rc+=[[str(s3.iloc[i][j+1]),str(s3.iloc[i+1][j+1]),str(s3.iloc[i+2][j+1]),str(s3.iloc[count+4][j+1])]]

                            c=int(len(rc))
                            c4=int(len(rc)/4)
                            for i in range(0,c):
                                for j in range(0,4):
                                    l=Entry(result,relief=RIDGE,font='bold')
                                    l.grid(row=i,column=j,sticky=NSEW)
                                    if j==1:
                                        result.grid_columnconfigure(j,minsize=600)
                                    else:
                                        result.grid_columnconfigure(j,minsize=10)
                                    result.grid_rowconfigure(i,minsize=5)
                                    l.insert(END,'%s'%rc[i][j])
                                    l.config(state='disabled')
                            p=Button(result,text="Close",command=Clo)
                            p.grid(row=c,column=0)
            
            def Cancel():
                search.destroy()
            
            sf3=Button(search,text="Submit",command=Submit)
            sf3.pack(side=LEFT,padx=50,pady=50)
            sf3=Button(search,text="Cancel",command=Cancel)
            sf3.pack(side=LEFT,padx=5,pady=20)
        except FileNotFoundError:
            tkinter.messagebox.showinfo(title="Info",message="Please Select a file first !")


##########################################################
#                      CREATING GUI                      #
##########################################################
def About(): 
    tkinter.messagebox.showinfo(title="Info",message="Created by Students @ UVPCE:\n=================\
    \nRaj Rao\nHerik Taily\nNinad Thaker\nDharmik Vyas\n=================\n\nVersion :8.1")

window=Tk()

# FILE MENU
menubar=Menu(window)

filemenu=Menu(menubar,tearoff=0)
filemenu.add_command(label="Select File",command=Main.File)
filemenu.add_command(label="Search",command=Main.Search)
filemenu.add_separator()
filemenu.add_command(label="Exit",command=Main.Close)
menubar.add_cascade(label="File",menu=filemenu)

aboutmenu=Menu(menubar,tearoff=0)
aboutmenu.add_command(label="About",command=About)
aboutmenu.add_command(label="Guide",command=Main.Guide)
menubar.add_cascade(label="Help",menu=aboutmenu)

window.configure(menu=menubar,background="#FFCE72")

# MAIN WINDOW
window.title("Login")
window.geometry("250x250")
frame=Frame(window)
frame.configure(background="#FFCE72")
pbar=Frame(window)
pbar.configure(background="#FFCE72")
l1=Label(window,text="Email :",background="#FFCE72")
l1.pack(padx=15,pady=5)
e1=Entry(window,bd=5,width=20)
e1.pack(padx=15,pady=5)
l2=Label(window,text="Password :",background="#FFCE72")
l2.pack(padx=10,pady=5)
e2=Entry(window,bd=5,width=20,show="*")
e2.pack(padx=10,pady=5)

# PROGRESS BAR
progress = ttk.Progressbar(orient="horizontal",
                           length=200, mode="determinate")
progress.pack(side=BOTTOM,padx=5,pady=10)

##########################################################
#            BUTTON GENERATION AND FUNCTION CALL         #
##########################################################

b1=Button(frame,text="Login",command=Main.Login)
b1.pack(side=LEFT,padx=5,pady=15)

frame.pack()
pbar.pack()
window.mainloop()
