import subprocess as sp
import smtplib
import time
import os
import tkinter.messagebox as tmsg
from email.message import EmailMessage
from pathlib import Path
from tkinter import *

try:
    import pandas as pd
except ModuleNotFoundError:
    sp.run("pip install pandas")
    sp.run("pip install openpyxl")
    sp.run("pip install xlrd")

os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

print(os.getcwd())
def excel_file(*file_name):
    changed_data = set()
    for File in file_name[0]:
        df = pd.read_excel(File, engine='openpyxl')
        if 'Phone' in df and 'Email' in df:
            data1 = {index:item['Email'] for index, item in df.iterrows()}
            data2 = {index:item['Phone'] for index, item in df.iterrows()}
            for i in data2.keys():
                changed_data.add(str(data2[i]))
            for i in data1.keys():
                changed_data.add(str(data1[i]))

        elif 'Email' in df:
            data = {index:item['Email'] for index, item in df.iterrows()}
            for i in data.keys():
                changed_data.add(str(data[i]))
        elif 'Phone' in df:
            data = {index:item['Phone'] for index, item in df.iterrows()}
            for i in data.keys():
                changed_data.add(str(data[i]))
        else:
            tmsg.showerror('Error', 'No Email or Phone field found in given files')
    newdf = pd.DataFrame(columns=["Email", "Phone"])

    for i, Data in enumerate(changed_data):
        if '@' in Data:
            newdf.loc[i, 'Email'] = Data
        else:
            newdf.loc[i, 'Phone'] = Data
    newdf.to_excel(r'Mail\result.xlsx', index=False)



def csv_file(*file_name):
    changed_data = set()
    for File in file_name[0]:
        df = pd.read_csv(File)
        if 'Phone' in df and 'Email' in df:
            data1 = {index:item['Email'] for index, item in df.iterrows()}
            data2 = {index:item['Phone'] for index, item in df.iterrows()}
            for i in data2.keys():
                changed_data.add(str(data2[i]))
            for i in data1.keys():
                changed_data.add(str(data1[i]))

        elif 'Email' in df:
            data = {index:item['Email'] for index, item in df.iterrows()}
            for i in data.keys():
                changed_data.add(str(data[i]))
        elif 'Phone' in df:
            data = {index:item['Phone'] for index, item in df.iterrows()}
            for i in data.keys():
                changed_data.add(str(data[i]))
        else:
            tmsg.showerror('Error', 'No Email or Phone field found in given files')
    newdf = pd.DataFrame(columns=["Email", "Phone"])

    for i, Data in enumerate(changed_data):
        if '@' in Data:
            newdf.loc[i, 'Email'] = Data
        else:
            newdf.loc[i, 'Phone'] = Data
    newdf.to_csv(r'Mail\result.csv', index=False)



def signin():
    global Id
    Id = smail_entry.get()
    global passwd
    passwd = spasswd_entry.get()
    global Subject
    Subject = sub_entry.get()
    global Omail
    Omail = omail_entry.get()
    root.destroy()

    global ask_file_root
    ask_file_root = Tk()


    Label(ask_file_root, text='How many files do you want to scan?     ', bg='#E8EAE0', font=('comicsan ms', 9, 'bold')).grid(sticky="W", row=0, column=0)

    global ask_file
    ask_file = IntVar()

    global file_entry
    file_entry = Entry(ask_file_root, textvariable= ask_file)
    file_entry.grid(sticky="W", row=1, column=0)

    Button(ask_file_root, text='Continue...', command=feed, bg='blue', fg='white').grid(sticky='W', row=2, column=1)

    ask_file_root.geometry('310x100')
    ask_file_root.maxsize(310, 100)
    ask_file_root.minsize(310, 100)
    ask_file_root.title('Alpha Mail')
    ask_file_root.mainloop()



def feed():
    global no_files_scan
    no_files_scan = file_entry.get()
    ask_file_root.destroy()


    if no_files_scan=="0":
        tmsg.showerror('Error', "Sorry you have given a wrong input")
        import sys; sys.exit()
    else:
        global root_feed
        root_feed = Tk()

        h1 = Frame(root_feed, bg='#E8EAE0')
        h1.pack(side='top', anchor='center')
        Label(h1, text='Welcome to Alpha Mail', fg='#34A853', font=('comicsan', 22, 'bold'), bg='#E8EAE0').pack()

        cont = Frame(root_feed, bg='#E8EAE0')
        cont.pack(side="left", anchor='n', pady=12)

        for i in range(int(no_files_scan)):
            exec(f'''
Label(cont, text="Enter location of {i+1} file that you want to scan  ", font=('comicsan', 10, 'bold'), bg='#E8EAE0').grid(sticky="W", row={i+1}, column=0)
file{i+1}_val = StringVar()
global file{i+1}_entry
file{i+1}_entry = Entry(cont, textvariable=file{i+1}_val)
file{i+1}_entry.grid(sticky="W", row={i+1}, column=1)
''')

        Button(cont, text='Feed me with data', command=fillter, bg="blue", fg="#E8EAE0").grid(sticky="W", row=i+2, column=1)

        root_feed.geometry('750x10000')
        root_feed.config(bg='#E8EAE0')
        root_feed.mainloop()



def mail():
    def sendmail(to, subject):
        mail = EmailMessage()
        mail['from'] = Id
        mail['subject'] = subject
        mail['to'] = to
        mail.set_content(Path(r'Mail\mail.html').read_text(), 'html')

        try:
            with smtplib.SMTP('SMTP.gmail.com', 587) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(Id, passwd)
                server.send_message(mail)
                print(f'MAil Sent! to {to}')
                # print(mail)
                # time.sleep(3)
        except Exception as e:
            print(e)

    df = pd.read_excel(r'Mail\result.xlsx', engine="openpyxl")
    remaining = df
    ral = pd.read_excel(r"Mail\remaining.xlsx", engine="openpyxl")

    sendmail(Omail, f"Server for {Subject} start")
    def engine(dd, remaining):
        for index, item in dd.iterrows():
            if index==1999 and gsuite.get():
                tmsg.showwarning('Limit exceed', 'You have exceed todays mailing limit but do not worry\n you left mails will be send tomorrow')
                break
            elif index==499 and gsuite.get()==0:
                tmsg.showwarning('Limit exceed', 'You have exceed todays mailing limit but do not worry\n you left mails will be send tomorrow')
                break
            else:
                while True:
                    try:
                        sendmail(item['Email'], Subject)
                    except Exception:
                        pass
                    else:
                        break
                remaining.drop([index], inplace=True)

    if len(ral.index)>0:
        engine(ral, ral)
        ral.to_excel(r"Mail\remaining.xlsx", index=False)

    else:
        engine(df, remaining)
        remaining.to_excel(r"Mail\result.xlsx", index=False)
        remaining.to_excel(r"Mail\remaining.xlsx", index=False)

    sendmail(Omail, f"Server for {Subject} completed")
    tmsg.showinfo('Thank you', 'Thank you for using Alpha Mailer\nHave a nice day...')

    root_feed.destroy()



def fillter():
    excel = []
    csv = []
    for i in range(int(no_files_scan)):
        exec(f'''
data = file{i+1}_entry.get()
if ".xlsx" in data:
    excel.append(data)
elif ".csv" in data:
    csv.append(data)
elif data=="":
    pass
else:
    tmsg.showwarning('Error', 'You have given an unsupported file...')
''')
    if len(excel)>0:
        excel_file(excel)
    elif len(csv)>0:
        csv_file(csv)
    else:
        retry = tmsg.askretrycancel('No file given', 'You have not given any file to scan\nWould you like to give some file again??')
        if retry:
            signin()
        else:
            tmsg.showinfo('Thank you', 'Thank you for using Alpha Mailer\nHave a nice day...')
    mail()
# all functions are created till here.


if __name__=='__main__':
    root = Tk()

    # Header
    Label(root, text='Welcome to Alpha Mail', fg= '#34A853', font=('comicsan', 22, 'bold'), bg='#E8EAE0').grid(sticky="W", row=0, column=4)

    # Entries
    Label(root, text='Please enter your current Email Id to get report', bg='#E8EAE0').grid(sticky="W", row=1, column=0)
    Label(root, text='Please enter details of Email through which you want to send mail', bg='#E8EAE0').grid(row=2, column=0)
    Label(root, text='Email id', bg='#E8EAE0').grid(sticky='W', row=4, column=0)
    Label(root, text='Password', bg='#E8EAE0').grid(sticky='W', row=5, column=0)
    Label(root, text='What subject you want to send??', bg='#E8EAE0').grid(sticky='W', row=6, column=0)

    omail = StringVar()
    smail = StringVar()
    spasswd = StringVar()
    sub = StringVar()
    gsuite = IntVar()

    omail_entry = Entry(root, textvariable=omail)
    smail_entry = Entry(root, textvariable=smail)
    spasswd_entry = Entry(root, textvariable=spasswd, show='*')
    sub_entry = Entry(root, textvariable=sub)

    omail_entry.grid(row=1, column=2)
    smail_entry.grid(row=4, column=2)
    spasswd_entry.grid(row=5, column=2)
    sub_entry.grid(row=6, column=2)

    # Buttons
    Checkbutton(root, text="Do you have a GSuite Account?", variable=gsuite, bg='#E8EAE0').grid(sticky='W', row=8, column=0)
    Button(root, text='Signin', command=signin, bg='blue', fg='#E8EAE0').grid(sticky='W', row=9, column=2)

    # Main Window settings
    min_hight = 400
    min_width = 900
    root.title("Alpha Mailer")
    root.configure(bg='#E8EAE0')
    root.minsize(min_width, min_hight)
    root.geometry('1000x500')
    root.mainloop()
