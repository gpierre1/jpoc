# -*- coding: utf-8 -*-
#########IMPORTS#################
try:
    from Tkinter import *
except:
    from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import csv
from datetime import datetime
import win32com.client
import os

root = Tk()
root.title("AUTOMATIC EMAIL GENERATION")

#########VARIABLES###############
#------DATA FROM CSV------------#
raw_data = []
task_number = []
request_number = []
requested_for = []
opened_by = []
assign_group = []
assign_to = []
assign_to_email = []
assign_manager = []
assign_manager_email = []
employee_status = []
location_region = []
request_opened = []
update_date = []


#-------OTHER VARIABLES---------#
date_time = []
outdated_days = []
index_outdated_tasks = []
outdated_recipients = dict()
#########FUNCTIONS###############
def import_parent():
    get_duration()
    get_email()
    viewsample_button.config(state="normal")
    sendall_button.config(state="normal")
    return
def open_csv():
    try:
        temp_task_number = []
        temp_request_number = []
        temp_requested_for = []
        temp_opened_by = []
        temp_assign_group = []
        temp_assign_to = []
        temp_assign_to_email = []
        temp_assign_manager = []
        temp_assign_manager_email = []
        temp_employee_status = []
        temp_location_region = []
        temp_request_opened = []
        temp_update_date = []
        
        temp_date_time = datetime.now()
        temp_date_time2 = temp_date_time.strftime("%m/%d/%Y %H:%M")
        date_time.append(temp_date_time2)
        file_format = [('CSV file','*.csv')]
        file_name = filedialog.askopenfilename(parent=root,title="Import CSV file",filetypes=file_format)
        file_open = open(file_name)
        import_data = csv.reader(file_open)
        import_entry.insert(0,str(file_name))
        import_entry.xview_moveto(1)
        import_entry.config(state="readonly",disabledforeground="black")
        for data in import_data:
            temp_task_number.append(data[0])
            temp_request_number.append(data[1])
            #temp_requested_for.append(data[2])
            #temp_opened_by.append(data[3])
            temp_assign_group.append(data[4])
            temp_assign_to.append(data[5])
            temp_assign_to_email.append(data[6])
            #temp_assign_manager.append(data[7])
            #temp_assign_manager_email.append(data[8])
            #temp_employee_status.append(data[9])
            #temp_location_region.append(data[10])
            temp_request_opened.append(data[14])
            temp_update_date.append(data[15])
        
        for i in range(len(temp_update_date)-1):
            task_number.append(temp_task_number[i+1])
            request_number.append(temp_request_number[i+1])
            #requested_for.append(temp_requested_for[i+1])
            #opened_by.append(temp_opened_by[i+1])
            assign_group.append(temp_assign_group[i+1])
            assign_to.append(temp_assign_to[i+1])
            assign_to_email.append(temp_assign_to_email[i+1])
            #assign_manager.append(temp_assign_manager[i+1])
            #assign_manager_email.append(temp_assign_manager_email[i+1])
            #employee_status.append(temp_employee_status[i+1])
            #location_region.append(temp_location_region[i+1])
            request_opened.append(temp_request_opened[i+1])
            update_date.append(temp_update_date[i+1])
        file_open.close()
        
        load_button.config(state="normal")
        
        del temp_task_number
        #del temp_requested_for
        #del temp_opened_by
        del temp_assign_group
        del temp_assign_to 
        del temp_assign_to_email
        #del temp_assign_manager
        #del temp_assign_manager_email
        #del temp_employee_status
        #del temp_location_region
        del temp_request_opened
        del temp_update_date
    except IOError:
        pass
    return
    
def get_duration():
    
    for i in range(len(update_date)):
        temp_duration = datetime.strptime(date_time[0],"%m/%d/%Y %H:%M") - datetime.strptime(update_date[i],"%m/%d/%Y %H:%M")
        outdated_days.append(temp_duration)
    
    for j in range(len(outdated_days)):
        duration = outdated_days[j]
        duration_days = duration.days
        if duration_days >= 90:
            index_outdated_tasks.append(j)
        else:
            pass
    return
    
def get_email():
    for index in index_outdated_tasks:
        email_address = assign_to_email[index]
        if email_address in outdated_recipients:
            outdated_recipients[email_address].append(index)
        else:
            outdated_recipients.update({email_address:[index]})
    return
    
def send_email_parent():
    make_email_content()
    address_array = list(outdated_recipients.keys())
    for address in address_array:
        send_email(address)
    return
    

    
def make_email_content():
    information = [task_number, request_number, assign_group, assign_to, assign_to_email, request_opened, update_date]
    for i in range(len(outdated_recipients.keys())):
        recipient_key = list(outdated_recipients.keys())[i]
        tasks_index = outdated_recipients.get(recipient_key)
        recipient_add = recipient_key.split('@')
        email_body = open(str(recipient_add[0])+'.txt','w')
        temp_recipient_name = assign_to[tasks_index[0]]
        temp_recipient_name = temp_recipient_name.split(' ')
        recipient_name = temp_recipient_name[0]
        email_body.write("Hi "+str(recipient_name)+","+'\r\n')
        email_body.write("Please find below the list of Aging Ticket(s) not updated for 90 or more days."+'\r\n')
        email_body.write("Could you please review and update the task(s) assigned to you or close if the request is no longer required?"+'\r\n')
        email_body.write('\r\n')
        for index in tasks_index:
            for params in information:
                email_body.write(str(params[index])+'\t')
            email_body.write('\r\n')
        email_body.write('\r\n')
        email_body.write("Thanks,"+'\r\n')
        email_body.write("Gina")
        email_body.close()
            
            
    
    return
    
def send_email(recepient):
    temp_file_name = recepient.split('@')
    file_name = temp_file_name[0]
    email_body_file = open(str(file_name)+'.txt','r')
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "Aging Service Request Task(s) - Not Updated For 90 Or More Days"
    newMail.Body = email_body_file.read()
    newMail.To = "gpierre@its.jnj.com"
    newMail.Send()
    email_body_file.close()
    filename = str(file_name)+'.txt'
    os.remove(filename)
    return
    
def show_sample(recepient):
    new_window2 = Toplevel()
    recipient_add = recepient.split('@')
    body = open(str(recipient_add[0])+'.txt','r')
    body_text = Text(new_window2,wrap="word",width=500,height=500)
    body_text.insert(END, body.read())
    body_text.grid(row=0,column=0)
    return
    
def view_sample():
    make_email_content()
    new_window = Toplevel()
    receiver_label = Label(new_window,text="Recipient Email Address")
    receiver_label.grid(row=0,column=0)
    for i in range(len(outdated_recipients.keys())):
        recipient_key = list(outdated_recipients.keys())[i]
        receiver_email_add = Label(new_window,text=recipient_key)
        receiver_email_add.grid(row=i,column=0,padx=5)
        email_view_button = Button(new_window,text="View Sample",command= lambda: show_sample(recipient_key))
        email_view_button.grid(row=i,column=1,padx=5)
    
    
    return

#########MAIN WINDOW#############
import_frame = Frame(root)
import_frame.grid(row=0,column=0,columnspan=6)

import_label = Label(import_frame,text="CSV FILE:")
import_label.grid(row=0,column=0,sticky=W)

import_entry = Entry(import_frame,width=50)
import_entry.grid(row=1,column=0,padx=2)

browse_button = Button(import_frame,text="Browse",width=7,command=open_csv)
browse_button.grid(row=1,column=1,padx=2)

load_button = Button(import_frame,text="Load",width=7,command=import_parent,state=DISABLED)
load_button.grid(row=1,column=2,padx=2)

#-----------------------------------------------------------------------------#
send_frame = Frame(root)
send_frame.grid(row=2,column=0)

viewsample_button = Button(send_frame, text="VIEW SAMPLE",command=view_sample,state=DISABLED)
viewsample_button.grid(row=0,column=0,padx=5)

sendall_button = Button(send_frame, text="SEND NOW", command=send_email_parent,state=DISABLED)
sendall_button.grid(row=0,column=1,padx=5)

root.mainloop()
