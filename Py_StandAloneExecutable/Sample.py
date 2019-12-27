# -*- coding: utf-8 -*-
"""
Created on Thu Dec 19 14:30:36 2019

@author: jmore
"""
import random 
import csv
import win32com.client as win32
import os

def cleanup_raw_list(inp_file):
    """
    Input: file name containing the participants list. 
    The input data is expected to be as each participant is seperated with a ; semicolon
    Returns a list of unique participants
    """
    try:
        participants_fp = open(inp_file, "r")
        raw_list = participants_fp.read().split(';')
        clean_list = list(set([p.strip() for p in raw_list if len(p) > 6]))
        #create unique list of participants
        return clean_list
    except FileNotFoundError:
        print("Make sure the input file {} exists in its proper location!".format(inp_file))
    except:
        print("Something went wrong!")
    finally:
        try:
            participants_fp.close()
        except:
            pass

def check_undeflow(participants):
    """
    Input: list of participants
    Checks if there are less than 2 participants in the list
    """
    return (len(participants) < 2)

def create_a_map(participants):
    """
    Input: List of participants
    Randomly creates a hashmap containing Participant:Santa
    """
    santa_mappings = dict() #resultant hasachmap
    mapped_participants_i = set() #to make sure every participant gets a santa
    counter = 0
    
    while(counter < len(participants)):            
        random_participant_i = random.randint(-len(participants),-1)
        #complexity of comparision with strings can be reduced if we use random.randint(-len(participants),-1)
        random_map_i = random.randint(-len(participants),-1)
        #Basic conditions
        ##If the participant is already mapped; Avoid symmetry; Ensure every participant gets a santa
        if((participants[random_participant_i] in santa_mappings) or (random_participant_i == random_map_i) or (random_map_i in mapped_participants_i)):
            continue
        else:
            santa_mappings[participants[random_participant_i]] = participants[random_map_i]
            mapped_participants_i.add(random_map_i)
            counter += 1
    
    return santa_mappings

def send_indv_mail(santa_mappings):
    """
    Input: Hashmap of Participant:Santa
    Triggers a local outlook application and sends an email to all the particiants in the hashmap with their Santa
    """
    #Retrieve mapping and send their secret santa's over mail
    
    #Create an instance of outlook application
    outlook = win32.Dispatch('outlook.application')

    for p,s in santa_mappings.items():
        # p: Participant  s: Santa
        mail = outlook.CreateItem(0) #new mail task
        mail.To = p #To field 
        mail.Subject = 'Your secret santa is here..'
        body = "Hello, {}! <br> Your \t secret santa is \t *{}*\n".format(p,s)
        footer = "<br><br><br>This is an automated mail to test the Python script. Please ignore it."
        mail.HTMLBody = '<h2>'+ body +'</h2>' + footer
        mail.Send()
        
        #Console output -- comment the next line if you dont' want to print the output in the console
        #print("{}'s \t Santa is \t *{}*\n".format(p,s))
    
def send_map_list(santa_mappings,organizer_id):    
    """
    Input: Hashmap of Participant:Santa and a default argument organizer_id
    Sends the hashmap list of to the organizer. So, only organizer knows the details
    """
    #output to external file
    output_file_name = "santa_maps.csv"
    #Create an instance of outlook application
    outlook = win32.Dispatch('outlook.application')
    try:
        f = open(output_file_name, 'w', newline='')
        writer = csv.DictWriter(f, fieldnames = ["Participant","Santa"])
        writer.writeheader()
        for p,s in santa_mappings.items():
            writer.writerow({'Participant': p, 'Santa': s})

    except PermissionError:
        print("Please close the file {} if its opened, or it doesn't have the permissions to write the file".format(output_file_name))
    except:
        print("Something went wrong!")    
    finally:
        f.close()
        #Send it to the organizer over mail
        #New_Email is at index 0 in the outlook app
        mail = outlook.CreateItem(0)
        #organizer_id = input("Organizer's mail ID: ")
        mail.To =  organizer_id #To field
        mail.Subject = 'Secret Santa | List'
        mail.HTMLBody = '<h2>List of mapings attached to the mail</h2>'
        file_dir = os.getcwd() #present working directory
        attachment  = file_dir + "/" + output_file_name
        mail.Attachments.Add(attachment)
        mail.Send()


#################
#---------------------Execution starts here-------------
#Input: 'participants_data.txt'. The data is expected to be as each participant is seperated with a ; semicolon

inp_file = "participants_data.txt"
participants = cleanup_raw_list(inp_file)
print("Participants: {}\n".format(participants))

if(check_undeflow(participants)):
    print("Participants not suffice. Make a big team and try again!")
    exit(0)
else:    
    santa_mappings = create_a_map(participants)
    send_indv_mail(santa_mappings)
    send_map_list(santa_mappings,input("Organizer's mail ID: "))
    #print("OUTSIDE function santa_mappings: {} ".format(santa_mappings))
print("End of the script")
input() #Just to hold the console


# To Create Standalone application of a Python Script 
###### pyinstaller.exe --onefile MyCode.py