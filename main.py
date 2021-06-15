import csv
import os

import openpyxl
import win32com.client
import ctypes
import pythoncom
import re
import time
import psutil
import pandas as pd


def edita_planilha(data):
    print('--> Salvando na planilha ...')
    dict_data = pd.read_excel('controle.xlsx', index_col="Veículo",
                              usecols=["Veículo", "Concessionária", "Hora", "Data"])
    dict_data = dict_data.append(data, ignore_index=True)
    print(dict_data)
    exist_veiculo = False
    for row in dict_data.iterrows():
        if str(row[1]['Veículo']) == str(data['Veículo']):
            exist_veiculo = True

    if not exist_veiculo:
        writer = pd.ExcelWriter('controle.xlsx', mode='w', engine='openpyxl')
        dict_data.to_excel(writer, 'controle.xlsx', encoding='utf-8')
        writer.save()


edita_planilha({
    'Concessionária': 'americana',
    'Data': '14:00',
    'Hora': '15:00',
    'Veículo': '36589',
})


class Handler_Class(object):

    def __init__(self):
        # First action to do when using the class in the DispatchWithEvents
        inbox = self.Application.GetNamespace("MAPI").GetDefaultFolder(6)
        messages = inbox.Items
        # Check for unread emails when starting the event
        print("--> Pesquisando e-mail's")
        for message in messages:
            if message.UnRead:
                subject = message.Subject
                if 'Notificação: Um veículo logo chegará' in subject:
                    print("--> Um novo veículo chegou, analisando corpo do e-mail....")
                    string = message.Body
                    lines = string.split('\r\n')
                    lines_stripped = []
                    for line in lines:
                        line = line.strip()
                        if line != '':
                            lines_stripped.append(line)

                    index = 0
                    data = {}
                    print(lines_stripped)
                    for string in lines_stripped:
                        if 'Número de registro' in string:
                            data['Veículo'] = lines_stripped[index + 1]
                            data['Concessionária'] = lines_stripped[index + 3].split(',')[1]
                            data['Data'] = lines_stripped[index + 6].split(' ')[0]
                            data['Hora'] = lines_stripped[index + 6].split(' ')[1]
                        index += 1
                    edita_planilha(data)
                    message.UnRead = False
                # print(message.Subject) # Or whatever code you wish to execute.

    def OnQuit(self):
        # To stop PumpMessages() when Outlook Quit
        # Note: Not sure it works when disconnecting!!
        ctypes.windll.user32.PostQuitMessage(0)

    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = self.Session.GetItemFromID(ID)
            subject = mail.Subject
            if 'Notificação: Um veículo logo chegará':
                print("--> Um novo veículo chegou, analisando corpo do e-mail....")
                string = mail.Body
                lines = string.split('\r\n')
                lines_stripped = []
                for line in lines:
                    line = line.strip()
                    if line != '':
                        lines_stripped.append(line)

                index = 0
                data = {}
                print(lines_stripped)
                for string in lines_stripped:
                    if 'Número de registro' in string:
                        data['Veículo'] = lines_stripped[index + 1]
                        data['Concessionária'] = lines_stripped[index + 3].split(',')[1]
                        data['Data'] = lines_stripped[index + 6].split(' ')[0]
                        data['Hora'] = lines_stripped[index + 6].split(' ')[1]
                    index += 1
                edita_planilha(data)
                mail.UnRead = False
            try:
                command = re.search(r"%(.*?)%", subject).group(1)
                print(command)  # Or whatever code you wish to execute.
            except:
                pass


# Function to check if outlook is open
def check_outlook_open():
    list_process = []
    for pid in psutil.pids():
        p = psutil.Process(pid)
        # Append to the list of process
        list_process.append(p.name())
    # If outlook open then return True
    if 'OUTLOOK.EXE' in list_process:
        return True
    else:
        return False


# Loop
while True:
    try:
        outlook_open = check_outlook_open()
    except:
        outlook_open = False
    # If outlook opened then it will start the DispatchWithEvents
    if outlook_open == True:
        outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
        pythoncom.PumpMessages()
    # To not check all the time (should increase 10 depending on your needs)
    time.sleep(10)
