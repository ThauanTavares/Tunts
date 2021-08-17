#!/usr/bin/python3
from time import sleep
from os import remove
import gspread
from google.oauth2.service_account import Credentials
from gspread.models import Worksheet
from openpyxl import Workbook

#link google sheets
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
credentials = {
  "type": "service_account",
  "project_id": "tunts-323115",
  "private_key_id": "a36d8dd06775b2d854b3d3ed79a04918cbc77f65",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDITo5VSAwjOB5Q\nH5TVbgljjvmj5z02z31wsqn3ALpFzH9qjPEP5EyFL7CXYcr8IENcAPbb5Ye4LTo7\n5XZr0iuVi2Qob/XChEbJ8JEqIDApPnZlw33cewDd/e2XkefvxviSTPNeKvH4/Vib\njw70TWS1xFRNyLtPHqxbNgX7NMktL9OciVqfX6RcBWQj66/NEN33sbUMSiJXdwBw\nsq4M6ly7HPH0Biuz7ZG4WjBo3+P7dSBFr4Y3oJHITC/BLuS7BPDDhatgIEMwy8rF\nLhHf6Mn4qcomFJrx9tVEMb6y8pvrrLwCjhzMkPXD768j0sQ9U5dSr98tt9VzCS9t\ngaNoiycLAgMBAAECggEAGK/fM0aNFxLLCvVMJXPe3LIvNVKsxOJqDjzkMr5DDD2J\nyfWInlH1AQiMJf7WWCBCfQaHZjmnR6fsq0YlbrYj9gCMfb6ywOBcTcuERfdanYgG\napYJ9JBaXoHLHxOgoiOD2PAlTHVUMOufLejsnxJ3d5Q0ubN5gQ+sOWpH7sUmMux1\nNMRNmmrReLDqRnA+5P/V62wtMFVUMqemchfHD3mPGzbwGCEcKLFxW1bstYJLJM6u\n8j4yA68gjorb3DKIQlGL5ek6A5P0RWEaqQtHuJ7yEsT6jCmrOM4xCQKgBNymkDBx\no0sI6J+S7NZC/MrzS72WhQqVWRVuMlh81iQvEzH0KQKBgQD3Sh9Y3M1wRN17cBmX\nh6oj7Y4SS5uNZb9qNHVuwB6bMVWNUWKf6AQWtGFsAovlDLD2eMPdsVxTwEV4hBUU\ndNX2ohG56K9Ikcfbr5OBo05pnk5LUKkIcf81r9Mt38JCsEvybT+DRoV1UwGuwzMr\nsrqBmnn5jBtaHXWQTby466xZSQKBgQDPXMcfnqCC/joF+gtCv+BkoDrOV7Gv1lqb\nxivErsJYc9GRV97v85BKjnuRZSkeg0kuci48C0cXqlMjmYJI2sBXYGjeUvpiPt/t\njvEiG4ZW5V3rqte3gPaOW575DVDqtAWqpGCh9F4016/sWEzkYTgxRLPmsBDKeLAl\nXHPZ2iPxswKBgBPSMCMKR7k4+9u6B4Maz6tjjiCvSL+TqT0VCVigM5PS532VSWdx\nzGoZeTmUFqx8UO7gjSqG3dSks6zOQXZLSx8irHMPUIVrke5s61DXyyAyHSpyQ6o2\noPMnrbCen86CnQPId+/Ixke6KIdehAp7n/FldWNoNIULmXHAmWDlGwJpAoGAXSZQ\nTOP1NuL9LbIlXNbU9l8EC61ZJKQmD9P11WSr6RAeFCxk+WVwbA2VdLr8vbg8J6Ej\nI9XtGbRppJRQ8mlR/e5RLY4Y7AlSjmSn7apvrplal4MnZEeKemiuATL6JnS4Fu/A\neQBNUW+Sl5kS8YS6uWcjaX2uDUYp2p6tbC+R9gECgYEA3dBmnggRlr8UlXUGZB1f\ndPrKWTfpHqHGFsHbZHmmMvTurPtw0ZQ45/45ZUq4QVr1zgVKoMh49mMuJHrtTE5c\nRlvjBrXSc9XpKUJYYjYjyX8Z15uqaMuJXkdhc6S2/Eeyho4VheJ6rEfrSn5yqOzX\n2Qm494CSHeMNzLZYiHtQMHo=\n-----END PRIVATE KEY-----\n",
  "client_email": "tunts-50@tunts-323115.iam.gserviceaccount.com",
  "client_id": "102096297259060589468",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/tunts-50%40tunts-323115.iam.gserviceaccount.com"
}

gc = gspread.service_account_from_dict(credentials)
sh = gc.open("Engenharia de Software – Desafio Thauan Tavares:")


def Check_Average(Students,line):#receive the student node, analyzes the data and calculates the result.
    line = line + 4
    limitattendance=(60*0.25)
    attendance=int(Students[2])
    average=(int(Students[3])+int(Students[4])+int(Students[5]))/3
    avarage=round(average)
    if(attendance <= limitattendance):
        if(avarage<50):
            sh.sheet1.update_cell(line,7,'Reprovado por Nota')
            sh.sheet1.update_cell(line,8,'0')

        if(avarage<70 and avarage >=50):
            sh.sheet1.update_cell(line,7,'Exame Final')
            naf=(100-avarage)
            sh.sheet1.update_cell(line,8,naf)
        if(avarage > 70):
            sh.sheet1.update_cell(line,7,'Aprovado')
            sh.sheet1.update_cell(line,8,'0')
    else:
        sh.sheet1.update_cell(line,7,'Reprovado por Falta')
        sh.sheet1.update_cell(line,8,'0')

def CompletTable(Students):#Run Students list and call Check_Average
    sleep(30)#because of limit Read and Write requests per minute per user of service
    for x in range(24): 
        Check_Average(Students[x],x)

def Main():
    lines=[] #catch lines
    temp=[]  #Temporary 
    if((sh.sheet1.acell('G4').value) and (sh.sheet1.acell('G5').value)) == '*': #Test if table was set, else set table and run program
        for i in range(4,28):
            temp=sh.sheet1.row_values(i)
            lines.append(temp)
    else:
        for i in range(4,28):    
            sh.sheet1.update_cell(i,8,'*')
        sleep(30)#because of limit Read and Write requests per minute per user of service
        for i in range(4,28):    
            sh.sheet1.update_cell(i,7,'*')
    sleep(30)#because of limit Read and Write requests per minute per user of service
    CompletTable(lines)
Main()