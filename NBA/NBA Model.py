import win32com.client
import time
import os
import csv
import bs4
import requests
import re
import datetime
import xlrd
import shutil


#Updates Database
print("Updating Database")

currentPath = "I:\\Coding Projects\\Sports Betting\\Model NBA\\NBA Database.xlsx"
xlapp = win32com.client.DispatchEx("Excel.Application")
wb = xlapp.Workbooks.Open(currentPath)
wb.RefreshAll()
time.sleep(45)
wb.Save()
time.sleep(5)
xlapp.Quit()

print("Database Updated")
##############


#Create Today's Games
os.remove("I:\\Coding Projects\\Sports Betting\\Model NBA\\Today's Games.csv")
league = ['nba']
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

today = str((datetime.datetime.today()))
today = today.split(' ')
month = today[0]
month = month.split('-')
year = month[0]
file_month = month[1]
month = int(month[1]) - 1
today = today[0]
print(today)

def date(month,nlist):
    global total, gdata
    if run == 3:
        total[0] = month + ' ' + gdata[nlist]

def teams(word):
    global gdata, run
    if word == 'Matchup':
        pass
    else:
        for current in months:
            if gdata[1] == current or gdata[2] == current or gdata[3] == current:
                word = current
    if gdata[1] == word:
        total[run] = gdata[0]
        date(word,2)
        gdata = gdata[2:]
    elif gdata[2] == word:
        total[run] = gdata[0] + ' ' + gdata[1]
        date(word,3)
        gdata = gdata[3:]
    elif gdata[3] == word:
        total[run] = gdata[0] + ' ' + gdata[1]  + ' ' + gdata[2]
        date(word,4)
        gdata = gdata[4:]

for current in league:
    link = 'https://www.oddsshark.com/' + current + '/computer-picks'
    sause = requests.get(link)
    soup = bs4.BeautifulSoup(sause.text, 'html.parser')
    for game in soup.find_all('table'):
        gdata = (game.text)
        gdata = re.split(' ', gdata)
        if gdata[0] == '':
            MD_T1 = gdata[1]
            gdata = gdata[2:]
            #Total = [Date, Team1, OS_Score1, Team2, OS_Score2, LV_Spread, LV_OU, MY_Team, MY_Spread]
            total = ['', '', '', '', '', '', '','','']
            run = 1
            teams('Matchup')
            run = 3
            MD_T2 = gdata[0]
            gdata = gdata[1:]
            teams(months[month])
            gdata = gdata[1:]
            total[2] = gdata[6][5:]
            if current != 'nba':
                total[4] = gdata[8][:4]
                if '.' not in total[4]:
                    total[4] = total[4][:2]
            else:
                total[4] = gdata[8][:5]
                if total[4].startswith('1'):
                    if '.' not in total[4]:
                        total[4] = total[4][:3]
                else:
                    if '.' not in total[4]:
                        total[4] = total[4][:2]
                    else:
                        total[4] = total[4][:4]
            if '-' not in gdata[12]:
                total[5] = gdata[11] + ' ' + gdata[12]
            else:
                if gdata[11].startswith(MD_T1):
                    total[5] = MD_T2 + ' (+' + gdata[12][2:]
                else:
                    total[5] = MD_T1 + ' (+' + gdata[12][2:]
            total[6] = gdata[14][:-6]
            spread = total[5]
            spread = re.split(' ',spread)
            if spread[0] == 'Push':
                total[7] = spread[0]
                total[8] = '0'
            else:
                if spread[0].startswith(MD_T1):
                    spread[1] = float(total[2]) + float(spread[1][2:-1]) - float(total[4])
                else:
                    spread[1] = float(total[4]) + float(spread[1][2:-1]) - float(total[2])
                total[7] = spread[0]
                total[8] = spread[1]
            with open("I:\\Coding Projects\\Sports Betting\\Model NBA\\Today's Games.csv", 'a', newline='') as file:
                wr = csv.writer(file, dialect='excel')
                wr.writerow(total)
print("Today's Games Created")
##############


#Create 4 Model Spreadsheet
Today_Path = "I:\\Coding Projects\\Sports Betting\\Model NBA\\Today's Games.csv"
Database_Path = "I:\\Coding Projects\\Sports Betting\\Model NBA\\NBA Database.xlsx"
workbook = xlrd.open_workbook(Database_Path)
school = workbook.sheet_by_index(0)
opponent = workbook.sheet_by_index(1)
record = workbook.sheet_by_index(2)
conf = workbook.sheet_by_index(3)
rankings = workbook.sheet_by_index(4)
factor = workbook.sheet_by_index(5)

def excel_info(number, sheet, other):
    for new_row in range(sheet.nrows):
            team = sheet.cell_value(new_row, other)
            if team == row[number]:
                team_excel = []
                for col in range(sheet.ncols):
                    team_excel.append(sheet.cell_value(new_row, col))
                return(team_excel)
    else:
        return('skip')

row1 = ['Team1','Team2','Spread','OS Spread','Dif Team','Difference','Game Link']
file_name = ["I:\\Coding Projects\\Sports Betting\\Model NBA\\Games\\", year, "-", file_month, "\\",today, "\\"]
file_name = ''.join(file_name)
if not os.path.exists(file_name):
    os.makedirs(file_name)
dest = file_name + "$ Index Model Data.csv"
with open(dest, 'a', newline='') as file:
    wr = csv.writer(file, dialect='excel')
    wr.writerow(row1)

game_number = 1
with open(Today_Path) as csv_file:
    csv_reader = csv.reader(csv_file)
    for row in csv_reader:
        
        team1_record = excel_info(1, record, 0)
        team2_record = excel_info(3, record, 0)
        team1_school = excel_info(1, school, 0)
        team1_opponent = excel_info(1, opponent, 0)
        team2_school = excel_info(3, school, 0)
        team2_opponent = excel_info(3, opponent, 0)
        team1_conf = excel_info(1, conf, 0)
        team2_conf = excel_info(3, conf, 0)
        team1_rankings = excel_info(1, rankings, 0)
        team2_rankings = excel_info(3, rankings, 0)
        team1_factor = excel_info(1, factor, 1)
        team2_factor = excel_info(3, factor, 1)

        #Total = [Date, Team1, OS_Score1, Team2, OS_Score2, LV_Spread, LV_OU, MY_Team, MY_Spread, T1 O, T1 D, Team1_4F, T2 O, T2 D Team2_4F, Dif Team, Difference]
        total = [row[0], row[1], row[2],row[3],row[4], row[5],row[6], row[7], row[8],'-','-','-','-','-','-','-','-',]
        
        if team1_school == 'skip' or team2_school == 'skip':
            pass
        else:
            total[9] = (float(team1_factor[17])*.4)+((float(team1_factor[18])/100)*.25)+((float(team1_factor[19])/100)*.2)+(float(team1_factor[20])*.15)
            total[10] = (float(team1_factor[21])*.4)+((float(team1_factor[22])/100)*.25)+((float(team1_factor[23])/100)*.2)+(float(team1_factor[24])*.15)
            total[12] = (float(team2_factor[17])*.4)+((float(team2_factor[18])/100)*.25)+((float(team2_factor[19])/100)*.2)+(float(team2_factor[20])*.15)
            total[13] = (float(team2_factor[21])*.4)+((float(team2_factor[22])/100)*.25)+((float(team2_factor[23])/100)*.2)+(float(team2_factor[24])*.15)
            total[11] = total[9]+total[10]
            total[14] = total[12]+total[13]
            if total[11] > total[14]:
                total[15] = total[1]
                total[16] = total[11]-total[14]
            else:
                total[15] = total[3]
                total[16] = total[14]-total[11]
        
        if team1_school == 'skip' or team2_school == 'skip':
            pass
        else:
            test = [game_number,total[1],total[3],
                    'Conf',team1_conf[1],team2_conf[1],
                    'Spread',total[5],total[6],
                    '','','',
                    'PR',team1_rankings[1],team2_rankings[1],
                    'LPR',team1_rankings[2],team2_rankings[2],
                    'Last Game','','',
                    'Record','"' + team1_record[1] + '"','"' + team2_record[1] + '"',
                    'ATS','"' + team1_record[2] + '"','"' + team2_record[2] + '"',
                    '','','',
                    'H/A','A','H',
                    'OS Score',total[2],total[4],
                    'OS Dif',total[7],total[8],
                    '','','',
                    'OFF%',total[9]*100,total[12]*100,
                    'DEF%',total[10]*100,total[13]*100,
                    'Total',total[11]*100,total[14]*100,
                    'Difference',total[15],total[16]*100,
                    '','','',
                    'FG%',float(team1_school[4]),float(team2_school[4]),
                    '3PT%',float(team1_school[7]),float(team2_school[7]),
                    'FT%',float(team1_school[5]),float(team2_school[5]),
                    'TOpG',float(team1_school[10]),float(team2_school[10]),
                    'OffREBpG',float(team1_school[8]),float(team2_school[8]),
                    'DefREBpG',float(team1_opponent[8]),float(team2_opponent[8]),
                    'STLpG',float(team1_opponent[10]),float(team2_opponent[10]),]

            new_dest = file_name + str(game_number) + " Game.csv"
            lines = [0,1,2]
            while lines[2] < 78:
                current_line = [test[lines[0]],test[lines[1]],test[lines[2]]]
                with open(new_dest, 'a', newline='') as file:
                    wr = csv.writer(file, dialect='excel')
                    wr.writerow(current_line)
                lines = [lines[0]+3,lines[1]+3,lines[2]+3]
            total = [total[1],total[3],total[5],total[8],total[15],total[16],game_number]
            with open(dest, 'a', newline='') as file:
                    wr = csv.writer(file, dialect='excel')
                    wr.writerow(total)
            game_number += 1
path = "I:\\Coding Projects\\Sports Betting\\Model NBA\\Games\\" + year + "-" + file_month + "\\" + today + "\\" + "\\edits"
if not os.path.exists(path):
    os.makedirs(path)
##############
