from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
import time
import os
import pandas as pd
import datetime


# usrnme = os.getenv('username')
# print(usrnme)
# usrnme=f'user-data-dir=C:/Users/{usrnme}/AppData/Local/Google/Chrome/User Data/Default'

meeting_link=input("Enter the meeting link")
gmail_id=input("Enter the gmail from which you have hosted the meeting: ")
password=input('Enter the password: ')
print('The file will be saved in the format"date filename.xlsx"')
file_name=input('Enter the file name: ')
no_of_present=[0]
no_of_absent=[0]


ch_options=Options()
driver_path='chromedriver.exe'
ch_options.add_argument("--start-maximized")
ch_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3641.0 Safari/537.36 ")
ch_options.add_experimental_option("prefs", { \
    "profile.default_content_setting_values.media_stream_mic": 1, 
    "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.geolocation": 1, 
    "profile.default_content_setting_values.notifications": 2
    
  })
# ch_options.add_argument(usrnme)
print("opening the driver")
driver=webdriver.Chrome(executable_path=driver_path,chrome_options=ch_options)

driver.get(r'https://accounts.google.com/o/oauth2/auth/identifier?client_id=717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com&scope=profile%20email&redirect_uri=https%3A%2F%2Fstackauth.com%2Fauth%2Foauth2%2Fgoogle&state=%7B%22sid%22%3A1%2C%22st%22%3A%2259%3A3%3Abbc%2C16%3Ac51291c820784476%2C10%3A1607098594%2C16%3A8181d90b08d0e2c4%2C670c70e18677663e4fc5e355bc806d81b97c1bdd94cc3583cbdd2b3cb02a1310%22%2C%22cdl%22%3Anull%2C%22cid%22%3A%22717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com%22%2C%22k%22%3A%22Google%22%2C%22ses%22%3A%22696a343211d44cd79cb26ed80ad4f9e7%22%7D&response_type=code&flowName=GeneralOAuthFlow')
# https://accounts.google.com/o/oauth2/auth/identifier?client_id=717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com&scope=profile%20email&redirect_uri=https%3A%2F%2Fstackauth.com%2Fauth%2Foauth2%2Fgoogle&state=%7B%22sid%22%3A1%2C%22st%22%3A%2259%3A3%3Abbc%2C16%3Ac51291c820784476%2C10%3A1607098594%2C16%3A8181d90b08d0e2c4%2C670c70e18677663e4fc5e355bc806d81b97c1bdd94cc3583cbdd2b3cb02a1310%22%2C%22cdl%22%3Anull%2C%22cid%22%3A%22717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com%22%2C%22k%22%3A%22Google%22%2C%22ses%22%3A%22696a343211d44cd79cb26ed80ad4f9e7%22%7D&response_type=code&flowName=GeneralOAuthFlow

email_id=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='whsOnd zHQkBf']")))
email_id.send_keys(gmail_id)

next_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ VfPpkd-LgbsSe-OWXEXe-dgl2Hf nCP5yc AjY5Oe DuMIQc qIypjc TrZEUc']")))
next_.click()

time.sleep(3)

email_id=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='whsOnd zHQkBf']")))
email_id.send_keys(password)

next_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ VfPpkd-LgbsSe-OWXEXe-dgl2Hf nCP5yc AjY5Oe DuMIQc qIypjc TrZEUc']")))
next_.click()
print('joining he meeting')
driver.get(r'https://meet.google.com/?hl=en')

meet_signin=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//a[@class='glue-header__link ']")))
meet_signin.click()

link_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='VfPpkd-fmcmS-wGMbrd B5oKfd']")))
link_.send_keys(meeting_link)

join=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-dgl2Hf ksBjEc lKxP2d cjtUbb']")))
join.click()

mic=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='U26fgb JRY2Pb mUbCce kpROve uJNmj HNeRed QmxbVb']")))
mic.click()

cam=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='U26fgb JRY2Pb mUbCce kpROve uJNmj QmxbVb FTMc0c N2RpBe jY9Dbb']")))
cam.click()

time.sleep(1.5)
join_now=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='uArJ5e UQuaGc Y5sE8d uyXBBb xKiqt']")))
join_now.click()





# time.sleep(5)
# driver.find_element_by_css_selector('div.uArJ5e.UQuaGc.kCyAyd.QU4Gid.foXzLb.IeuGXd').click()  # participant list
# time.sleep(1)

join_now=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='uArJ5e UQuaGc kCyAyd QU4Gid foXzLb IeuGXd']")))
join_now.click()
print('retriving participants')
time.sleep(2)
participants=driver.find_elements_by_class_name("ZjFb7c")

# for participant in participants:
#   print('Names:  ',participant.text)

participants_list=[participant.text for participant in participants]

participants_df=pd.DataFrame({'Participants':participants_list})
temp_path=r'E:\Google meet\temp.xlsx'
participants_df.to_excel(temp_path)
# print('saved')
print('Acessing the name list and preparing the attendance')
Name_list_df=pd.read_excel('namelist.xlsx')
temp_df=pd.read_excel(temp_path)

# print(Name_list_df)
# print(Name_list_df.columns)

name_list=Name_list_df['Names'].values


temp_list=temp_df['Participants'].values


Attendance=[]
Present_name_list=[]
Absent_name_list=[]
Unknown_name_list=[]

for one in name_list:
  if one in temp_list:
    Attendance.append('Present')
    no_of_present[0]+=1
    Present_name_list.append(one)
  else:
    Attendance.append('Absent')
    no_of_absent[0]+=1
    Absent_name_list.append(one)

for one_ in temp_list:
  if one_ not in name_list:
    Unknown_name_list.append(one_)
    
# credit
# https://stackoverflow.com/questions/27126511/add-columns-different-length-pandas/33404243

lname_list,lAttendance,lPresent_name_list,lAbsent_name_list,lno_of_present,lno_of_absent,lUnknown_name_list = len(name_list),len(Attendance),len(Present_name_list),len(Absent_name_list),1,1,len(Unknown_name_list)
max_len = max(lname_list,lAttendance,lPresent_name_list,lAbsent_name_list,lno_of_present,lno_of_absent)

if not max_len == lname_list:
  name_list.extend(['']*(max_len-lname_list))
if not max_len ==lAttendance:
  Attendance.extend(['']*(max_len-lAttendance))
if not max_len == lPresent_name_list:
  Present_name_list.extend(['']*(max_len-lPresent_name_list))
if not max_len == lAbsent_name_list:
  Absent_name_list.extend(['']*(max_len-lAbsent_name_list))
if not max_len == lno_of_present:
  no_of_present.extend(['']*(max_len-lno_of_present))
if not max_len == lno_of_absent:
  no_of_absent.extend(['']*(max_len-lno_of_absent))
if not max_len == lUnknown_name_list:
  Unknown_name_list.extend(['']*(max_len-lUnknown_name_list))



Final=pd.DataFrame({
  'Names':name_list,
  'Attendance':Attendance,
  'Attendees':Present_name_list,
  'Absentees':Absent_name_list,
  'Unkown Participants':Unknown_name_list,
  'No of people Present':no_of_present,
  'No of people Absent':no_of_absent
})


print('saving the attendance')
import datetime
x = datetime.datetime.now()
# print(x.time())
Filename=str(x.day)+'-'+str(x.month)+'-'+str(x.year)+' '+ f'{file_name}.xlsx'
Filepath='E:\\Google meet\\Attendance\\'+Filename 

# print(Filepath)
Final.to_excel(Filepath)

print('completed')

