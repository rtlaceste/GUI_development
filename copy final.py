# -*- coding: utf-8 -*-
"""
Created on Mon Nov 30 06:42:19 2020

@author: RLaceste
"""

from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import pyautogui
import os
#from simple_salesforce import Salesforce
from PIL import ImageTk,Image
import json
import win32com.client as win32
from datetime import datetime, timedelta
from datetime import date

# =============================================================================
# options = webdriver.ChromeOptions()
# options.add_experimental_option("prefs", {
#   "download.default_directory": r"C:\Users\Troy\Desktop\Download Testing",
#   "download.prompt_for_download": False,
#   "download.directory_upgrade": True,
#   "safebrowsing.enabled": True
# })
# 
# chrome_browser = webdriver.Chrome(r'C:\Users\Troy\Desktop\chromedriver.exe', options = options)
# chrome_browser.get("https://chromedriver.storage.googleapis.com/index.html?path=87.0.4280.20/")
# time.sleep(3)
# try:
#     chrome_browser.find_element_by_xpath("//a[contains(text(),'win32')]").click()
# except Exception:
#     print('No Such Element found')
# =============================================================================

# Send Emails 
# =============================================================================
# outlook = win32.Dispatch('outlook.application')
# mail = outlook.CreateItem(0)
# mail.To = 'abaez@axosbank.com'
# mail.CC = ""
# mail.Subject = "test"
# mail.Body = "Cheese"
# mail.Send()
# =============================================================================
############JSON################
# =============================================================================
# json_file = open(r"C:\Users\rlaceste\Documents\Custom Office Templates\script.json","r",encoding='utf-8')
# info = json.load(json_file)
# json_file.close()
# 
# =============================================================================

# =============================================================================
# print(info['SalesForce Username'])
# =============================================================================


#driver = webdriver.Chrome(r'C:\Users\Troy\Desktop\chromedriver.exe')    
   
#driver.get("https://www.google.com/") 
#driver.maximize_window
  
# get  element  
#element = driver.find_element_by_name('q')
#element.send_keys("Hi")
  
# create action chain object 
#action = ActionChains(driver) 
  
# perform the operation 
#action.move_to_element_with_offset(element, 100, 50).click().perform() 



    
    
    
    

def UploadAction1(event=None):
    filename1 = filedialog.askopenfilename()
    print('Selected:', filename1)
    
def UploadAction2(event=None):
    filename2 = filedialog.askopenfilename()
    print('Selected:', filename2)
    
def UploadAction3(event=None):
    filename3 = filedialog.askopenfilename()
    print('Selected:', filename3)

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    



root = Tk()
root.geometry("1400x650")
root.title("Loan Applicant Information ver. 1.0")


OPTIONSi = ["None","Simple","Complex"]
variable = StringVar(root)
variable.set(OPTIONSi[0]) # default value
#print(variable.get()) = yes
w = OptionMenu(root, variable, *OPTIONSi)
w.pack()
w.place(relx=.75,rely=.45)


folder_path = StringVar()
lbl1 = Label(root,textvariable=folder_path)
lbl1.pack()
button2 = Button(text="Send to what folder?", command=browse_button)
button2.pack()
button2.place(relx=.4,rely=.6)

# =============================================================================
# image1 = Image.open(r"C:\Users\rlaceste\Downloads\Axos logo.png")
# image2 = image1.resize((100,100))
# test = ImageTk.PhotoImage(image2)
#label1 = Label(image=test)
#label1.image = test
#label1.grid(row=0,column=0)
# =============================================================================

#root.grid_columnconfigure((0,14), weight=1)




labeltitle = Label(root, text = "Axos Bank", font=("Courier", 25),fg='DarkGoldenrod2', bg='navy')
labeltitle.pack()
labeltitle.place(relx=.68,rely=.0001)


# =============================================================================
# labeltitle2 = Label(root, text = "", font=("Courier", 25))
# labeltitle2.pack()
# labeltitle2.place(relx=.75,rely=.0001)
# =============================================================================
#labeltitle3 = Label(root, text = "If info not known, leave blank. Thanks.\n Axos Bank IPL Department", font=("Courier", 8)).grid(row=35, column=0)




applabel1 = Label(root,text = "Applicant 1")
applabel1.pack()
applabel1.place(relx=.25,rely=.05)


applabel2 = Label(root,text = "Applicant 2")
applabel2.pack()
applabel2.place(relx=.4,rely=.05)


Label3 = Label(root, text="First Name")
Label3.pack()
Label3.place(relx=.15,rely=.1)

Label4 = Label(root, text="Last Name")
Label4.pack()
Label4.place(relx=.15,rely=.15)


Label5 = Label(root, text="City")
Label5.pack()
Label5.place(relx=.15,rely=.2)


Label6 = Label(root, text="Loan Number")
Label6.pack()
Label6.place(relx=.15,rely=.25)


Label7 = Label(root, text ="SSN")
Label7.pack()
Label7.place(relx=.15,rely=.3)


Label8 = Label(root, text = "Address No.")
Label8.pack()
Label8.place(relx=.15,rely=.35)

Label9= Label(root, text = "Street Name")
Label9.pack()
Label9.place(relx=.15,rely=.4)


Label10 = Label(root, text = "State (Ex: CA, NY)")
Label10.pack()
Label10.place(relx=.15,rely=.45)


Label11 = Label(root, text = "Zip Code")
Label11.pack()
Label11.place(relx=.15,rely=.5)


Labela = Label(root, text = "L.O. First Name")
Labela.pack()
Labela.place(relx=.15,rely=.55)


Labelb = Label(root, text = "L.O. Last Name")
Labelb.pack()
Labelb.place(relx=.15,rely=.6)


Labelc = Label(root, text = "JCA First Name")
Labelc.pack()
Labelc.place(relx=.15,rely=.65)



Labeld = Label(root, text = "JCA Last Name")
Labeld.pack()
Labeld.place(relx=.15,rely=.7)


Labele = Label(root, text = "L.P. First Name")
Labele.pack()
Labele.place(relx=.15,rely=.75)



Labelf = Label(root, text = "L.P. Last Name")
Labelf.pack()
Labelf.place(relx=.15,rely=.8)

Labelg = Label(root, text = "L.C. First Name")
Labelg.pack()
Labelg.place(relx=.15,rely=.85)


Labelh = Label(root, text = "L.C. Last Name")
Labelh.pack()
Labelh.place(relx=.15,rely=.9)


# =============================================================================
# button_LOI = ttk.Button(root, text='LOI', command=UploadAction1)
# button_LOI.grid(row=0,column=5)
# 
# button_Credit_auth = ttk.Button(root, text='Credit Auth', command=UploadAction2)
# button_Credit_auth.grid(row=0,column=6)
# 
# button_App_D = ttk.Button(root, text='App Deposit', command=UploadAction3)
# button_App_D.grid(row=0, column=7)
# 
# =============================================================================

Label26 = Label(root, text ="CoStar")
Label26.pack()
Label26.place(relx=.6,rely=.15)

Label12 = Label(root, text = "OFAC")
Label12.pack()
Label12.place(relx=.6,rely=.2)

Label14 = Label(root, text = 'Google Search')
Label14.pack()
Label14.place(relx=.6,rely=.25)

Label13 = Label(root, text = 'Credit Report')
Label13.pack()
Label13.place(relx=.6,rely=.3)

Label15 = Label(root, text = "Lexis Nexis")
Label15.pack()
Label15.place(relx=.6,rely=.35)

Label16 = Label(root, text = "IRS TIN (n/a)")
Label16.pack()
Label16.place(relx=.6,rely=.4)

Label17 = Label(root, text = "Flood Certification")
Label17.pack()
Label17.place(relx=.6,rely=.45)



Label19 = Label(root, text = "Environmental Report")
Label19.pack()
Label19.place(relx=.6,rely=.5)


Label28 = Label(root, text = "Operation and Maintenance Plan")
Label28.pack()
Label28.place(relx=.6,rely=.55)


#Label20 = Label(root, text = "Phase 1").grid(row =10, column = 6)
#Label21 = Label(root, text = "Inspection Report").pack()


Label23 = Label(root, text = "Legal Ticket")
Label23.pack()
Label23.place(relx=.6,rely=.65)

Label24 = Label(root, text = "Salesforce Autofill")
Label24.pack()
Label24.place(relx=.6,rely=.75)


Label25 = Label(root, text = "Confirmation Email")
Label25.pack()
Label25.place(relx=.6,rely=.7)

Label22 = Label(root, text = "Order Appraisal")
Label22.pack()
Label22.place(relx=.6,rely=.6)


#Label27 = Label(root, text = "3rd Party Report E-mail").pack()


c1 = BooleanVar(value=0)
c2 = BooleanVar(value=0)
c3 = BooleanVar(value=0)
c4 = BooleanVar(value=0)
c5 = BooleanVar(value=0)
c6 = BooleanVar(value=0)
c7 = BooleanVar(value=0)
c8 = BooleanVar(value=0)
c9 = BooleanVar(value=0)
c10 = BooleanVar(value=0)
c11= BooleanVar(value=0)
c12= BooleanVar(value=0)
c12= BooleanVar(value=0)
c13 = BooleanVar(value=0)
c14 = BooleanVar(value=0)
c15 = BooleanVar(value=0)
c16 = BooleanVar(value=0) # Email Confirmation
c17 = BooleanVar(value=0) #costar
c18 = BooleanVar(value=0) # Third Party Report Confirmation
c19 = BooleanVar(value=0) # ACM
c20 = BooleanVar(value=0) #LBP
c21 = BooleanVar(value=0) #ACM and LBP
c22 = BooleanVar(value=0) #Appraisal External


costar = Checkbutton(root, text="", variable=c17)
costar.pack()
costar.place(relx=.75,rely=.15)

OFAC = Checkbutton(root, text="", variable=c1)
OFAC.pack()
OFAC.place(relx=.75,rely=.2)

Google_Search = Checkbutton(root, text="", variable=c3)
Google_Search.pack()
Google_Search.place(relx=.75,rely=.25)

Credit_Report = Checkbutton(root, text="individual", variable=c2)
Credit_Report.pack()
Credit_Report.place(relx=.75,rely=.3)


Credit_Report_Joint = Checkbutton(root, text="joint", variable=c15)
Credit_Report_Joint.pack()
Credit_Report_Joint.place(relx=.81,rely=.3)

LN = Checkbutton(root, text="", variable=c4)
LN.pack()
LN.place(relx=.75,rely=.35)

IRS_Tin = Checkbutton(root, text="", variable=c5)
IRS_Tin.pack()
IRS_Tin.place(relx=.75,rely=.4)

# =============================================================================
# f_cert_8 = Checkbutton(root, text="simple", variable=c6)
# f_cert_8.pack()
# f_cert_8.place(relx=.75,rely=.45)
# 
# f_cert_30 = Checkbutton(root, text="complex", variable=c7)
# f_cert_30.pack()
# f_cert_30.place(relx=.81,rely=.45)
# =============================================================================


ETS = Checkbutton(root, text="ETS", variable=c8)
ETS.pack()
ETS.place(relx=.75,rely=.5)


EDR = Checkbutton(root, text="EDR", variable=c13)
EDR.pack()
EDR.place(relx=.81,rely=.5)

P_1 = Checkbutton(root, text="Phase 1", variable=c9)
P_1.pack()
P_1.place(relx=.87,rely=.5)

Operation_and_maintenance_ACM = Checkbutton(root, text = "ACM", variable = c19)
Operation_and_maintenance_ACM.pack()
Operation_and_maintenance_ACM.place(relx=.75,rely=.55)

Operation_and_maintenance_LBP = Checkbutton(root, text = "LBP", variable = c20)
Operation_and_maintenance_LBP.pack()
Operation_and_maintenance_LBP.place(relx=.81,rely=.55)
    
Operation_and_maintenance_ACM_LBP = Checkbutton(root, text = "ACM & LBP", variable = c21)
Operation_and_maintenance_ACM_LBP.pack()
Operation_and_maintenance_ACM_LBP.place(relx=.87,rely=.55)



IR = Checkbutton(root, text="IR", variable=c10)
IR.pack()
IR.place(relx=.75,rely=.9)


OA = Checkbutton(root, text='Internal',variable=c11)
OA.pack()
OA.place(relx=.75,rely=.6)

OA_ex = Checkbutton(root, text='External',variable=c22)
OA_ex.pack()
OA_ex.place(relx=.81,rely=.6)

LT = Checkbutton(root, text='',variable=c12)
LT.pack()
LT.place(relx=.75,rely=.65)

SF = Checkbutton(root, text ='', variable =c14)
SF.pack()
SF.place(relx=.75,rely=.75)


email_confirm = Checkbutton(root, text = "Add comments: ", variable=c16)
email_confirm.pack()
email_confirm.place(relx=.75,rely=.7)

#party_email_confirm = Checkbutton(root, text = "Missing Items: ", variable=c18).pack() #Confirm 3rd Party Reports

# =============================================================================
# variable = StringVar(root)
# variable.set("one") # default value
# 
# w = OptionMenu(root, variable,text, "one", "two", "three")
# w.grid(row=10,column=7)
# 
# =============================================================================



def Get_OFAC():    
    
    #try:
        #os.mkdir(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}")
    #except Exception:
        #pass
    
    chrome_browser = webdriver.Chrome(r'C:\Users\rlaceste\Desktop\chromedriver.exe')
    chrome_browser.maximize_window()
    chrome_browser.get('https://sanctionssearch.ofac.treas.gov/default.aspx')
    chrome_browser.find_element_by_id("ctl00_MainContent_txtLastName").send_keys(Entry3.get())
    
    chrome_browser.find_element_by_id("ctl00_MainContent_txtAddress").send_keys(Entry4.get())
    
    chrome_browser.find_element_by_id("ctl00_MainContent_txtCity").send_keys(Entry5.get())
    
    chrome_browser.find_element_by_id("ctl00_MainContent_btnSearch").click()
    
    chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.CONTROL, 'a')
    chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.BACKSPACE)
    chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys('93')
    
    time.sleep(.5)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}\{Entry3.get()} OFAC.png")
   
    
   
    
    chrome_browser.get("http://www.google.com")
    search = chrome_browser.find_element_by_name('q')
    search.send_keys(f'"{Entry3.get()} {Entry4.get()}" AND "money laundering" OR "fraud" OR "lawsuits"')
    search.send_keys(Keys.RETURN) # hit return after you enter search text
    myScreenshot1 = pyautogui.screenshot()
    myScreenshot1.save(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}\{Entry3.get()} google search 1.pdf")
    chrome_browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    myScreenshot1 = pyautogui.screenshot()
    myScreenshot1.save(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}\{Entry3.get()} google search 2.pdf")
    time.sleep(1) # sleep for 5 seconds so you can see the results
   
    
    chrome_browser.get('https://www.credco.com/ecredco/security/login.aspx')
    time.sleep(1)
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Username").send_keys("rlaceste20")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Password").send_keys("Panda544!")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_LoginButton").click()
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_ctl00_liOrder").click()
    
    
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtLastName").send_keys(f"{Entry4.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtFirstName").send_keys(f"{Entry3.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtSSN").send_keys(f"{Entry7.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtNum").send_keys(f"{Entry8.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtStreetName").send_keys(f"{Entry9.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtCity").send_keys(f"{Entry5.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_ddlState").send_keys(f"{Entry10.get()}")
    chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtZip").send_keys(f"{Entry11.get()}")
    
    
    chrome_browser.get("https://la1.www4.irs.gov/eauth/pub/login.jsp?Data=VGFyZ2V0TG9BPUY%253D&TYPE=33554433&REALMOID=06-0004b429-9e8b-1a23-9e85-163b0acf4037&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=UOkC7yx4eMTO24FGxPfBRb5q3Mj3Xh3pyXfBEjYyHJ97nGCXu16wx5MzFHjfZmlG&TARGET=-SM-https%3a%2f%2fla1%2ewww4%2eirs%2egov%2fesrv%2ftinm%2fproauth%2efaces")
    
    
    chrome_browser.get("https://amlinsight.lexisnexis.com/")
    chrome_browser.find_element_by_id("LOGINID").send_keys(f"RLaceste558")
    chrome_browser.find_element_by_id("PASSWORD").send_keys(f"Panda544")
    chrome_browser.find_element_by_id("SIGNON").click()
    
    
    
    
    chrome_browser.get("https://bofi.my.salesforce.com/?ec=302&startURL=%2Fvisualforce%2Fsession%3Furl%3Dhttps%253A%252F%252Fbofi.lightning.force.com%252Flightning%252Fr%252FReport%252F00O3o0000055toBEAQ%252Fview%253FqueryScope%253DuserFolders")
    chrome_browser.find_element_by_id("username").send_keys("rlaceste@axosbank.com")
    chrome_browser.find_element_by_id("password").send_keys("Troyboyoy544!")
    chrome_browser.find_element_by_id("Login").click()
    loan_search = chrome_browser.find_element_by_id("169:0;p")
    loan_search.send_keys(f"{Entry6.get()}")
    loan_search.send_keys(Keys.RETURN)
    
    
    os.startfile(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}")
    


def Master_Function():
   
    startTime = datetime.now().second
    
    
   
    
    global status
    status = {"OFAC":"No", "Individual Credit Report":"No", "Joint Credit Report":"No", "Google Search": "No", "Lexis Nexis": "No", "IRS-TIN":"No"
              , "Flood Cert - Simple": "No", "Flood Cert - Complex": "No", "Environmental Report (ETS)":"No", "Environmental Report (Phase 1)":"No",
              'Environmental Report (EDR)': "No", "Inspection Report": "No", "Order Appraisal": "No", "Legal Ticket": "No"}
    
# =============================================================================
#     try:
#         os.mkdir(fr"C:\Users\rlaceste\Desktop\Intake Checks\Loan #{Entry6.get()} {Entry3.get()} {Entry4.get()}")
#     except Exception:
#         pass
#     
# =============================================================================
    
    try:
        filename1.save(fr"{folder_path.get()}")
    except Exception:
        pass
        
    try:
        filename2.save(fr"{folder_path.get()}")
    except Exception:
        pass
        
    try:
        filename3.save(fr"{folder_path.get()}")
    except Exception:
        pass
    
  
    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
    "download.default_directory": r"C:\Users\rtlac\downloads", #Change default directory for downloads
    "download.prompt_for_download": False, #To auto download the file
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
    })
  
    
  
    

    options.add_argument("--disable-notifications")

    chrome_browser = webdriver.Chrome(r'C:\Users\Troy\Desktop\chromedriver.exe', options = options)
    chrome_browser.maximize_window()
    
    
    
    #OFAC
    
    def co_verify():
        pass
    
    
    if c17.get():
        # Costar
        
        chrome_browser.get("https://www.costar.com/")
        time.sleep(1)
        chrome_browser.find_element_by_xpath('//*[@id="__next"]/div/div[1]/div/div[1]/header/div[2]/div[1]/div/ul/li[12]/a').click()
        time.sleep(1.5)
        chrome_browser.find_element_by_id("username").send_keys("abaez@axosbank.com")
        chrome_browser.find_element_by_id("password").send_keys("YellowStar0709!!")
        chrome_browser.find_element_by_id("loginButton").click()
        time.sleep(8)
        chrome_browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/input').send_keys("testing")
        
        chrome_browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[3]/a').click()
        
        
# =============================================================================
#         co_pw = Toplevel(root)
#         co_pw.title("Costar Password")
#         co_pw.geometry("100x100")
# =============================================================================
        
# =============================================================================
#         global co_pw_prompt
#         
#         co_pw_prompt = Entry(co_pw)
#         co_pw_prompt.grid(row=1,column=0)
#         co_pw_prompt.configure( highlightcolor="blue")
#         Label_co_pw = Label(co_pw, text="Costar Verification #")
#         Label_co_pw.grid(row=0,column=0)
#         co_go = ttk.Button(co_pw, text='Submit')
#         co_go.grid(row=3,column=0)
#         co_go.config(command=co_verify)
# =============================================================================
        pass
    
    if c2.get() and c15.get():
        chrome_browser.close()
        messagebox.showwarning("Error", "Error - You can not pull both an individual and joint credit report")
        
    if c6.get() and c7.get():
        chrome_browser.close()
        messagebox.showwarning("Error", "Error - You can not order both :/")
        
        
    if c1.get():
        
        chrome_browser.get('https://sanctionssearch.ofac.treas.gov/default.aspx')
        chrome_browser.find_element_by_id("ctl00_MainContent_txtLastName").send_keys(Entry3.get() +" " + Entry4.get())
    
        chrome_browser.find_element_by_id("ctl00_MainContent_txtAddress").send_keys(Entry8.get() + " " + Entry9.get())
    
        chrome_browser.find_element_by_id("ctl00_MainContent_txtCity").send_keys(Entry5.get())
        
        
        
        chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.CONTROL, 'a')
        chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.BACKSPACE)
        chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys('93')
    
        chrome_browser.find_element_by_id("ctl00_MainContent_btnSearch").click()
        
        try:
            myScreenshot = pyautogui.screenshot()
            myScreenshot.save(fr"{folder_path.get()}/OFAC for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png")
        except Exception:
            chrome_browser.close()
            messagebox.showwarning("Error", "You must specify a folder destination")
        
        try:
           
            chrome_browser.find_element_by_xpath('//*[@id="btnDetails"]').click()
            print("OFAC Hit!")
            OFAC_hit_screen = pyautogui.screenshot()
            OFAC_hit_screen.save(fr"{folder_path.get()}/OFAC HIT for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png")
            
            outlook1 = win32.Dispatch('outlook.application')
            mail1 = outlook1.CreateItem(0)
            mail1.To = 'rtlaceste@gmail.com'
            mail1.Recipients.Add("troyboyoy@gmail.com")
            mail1.CC = "rtlaceste@ucdavis.edu"
            mail1.Subject = (f"URGENT - OFAC HIT for Loan #{Entry6.get()} for Applicant(s): {Entry3.get()} {Entry4.get()}, {Entry16.get()} {Entry17.get()}  ")
            mail1.Body = (f'**This is an automated message** \n There has been an OFAC hit regarding applicant(s): {Entry3.get()} {Entry4.get()}, {Entry6.get()} {Entry7.get()} \n Loan number: {Entry6.get()} ')
            attachment = fr"{folder_path.get()}/OFAC HIT for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png"
            mail1.Attachments.Add(attachment)             
            mail1.Send()
            
            
            messagebox.showwarning("Warning", "URGENT - OFAC HIT. CONTACT MANAGER")
        
        
        except NoSuchElementException:
            pass
            
            
        
        if len(Entry16.get()) > 0:
            chrome_browser.get('https://sanctionssearch.ofac.treas.gov/default.aspx')
                    
            chrome_browser.find_element_by_id("ctl00_MainContent_txtLastName").send_keys(Entry16.get() +" "+ Entry17.get())
    
            chrome_browser.find_element_by_id("ctl00_MainContent_txtAddress").send_keys(Entry21.get() + " " + Entry22.get())
    
            chrome_browser.find_element_by_id("ctl00_MainContent_txtCity").send_keys(Entry18.get())
        
        
            chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.CONTROL, 'a')
            chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.BACKSPACE)
            chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys('93')
            
            chrome_browser.find_element_by_id("ctl00_MainContent_btnSearch").click()
            
            myScreenshot1 = pyautogui.screenshot()
            myScreenshot1.save(fr"{folder_path.get()}/OFAC for {Entry16.get()} {Entry17.get()} Loan #{Entry6.get()}.png")
            
            
            try:
                
                chrome_browser.find_element_by_xpath('//*[@id="btnDetails"]').click()
                
                
                myScreenshot1 = pyautogui.screenshot()
                myScreenshot1.save(fr"{folder_path.get()}/OFAC HIT for {Entry16.get()} {Entry17.get()} Loan #{Entry6.get()}.png")
                
                outlook1 = win32.Dispatch('outlook.application')
                mail1 = outlook.CreateItem(0)
                mail1.To = 'rtlaceste@gmail.com'
                mail1.Recipients.Add("troyboyoy@gmail.com")
                mail1.CC = "rtlaceste@ucdavis.edu"
                mail1.Subject = (f"Update on Loan #{Entry6.get()} for Applicant(s): {Entry3.get()} {Entry4.get()}, {Entry16.get()} {Entry17.get()}  ")
                mail1.Body = (f'** This is an automated message** \n \n Information regarding what was pulled on Loan #{Entry6.get()} through the script: \n\n OFAC: {status["OFAC"]} \n Credit Report (Individual): {status["Individual Credit Report"]} \n Credit Report (Joint): {status["Joint Credit Report"]} \n Google Search: {status["Google Search"]} \n Lexis Nexis: {status["Lexis Nexis"]} \n IRS-TIN: {status["IRS-TIN"]} \n Flood Cert (Simple): {status["Flood Cert - Simple"]} \n Flood Cert (Complex): {status["Flood Cert - Complex"]} \n Environmental Report (ETS): {status["Environmental Report (ETS)"]} \n Environmental Report (EDR): {status["Environmental Report (EDR)"]} \n Environmental Report (Phase 1): {status["Environmental Report (Phase 1)"]} \n Inspection Report: {status["Inspection Report"]} \n Order Appraisal: {status["Order Appraisal"]} \n Legal Ticket: {status["Legal Ticket"]} \n \n User Comments: {email_comments.get()} \n \n Script total run time: {run_time_sec} seconds')
                attachment = fr"{folder_path.get()}/OFAC HIT for {Entry16.get()} {Entry17.get()} Loan #{Entry6.get()}.png"             
                mail1.Send()
                
            
                messagebox.showwarning("Warning", "URGENT - OFAC HIT. CONTACT MANAGER")
        
        
            except NoSuchElementException:
                pass
    
    
        else:
            pass
            
            
        
        status["OFAC"] = "Yes"    
        
        
        
    else:
        pass
    
    
    #Credit Report Pull (Individual)
    if c2.get():
         
        
        
        chrome_browser.get('https://www.credco.com/ecredco/security/login.aspx')
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Username").send_keys(info["CredCo Username"])
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Password").send_keys(info["Credco Password"])
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_LoginButton").click()
        try:
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_ctl00_liOrder").click()
        except Exception:
            messagebox.showwarning("Error", "Please make sure your CredCo information is accurate! ")
    
        #Instant Merge Selection
        try:
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk800").click()
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk800").click()
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk2147483591").click()
        except Exception:
            messagebox.showwarning("Error", "Website HTML updated, please contact Troy.")
        
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLReaccess_txtLoanNumber").send_keys(f"{Entry6.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtLastName").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtFirstName").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtSSN").send_keys(f"{Entry7.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtNum").send_keys(f"{Entry8.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtStreetName").send_keys(f"{Entry9.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtCity").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_ddlState").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtZip").send_keys(f"{Entry11.get()}")
        
        #Click Order
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_btnOrder").click()
        
        
        
        if len(Entry20.get()) > 0:
            print("Go")
            
            chrome_browser.get('https://www.credco.com/ecredco/order/order.aspx')
            
            
            try:
                chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk800").click()
                chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk800").click()
                chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk2147483591").click()
            except Exception:
                messagebox.showwarning("Error", "Website HTML updated, please contact Troy.")
        
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLReaccess_txtLoanNumber").send_keys(f"{Entry6.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtLastName").send_keys(f"{Entry17.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtFirstName").send_keys(f"{Entry16.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtSSN").send_keys(f"{Entry20.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtNum").send_keys(f"{Entry21.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtStreetName").send_keys(f"{Entry22.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtCity").send_keys(f"{Entry18.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_ddlState").send_keys(f"{Entry23.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtZip").send_keys(f"{Entry24.get()}")
            
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_btnOrder").click()
            
            
            
            
        else:
            pass
            
        status["Individual Credit Report"] = "Yes"
    else:
        pass
    if c15.get():
        
        #Joint Credit Report
        
        chrome_browser.get('https://www.credco.com/ecredco/security/login.aspx')
        time.sleep(1)
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Username").send_keys(info["CredCo Username"])
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Password").send_keys(info["Credco Password"])
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_LoginButton").click()
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_ctl00_liOrder").click()
        
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_phM_phB_DHQControl_CRLApplicantType_rdoJoint"]').click()
        
        try:
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk800").click()
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk800").click()
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_chk2147483591").click()
        except Exception:
            messagebox.showwarning("Error", "Website HTML updated, please contact Troy.")
        
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLReaccess_txtLoanNumber").send_keys(f"{Entry6.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtLastName").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtFirstName").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtSSN").send_keys(f"{Entry7.get()}")
        
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLCOApplicantDetails1_Applicant_txtLastName").send_keys(f"{Entry17.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLCOApplicantDetails1_Applicant_txtFirstName").send_keys(f"{Entry16.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLCOApplicantDetails1_Applicant_txtSSN").send_keys(f"{Entry20.get()}")
        
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtNum").send_keys(f"{Entry8.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtStreetName").send_keys(f"{Entry9.get()}")
        
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtCity").send_keys(f"{Entry5.get()}")
        
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_ddlState").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtZip").send_keys(f"{Entry11.get()}")
    
        chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_btnOrder").click()
        
        status["Individual Credit Report"] = "Yes"         

      

              
    else:
        pass
    
    if c3.get():
        
        
        chrome_browser.get("http://www.google.com")
        search = chrome_browser.find_element_by_name('q')
        search.send_keys(f'"{Entry3.get()} {Entry4.get()}" AND "money laundering" OR "fraud" OR "lawsuits"')
        search.send_keys(Keys.RETURN) # hit return after you enter search text
        myScreenshot2 = pyautogui.screenshot()
        myScreenshot2.save(fr"{folder_path.get()}/Google Search 1 for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png")
        chrome_browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        myScreenshot3 = pyautogui.screenshot()
        myScreenshot3.save(fr"{folder_path.get()}/Google Search 2 for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png")
        
        chrome_browser.find_element_by_id('pnnext').click()
        
        status["Google Search"] = "Yes"
        
        
        if len(Entry16.get()) > 0:
             chrome_browser.get("http://www.google.com")
             search = chrome_browser.find_element_by_name('q')
             search.send_keys(f'"{Entry16.get()} {Entry17.get()}" AND "money laundering" OR "fraud" OR "lawsuits"')
             search.send_keys(Keys.RETURN) 
             myScreenshot4 = pyautogui.screenshot()
             myScreenshot4.save(fr"{folder_path.get()}/Google Search 1 for {Entry16.get()} {Entry17.get()} Loan #{Entry6.get()}.png")
             chrome_browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
             time.sleep(1)
             myScreenshot5 = pyautogui.screenshot()
             myScreenshot5.save(fr"{folder_path.get()}/Google Search 2 for {Entry16.get()} {Entry17.get()} Loan #{Entry6.get()}.png")
        
        else:
            pass
        
    else:
        pass
    
    if c4.get():
        
        
        chrome_browser.get("https://amlinsight.lexisnexis.com/")
        chrome_browser.find_element_by_id("LOGINID").send_keys(f"RLaceste558")
        chrome_browser.find_element_by_id("PASSWORD").send_keys(f"Panda544")
        chrome_browser.find_element_by_id("SIGNON").click()
        
        chrome_browser.find_element_by_name("confirm_button").click()
        chrome_browser.find_element_by_xpath('//*[@id="IDX_0"]/a').click()
        time.sleep(2)
        chrome_browser.find_element_by_id("SSN").send_keys(f"{Entry7.get()}")
        chrome_browser.find_element_by_id("LAST_NAME").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("FIRST_NAME").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("STREET_ADDRESS").send_keys(f"{Entry8.get()} {Entry9.get()}")
        chrome_browser.find_element_by_id("CITY").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("ZIP").send_keys(f"{Entry11.get()}")
        
        
        
        status['Lexis Nexis'] = 'Yes'
        #create action chain object 
        #action = ActionChains(chrome_browser) 
  
    # perform the operation 
        #action.move_to_element_with_offset(element, 0, 200).send_keys("test").perform() 
        
     
        
        
    else:
        pass
        
    if c5.get():
        
        chrome_browser.get("https://la1.www4.irs.gov/eauth/pub/login.jsp?Data=VGFyZ2V0TG9BPUY%253D&TYPE=33554433&REALMOID=06-0004b429-9e8b-1a23-9e85-163b0acf4037&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=UOkC7yx4eMTO24FGxPfBRb5q3Mj3Xh3pyXfBEjYyHJ97nGCXu16wx5MzFHjfZmlG&TARGET=-SM-https%3a%2f%2fla1%2ewww4%2eirs%2egov%2fesrv%2ftinm%2fproauth%2efaces")
        status['IRS-TIN'] = 'Yes'
        
        
    else:
        pass
    
    
    if c6.get():
        
        chrome_browser.get("https://lender.floodapp.com/Login.aspx")
        chrome_browser.find_element_by_id("ctl19_txtUserName").send_keys("IPL-JC")
        chrome_browser.find_element_by_id("ctl19_txtPassword").send_keys("Fl00d2020!")
        chrome_browser.find_element_by_id("ctl19_btnLogin").click()
        chrome_browser.get("https://lender.floodapp.com/Content/neworder.aspx")
        
        chrome_browser.find_element_by_id("Content_Main_txtLoanNum").send_keys(f"{Entry6.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtBrwrFstNme").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtBrwrLastNme").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtStreet").send_keys(f"{Entry8.get()} {Entry9.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtCity").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtState").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtZip").send_keys(f"{Entry11.get()}")
        #chrome_browser.find_element_by_id("Content_Main_txtTaxID").send_keys(f"{TIN_number.get()}")
# =============================================================================
#         chrome_browser.find_element_by_id("Content_Main_btnVerifyOrder").click()
#         
#         chrome_browser.find_element_by_id("Content_Main_btnSubmitOrder").click()
#         
#         #View and print cert
#         chrome_browser.find_element_by_id("Content_Main_hlinkViewCert").click()
#         time.sleep(7)
#         
#         chrome_browser.find_element_by_id("Content_Main_imgEditBttn_PrpAddr").click()
#         
#         
#         #Add in code to remove everything from line!
#         chrome_browser.find_element_by_id("Content_Main_txtAddress").send_keys("...")
#         
#         #Submit Corrections
#         chrome_browser.find_element_by_id("Content_Main_btnVerify").click()
#         
#         chrome_browser.find_element_by_id("Content_Main_btnChooseAddress").click()
# =============================================================================
        time.sleep(1)
        
        status['Flood Cert - Simple'] = 'Yes'
        
        #Submit Order
        #chrome_browser.find_element_by_id("Content_Main_btnSubmitOrder").click()
        #Correct Order Here
        
    else:
        pass
    
    
    if variable.get() == 'None':
        pass
    
    elif variable.get() == 'Simple':
        chrome_browser.get("https://lender.floodapp.com/Login.aspx")
        chrome_browser.find_element_by_id("ctl19_txtUserName").send_keys("IPL-JC")
        chrome_browser.find_element_by_id("ctl19_txtPassword").send_keys("Fl00d2020!")
        chrome_browser.find_element_by_id("ctl19_btnLogin").click()
        chrome_browser.get("https://lender.floodapp.com/Content/neworder.aspx")
        
        chrome_browser.find_element_by_id("Content_Main_txtLoanNum").send_keys(f"{Entry6.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtBrwrFstNme").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtBrwrLastNme").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtStreet").send_keys(f"{Entry8.get()} {Entry9.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtCity").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtState").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("Content_Main_txtZip").send_keys(f"{Entry11.get()}")
    
        time.sleep(1)
        
        status['Flood Cert - Simple'] = 'Yes'
        
    else:
        
        chrome_browser.get("https://weborders.floodapp.com/Login ")
        chrome_browser.find_element_by_id("acceptBttn").click()
        chrome_browser.find_element_by_id("UserName").send_keys("BOFI-MCS")
        chrome_browser.find_element_by_id("Password").send_keys("Axos@2020!")
        
        chrome_browser.find_element_by_id("loginSubmit").click()
        chrome_browser.get("https://weborders.floodapp.com/Order/PlaceOrder")
        
        chrome_browser.find_element_by_id("LoanNumber").send_keys(f"{Entry6.get()}")
        chrome_browser.find_element_by_id("BrwrFirstName").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("BrwrLastName").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("StreetAddress").send_keys(f"{Entry8.get()} {Entry9.get()}")
        chrome_browser.find_element_by_id("City").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("borrowerStInput").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("borrowerZipInput").send_keys(f"{Entry11.get()}")
        chrome_browser.find_element_by_id("newOrderVerify").click()
        
        
        status["Flood Cert - Complex"] = 'Yes'
    
    if c7.get():
        
        chrome_browser.get("https://weborders.floodapp.com/Login ")
        chrome_browser.find_element_by_id("acceptBttn").click()
        chrome_browser.find_element_by_id("UserName").send_keys("BOFI-MCS")
        chrome_browser.find_element_by_id("Password").send_keys("Axos@2020!")
        
        chrome_browser.find_element_by_id("loginSubmit").click()
        chrome_browser.get("https://weborders.floodapp.com/Order/PlaceOrder")
        
        chrome_browser.find_element_by_id("LoanNumber").send_keys(f"{Entry6.get()}")
        chrome_browser.find_element_by_id("BrwrFirstName").send_keys(f"{Entry3.get()}")
        chrome_browser.find_element_by_id("BrwrLastName").send_keys(f"{Entry4.get()}")
        chrome_browser.find_element_by_id("StreetAddress").send_keys(f"{Entry8.get()} {Entry9.get()}")
        chrome_browser.find_element_by_id("City").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("borrowerStInput").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("borrowerZipInput").send_keys(f"{Entry11.get()}")
        chrome_browser.find_element_by_id("newOrderVerify").click()
        
        
        status["Flood Cert - Complex"] = 'Yes'
        
    if c8.get():
        
        chrome_browser.get("https://www.sitelynx.net/admin/login ")
        time.sleep(1.5)
        chrome_browser.find_element_by_id("name").send_keys(info["Sitelynx Username"])
        chrome_browser.find_element_by_id("password").send_keys(info["Sitelynx Password"])
        chrome_browser.find_element_by_xpath("/html/body/div/div/form/fieldset/div[4]/button").click()
        
        chrome_browser.get("https://www.sitelynx.net/projects/new")
        chrome_browser.find_element_by_id("client_address").send_keys(f"4350 La Jolla Village Drive")
        chrome_browser.find_element_by_id("client_city").send_keys(f"San Diego")
        
        chrome_browser.find_element_by_id("client_state").send_keys(f"CA")
        chrome_browser.find_element_by_id("client_zip").send_keys(f"92122")
    
        chrome_browser.find_element_by_id("continue-1").click()
        
        chrome_browser.find_element_by_id("property_name").send_keys(f"Placeholder LLC")
        chrome_browser.find_element_by_id("address").send_keys(f"{Entry8.get()} {Entry9.get()}")
        chrome_browser.find_element_by_id("city").send_keys(f"{Entry5.get()}")
        chrome_browser.find_element_by_id("state").send_keys(f"{Entry10.get()}")
        chrome_browser.find_element_by_id("zip").send_keys(f"{Entry11.get()}")
        
        chrome_browser.find_element_by_id("continue-2").click()
        chrome_browser.find_element_by_id("client_number").click()
# =============================================================================
#         chrome_browser.find_element_by_id("continue-3").click()
# =============================================================================
        
        status["Environmental Report (ETS)"] = 'Yes'
        
        
    if c9.get():
        #Phase 1
        
        status["Environmental Report (Phase 1)"] = 'Yes'
        
        pass
    
    
    if c10.get():
        
        chrome_browser.get("https://axos.exactbid.com/Account/Login?msg=m")
        chrome_browser.find_element_by_id("LoginModel_UserName").send_keys("Rlaceste")
        chrome_browser.find_element_by_id("LoginModel_Password").send_keys("Panda544!")
        chrome_browser.find_element_by_id("LoginButton").click()
        chrome_browser.get("https://axos.exactbid.com/Project/NewServiceRequest")
        iframe = chrome_browser.find_element_by_tag_name('iframe')
        chrome_browser.switch_to.frame(iframe)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_txtFirstName").send_keys(f"{Entry12.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_txtLastName").send_keys(f"{Entry13.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_lblSearch").click()
        time.sleep(1)
        
        try:
            #chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_repResults_ctl01_ReturnLink").click()
            chrome_browser.find_element_by_xpath(f"//a[contains(text(),'{Entry12.get()} {Entry13.get()}')]").click()
        except Exception:
            pass
        time.sleep(1.5)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalSrvType_rptGroups_ctl00_rptServices_ctl04_chkServiceName").click()
        chrome_browser.find_element_by_xpath('/html/body/form/div[3]/div/table/tfoot/tr/td/button/span[1]').click()
       
        
        
         
        
# =============================================================================
#         # Last Name / Entity
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPBRLastName").send_keys(f"{Entry4.get()}")
        
#         #Drop down menu (Prevous Report)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_PrevRpt")
#         
#         #Drop Down Menu (Is this an existing Axos Bank Loan)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanType")
#         
#         # Loan Amount
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanAmount").send_keys("0")
#         
#         # Loan Purpose (drop down)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID")
#         
#         #Loan Number
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanNumber").send_keys(f"{Entry6.get()}")
#         
#         #Loan Property Type dropdown
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_MajorTypeID")
        
        time.sleep(15)
        
#         #Save and continue
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblContinue").click()
#       
        time.sleep(8)
        
#         #Property Type Drop-down
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_drpPropertyMajorTypeID")
#         
#         #status drop down
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_drpStatus")
#         
#         #tenancy drop down
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_drpTenancy")
#         
#         #Property type-type
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_drpPropertyTypeID")
#         
#         #Update
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_btnUpdatePropertyData")
#         
        time.sleep(2)
#         
#         
#         
#         #Property Contact - affiliation
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPPC1Affiliation")
#         #Property Contact - Last Name
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPPC1LastName")
#         #Property Contact - Phone number
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPPC1Phone")
#         
#         #Pending/Recent Sale Drop down
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_IsPendingSale")
#         
#         #Send Selected Service
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblSendService").click()
#         
#         #Save Details for all
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblSaveAllDetails").click()
#         
#         #Send Selected Services
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblSendService").click()
#         
#         time.sleep(10)
#         # Advanced Button
        chrome_browser.find_element_by_xpath("/html/body/div[2]/header/div[2]/div[1]/form/div[1]/div/a/span").click()
#         
#         #Search by Property Filter (Street Number)
        chrome_browser.find_element_by_id("ProjectSearchCriteria_Address_StreetNumber").send_keys("...")
#         #Search by Property Filter (Street Name)
        chrome_browser.find_element_by_id("ProjectSearchCriteria_Address_StreetName")
#          #Search by Property Filter (City)
        chrome_browser.find_element_by_id("ProjectSearchCriteria_Address_City")
#         
#         #Click Element 
        chrome_browser.find_element_by_xpath("/html/body/div[2]/div/div/div/div[2]/div[1]/div/div[4]/table/tbody/tr[2]/td[2]/a").click()
#         
#         #Expand All Tasks
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[1]/a[2]').click()
#         time.sleep(2)
        chrome_browser.find_element_by_id("edit-task").click()
#         
#         #Uncheck Tentative
        chrome_browser.find_element_by_id("IsTypeRequired").click()
#         
#         #Due date
        chrome_browser.find_element_by_xpath("/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/form/div[2]/div/div[2]/div/div[2]/span[1]/span/input").send_keys("testing...")
#         
#         
        chrome_browser.find_element_by_id('pjt-task-save-button').click()
#         
#         
#         #Click RFP Manager
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/div[1]/a[5]').click()
#         
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/div[1]/a[5]')
#         
#         #Click Direct Award
        chrome_browser.find_element_by_id("direct-award").click()
#         
#         #Add fee
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_txtBidFee").send_keys("200")
#         
#         #Select/Edit Recipients
        chrome_browser.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/a/b').click()
#         
#         #Deselect Show Pre-selected Vendors
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_chkShowPreselected').click()
#         
#         #Deselect Certified Vendors Only
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_chkCertifiedOnly').click()
#         
#         #Search
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_Button1').click()
#         
#         #Click "Include Vendor Information"
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_chkIncludeVendorInfo').click()
#         
#         #Input Last Name
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_searchLastName').send_keys('Johnson')
#         
#         #Input First Name
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_searchLastName').send_keys('Kimberly')
#         
#         #Search
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_Button1').click()
#         
#         #Click K Johnson
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_grdRFPRecipients_ctl02_chkRecipient').click()
#         
#         #Click Select
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/button[1]/span').click()
#         
#         #Direct Award
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_cmdDirectAward').click()
#         
#         #Press ok
        chrome_browser.find_element_by_xpath('/html/body/div[1]/div[3]/div/button/span').click()
#         
#         #close window
        chrome_browser.find_element_by_xpath('/html/body/div[39]/div[1]/div/a[8]/span').click()
# =============================================================================
        
        status["Inspection Report"] = "Yes"
        
        

        
    if c11.get():
        chrome_browser.get("https://theflynngroup.exactbid.com/default.asp?ini=2&sce=1")
        chrome_browser.find_element_by_name("loginid").send_keys(f"AxosAO")
        
        chrome_browser.find_element_by_name("password").send_keys(f"#Bofi2018")
        chrome_browser.find_element_by_name("Action").click()
        
        
        chrome_browser.get("https://theflynngroup.exactbid.com/AccountOfficers/AODesktop/Default.aspx?Dest=NR")
        
        time.sleep(.75)
        iframe2 = chrome_browser.find_element_by_tag_name('iframe')
        chrome_browser.switch_to.frame(iframe2)
        
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalSrvType_tcblSrvTypes_1").click()
        chrome_browser.find_element_by_name("ctl00$ctl00$contentBody$contentBody$modalSrvType$btnNext").click()
        
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_SRFPBRLastName"]').send_keys(f"{Entry4.get()}")
        
        time.sleep(.5)
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_LoanAmount"]').send_keys(f"0")
        
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_LoanNumber"]').send_keys(Keys.CONTROL, 'a')
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_LoanNumber"]').send_keys(Keys.CONTROL, 'a')
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_LoanNumber"]').send_keys(Keys.BACKSPACE)
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_LoanNumber"]').send_keys(f'{Entry6.get()}')
        
        chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_ProposedLTV"]').send_keys("0")
        
# =============================================================================
#         chrome_browser.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_contentBody_btnContinue"]').click()
# =============================================================================
        
        status["Order Appraisal"] = "Yes"
        
        
    if c12.get():
        try:
            chrome_browser.get("https://bofi.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D2e76be11db5d76405661b96c4e961985%26sysparm_link_parent%3D59de44e4db19f2405661b96c4e961965%26sysparm_catalog%3D3910bd12df132100dca6a5")
            iframe = chrome_browser.find_element_by_xpath("//iframe[@name='gsft_main']")
            chrome_browser.switch_to.frame(iframe)
            chrome_browser.find_element_by_id("sys_display.IO:a63c2c68db59f2405661b96c4e961920").send_keys(f"{Entry27.get()}")
            chrome_browser.find_element_by_id("sys_display.IO:d2f0fd01dbf5ba805661b96c4e96195b").send_keys(f"Tom Constantine")
            chrome_browser.find_element_by_id("IO:7f421c28db19f2405661b96c4e9619e1").send_keys(f"{Entry6.get()}")
            chrome_browser.find_element_by_id("sys_display.IO:46655ce8db19f2405661b96c4e961939").send_keys(f"IPL Intake and Processing")
            chrome_browser.find_element_by_id("IO:e99354f2db557a405661b96c4e961931").send_keys(f"O")
            chrome_browser.find_element_by_id("IO:9a7bd46cdb19f2405661b96c4e9619e0").send_keys(f"R")
            chrome_browser.find_element_by_id("IO:aca1e4e0db59f2405661b96c4e961943").send_keys(f"{Entry3.get()} {Entry4.get()} #{Entry6.get()}")
            
            chrome_browser.switch_to.default_content
        
        
            status["Legal Ticket"] = "Yes"
        except Exception:
            pass
        
    else:
        pass
    
    
        
        
#        pyautogui.moveTo(3000, 500)  # moves mouse to X of 600, Y of 500.
#        chrome_browser.get("https://bofi.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D2e76be11db5d76405661b96c4e961985%26sysparm_link_parent%3D59de44e4db19f2405661b96c4e961965%26sysparm_catalog%3D3910bd12df132100dca6a5")
#
#        shell = win32com.client.Dispatch("WScript.Shell")
#        shell.SendKeys("{TAB}") #Press tab... to change focus or whatever
#        #   pyautogui.press("tab").send_keys("aa")
#        shell.SendKeys("Tom Constantine")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys(f"{Entry6.get()}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("O")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("IPL Intake and Processing")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("R")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys("{TAB}")
#        shell.SendKeys(f"{Entry3.get()} {Entry4.get()} #{Entry6.get()}")
#        
#        time.sleep(2)
        
        
        
    if c13.get():
        
        # EDR
        
        time.sleep(1)
        chrome_browser.get("https://axos.exactbid.com/Home/Dashboard")
        chrome_browser.find_element_by_id("LoginModel_UserName").send_keys("Rlaceste")
        chrome_browser.find_element_by_id("LoginModel_Password").send_keys("Panda544!")
        chrome_browser.find_element_by_id("LoginButton").click()
    
    if c19.get():
        #ACM
        chrome_browser.get('https://www.sitelynx.net/admin/login ')
        pass
        
    if c14.get():
        
        
        #SALESFORCE

        today = datetime.today()

        tod = datetime.strftime(today, "%Y/%m/%d")

        mod = today + timedelta(days=7)

        mod1 = today + timedelta(days=14)

        fut = datetime.strftime(mod, "%Y/%m/%d")
        
        Fut_14 = datetime.strftime(mod1, "%Y/%m/%d")



        today_month = tod[5:7]

        today_day = tod[8:]

        today_year = tod[:4]

        fut_7_days = fut[8:]

        fut_7_days_month = fut[5:7]

        fut_7_days_year = fut[:4]

        fut_14_days = Fut_14[8:]

        fut_14_days_month = Fut_14[5:7]

        fut_14_days_year = Fut_14[:4]

        
        
        
        chrome_browser.get("https://bofi.my.salesforce.com/?ec=302&startURL=%2Fvisualforce%2Fsession%3Furl%3Dhttps%253A%252F%252Fbofi.lightning.force.com%252Flightning%252Fr%252FReport%252F00O3o0000055toBEAQ%252Fview%253FqueryScope%253DuserFolders")    
    
        chrome_browser.find_element_by_id("username").send_keys("rlaceste@axosbank.com")
        chrome_browser.find_element_by_id("password").send_keys("Troyboyoy544!")
        chrome_browser.find_element_by_id("Login").click()
        chrome_browser.get("https://bofi.lightning.force.com/lightning/r/Report/00O3o0000055toBEAQ/view")
        
        
        
# =============================================================================
        time.sleep(2)
# =============================================================================
        loan_search = chrome_browser.find_element_by_id("169:0;p")
        loan_search.send_keys(f"90099010409")
        loan_search.send_keys(Keys.RETURN)
# =============================================================================

        time.sleep(2.75)
# =============================================================================
#         iframe = chrome_browser.find_element_by_tag_name("iframe")    
#         chrome_browser.switch_to.frame(iframe)
# =============================================================================
        
# =============================================================================
# =============================================================================
        
        
        chrome_browser.find_element_by_xpath("//a[contains(text(),'90099010409')]").click()
        time.sleep(6)
        
# =============================================================================
# =============================================================================
#         iframe = chrome_browser.find_element_by_tag_name("iframe")    
#         chrome_browser.switch_to.frame(iframe)
# =============================================================================
#         time.sleep(10)
# =============================================================================

#chrome_browser.find_element_by_xpath("//span[@ng-bind='$ctrl.item.displayName']").click()
        
        time.sleep(3)
        iframe = chrome_browser.find_element_by_xpath("//iframe[@title='accessibility title']")
        chrome_browser.switch_to.frame(iframe)
        time.sleep(2.5)
        chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[6]/a').click()
        #chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[2]/span/span[2]/div/table/tbody/tr[1]/td[4]/button[2]').click()
        time.sleep(5)
        
        # ADD JCA
        
# =============================================================================
#         chrome_browser.find_element_by_xpath("//button[@data-bind = 'click: getViewModel().create' ]").click()
#         time.sleep(3)
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[18]/div/input').send_keys(f'{Entry14.get()} {Entry15.get()}')
#         #chrome_browser.switch_to_default_content
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath(f"//td[contains(text(),'{Entry14.get()} {Entry15.get()}')]").click()
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys("IPL")
#         time.sleep(1.5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys(Keys.SPACE)
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//td[contains(text(),'IPL JCA')]").click()
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[contains(text(),'Save')]").click()
#         
# =============================================================================
# =============================================================================
#         # ADD LOAN OFFICER
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[@data-bind = 'click: getViewModel().create' ]").click()
#         time.sleep(3)
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[18]/div/input').send_keys(f'{Entry12.get()} {Entry13.get()}')
#        
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath(f"//td[contains(text(),'{Entry12.get()} {Entry13.get()}')]").click()
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys("IPL")
#         time.sleep(2.5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys(Keys.SPACE)
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/div/div/table/tbody[2]/tr[7]/td').click()
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[contains(text(),'Save')]").click()
#         
#         
#         # ADD Processor
#         
#                 
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[@data-bind = 'click: getViewModel().create' ]").click()
#         time.sleep(3)
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[18]/div/input').send_keys(f'{L_P_FN.get()} {L_P_LN.get()}')
#        
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath(f"//td[contains(text(),'{L_P_FN.get()} {L_P_LN.get()}')]").click()
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys("IPL")
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys(Keys.SPACE)
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/div/div/table/tbody[2]/tr[9]/td").click()
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[contains(text(),'Save')]").click()
#         
#         # ADD CLOSER
#         
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[@data-bind = 'click: getViewModel().create' ]").click()
#         time.sleep(3)
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[18]/div/input').send_keys(f'{closer_fn.get()} {closer_ln.get()}')
#        
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath(f"//td[contains(text(),'{closer_fn.get()} {closer_ln.get()}')]").click()
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys("IPL")
#         time.sleep(1)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[18]/div/input").send_keys(Keys.SPACE)
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/div/div/table/tbody[2]/tr[4]/td").click()
#         time.sleep(2)
#         chrome_browser.find_element_by_xpath("//button[contains(text(),'Save')]").click()
# =============================================================================
        
        time.sleep(1)
        
        chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[9]/a/span').click()
        time.sleep(5.5)
        
        ################### Add Flood Cert #######################
        if c6.get() or c7.get():
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[1]/div/button').click()
            time.sleep(3)
            select = Select(chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[16]/select'))
            select.select_by_index(15)
            

            
            
            
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[8]/div/div[2]/span[2]/span[6]/div/input').send_keys("")
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[8]/div/div[2]/span[2]/span[6]/div/input').send_keys(f"{today_month}/{today_day}/{today_year}")
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[2]/div/div[2]/span[2]/span[6]/div/input').send_keys(f"{fut_7_days_month}/{fut_7_days}/{fut_7_days_year}")
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[19]/textarea').send_keys("Flood Zone X")
        
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[4]/div/button[1]').click()
        
        # ADD Appraisal
        
        
        if c11.get():
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[1]/div/button').click()
            time.sleep(5)
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[19]/textarea').send_keys("Appraisal")
            
            
            select = Select(chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[16]/select'))
            select.select_by_index(2)        
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[8]/div/div[2]/span[2]/span[6]/div/input').send_keys(f'{today_month}/{today_day}/{today_year}')
                
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[4]/div/button[1]').click()
            
            
         # Property Inspection   #Due 14 Days
        if c10.get():
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[1]/div/button').click()
            time.sleep(4)
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[19]/textarea').send_keys('Inspection')
            
            select = Select(chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[16]/select'))
            select.select_by_index(25)
            
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[8]/div/div[2]/span[2]/span[6]/div/input').send_keys(f'{today_month}/{today_day}/{today_year}')
    
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[2]/div/div[2]/span[2]/span[6]/div/input').send_keys(f'{fut_14_days_month}/{fut_14_days}/{fut_14_days_year}')
    
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[4]/div/button[1]').click()
            
        # Environmental Report
        
        
        
        
        
        
        
        #Legal Ticket
        if c12.get():
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[1]/div/button').click()
            time.sleep(4)
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[19]/textarea').send_keys(f'Legal ticket submitted {today}')
            
            select = Select(chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[16]/select'))
            select.select_by_index(18)
             
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[2]/div/div[2]/span[2]/span[6]/div/input').send_keys(f'{today_month}/{today_day}/{today_year}')
            time.sleep(1)
            chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[4]/div/button[1]').click()
            
          
            
            #Log Email
            
        chrome_browser.switch_to_default_content()
        time.sleep(1.5)
        chrome_browser.find_element_by_xpath('/html/body/div[4]/div[1]/section/div[1]/div[2]/div[2]/div[1]/div/div/div/div[2]/div/one-record-home-flexipage2/forcegenerated-adgrollup_component___forcegenerated__flexipage_recordpage___commercial_loan_page___llc_bi__loan__c___view/forcegenerated-flexipage_commercial_loan_page_llc_bi__loan__c__view_js/record_flexipage-record-page-decorator/div[1]/slot/flexipage-record-home-single-col-template-desktop2/div/div[2]/div/slot/slot/flexipage-component2/slot/flexipage-tabset2/div/lightning-tabset/div/lightning-tab-bar/ul/li[5]/a').click()
        time.sleep(2.5)
        chrome_browser.find_element_by_xpath("//buttin[@class='slds-button slds-button--brand testid__dummy-button-submit-action slds-col slds-no-space dummyButtonSubmitAction uiButton']").click()

        #chrome_browser.find_element_by_xpath('/html/body/div[4]/div[1]/section/div[1]/div[2]/div[2]/div[1]/div/div/div/div[2]/div/one-record-home-flexipage2/forcegenerated-adgrollup_component___forcegenerated__flexipage_recordpage___commercial_loan_page___llc_bi__loan__c___view/forcegenerated-flexipage_commercial_loan_page_llc_bi__loan__c__view_js/record_flexipage-record-page-decorator/div[1]/slot/flexipage-record-home-single-col-template-desktop2/div/div[2]/div/slot/slot/flexipage-component2/slot/flexipage-tabset2/div/lightning-tabset/div/slot/slot/slot/flexipage-tab2[5]/slot/flexipage-component2/slot/flexipage-aura-wrapper/div/div/div[1]/section/div/div[1]/button[2]/span').click()
        
        #chrome_browser.find_element_by_id('input-212').send_keys('Email')
        
        
        
# =============================================================================
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[1]/div/button').click()
#         time.sleep(5)
#         chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[3]/div/div[2]/span[2]/span[19]/textarea').send_keys("beep")
#         select = Select(chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[16]/select'))
#         select.select_by_index(2)
#         
# =============================================================================
        
        
        
        
        
        
        
# =============================================================================
#         # Go to Loan Details
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[4]/a").click()
#         time.sleep(6)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[1]/div/span/div/nc-tertiary-navigation/div/nc-tertiary-navigation-item[3]/a/span").click()
#         time.sleep(5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[1]/div/span/div/nc-tertiary-navigation/div/nc-tertiary-navigation-item[4]/a/span").click()
#     
#         time.sleep(5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[1]/div/span/div/nc-tertiary-navigation/div/nc-tertiary-navigation-item[5]/a/span").click()
#         time.sleep(5.5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[10]/a/span").click()
#         time.sleep(3.5)
#         chrome_browser.switch_to_default_content
#         time.sleep(2)
#         #collateral
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[7]/a/span").click()
#         time.sleep(3.5)
#         #3rd party reports
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[9]/a/span").click()
#         time.sleep(6.5)
#         # Fees
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[10]/a/span").click()
#         time.sleep(4.5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[10]/a/span").click()
#         time.sleep(5)
#         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[1]/span/div/nc-secondary-navigation/div/div/div[2]/nc-secondary-navigation-item[11]/a/span").click()
#        # time.sleep(3.5)
# # =============================================================================
# #         frame3 = chrome_browser.find_element_by_xpath("//iframe[@title='accessibility title']")
# #         chrome_browser.switch_to.frame(frame3)
# #         chrome_browser.find_element_by_xpath("/html/body/div[3]/div/div[3]/div/div[3]/div[1]/div/span/div/nc-tertiary-navigation/div/nc-tertiary-navigation-item[3]/a/span").click()
# #         
# # =============================================================================
#         messagebox.showinfo("Message", "Hi, please fill out the compliance survey")
# =============================================================================
# =============================================================================
#         time.sleep(1)
#         chrome_browser.get("https://bofi.my.salesforce.com/?ec=302&startURL=%2Fvisualforce%2Fsession%3Furl%3Dhttps%253A%252F%252Fbofi.lightning.force.com%252Flightning%252Fr%252FReport%252F00O3o0000055toBEAQ%252Fview%253FqueryScope%253DuserFolders")
#         chrome_browser.find_element_by_id("username").send_keys("rlaceste@axosbank.com")
#         chrome_browser.find_element_by_id("password").send_keys("Troyboyoy544!")
#         chrome_browser.find_element_by_id("Login").click()
#         chrome_browser.get("https://bofi.lightning.force.com/lightning/r/Report/00O3o0000055toBEAQ/view")
#         
#         
#         
# # =============================================================================
#         time.sleep(2)
#         loan_search = chrome_browser.find_element_by_id("169:0;p")
#         loan_search.send_keys(f"90099029279")
#         loan_search.send_keys(Keys.RETURN)
# #         pyautogui.moveTo(500, 350)
# #         time.sleep(3)
# #         pyautogui.click()
# #         pyautogui.moveTo(125, 630)
# # # =============================================================================
# #         time.sleep(8)
# # =============================================================================
# #         pyautogui.moveTo(120, 630)
# #         time.sleep(1.5)
# #         pyautogui.click()
# # =============================================================================
# #         time.sleep(2)
# # =============================================================================
# # =============================================================================
# # =============================================================================
# # =============================================================================
#         
#         time.sleep(3)
# # =============================================================================
# #         iframe = chrome_browser.find_element_by_tag_name("iframe")    
# #         chrome_browser.switch_to.frame(iframe)
# # =============================================================================
#         
# # =============================================================================
# # =============================================================================
#         
#         
#         chrome_browser.find_element_by_xpath("//a[contains(text(),'90099029279')]").click()
#         time.sleep(9)
#         
# # =============================================================================
# # =============================================================================
# #         iframe = chrome_browser.find_element_by_tag_name("iframe")    
# #         chrome_browser.switch_to.frame(iframe)
# # =============================================================================
# #         time.sleep(10)
# # =============================================================================
#         
#         chrome_browser.find_element_by_xpath("//a[contains(text(),'DocMan')]").click()
#         
#         iframe = chrome_browser.find_element_by_xpath("//iframe[@title='Salesforce - Performance Edition']")
#         chrome_browser.switch_to.frame(iframe)
#         
#         chrome_browser.find_element_by_xpath("//div[@data-target'#a5S3o000001I1jvEAC']")
# =============================================================================
        
        #time.sleep(10)
        
    
        #chrome_browser.find_element_by_xpath("//img[contains(text(),'/resource/1599091587000/nforce__SLDS0102/assets/icons/utility/chevrondown_60.png')]")
        
        
        
        
        
        
        #chrome_browser.find_element_by_xpath("//a[contains(text(),'slds-truncate pull-left ng-binding')]")
# =============================================================================
#         chrome_browser.switchTo().defaultContent()
# =============================================================================
  #      chrome_browser.find_element_by_xpath(f'//*[@id="tertiary-navigation"]/nc-tertiary-navigation-item[3]/a/span').click()
        
# =============================================================================
#         iframe1 = chrome_browser.find_element_by_tag_name("iframe")    
#         chrome_browser.switch_to.frame(iframe1)
#         
# =============================================================================

        
# =============================================================================
  #      chrome_browser.get("https://bofi.my.salesforce.com/home/home.jsp?source=lex")
  #      chrome_browser.find_element_by_id("phSearchInput").send_keys("90099029534")
 #       chrome_browser.find_element_by_id("phSearchButton").click()
 #       time.sleep(5)
#        chrome_browser.find_element_by_xpath(f"//a[contains(text(),'90099029534')]").click()
# =============================================================================
        
# =============================================================================
#         loan_search = chrome_browser.find_element_by_id("169:0;p")
#         loan_search.send_keys(f"{Entry6.get()}")
#         loan_search.send_keys(Keys.RETURN)
#         
# =============================================================================
        

        
        pass
    
    
    if c16.get():
        # Email Confirmation
        
        import win32com.client as comclt
        wsh= comclt.Dispatch("WScript.Shell")
         # send the keys you want
        #Zip Code Lookup
        chrome_browser.get("http://reportingportal.prod.axosbank.com/#/report/1078")
                           
        time.sleep(1)
        wsh.SendKeys("RLaceste")
        wsh.SendKeys("{TAB}")
        wsh.SendKeys("LakeGirlBirdDoor")
        wsh.SendKeys("{ENTER}", 0)
        time.sleep(4)
        chrome_browser.find_element_by_id("search-input").send_keys("11103")
        chrome_browser.find_element_by_id("search-btn").click()
        time.sleep(2)
        
        myScreenshot3 = pyautogui.screenshot()
        myScreenshot3.save(fr"{folder_path.get()}/Zip Code Lookup for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png")
        
        time.sleep(1)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'RLaceste@axosbank.com'
        #mail.Recipients.Add("troyboyoy@gmail.com")
        #mail.CC = "rtlaceste@ucdavis.edu"
        #mail.Subject = (f"Update on Loan #{Entry6.get()} for Applicant(s): {Entry3.get()} {Entry4.get()}, {Entry16.get()} {Entry17.get()}  ")
        mail.Subject = (f"Update on Loan #{Entry6.get()} for Applicant(s): {Entry3.get()} {Entry4.get()}, {Entry16.get()} {Entry17.get()}  ")
        mail.Body = (f'** This is an automated message** \n \n Hello team, \n\n All third party reports have been ordered. We will be ordering the appraisal internally/externally \n Missing Items: (items here) \n Loan Officer: {Entry12.get()} {Entry13.get()} \n Loan Processor: \n JCA: {Entry14.get()} {Entry15.get()} \n Loan Closer: \n Flood Zone: X   \n\n OFAC: {status["OFAC"]} \n Credit Report (Individual): {status["Individual Credit Report"]} \n Credit Report (Joint): {status["Joint Credit Report"]} \n Google Search: {status["Google Search"]} \n Lexis Nexis: {status["Lexis Nexis"]} \n IRS-TIN: {status["IRS-TIN"]} \n Flood Cert (Simple): {status["Flood Cert - Simple"]} \n Flood Cert (Complex): {status["Flood Cert - Complex"]} \n Environmental Report (ETS): {status["Environmental Report (ETS)"]} \n Environmental Report (EDR): {status["Environmental Report (EDR)"]} \n Environmental Report (Phase 1): {status["Environmental Report (Phase 1)"]} \n Inspection Report: {status["Inspection Report"]} \n Order Appraisal: {status["Order Appraisal"]} \n Legal Ticket: {status["Legal Ticket"]} \n \n User Comments: {email_comments.get()} \n \n Loan #: {Entry6.get()}')
        attachment = fr"{folder_path.get()}/Zip Code Lookup for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png"
        mail.Attachments.Add(attachment)   
        mail.Send()
        
        
        
    if c18.get():
        
        import win32com.client as comclt
        wsh= comclt.Dispatch("WScript.Shell")
         # send the keys you want
        #Zip Code Lookup
        chrome_browser.get("http://reportingportal.prod.axosbank.com/#/report/1078")
                           
        time.sleep(1)
        wsh.SendKeys("RLaceste")
        wsh.SendKeys("{TAB}")
        wsh.SendKeys("LakeGirlBirdDoor")
        wsh.SendKeys("{ENTER}", 0)
        time.sleep(4)
        chrome_browser.find_element_by_id("search-input").send_keys("11103")
        chrome_browser.find_element_by_id("search-btn").click()
        #chrome_browser.find_element_by_xpath('/html/body/app-root/app-layout/div/mat-sidenav-container/mat-sidenav-content/div/div/a[6]').click()
# =============================================================================
#         outlook = win32.Dispatch('outlook.application')
#         mail = outlook.CreateItem(0)
#         mail.To = 'rtlaceste@gmail.com'
#         mail.Recipients.Add("troyboyoy@gmail.com")
#         mail.CC = "rtlaceste@ucdavis.edu"
#         mail.Subject = (f"Example of a sample email for missing items and confirming reports ")
# =============================================================================
# =============================================================================
#         mail.Body = (f'** This is an automated message** Hello team, \n \n All third party reports have been ordered. We will be ordering the appraisal internally/externally \n \n Missing Items: \n1. \n2. \n3. Loan team: \n \n Loan team: \n Loan Officer: \n Processor: \n Closer: \nJCA:')
#         
# =============================================================================
        
        

    
    
    #os.startfile(fr"{folder_path.get()}")
    #chrome_browser.quit()        
        
        
global Entry3
Entry3 = Entry(root)
Entry3.pack()
Entry3.place(relx=.25,rely=.1)

global Entry4
Entry4 = Entry(root)
Entry4.pack()
Entry4.place(relx=.25,rely=.15)

global Entry5
Entry5 = Entry(root)
Entry5.pack()
Entry5.place(relx=.25,rely=.2)

global Entry6
Entry6 = Entry(root)
Entry6.pack()
Entry6.place(relx=.25,rely=.25)

global Entry7
Entry7 = Entry(root)
Entry7.pack()
Entry7.place(relx=.25,rely=.3)

global Entry8
Entry8 = Entry(root)
Entry8.pack()
Entry8.place(relx=.25,rely=.35)

global Entry9
Entry9 = Entry(root)
Entry9.pack()
Entry9.place(relx=.25,rely=.4)

global Entry10
Entry10=Entry(root)
Entry10.pack()
Entry10.place(relx=.25,rely=.45)

global Entry11
Entry11=Entry(root)
Entry11.pack()
Entry11.place(relx=.25,rely=.5)

global Entry12
Entry12=Entry(root)
Entry12.pack()
Entry12.place(relx=.25,rely=.55)

global Entry13
Entry13=Entry(root)
Entry13.pack()
Entry13.place(relx=.25,rely=.6)

global Entry14
Entry14=Entry(root)
Entry14.pack()
Entry14.place(relx=.25,rely=.65)

global Entry15
Entry15=Entry(root)
Entry15.pack()
Entry15.place(relx=.25,rely=.7)

global Entry16 
Entry16 = Entry(root)
Entry16.pack()
Entry16.place(relx=.4,rely=.1)


global Entry17
Entry17 = Entry(root)
Entry17.pack()
Entry17.place(relx=.4,rely=.15)

global Entry18
Entry18 = Entry(root)
Entry18.pack()
Entry18.place(relx=.4,rely=.2)

global Entry19
Entry19 = Entry(root)
Entry19.pack()
Entry19.place(relx=.4,rely=.25)

global Entry20
Entry20 = Entry(root)
Entry20.pack()
Entry20.place(relx=.4,rely=.3)

global Entry21
Entry21 = Entry(root)
Entry21.pack()
Entry21.place(relx=.4,rely=.35)

global Entry22
Entry22 = Entry(root)
Entry22.pack()
Entry22.place(relx=.4,rely=.4)

global Entry23
Entry23 = Entry(root)
Entry23.pack()
Entry23.place(relx=.4,rely=.45)

global Entry24
Entry24 = Entry(root)
Entry24.pack()
Entry24.place(relx=.4,rely=.5)

# =============================================================================
# global Entry25
# Entry25 = Entry(root)
# Entry25.grid(row=12, column=3, sticky="ew")
# 
# global Entry26
# Entry26 = Entry(root)
# Entry26.grid(row=13, column=3, sticky="ew")
# 
# global Entry27
# Entry27 = Entry(root)
# Entry27.grid(row=14, column=3, sticky="ew")
# 
# global Entry28
# Entry28 = Entry(root)
# Entry28.grid(row=15, column=3, sticky="ew")
# 
# =============================================================================


global email_comments
email_comments = Entry(root)
email_comments.pack()
email_comments.place(relx=.85,rely=.7)

# =============================================================================
# global email_confirm_report
# email_confirm = Entry(root)
# email_confirm.pack()
# =============================================================================

global L_P_FN
L_P_FN = Entry(root)
L_P_FN.pack()
L_P_FN.place(relx=.25,rely=.75)


global L_P_LN
L_P_LN = Entry(root)
L_P_LN.pack()
L_P_LN.place(relx=.25,rely=.8)

global closer_fn
closer_fn = Entry(root)
closer_fn.pack()
closer_fn.place(relx=.25,rely=.85)


global closer_ln
closer_ln = Entry(root)
closer_ln.pack()
closer_ln.place(relx=.25,rely=.9)

button = ttk.Button(root, text='Get Documents')
button.config(command=Get_OFAC)
button2=ttk.Button(root, text='Submit')
button2.config(command=Master_Function)
button2.place(relx=.4,rely=.67)

button3 = ttk.Button(root, text ="Inspection Report")
button3.pack()
button3.place(relx=.4,rely=.74)






OPTIONSi = ["None","Simple","Complex"]
variable = StringVar(root)
variable.set(OPTIONSi[0]) # default value
#print(variable.get()) = yes
w = OptionMenu(root, variable, *OPTIONSi)
w.pack()
w.place(relx=.75,rely=.45)


# ADD CONFIG HERE


# =============================================================================
# Label3.pack()
# Entry3.pack()
# Label4.pack()
# Entry4.pack()
# Label5.pack()
# Entry5.pack()
# Label6.pack()
# Entry6.pack()
# #button.grid(row=20, column=1, sticky="ew")
# Label7.pack()
# Entry7.pack()
# Label8.pack()
# Label9.pack()
# Label10.pack()
# Label11.pack()
# Entry8.pack()
# Entry9.pack()
# Entry10.pack()
# Entry11.pack()
# Label12.pack()
# button2.pack()
# Entry12.pack()
# Entry13.pack()
# Entry14.pack()
# Entry15.pack()
# =============================================================================


Entry3.insert(0, "John")
Entry4.insert(0, "Doe")
Entry5.insert(0, "San Diego")
Entry6.insert(0, "XXXXXXXXXXX")
Entry7.insert(0, "123456789")
Entry8.insert(0, "452")
Entry9.insert(0, "Smith Lane")
Entry10.insert(0, "CA")
Entry11.insert(0, "90210")    
Entry12.insert(0, "Fred")
Entry13.insert(0,"Ornelas")
Entry14.insert(0,"Rainier")
Entry15.insert(0,"Laceste")

L_P_FN.insert(0,"Emily")
L_P_LN.insert(0,"Altoro")

closer_fn.insert(0, "Samantha")
closer_ln.insert(0, "Reeves")

def f_cert_mult():
    win2 = tk.Toplevel()
    win2.wm_title("Flood Cert Multiple Addresses")
    win2.geometry("800x500")
    
    addy = tk.Label(win2, text = "Address 1:", font=('Helvetica',15))
    addy.pack()
    addy.place(relx=.1,rely=.05)
    
    addy_no = tk.Label(win2, text="Address No.")
    addy_no.pack()
    addy_no.place(relx=.1,rely=.15)

button5 = ttk.Button(root, text ="Multiple Addresses?", command = f_cert_mult)
button5.pack()
button5.place(relx=.83,rely=.45)



def SF_fill():
    win1 = tk.Toplevel()
    win1.wm_title("SalesForce Autofill")
    win1.geometry("800x500")
    
    per_1 = tk.Label(win1, text = "%")
    per_1.pack()
    per_1.place(relx=.36,rely=.15)
    
    l = tk.Label(win1, text="Broker Fee")
    l.pack()
    l.place(relx=.1,rely=.15)
    
    global B_fee
    B_fee = Entry(win1)
    B_fee.pack()
    B_fee.place(relx=.2,rely=.15)
    
    l1 = tk.Label(win1, text="Broker YSP")
    l1.pack()
    l1.place(relx=.1,rely=.25)
    
    global B_ysp
    B_ysp = Entry(win1)
    B_ysp.pack()
    B_ysp.place(relx=.2,rely=.25)
    

    per_1 = tk.Label(win1, text = "%")
    per_1.pack()
    per_1.place(relx=.36,rely=.25)

    l2 = tk.Label(win1,text="Lender Fee")
    l2.pack()
    l2.place(relx=.1,rely=.35)
    
    
    global l_fee
    l_fee = Entry(win1)
    l_fee.pack()
    l_fee.place(relx=.2,rely=.35)
    
    per_2 = tk.Label(win1, text = "%")
    per_2.pack()
    per_2.place(relx=.36,rely=.35)
    
    

    b = ttk.Button(win1, text="Quit", command=win1.destroy)
    b.pack()
    b.place(relx=.55,rely=.85)
    
    b1 = ttk.Button(win1, text="Go")
    b1.pack()
    b1.place(relx=.4,rely=.85)
    

    
    
    
##END SALEFORCE AUTOFILL##
button4 = ttk.Button(root, text ="SalesForce",command=SF_fill)
button4.pack()
button4.place(relx=.4,rely=.81)


def popup_bonus():
    win = tk.Toplevel()
    win.wm_title("Inspection Report")
    win.geometry("800x500")

    l = tk.Label(win, text="")
    l.pack()

    b = ttk.Button(win, text="Quit", command=win.destroy)
    b.pack()
    b.place(relx=.4,rely=.85)
    
    
        
# =============================================================================
    lab_warn = Label(win, text = "* Please refresh after selecting inputs",font=('bold'))
    lab_warn.pack()
    lab_warn.place(relx=.05,rely=.005)
    
    
    lab0 = Label(win, text = "Previous Report?")
    lab0.pack()
    lab0.place(relx=.05,rely=.10)
    OPTIONS = ["Yes","No"]
    variable = StringVar(root)
    variable.set(OPTIONS[1]) # default value
    #print(variable.get()) = yes
    w = OptionMenu(win, variable, *OPTIONS)
    w.pack()
    w.place(relx=.35,rely=.1)
#     
#     
    lab1 = Label(win, text = "New or Existing Loan?")
    lab1.pack()
    lab1.place(relx=.05, rely=.18)
    OPTIONS2 = ["New","Existing"]
    variable1 = StringVar(root)
    variable1.set(OPTIONS2[0])
    w1 = OptionMenu(win, variable1, *OPTIONS2)
    w1.pack()
    w1.place(relx=.35, rely=.18)
    
#     

#     

#     

#     
 
#     
    lab_LP = Label(win, text = "Loan Purpose?")
    lab_LP.pack()
    lab_LP.place(relx=.05,rely=.26)
    OPTIONS3 = ["New Loan","Renewal/Subsequent Transaction","Change in collateral", "Foreclosure", "OREO", "Refinance", "Asset Valuation"]
    variable2 = StringVar(root)
    variable2.set(OPTIONS3[0])
    w2 = OptionMenu(win, variable2, *OPTIONS3)
    w2.pack()
    w2.place(relx=.35,rely=.26)
    
    lab_pt = Label(win, text = "Property Type")
    lab_pt.pack()
    lab_pt.place(relx=.05,rely=.34)
    OPTIONS4 = ["Land", "Lodging/Hospitality", "Multi-Family", "Office", "Residential", "Retail-Commercial"]
    variable3 = StringVar(root)
    variable3.set(OPTIONS4[0])
    w3 = OptionMenu(win, variable3, *OPTIONS4)
    w3.pack()
    w3.place(relx=.35,rely=.34)
    
    lab_stat = Label(win, text = "Status")
    lab_stat.pack()
    lab_stat.place(relx=.05,rely=.41)
    OPTIONS7 = ["Existing", "Land Only", "New Addition", "Other (Retired)", "Proposed Construction", "Under Construction", "Under Renovation"]
    variable6 = StringVar(root)
    variable6.set(OPTIONS7[0])
    w5 = OptionMenu(win, variable6, *OPTIONS7)
    w5.pack()
    w5.place(relx=.35,rely=.41)
    
    lab_ten = Label(win, text = "Tenancy")
    lab_ten.pack()
    lab_ten.place(relx=.05,rely=.48)
    OPTIONS8 = ["Vacant", "NA", "Single Tenant Investor", "Multi Tenant Investor", "Owner Occupied 100%", "Owner Occupied > 51%", "Owner Occupied <= 50%"]
    variable7 = StringVar(root)
    variable7.set(OPTIONS8[0])
    w5 = OptionMenu(win, variable7, *OPTIONS8)
    w5.pack()
    w5.place(relx=.35,rely=.48)   


        
    def refresh():
        
        global variable4
        
        variable4 = StringVar(root)
        
        if variable3.get() == 'Land':
            lab_pt_2 = Label(win, text = "Type?")
            lab_pt_2.pack()
            lab_pt_2.place(relx=.5,rely=.34)
            OPTIONS5 = ["Commercial", "Commercial Land", "Multi-Family Land", "Multi-Family-Apartment", "Residential Subdivision Land"]
            variable4 = StringVar(root)
            variable4.set(OPTIONS5[0])
            w4 = OptionMenu(win, variable4, *OPTIONS5)
            w4.pack()
            w4.place(relx=.6,rely=.34)
            
        elif variable3.get() == 'Lodging/Hospitality':
            lab_pt_2 = Label(win, text = "Type?")
            lab_pt_2.pack()
            lab_pt_2.place(relx=.5,rely=.34)
            OPTIONS5 = ["Full Service", "Hotel"]
            variable4 = StringVar(root)
            variable4.set(OPTIONS5[0])
            w4 = OptionMenu(win, variable4, *OPTIONS5)
            w4.pack()
            w4.place(relx=.6,rely=.34)
            
        elif variable3.get() == 'Multi-Family':
            lab_pt_2 = Label(win, text = "Type?")
            lab_pt_2.pack()
            lab_pt_2.place(relx=.5,rely=.34)
            OPTIONS5 = ["Condominium/PUD BLDG(s)", "Garden/Low-Rise Apartments", "Mid/High-Rise", "Mobile/Manufactured Home Park", "Student-Oriented Housing-Student Oriented Apartment"]
            variable4 = StringVar(root)
            variable4.set(OPTIONS5[0])
            w4 = OptionMenu(win, variable4, *OPTIONS5)
            w4.pack()
            w4.place(relx=.6,rely=.34)
            
            
        elif variable3.get() == 'Office':
            lab_pt_2 = Label(win, text = "Type?")
            lab_pt_2.pack()
            lab_pt_2.place(relx=.5,rely=.34)
            OPTIONS5 = ["Condominium Unit(s)", "Creative/Loft", "Medical Office", "Mixed Use-Office-Multi-Family", "Mixed Use-Office-Retail", "Office Building-High Rise", "Office Building-Low Rise", "Office Building-Mid Rise"]
            variable4 = StringVar(root)
            variable4.set(OPTIONS5[0])
            w4 = OptionMenu(win, variable4, *OPTIONS5)
            w4.pack()
            w4.place(relx=.6,rely=.34)
            
            
        elif variable3.get() == 'Residential':
            lab_pt_2 = Label(win, text = "Type?")
            lab_pt_2.pack()
            lab_pt_2.place(relx=.5,rely=.34)
            OPTIONS5 = ["2-4 Units"]
            variable4 = StringVar(root)
            variable4.set(OPTIONS5[0])
            w4 = OptionMenu(win, variable4, *OPTIONS5)
            w4.pack()
            w4.place(relx=.6,rely=.34)
        
        
        elif variable3.get() == 'Retail-Commercial':
            lab_pt_2 = Label(win, text = "Type?")
            lab_pt_2.pack()
            lab_pt_2.place(relx=.5,rely=.34)
            OPTIONS5 = ["Condominium Unit(s)", "Convenience Store", "Free Standing Building-Bank Branch", "Free Standing Building-Free Standing", "Mixed Use-Retail-Office", "Mixed Use-Retail-Residential", "Restaurant-fast food", "Street Retail"]
            variable4 = StringVar(root)
            variable4.set(OPTIONS5[0])
            w4 = OptionMenu(win, variable4, *OPTIONS5)
            w4.pack()
            w4.place(relx=.6,rely=.34)
      
        
    def IR():
        
        options = webdriver.ChromeOptions()
        options.add_experimental_option('prefs', {
                "download.default_directory": r"C:\Users\rtlac\downloads", #Change default directory for downloads
                "download.prompt_for_download": False, #To auto download the file
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
            })
  
    
  
    

        options.add_argument("--disable-notifications")

        chrome_browser = webdriver.Chrome(r'C:\Users\Troy\Desktop\chromedriver.exe', options = options)
        chrome_browser.maximize_window()
    
        
        chrome_browser.get("https://axos.exactbid.com/Account/Login?msg=m")
        chrome_browser.find_element_by_id("LoginModel_UserName").send_keys("Rlaceste")
        chrome_browser.find_element_by_id("LoginModel_Password").send_keys("Panda544!")
        chrome_browser.find_element_by_id("LoginButton").click()
        chrome_browser.get("https://axos.exactbid.com/Project/NewServiceRequest")
        iframe = chrome_browser.find_element_by_tag_name('iframe')
        chrome_browser.switch_to.frame(iframe)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_txtFirstName").send_keys(f"{Entry12.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_txtLastName").send_keys(f"{Entry13.get()}")
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_lblSearch").click()
        time.sleep(1)
        
        try:
            #chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalUsrSearch_repResults_ctl01_ReturnLink").click()
            chrome_browser.find_element_by_xpath(f"//a[contains(text(),'{Entry12.get()} {Entry13.get()}')]").click()
        except Exception:
            pass
        time.sleep(1.5)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_modalSrvType_rptGroups_ctl00_rptServices_ctl04_chkServiceName").click()
        chrome_browser.find_element_by_xpath('/html/body/form/div[3]/div/table/tfoot/tr/td/button/span[1]').click()
       
        
        
# =============================================================================
#     select = Select(chrome_browser.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div[3]/div[2]/div/span/form/span[1]/div[2]/span/div[2]/div/div/div/div/div/div/span/span[3]/span/span/div[1]/div/div[2]/span[2]/span[16]/select'))
#             select.select_by_index(25)
# =============================================================================
        
# =============================================================================
#         # Last Name / Entity
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPBRLastName").send_keys(f"{Entry4.get()}")
        
#         #Drop down menu (Prevous Report)
        
# =============================================================================
#         iframe = chrome_browser.find_element_by_xpath("/html/body/form/iframe")
#         chrome_browser.switch_to.frame(iframe)
# =============================================================================
        time.sleep(1)
        
        if variable.get() == 'Yes':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[1]/div[3]/div[2]/table[3]/tbody/tr[1]/td[2]/select"))
            select.select_by_index(1)
        else:
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[1]/div[3]/div[2]/table[3]/tbody/tr[1]/td[2]/select"))
            select.select_by_index(2)
            
            
#         #Drop Down Menu (Is this an existing Axos Bank Loan)
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanType")
        
        if variable1.get() == 'New':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanType"))
            select.select_by_index(1)
        else:
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanType"))
            select.select_by_index(2)
            
        #         # Loan Purpose (drop down)
        
        if variable2.get() == 'New Loan':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(1)
        elif variable2.get() == 'Renewal/Subsequent Transaction':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(2)
        elif variable2.get() == 'Change in Collateral':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(3)
        elif variable2.get() == 'Foreclosure':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(4)
        elif variable2.get() == 'OREO':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(5)
        elif variable2.get() == 'Refinance':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(6)
        elif variable2.get() == 'Asset Valuation':
            select = Select(chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_ProjectPurposeID"))
            select.select_by_index(7)    
# =============================================================================
#         iframe = chrome_browser.find_element_by_xpath("/html/body/form/iframe")
#         chrome_browser.switch_to.frame(iframe)
# =============================================================================
        time.sleep(1)
        #Loan Property Type dropdown
        #OPTIONS4 = ["Land", "Lodging/Hospitality", "Multi-Family", "Office", "Residential", "Retail-Commercial"]
        if variable3.get() == 'Land':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[1]"))
            select.select_by_index(6)
        
        elif variable3.get() == 'Lodging/Hospitality':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[1]"))
            select.select_by_index(7)
        elif variable3.get() == 'Multi-Family':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[1]"))
            select.select_by_index(8)
        elif variable3.get() == 'Office':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[1]"))
            select.select_by_index(9)    
        elif variable3.get() == 'Residential':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[1]"))
            select.select_by_index(10)        
        elif variable3.get() == 'Retail-Commercial':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[1]"))
            select.select_by_index(11)  
        
        time.sleep(2)
        
        #         #Property Type Drop-down
        if variable4.get() == 'Commercial':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(1)
        elif variable4.get() == 'Commercial Land':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(2)
        elif variable4.get() == 'Multi-Family Land':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(3)
        elif variable4.get() == 'Multi-Family-Apartment':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(4) 
        elif variable4.get() == 'Residential Subdivision Land':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(5)
        elif variable4.get() == 'Full Service':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(1)
        # Multi Family Drop Down
        elif variable4.get() == 'Hotel':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(2)
        elif variable4.get() == 'Condominium/PUD Bldg(s)':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(1)
        elif variable4.get() == 'Garden/Low-Rise Apartments':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(2)
        elif variable4.get() == 'Mid/High-Rise':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(3)
            
        elif variable4.get() == 'Mobile/Manufactured Home Park':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(4)  
        
        elif variable4.get() == 'Student-Oriented Housing-Student-Oriented Apartment':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(5)     
        
        # Office Drop Down
        elif variable4.get() == 'Condominium Unit(s)':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(1)
        
        elif variable4.get() == 'Creative/Loft':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(2)
            
        elif variable4.get() == 'Medical Office':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(3)
            
        elif variable4.get() == 'Mixed Use-Office-Multi Family':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(4) 
            
        elif variable4.get() == 'Mixed Use-Office-Retail':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(5)
            
        elif variable4.get() == 'Office Building-High-Rise':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(6)
            
        elif variable4.get() == 'Office Building-Low-Rise':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(7)
            
        elif variable4.get() == 'Office Building-Mid-Rise':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(8)
            
        elif variable4.get() == '2-4 Units':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(1)
            
        elif variable4.get() == 'Condominium Unit(s)':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(1)
            
        elif variable4.get() == 'Convenience Store':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(2)
            
        elif variable4.get() == 'Free Standing Building-Bank Branch':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(3)
            
        elif variable4.get() == 'Free Standing Building-Free Standing':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(4)
            
        elif variable4.get() == 'Mixed Use-Retail-Office':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(5)
            
        elif variable4.get() == 'Mixed Use-Retail-Residential':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(6)
            
        elif variable4.get() == 'Restaurant-Fast Food':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(7)
            
        elif variable4.get() == 'Street Retail':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[5]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[2]/div/select[2]"))
            select.select_by_index(8)
            
        #chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_drpPropertyMajorTypeID")    
        
#         
#         # Loan Amount
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanAmount").send_keys("0")
#         

#         
#         #Loan Number
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_LoanNumber").send_keys(f"{Entry6.get()}")
#         
#         

#         #Save and continue
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblContinue").click()
#       
        time.sleep(8)
        
# =============================================================================
#         iframe = chrome_browser.find_element_by_xpath("/html/body/form/iframe")
#         chrome_browser.switch_to.frame(iframe)
# =============================================================================
#         OPTIONS7 = ["Existing", "Land Only", "New Addition", "Other (Retired)", "Proposed Construction", "Under Construction", "Under Renovation"]
    #######variable6 = StringVar(root)
#         #status drop down
        if variable6.get() == 'Existing':
            
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(1)
        
        elif variable6.get() == 'Land Only':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(2)
            
        elif variable6.get() == 'New Addition':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(3)
        elif variable6.get() == 'Other (Retired)':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(4)
        elif variable6.get() == 'Proposed Construction':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(5)
            chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_drpStatus")
        elif variable6.get() == 'Under Construction':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(6)  
        elif variable6.get() == 'Under Renovation':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select"))
            select.select_by_index(7)
            
            
#         #tenancy drop down
        if variable7.get() == 'Vacant':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(1)
        elif variable7.get() == 'NA':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(2)  
        elif variable7.get() == 'Single Tenant Investor':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(3)
        elif variable7.get() == 'Multi Tenant Investor':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(4)
        elif variable7.get() == 'Owner Occupied 100%':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(5)
        elif variable7.get() == 'Owner Occupied > 51%':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(6)
            
        elif variable7.get() == 'Owner Occupied <= 50%':
            select = Select(chrome_browser.find_element_by_xpath("/html/body/form/div[6]/div/table/tbody/tr/td/table/tbody/tr[3]/td[2]/select"))
            select.select_by_index(7)
            
        
# =============================================================================
#     OPTIONS8 = ["Vacant", "NA", "Single Tenant Investor", "Multi Tenant Investor", "Owner Occupied 100%", "Owner Occupied > 51%", "Owner Occupied <= 50%"]
#     variable7 = StringVar(root)
# =============================================================================
    

#         
#         #Update
        chrome_browser.find_element_by_xpath('/html/body/form/div[6]/div/table/tfoot/tr/td/div[1]/input').click()
#         
        time.sleep(2)
#         
#         
#         
#         #Property Contact - affiliation
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPPC1Affiliation")
#         #Property Contact - Last Name
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPPC1LastName")
#         #Property Contact - Phone number
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_SRFPPC1Phone")
#         
#         #Pending/Recent Sale Drop down
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_IsPendingSale")
#         
#         #Send Selected Service
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblSendService").click()
#         
#         #Save Details for all
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblSaveAllDetails").click()
#         
#         #Send Selected Services
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_lblSendService").click()
#         
#         time.sleep(10)
#         # Advanced Button
        chrome_browser.find_element_by_xpath("/html/body/div[2]/header/div[2]/div[1]/form/div[1]/div/a/span").click()
#         
#         #Search by Property Filter (Street Number)
        chrome_browser.find_element_by_id("ProjectSearchCriteria_Address_StreetNumber").send_keys("...")
#         #Search by Property Filter (Street Name)
        chrome_browser.find_element_by_id("ProjectSearchCriteria_Address_StreetName")
#          #Search by Property Filter (City)
        chrome_browser.find_element_by_id("ProjectSearchCriteria_Address_City")
#         
#         #Click Element 
        chrome_browser.find_element_by_xpath("/html/body/div[2]/div/div/div/div[2]/div[1]/div/div[4]/table/tbody/tr[2]/td[2]/a").click()
#         
#         #Expand All Tasks
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[1]/a[2]').click()
#         time.sleep(2)
        chrome_browser.find_element_by_id("edit-task").click()
#         
#         #Uncheck Tentative
        chrome_browser.find_element_by_id("IsTypeRequired").click()
#         
#         #Due date
        chrome_browser.find_element_by_xpath("/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/form/div[2]/div/div[2]/div/div[2]/span[1]/span/input").send_keys("testing...")
#         
#         
        chrome_browser.find_element_by_id('pjt-task-save-button').click()
#         
#         
#         #Click RFP Manager
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/div[1]/a[5]').click()
#         
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[1]/ul/li[4]/div/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/div[1]/a[5]')
#         
#         #Click Direct Award
        chrome_browser.find_element_by_id("direct-award").click()
#         
#         #Add fee
        chrome_browser.find_element_by_id("ctl00_ctl00_contentBody_contentBody_txtBidFee").send_keys("200")
#         
#         #Select/Edit Recipients
        chrome_browser.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/a/b').click()
#         
#         #Deselect Show Pre-selected Vendors
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_chkShowPreselected').click()
#         
#         #Deselect Certified Vendors Only
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_chkCertifiedOnly').click()
#         
#         #Search
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_Button1').click()
#         
#         #Click "Include Vendor Information"
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_chkIncludeVendorInfo').click()
#         
#         #Input Last Name
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_searchLastName').send_keys('Johnson')
#         
#         #Input First Name
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_searchLastName').send_keys('Kimberly')
#         
#         #Search
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_Button1').click()
#         
#         #Click K Johnson
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_grdRFPRecipients_ctl02_chkRecipient').click()
#         
#         #Click Select
        chrome_browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/button[1]/span').click()
#         
#         #Direct Award
        chrome_browser.find_element_by_id('ctl00_ctl00_contentBody_contentBody_cmdDirectAward').click()
#         
#         #Press ok
        chrome_browser.find_element_by_xpath('/html/body/div[1]/div[3]/div/button/span').click()
#         
#         #close window
        chrome_browser.find_element_by_xpath('/html/body/div[39]/div[1]/div/a[8]/span').click()
        
        
        
        
    IR_go =ttk.Button(win, text="Go", command=IR)
    IR_go.pack()
    IR_go.place(relx=.4,rely=.75)
    
    refresh_button = ttk.Button(win, text="Refresh", command=refresh)
    refresh_button.pack()
    refresh_button.place(relx=.75,rely=.75)
    


    
    
#     
# =============================================================================
    

button3.config(command=popup_bonus)




root.mainloop()