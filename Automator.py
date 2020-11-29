# -*- coding: utf-8 -*-
"""
Created on Fri Nov 27 08:24:59 2020

@author: RLaceste
"""

from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time
import pyautogui
import os
#from simple_salesforce import Salesforce
from PIL import ImageTk,Image
import json
import win32com.client as win32
from datetime import datetime

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

json_file = open(r"C:\Users\Troy\Desktop\idk.json","r",encoding='utf-8')
info = json.load(json_file)
json_file.close()

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


# json format
# =============================================================================
# {
#     "SalesForce Username": "Test",
#     "SalesForce Password": "jackson",
#     "CredCo Username": "rlaceste20",
#     "Credco Password": "Panda544!",
#     "AML Solutions Username": "RLaceste558",
#     "Sitelynx User ID":"Rlaceste@axosbank.com",
#     "Sitelynx Password":"/3T@jchE"
# }
# =============================================================================


def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    print(filename)


root = Tk()
root.geometry("1200x500")
root.title("Loan Applicant Information ver. 1.0")


folder_path = StringVar()
lbl1 = Label(root,textvariable=folder_path)
lbl1.grid(row=0, column=1)
button2 = Button(text="Send to what folder?", command=browse_button)
button2.grid(row=30, column=5)


# =============================================================================
# image1 = Image.open(r"C:\Users\rlaceste\Downloads\Axos logo.png")
# image2 = image1.resize((100,100))
# test = ImageTk.PhotoImage(image2)
# label1 = Label(image=test)
# label1.image = test
# label1.grid(row=0,column=0)
# =============================================================================

root.grid_columnconfigure((0,14), weight=1)


labeltitle = Label(root, text = "IPL", font=("Courier", 25)).grid(row=0, column=1)
labeltitle2 = Label(root, text = "Dept", font=("Courier", 25)).grid(row=0, column=3)
label_spacer = Label(root, text = "", font=("Courier", 8)).grid(row=29, column=12)
labeltitle3 = Label(root, text = "If info not known, leave blank. Thanks.\n Axos Bank IPL Department", font=("Courier", 8)).grid(row=0, column=0)


spacer1 = Label(root, text = "           ", font=("Courier", 8)).grid(row=5, column=4)
spacer2 = Label(root, text = "           ", font=("Courier", 8)).grid(row=43, column=1)
spacer3 = Label(root, text = "           ", font=("Courier", 8)).grid(row=38, column=1)


applabel1 = Label(root,text = "Applicant 1").grid(row=1, column=1)
applabel2 = Label(root,text = "Applicant 2").grid(row=1, column=3)


Label3 = Label(root, text="First Name")
Label4 = Label(root, text="Last Name")
Label5 = Label(root, text="City")
Label6 = Label(root, text="Loan Number")
Label7 = Label(root, text ="SSN")
Label8 = Label(root, text = "Address No.")
Label9= Label(root, text = "Street Name")
Label10 = Label(root, text = "State (Ex: CA, NY)")
Label11 = Label(root, text = "Zip Code")
Labela = Label(root, text = "Loan Officer First Name").grid(row=12, column = 0)
Labelb = Label(root, text = "Loan Officer Last Name").grid(row=13, column=0)
Labelc = Label(root, text = "JCA First Name").grid(row=14, column=0)
Labeld = Label(root, text = "JCA Last Name").grid(row=15, column=0)
Labele = Label(root, text = "Tax I.D. #").grid(row=16,column=0)





Label12 = Label(root, text = "OFAC")
Label13 = Label(root, text = 'Credit Report').grid(row=4, column=5)
Label14 = Label(root, text = 'Google Search').grid(row=5, column=5)
Label15 = Label(root, text = "Lexis Nexis").grid(row=6, column=5)
Label16 = Label(root, text = "IRS TIN (n/a)").grid(row=7, column=5)
Label17 = Label(root, text = "Flood Certification").grid(row=8, column = 5)
Label19 = Label(root, text = "Environmental Report").grid(row=9, column=5)
#Label20 = Label(root, text = "Phase 1").grid(row =10, column = 6)
Label21 = Label(root, text = "Inspection Report").grid(row=10, column = 5)
Label22 = Label(root, text = "Order Appraisal").grid(row=11,column=5)
Label23 = Label(root, text = "Legal Ticket").grid(row=12, column=5)
Label24 = Label(root, text = "Populate Salesforce").grid(row=13, column=5)
Label25 = Label(root, text = "Confirmation Email").grid(row=14, column = 5)

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
c12=BooleanVar(value=0)
c13 = BooleanVar(value=0)
c14 = BooleanVar(value=0)
c15 = BooleanVar(value=0)
c16 = BooleanVar(value=0) # Email Confirmation

check1 = BooleanVar(value=0)


OFAC = Checkbutton(root, text="", variable=c1).grid(row=3, column = 6, sticky=W)
Credit_Report = Checkbutton(root, text="individual", variable=c2).grid(row=4, column = 6, sticky=W)
Credit_Report_Joint = Checkbutton(root, text="joint", variable=c15).grid(row=4, column = 7, sticky=W)
Google_Search = Checkbutton(root, text="", variable=c3).grid(row=5, column = 6, sticky=W)
LN = Checkbutton(root, text="", variable=c4).grid(row=6, column = 6, sticky=W)
IRS_Tin = Checkbutton(root, text="", variable=c5).grid(row=7, column = 6, sticky=W)
f_cert_8 = Checkbutton(root, text="simple", variable=c6).grid(row=8, column = 6, sticky=W)
f_cert_30 = Checkbutton(root, text="complex", variable=c7).grid(row=8, column = 7, sticky=W)
ETS = Checkbutton(root, text="ETS", variable=c8).grid(row=9, column = 6, sticky=W)
EDR = Checkbutton(root, text="EDR", variable=c13).grid(row=9, column = 8, sticky=W)
P_1 = Checkbutton(root, text="Phase 1", variable=c9).grid(row=9, column = 7, sticky=W)
IR = Checkbutton(root, text="", variable=c10).grid(row=10, column = 6, sticky=W)
OA = Checkbutton(root, text='',variable=c11).grid(row=11, column=6, sticky=W)
LT = Checkbutton(root, text='',variable=c12).grid(row=12, column=6, sticky=W)
SF = Checkbutton(root, text ='', variable =c14).grid(row=13, column=6, sticky=W)
email_confirm = Checkbutton(root, text = "Add comments: ", variable=c16).grid(row=14, column=6, sticky=W)






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
    
    #try:
        #os.mkdir(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}")
    #except Exception:
        #pass
    
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
   "download.default_directory": fr"{folder_path.get()}",
   "download.prompt_for_download": False,
   "download.directory_upgrade": True,
   "safebrowsing.enabled": True
 })
   
    chrome_browser = webdriver.Chrome(r'C:\Users\Troy\Desktop\chromedriver.exe', options = options)
    chrome_browser.maximize_window()
    
    #OFAC
    
    if c2.get() and c15.get():
        chrome_browser.close()
        messagebox.showwarning("Error", "Error - You can not pull both an individual and joint credit report")
        
    if c6.get() and c7.get():
        chrome_browser.close()
        messagebox.showwarning("Error", "Error - You can not order both :D")
        
        
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
        chrome_browser.find_element_by_id("Content_Main_txtTaxID").send_keys(f"{TIN_number.get()}")
        chrome_browser.find_element_by_id("Content_Main_btnVerifyOrder").click()
        
        
        status['Flood Cert - Simple'] = 'Yes'
        
        #Submit Order
        #chrome_browser.find_element_by_id("Content_Main_btnSubmitOrder").click()
        #Correct Order Here
        
    else:
        pass
    
    
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
        chrome_browser.find_element_by_id("name").send_keys(info["Sitelynx User ID"])
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
        chrome_browser.find_element_by_id("continue-3").click()
        
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
        #iframe = chrome_browser.find_element_by_id('ctl00_ctl00_HeadTag')
        #chrome_browser.switch_to_frame(iframe)
        
        #chrome_browser.find_element_by_name("ctl00$ctl00$contentBody$contentBody$modalUsrSearch$txtFirstName").send_keys("TEST")
        
# =============================================================================
#         element = chrome_browser.find_element_by_id("ProjectSearchCriteria_QuickSearch")
#         action = ActionChains(chrome_browser) 
#   
#         #perform the operation 
#         action.move_to_element_with_offset(element, 0, 50).click().perform()
# =============================================================================
        
        status["Inspection Report"] = "Yes"
        
        

        
    if c11.get():
        chrome_browser.get("http://jhaknow/reports/report/Credit/Underwriting/Income%20Property%20Lending%20(IPL)/IPL%20Appraisal%20Assignment%20ZIP%20Lookup")
        
        
        status["Order Appraisal"] = "Yes"
        
        
    if c12.get():
        
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
        
        chrome_browser.switchTo().defaultContent()
        
        
        status["Legal Ticket"] = "Yes"
        
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
        
        time.sleep(1)
        chrome_browser.get("https://axos.exactbid.com/Home/Dashboard")
        chrome_browser.find_element_by_id("LoginModel_UserName").send_keys("Rlaceste")
        chrome_browser.find_element_by_id("LoginModel_Password").send_keys("Panda544!")
        chrome_browser.find_element_by_id("LoginButton").click()
        
        
    if c14.get():
        

        time.sleep(1)
        chrome_browser.get("https://bofi.my.salesforce.com/?ec=302&startURL=%2Fvisualforce%2Fsession%3Furl%3Dhttps%253A%252F%252Fbofi.lightning.force.com%252Flightning%252Fr%252FReport%252F00O3o0000055toBEAQ%252Fview%253FqueryScope%253DuserFolders")
        chrome_browser.find_element_by_id("username").send_keys("rlaceste@axosbank.com")
        chrome_browser.find_element_by_id("password").send_keys("Troyboyoy544!")
        chrome_browser.find_element_by_id("Login").click()
        loan_search = chrome_browser.find_element_by_id("169:0;p")
        loan_search.send_keys(f"{Entry6.get()}")
        loan_search.send_keys(Keys.RETURN)
        
        

        
        pass
    
    
    if c16.get():
        # Email Confirmation
        
        
        
        run_time_sec = (datetime.now().second - startTime)
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'rtlaceste@gmail.com'
        mail.Recipients.Add("troyboyoy@gmail.com")
        mail.CC = "rtlaceste@ucdavis.edu"
        mail.Subject = (f"Update on Loan #{Entry6.get()} for Applicant(s): {Entry3.get()} {Entry4.get()}, {Entry16.get()} {Entry17.get()}  ")
        mail.Body = (f'** This is an automated message** \n \n Information regarding what was pulled on Loan #{Entry6.get()} through the script: \n\n OFAC: {status["OFAC"]} \n Credit Report (Individual): {status["Individual Credit Report"]} \n Credit Report (Joint): {status["Joint Credit Report"]} \n Google Search: {status["Google Search"]} \n Lexis Nexis: {status["Lexis Nexis"]} \n IRS-TIN: {status["IRS-TIN"]} \n Flood Cert (Simple): {status["Flood Cert - Simple"]} \n Flood Cert (Complex): {status["Flood Cert - Complex"]} \n Environmental Report (ETS): {status["Environmental Report (ETS)"]} \n Environmental Report (EDR): {status["Environmental Report (EDR)"]} \n Environmental Report (Phase 1): {status["Environmental Report (Phase 1)"]} \n Inspection Report: {status["Inspection Report"]} \n Order Appraisal: {status["Order Appraisal"]} \n Legal Ticket: {status["Legal Ticket"]} \n \n User Comments: {email_comments.get()} \n \n Script total run time: {run_time_sec} seconds')
                     
        mail.Send()
        

    
    
    #os.startfile(fr"{folder_path.get()}")
    #chrome_browser.quit()        
        
        
global Entry3
Entry3 = Entry(root)

global Entry4
Entry4 = Entry(root)

global Entry5
Entry5 = Entry(root)

global Entry6
Entry6 = Entry(root)

global Entry7
Entry7 = Entry(root)

global Entry8
Entry8 = Entry(root)

global Entry9
Entry9 = Entry(root)

global Entry10
Entry10=Entry(root)

global Entry11
Entry11=Entry(root)

global Entry12
Entry12=Entry(root)

global Entry13
Entry13=Entry(root)

global Entry14
Entry14=Entry(root)

global Entry15
Entry15=Entry(root)

global Entry16 
Entry16 = Entry(root)
Entry16.grid(row=3, column = 3, sticky="ew")

global Entry17
Entry17 = Entry(root)
Entry17.grid(row=4, column=3, sticky="ew")

global Entry18
Entry18 = Entry(root)
Entry18.grid(row=5, column=3, sticky="ew")

global Entry19
Entry19 = Entry(root)
Entry19.grid(row=6, column=3, sticky="ew")

global Entry20
Entry20 = Entry(root)
Entry20.grid(row=7, column=3, sticky="ew") 

global Entry21
Entry21 = Entry(root)
Entry21.grid(row=8, column=3, sticky="ew") 

global Entry22
Entry22 = Entry(root)
Entry22.grid(row=9, column=3, sticky="ew")

global Entry23
Entry23 = Entry(root)
Entry23.grid(row=10, column=3, sticky="ew")

global Entry24
Entry24 = Entry(root)
Entry24.grid(row=11, column=3, sticky="ew")

global Entry25
Entry25 = Entry(root)
Entry25.grid(row=12, column=3, sticky="ew")

global Entry26
Entry26 = Entry(root)
Entry26.grid(row=13, column=3, sticky="ew")

global Entry27
Entry27 = Entry(root)
Entry27.grid(row=14, column=3, sticky="ew")

global Entry28
Entry28 = Entry(root)
Entry28.grid(row=15, column=3, sticky="ew")

global TIN_number
TIN_number = Entry(root)
TIN_number.grid(row=16,column=1, sticky="ew")

global email_comments
email_comments = Entry(root)
email_comments.grid(row=14, column =7, sticky = "ew")



button = ttk.Button(root, text='Get Documents')
button.config(command=Get_OFAC)
button2=ttk.Button(root, text='Submit')
button2.config(command=Master_Function)


# ADD CONFIG HERE


Label3.grid(row=3, column=0)
Entry3.grid(row=3, column=1, sticky="ew")
Label4.grid(row=4, column=0)
Entry4.grid(row=4, column=1, sticky="ew")
Label5.grid(row=5, column=0)
Entry5.grid(row=5, column=1, sticky="ew")
Label6.grid(row=6, column=0, sticky="ew")
Entry6.grid(row=6, column=1, sticky="ew")
#button.grid(row=20, column=1, sticky="ew")
Label7.grid(row=7,column=0, sticky="ew")
Entry7.grid(row=7,column=1, stick="ew")
Label8.grid(row=8,column=0,sticky="ew")
Label9.grid(row=9,column=0,sticky="ew")
Label10.grid(row=10, column=0, sticky="ew")
Label11.grid(row=11, column=0, sticky="ew")
Entry8.grid(row=8,column=1,sticky="ew")
Entry9.grid(row=9,column=1,sticky="ew")
Entry10.grid(row=10,column=1,sticky="ew")
Entry11.grid(row=11,column=1,sticky="ew")
Label12.grid(row=3, column=5)
button2.grid(row=30, column=7, sticky="ew")
Entry12.grid(row=12, column=1, sticky="ew")
Entry13.grid(row=13, column=1, sticky="ew")
Entry14.grid(row=14, column=1, sticky="ew")
Entry15.grid(row=15, column=1, sticky="ew")


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
Entry14.insert(0,"Troy")
Entry15.insert(0,"Laceste")




root.mainloop()
