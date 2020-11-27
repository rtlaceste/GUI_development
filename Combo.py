# -*- coding: utf-8 -*-
"""
Created on Mon Nov 23 16:54:02 2020

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
import time
import pyautogui
import os
#from simple_salesforce import Salesforce
from PIL import ImageTk,Image
import urllib.request
import win32com.client
from bs4 import BeautifulSoup



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



def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    print(filename)


root = Tk()
root.geometry("1300x500")
root.title("Loan Applicant Information ver. 1.0")


folder_path = StringVar()
lbl1 = Label(root,textvariable=folder_path)
lbl1.grid(row=0, column=1)
button2 = Button(text="Send to what folder?", command=browse_button)
button2.grid(row=30, column=5)


#image1 = Image.open(r"C:\Users\rlaceste\Downloads\Axos logo.png")
#image2 = image1.resize((100,100))
#test = ImageTk.PhotoImage(image2)
#label1 = Label(image=test)
#label1.image = test
#label1.grid(row=0,column=0)

root.grid_columnconfigure((0,14), weight=1)


labeltitle = Label(root, text = "IPL", font=("Courier", 25)).grid(row=0, column=1)
labeltitle2 = Label(root, text = "Dept", font=("Courier", 25)).grid(row=0, column=3)
label_spacer = Label(root, text = "", font=("Courier", 8)).grid(row=29, column=12)
labeltitle3 = Label(root, text = "If info not known, leave blank. Thanks.\n Axos Bank IPL Department", font=("Courier", 8)).grid(row=0, column=10)


spacer1 = Label(root, text = "           ", font=("Courier", 8)).grid(row=5, column=4)

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



LoginSf = Label(root, text = "Salesforce Username").grid(row=3, column = 10)
PassSf = Label(root, text = "Salesforce Password").grid(row=4, column = 10)

LoginCr = Label(root, text = "CredCo Username").grid(row=6, column=10)
PassCr = Label(root, text = "CredCo Password").grid(row=7, column=10)

LoginAml = Label(root, text = "AML Soln Username").grid(row=9, column=10)
PassAml = Label(root, text ="AML Soln Password").grid(row=10, column=10)

LoginRims = Label(root, text = "RIMS Username").grid(row=12, column=10)
PassRims = Label(root, text = "RIMS Password").grid(row=13, column=10)

LoginLynx = Label(root, text = "Sitelynx Username").grid(row=15, column=10)
PassLynx = Label(root, text = "Sitelynx Password").grid(row=16, column=10)

Label12 = Label(root, text = "OFAC")
Label13 = Label(root, text = 'Credit Report').grid(row=4, column=5)
Label14 = Label(root, text = 'Google Search').grid(row=5, column=5)
Label15 = Label(root, text = "Lexis Nexis").grid(row=6, column=5)
Label16 = Label(root, text = "IRS TIN").grid(row=7, column=5)
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
   
    
    #try:
        #os.mkdir(fr"C:\Users\rlaceste\Desktop\Intake Checks\{Entry3.get()} {Entry4.get()}")
    #except Exception:
        #pass
    
   
    chrome_browser = webdriver.Chrome(r'C:\Users\Troy\Desktop\chromedriver.exe')
    chrome_browser.maximize_window()
    
    #OFAC
    
    if c1.get():
        
        chrome_browser.get('https://sanctionssearch.ofac.treas.gov/default.aspx')
        chrome_browser.find_element_by_id("ctl00_MainContent_txtLastName").send_keys(Entry3.get() +" " + Entry4.get())
    
        chrome_browser.find_element_by_id("ctl00_MainContent_txtAddress").send_keys(Entry8.get() + " " + Entry9.get())
    
        chrome_browser.find_element_by_id("ctl00_MainContent_txtCity").send_keys(Entry5.get())
        
        
        chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.CONTROL, 'a')
        chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys(Keys.BACKSPACE)
        chrome_browser.find_element_by_id("ctl00_MainContent_Slider1_Boundcontrol").send_keys('93')
    
        chrome_browser.find_element_by_id("ctl00_MainContent_btnSearch").click()
        
        myScreenshot = pyautogui.screenshot()
        myScreenshot.save(fr"{folder_path.get()}/OFAC for {Entry3.get()} {Entry4.get()} Loan #{Entry6.get()}.png")
        
        
        try:
           
            chrome_browser.find_element_by_xpath('//*[@id="btnDetails"]').click()
            print("OFAC Hit!")
            
            
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
                
            
            
                messagebox.showwarning("Warning", "URGENT - OFAC HIT. CONTACT MANAGER")
        
        
            except NoSuchElementException:
                pass
            
        else:
            pass
            
            
        
        
        
        
        
    else:
        pass
    
    
    #Credit Report Pull (Individual)
    if c2.get():
         
        
        
        chrome_browser.get('https://www.credco.com/ecredco/security/login.aspx')
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
        
        
        
        if len(Entry20.get()) > 0:
            print("Go")
            
            chrome_browser.get('https://www.credco.com/ecredco/security/login.aspx')
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Username").send_keys("rlaceste20")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_Password").send_keys("Panda544!")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_LoginButton").click()
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_ctl00_liOrder").click()
        
    
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtLastName").send_keys(f"{Entry17.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtFirstName").send_keys(f"{Entry16.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLApplicantDetails_Applicant_txtSSN").send_keys(f"{Entry20.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtNum").send_keys(f"{Entry21.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtStreetName").send_keys(f"{Entry22.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtCity").send_keys(f"{Entry18.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_ddlState").send_keys(f"{Entry23.get()}")
            chrome_browser.find_element_by_id("ctl00_ctl00_phM_phB_DHQControl_CRLAddressDetails_txtZip").send_keys(f"{Entry24.get()}")
            
        else:
            pass
            
            
            
                    
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
        
        
        if len(Entry16.get()) > 0:
             chrome_browser.get("http://www.google.com")
             search = chrome_browser.find_element_by_name('q')
             search.send_keys(f'"{Entry16.get()} {Entry17.get()}" AND "money laundering" OR "fraud" OR "lawsuits"')
             search.send_keys(Keys.RETURN) # hit return after you enter search text
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
        
        
    else:
        pass
        
    if c5.get():
        
        chrome_browser.get("https://la1.www4.irs.gov/eauth/pub/login.jsp?Data=VGFyZ2V0TG9BPUY%253D&TYPE=33554433&REALMOID=06-0004b429-9e8b-1a23-9e85-163b0acf4037&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=UOkC7yx4eMTO24FGxPfBRb5q3Mj3Xh3pyXfBEjYyHJ97nGCXu16wx5MzFHjfZmlG&TARGET=-SM-https%3a%2f%2fla1%2ewww4%2eirs%2egov%2fesrv%2ftinm%2fproauth%2efaces")
        
    else:
        pass
    
    
    if c6.get():
        
        chrome_browser.get("https://weborders.floodapp.com/")
        chrome_browser.find_element_by_id("acceptBttn").click()
        chrome_browser.find_element_by_id("UserName").send_keys("IPL-JC")
        chrome_browser.find_element_by_id("Password").send_keys("Axos2021!")
        
    else:
        pass
    
    
    if c7.get():
        
        chrome_browser.get("https://weborders.floodapp.com/Login ")
        chrome_browser.find_element_by_id("acceptBttn").click()
        chrome_browser.find_element_by_id("UserName").send_keys("Placeholder")
        chrome_browser.find_element_by_id("Password").send_keys("Placeholder")
        
    if c8.get():
        
        chrome_browser.get("https://www.sitelynx.net/admin/login ")
        chrome_browser.find_element_by_id("name").send_keys("Rlaceste@axosbank.com")
        chrome_browser.find_element_by_id("password").send_keys("/3T@jchE")
        chrome_browser.find_element_by_xpath("/html/body/div/div/form/fieldset/div[4]/button").click()
        
    
    if c9.get():
        #Phase 1
        
        pass
    
    
    if c10.get():
        
        chrome_browser.get("https://axos.exactbid.com/Account/Login?msg=m")
        chrome_browser.find_element_by_id("LoginModel_UserName").send_keys("Rlaceste")
        chrome_browser.find_element_by_id("LoginModel_Password").send_keys("Panda544!")
        chrome_browser.find_element_by_id("LoginButton").click()
        chrome_browser.get("https://axos.exactbid.com/Project/NewServiceRequest")
        chrome_browser.find_element_by_xpath("//input[@name='ctl00$ctl00$contentBody$contentBody$modalUsrSearch$txtFirstName']").send_keys("TEST")
        
        
    if c11.get():
        url = 'http://i3.ytimg.com/vi/J---aiyznGQ/mqdefault.jpg'
        urllib.request.urlretrieve(url, '/Users/scott/Downloads/cat.jpg')
    
    if c12.get():
        
        
        pyautogui.moveTo(3000, 500)  # moves mouse to X of 600, Y of 500.
        chrome_browser.get("https://bofi.service-now.com/nav_to.do?uri=%2Fcom.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1%26sysparm_id%3D2e76be11db5d76405661b96c4e961985%26sysparm_link_parent%3D59de44e4db19f2405661b96c4e961965%26sysparm_catalog%3D3910bd12df132100dca6a5")
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys("{TAB}") #Press tab... to change focus or whatever
        #   pyautogui.press("tab").send_keys("aa")
        shell.SendKeys("Tom Constantine")
        shell.SendKeys("{TAB}")
        shell.SendKeys(f"{Entry6.get()}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("O")
        shell.SendKeys("{TAB}")
        shell.SendKeys("IPL Intake and Processing")
        shell.SendKeys("{TAB}")
        shell.SendKeys("R")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys("{TAB}")
        shell.SendKeys(f"{Entry3.get()} {Entry4.get()} #{Entry6.get()}")
        
        
        
    if c13.get():
        
        chrome_browser.get("https://axos.exactbid.com/Home/Dashboard")
        chrome_browser.find_element_by_id("LoginModel_UserName").send_keys("Rlaceste")
        chrome_browser.find_element_by_id("LoginModel_Password").send_keys("Panda544!")
        chrome_browser.find_element_by_id("LoginButton").click()
        
        
    if c14.get():
        
        chrome_browser.get("https://bofi.my.salesforce.com/?ec=302&startURL=%2Fvisualforce%2Fsession%3Furl%3Dhttps%253A%252F%252Fbofi.lightning.force.com%252Flightning%252Fr%252FReport%252F00O3o0000055toBEAQ%252Fview%253FqueryScope%253DuserFolders")
        chrome_browser.find_element_by_id("username").send_keys("rlaceste@axosbank.com")
        chrome_browser.find_element_by_id("password").send_keys("Troyboyoy544!")
        chrome_browser.find_element_by_id("Login").click()
        loan_search = chrome_browser.find_element_by_id("169:0;p")
        loan_search.send_keys(f"{Entry6.get()}")
        loan_search.send_keys(Keys.RETURN)
        
        chrome_browser.find_elements_by_xpath("//*[contains(text(), '900')]").click()
        
        
        
    if c15.get():
        
        #Joint Credit Report
        
        pass
    
    
    chrome_browser.quit()        
        
        
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


def toggle_passwords():
    
    if SF_pass_check.var.get():
        sf2Entry['show'] = ""
    else:
        sf2Entry['show'] = "*"
        
        
    if CR_pass_check.var.get():
        crEntry2['show'] = ""
    else:
        crEntry2['show'] = "*"
        
    if AML_pass_check.var.get():
        amlEntry2['show'] = ""
    else:
        amlEntry2['show'] = "*"
        
    if RIMS_pass_check.var.get():
        rimsEntry2['show'] = ""
    else:
        rimsEntry2["show"] = "*"
        
    if LYNX_pass_check.var.get():
        lynxEntry2['show'] = ""
    else:
        lynxEntry2['show'] = "*"



global sfEntry
sfEntry = Entry(root).grid(row=3, column=11, sticky="ew")
global sf2Entry
sf2Entry = Entry(root)

sf2Entry.default_show_val = sf2Entry["show"]
sf2Entry["show"] = "*"
SF_pass_check = Checkbutton(root, text = "show", onvalue=True,offvalue=False,command = toggle_passwords)
SF_pass_check.var = BooleanVar(value=False)
SF_pass_check['variable'] = SF_pass_check.var
SF_pass_check.grid(row=4,column=12,sticky=W)
sf2Entry.grid(row=4, column=11, sticky="ew")


global crEntry
crEntry = Entry(root).grid(row=6, column=11, sticky="ew")
global crEntry2
crEntry2 = Entry(root)

crEntry2.default_show_val = crEntry2["show"]
crEntry2["show"] = "*"
CR_pass_check = Checkbutton(root, text = "show", onvalue=True,offvalue=False,command = toggle_passwords)
CR_pass_check.var = BooleanVar(value=False)
CR_pass_check['variable'] = CR_pass_check.var
CR_pass_check.grid(row=7,column=12,sticky=W)

crEntry2.grid(row=7, column=11, sticky="ew")


global amlEntry
amlEntry = Entry(root).grid(row=9, column=11, sticky="ew")

global amlEntry2
amlEntry2 = Entry(root)

amlEntry2.default_show_val = amlEntry2["show"]
amlEntry2["show"] = "*"
AML_pass_check = Checkbutton(root, text = "show", onvalue=True,offvalue=False,command = toggle_passwords)
AML_pass_check.var = BooleanVar(value=False)
AML_pass_check['variable'] = AML_pass_check.var
AML_pass_check.grid(row=10,column=12,sticky=W)

amlEntry2.grid(row=10, column=11, sticky="ew")

global rimsEntry
rimsEntry = Entry(root).grid(row=12, column=11, sticky="ew")

global rimsEntry2
rimsEntry2 = Entry(root)

rimsEntry2.default_show_val = rimsEntry2["show"]
rimsEntry2["show"] = "*"
RIMS_pass_check = Checkbutton(root, text = "show", onvalue=True,offvalue=False,command = toggle_passwords)
RIMS_pass_check.var = BooleanVar(value=False)
RIMS_pass_check['variable'] = RIMS_pass_check.var
RIMS_pass_check.grid(row=13,column=12,sticky=W)

rimsEntry2.grid(row=13, column=11, sticky="ew")

global lynxEntry
lynxEntry = Entry(root).grid(row=15, column=11, sticky="ew")

global lynxEntry2
lynxEntry2 = Entry(root)

lynxEntry2.default_show_val = lynxEntry2["show"]
lynxEntry2["show"] = "*"
LYNX_pass_check = Checkbutton(root, text = "show", onvalue=True,offvalue=False,command = toggle_passwords)
LYNX_pass_check.var = BooleanVar(value=False)
LYNX_pass_check['variable'] = LYNX_pass_check.var
LYNX_pass_check.grid(row=16,column=12,sticky=W)


lynxEntry2.grid(row=16, column=11, sticky="ew")





button = ttk.Button(root, text='Get Documents')
button.config(command=Get_OFAC)
button2=ttk.Button(root, text='Submit')
button2.config(command=Master_Function)


##ADD CONFIG HERE


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