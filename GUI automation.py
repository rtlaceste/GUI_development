from tkinter import *
from tkinter import ttk
from selenium import webdriver
import time
root = Tk()
label = ttk.Label(root, text="Hello, Tkinter")
label.pack()
label.config(foreground='black', background='white')
label.config(font=('Courier', 15, 'bold'))
label.config(text="\tHello Boss! \nWhat would you like to do today?", anchor='n')
logo = PhotoImage(file=r'C:/Users/Troy/Desktop/Python Projects/GUI Development/A.i..gif')
small_logo = logo.subsample(6, 6)
label.config(image=small_logo)
label.config(compound='left')

button = ttk.Button(root, text='Youtube?')
button2 = ttk.Button(root, text='Reddit?')
button3 = ttk.Button(root, text='Excel?')
button.pack()
button2.pack()
button3.pack()
button.place(relx=.46, rely=.75)
button2.place(relx=.63, rely=.75)
button3.place(relx=.8, rely=.75)

def go_to_youtube():
    chrome_browser = webdriver.Chrome(r'C:\Users\Troy\Desktop\WebDev\Misc\chromedriver.exe')
    chrome_browser.get('https://www.youtube.com')



def go_to_reddit():
    chromebrowser = webdriver.Chrome(r'C:\Users\Troy\Desktop\WebDev\Misc\chromedriver.exe')
    chromebrowser.get('https://www.reddit.com/')

def go_to_excel():
    pass

button.config(command=go_to_youtube)
button2.config(command=go_to_reddit)


root.mainloop()
