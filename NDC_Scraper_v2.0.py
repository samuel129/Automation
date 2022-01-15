import selenium
from selenium import webdriver
import time
from time import sleep
import select
from selenium.webdriver.support.ui import Select
import os.path
from selenium.webdriver.common.keys import Keys
from pynput.keyboard import Key, Controller
from tkinter import *
import tkinter as tk

apn = {
'subject': '',
'county': '',
1: '',
2: '',
3: '',
4: '', 
5: '',
6: 'x',
}

a = tk.Tk()
a.title('Ndc Scraper')

subv = tk.StringVar()
countyv = tk.StringVar()
apnv1 = tk.StringVar()
apnv2 = tk.StringVar()
apnv3 = tk.StringVar()
apnv4 = tk.StringVar()
apnv5 = tk.StringVar()
apnv6 = tk.StringVar()

Label(a, text = 'County: ').grid(row = 0)
Label(a, text = 'Subject APN: ').grid(row = 1)
Label(a, text = 'Comp 1 APN: ').grid(row = 2)
Label(a, text = 'Comp 2 APN: ').grid(row = 3)
Label(a, text = 'Comp 3 APN: ').grid(row = 4)
Label(a, text = 'Comp 4 APN: ').grid(row = 5)
Label(a, text = 'Comp 5 APN: ').grid(row = 6)
Label(a, text = 'Comp 6 APN: ').grid(row = 7)
county = Entry(a, textvariable = countyv).grid(row = 0, column = 1)
sub = Entry(a, textvariable = subv).grid(row = 1, column = 1)
apn1 = Entry(a, textvariable = apnv1).grid(row = 2, column = 1)
apn2 = Entry(a, textvariable = apnv2).grid(row = 3, column = 1)
apn3 = Entry(a, textvariable = apnv3).grid(row = 4, column = 1)
apn4 = Entry(a, textvariable = apnv4).grid(row = 5, column = 1)
apn5 = Entry(a, textvariable = apnv5).grid(row = 6, column = 1)
apn6 = Entry(a, textvariable = apnv6).grid(row = 7, column = 1)

keyboard = Controller()

def getEntry():
	apn['county'] = countyv.get()
	apn['subject'] = subv.get()
	apn[1] = (apnv1).get()
	apn[2] = (apnv2).get()
	apn[3] = (apnv3).get()
	apn[4] = (apnv4).get()
	apn[5] = (apnv5).get()
	apn[6] = (apnv6).get()
	if apn[6] == '':
		apn[6] = 'x'

def login(browser): #Logging in to NDC
	browser.get('https://www.parcelquestappraise.com/PQAppraise/Login')
	username = browser.find_element_by_id('ctl00_ContentPlaceHolder1_txtUserName')
	username.send_keys('kenny')
	next = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtUserName"]')
	next.click()
	password = browser.find_element_by_id('ctl00_ContentPlaceHolder1_txtPassword')
	password.send_keys('315chong')
	sign_in = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnLogin"]')
	sign_in.click()
	url = browser.current_url
	if url == 'https://www.parcelquestappraise.com/Duplicate-Login':
		dupe = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnContinue"]')
		dupe.click()
	browser.get('https://www.parcelquestappraise.com/search/search_property_apn.aspx')
	countySelect = Select(browser.find_element_by_id('ctl00_ContentPlaceHolder1_FindCounty1_ddlCounty'))
	countySelect.select_by_visible_text(apn['county'])

def compSearch():
	getEntry()
	browser = webdriver.Chrome('C:\Program Files\Google\Chrome\Application\chromedriver.exe')
	login(browser)
	browser.get('https://www.parcelquestappraise.com/export/Export.aspx')
	countySelect = Select(browser.find_element_by_id('ctl00_ContentPlaceHolder1_FindCounty1_ddlCounty'))
	countySelect.select_by_visible_text(apn['county'])
	softwareSelect = Select(browser.find_element_by_id('ctl00_ContentPlaceHolder1_ddlSoftware'))
	softwareSelect.select_by_visible_text('a la mode TOTAL')
	sub = browser.find_element_by_id('ctl00_ContentPlaceHolder1_msAPN0')
	sub.send_keys(apn['subject'])
	for num in range(1,6):
		if apn[num] != '':
			search = browser.find_element_by_id('ctl00_ContentPlaceHolder1_msAPN' + str(num))
			search.send_keys(apn[num])
	if apn[6] != 'x':
		if apn[num] != '':
			search = browser.find_element_by_id('ctl00_ContentPlaceHolder1_msAPN' + str(6))
			search.send_keys(apn[6])
	clickExport = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSubmitFormW"]')
	clickExport.click()

def printComps():
	getEntry()
	browser = webdriver.Chrome('C:\Program Files\Google\Chrome\Application\chromedriver.exe') 
	login(browser)
	for num in range(1,6):
		if apn[num] != '':
			APN = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtAPNNumber"]').clear()
			APN = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtAPNNumber"]')
			APN.send_keys(apn[num])
			keyboard.press(Key.enter) 
			keyboard.release(Key.enter)
			#abc = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSubmitForm"]').click()
			printButton = browser.find_element_by_css_selector("img#ctl00_imgPrint").click()
			time.sleep(2)
			keyboard.press(Key.enter)
			keyboard.release(Key.enter)
			browser.get('https://www.parcelquestappraise.com/search/search_property_apn.aspx')
			time.sleep(1)

	if apn[6] != 'x':
			APN = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtAPNNumber"]').clear()
			APN = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtAPNNumber"]')
			APN.send_keys(apn[6])
			search = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSubmitForm"]').click()
			printButton = browser.find_element_by_css_selector("img#ctl00_imgPrint").click()
			time.sleep(5)
			keyboard.press(Key.enter)
			keyboard.release(Key.enter)

printButton = tk.Button(a, text = 'Print NDC Files', width = 15, command = printComps).grid(row = 8)
exportButton = tk.Button(a, text = 'Export NDC Files', width = 15, command = compSearch).grid(row = 8, column = 1)

a.mainloop()