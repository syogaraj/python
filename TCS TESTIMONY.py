'''Write an automation script using selenium webdriver which will give the least fare of
the bus in www.redbus.in by giving the from, to, departure date. Once search button is
clicked then the script should delimit the search by giving bus type and bus rating. The
details should be parameterized from the Excel sheet and result should be updated with
the price and bus details'''

# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re

class Redbus(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.implicitly_wait(30)
        self.base_url = "https://www.redbus.in/"
        self.verificationErrors = []
        self.accept_next_alert = True
	self.availability = []
	self.departure = []
	self.price = []
	self.arrival = []
	self.ratings = []
	self.bus_type = []

    def test_redbus(self):
	
	#open workbook
	import openpyxl
	workbook = openpyxl.load_workbook('demo.xlsx')
	cell = workbook.active
	sheet = workbook.worksheets[0]
	ws2 = workbook.create_sheet(1)	
	#rows = int(sheet.get_highest_row())-1
	
	for row in range(1,4):
		native_place = cell['A'+str(row)].value
		arrival_place = cell['B'+str(row)].value
		from_date = str(cell['C'+str(row)].value)
		frd = (from_date[:2])
		
		#web_page
		value = []
		adate = []
		qwe = []
        	driver = self.driver
        	driver.get(self.base_url)
        	driver.find_element_by_id("txtSource").clear()
        	driver.find_element_by_id("txtSource").send_keys(native_place)
        	driver.find_element_by_id("txtDestination").clear()
        	driver.find_element_by_id("txtDestination").send_keys(arrival_place)
        	driver.find_element_by_xpath(".//*[@id='txtOnwardCalendar']").click()

		#iterate through the datepicker
		for i in range(1,8):
			adate.append((driver.find_element_by_xpath(".//*[@id='rbcal_txtOnwardCalendar']/table[1]/tbody/tr[7]/td[%s]"%i)).text)	
		ind = adate.index(frd)
		driver.find_element_by_xpath(".//*[@id='rbcal_txtOnwardCalendar']/table[1]/tbody/tr[7]/td[%s]"%(ind+1)).click()
        	driver.find_element_by_id("txtReturnCalendar").click()
        	driver.find_element_by_id("searchBtn").click()

		#Retrieval of data from the result
		self.availability = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[2]/h4")
		self.departure = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[4]/div[1]/div[1]/a")
		self.price = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[8]/span")
		self.seats = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[6]/div[1]")
		self.arrival = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[4]/div[1]/div[3]/a")
		self.ratings = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[7]/div[2]")
		self.bus_type = driver.find_elements_by_xpath(".//*[@id='onwardTrip']/div[2]/ul/li[*]/div/div[2]/span")
			
		#finding the min_value	
		for i in self.price:
			value.append(i.text)
		min_value = min(value)
		min_price = value.index(min_value)
	
		#export to workbook
		av = self.availability[min_price]
		de = self.departure[min_price]
		arr = self.arrival[min_price]
		seat = self.seats[min_price]
		pr = self.price[min_price]
		rate = self.ratings[min_price]
		btype = self.bus_type[min_price]
		ws2['A'+str(row)] = av.text
		ws2['B'+str(row)] = de.text
		ws2['C'+str(row)] = arr.text	
		ws2['D'+str(row)] = seat.text
		ws2['E'+str(row)] = pr.text
		ws2['F'+str(row)] = rate.text
		ws2['G'+str(row)] = btype.text
		workbook.save('demo.xlsx')
	
	self.driver.quit()

    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException, e: return False
        return True

    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException, e: return False
        return True

    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True

    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
