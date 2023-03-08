import mechanize
import re
import requests
import sys
from bs4 import BeautifulSoup
import openpyxl
from lxml import html
import os
import ssl
import tkinter as tk


# Disable SSL certificate verification globally
os.environ['PYTHONHTTPSVERIFY'] = '0'

class LoginForm:
    def __init__(self, master):
        self.master = master
        master.title("Login Form")

        self.label_username = tk.Label(master, text="Username")
        self.label_username.pack()

        self.entry_username = tk.Entry(master)
        self.entry_username.pack()

        self.label_password = tk.Label(master, text="Password")
        self.label_password.pack()

        self.entry_password = tk.Entry(master, show="*")
        self.entry_password.pack()

        self.submit_button = tk.Button(master, text="Submit", command=self.submit)
        self.submit_button.pack()
        
        # Define instance variables to store username and password
        self.username = None
        self.password = None

    def submit(self):
        self.username = self.entry_username.get()
        self.password = self.entry_password.get()

        #print(f"Username: {username}")
        #print(f"Password: {password}")
        self.master.destroy()

# Done
def check_login(username:str, password:str) -> bool:
    ssl._create_default_https_context = ssl._create_unverified_context
    br = mechanize.Browser()

    # ignore robots.txt rules
    br.set_handle_robots(False)

    # navigate to the login page
    br.open('https://www.gpro.net/gb/Login.asp')

    # select the login form and fill in the fields
    br.select_form(nr=0)
    br.form['textLogin'] = username
    br.form['textPassword'] = password

    # submit the login form
    response = br.submit()
    
    response = list(br.links(url_regex=re.compile("DriverProfile")))
    
    # check the response for an error message indicating invalid credentials
    if len(response) == 1:
        print('Login successful.')
        return True
    # handle the error, e.g. retry with different credentials or exit the program
    else:
        print('Invalid username or password.')
        return False

# Done 
def browser_open(username:str, password:str):
    ssl._create_default_https_context = ssl._create_unverified_context
    br = mechanize.Browser()

    # ignore robots.txt rules
    br.set_handle_robots(False)

    # navigate to the login page
    br.open('https://www.gpro.net/gb/Login.asp')

    # select the login form and fill in the fields
    br.select_form(nr=0)
    br.form['textLogin'] = username
    br.form['textPassword'] = password

    # submit the login form
    br.submit()
    
    return br

# Done
def extract_row_data(tree, first_arg:str, second_arg:str, third_arg:str, row_num:str):
    return tree.xpath(f"//{first_arg}[contains(@{second_arg}, '{third_arg}')]/tr[{row_num}]/td[1]/text()")[0]

# Done     
def open_workbook_worksheet(workbook_path:str , worksheet_name:str):
    
    workbook = openpyxl.load_workbook(workbook_path)
    worksheet = workbook[worksheet_name]
    return workbook, worksheet  
    
# Done     
def save_workbook(workbook, workbook_path):
    
    workbook.save(workbook_path)

# Done   
def write_to_excel(data, gp_list, actual_gp:str , worksheet, start_row, end_row, season_calendar=False):
    
    if not season_calendar:
        gp_index = gp_list.index(actual_gp)
        col_num = gp_index + 2
    else:
        col_num = gp_list
    
    for item in data.items():
        worksheet.cell(row=start_row, column=col_num).value = item[1]
        start_row = start_row + 1
        if start_row == end_row:
            break

# Done
def fill_season_calendar():
    # send a GET request to the calendar page
    response = requests.get("https://www.gpro.net/gb/Calendar.asp")

    # create a BeautifulSoup object
    soup = BeautifulSoup(response.content, "html.parser")

    gp_list = []
    for row in soup.find_all('tr'):
        cols = row.find_all('td')
        if len(cols) >= 3:
            gp_name = cols[2].text.strip()
            gp_name = gp_name.split(' GP')[0]
            gp_list.append(gp_name)
            
    # Build a dictionary that maps GP names to their corresponding numbers
    gp_dict = {}
    for i, gp_name in enumerate(gp_list, start=1):
        gp_dict[gp_name] = str(i)
    
    # Find the element with class 'yellow'
    yellow_elem = soup.find(class_='yellow')

    # Get the text content of the element, strip leading/trailing white space
    yellow_text = yellow_elem.get_text().strip()

    # Remove any '0' characters and '.' characters from the text
    yellow_text = yellow_text.replace('0', '').replace('.', '')

    # Look up the GP name in the dictionary based on the modified text
    gp_name = next((name for name, num in gp_dict.items() if num == yellow_text), None)

    # Print the GP name if it was found, or a message if it wasn't found
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '1 Season')
    
    new_gp_dict = {v: k for k, v in gp_dict.items()}
    
    write_to_excel(new_gp_dict, 2, gp_name, worksheet, 6, 24, True)
    
    save_workbook(workbook, 'me_s80.xlsx')
    
    if gp_name:
        return gp_list, gp_name
    else:
        print(f"No GP name found for '{yellow_text}'")
        sys.exit()
    
# Done
def fill_driver_profile(br, gp_list , actual_gp):
    
    # navigate to the login page
    br.open('https://www.gpro.net/gb/DriverContract.asp')
    
    tree = html.fromstring(br.response().get_data())
    
    driver_data = {
        'overall': int(tree.xpath("normalize-space(//tr[contains(@data-step, '2')]//td/text())")),
        'concentration': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '3')),
        'talent': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '4')),
        'aggressiveness': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '5')),
        'experience': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '6')),
        'technical_insight': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '7')),
        'stamina': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '8')),
        'charisma': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '9')),
        'motivation': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '10')),
        'reputation': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '12')),
        'weight': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '14')),
        'age': int(extract_row_data(tree, 'table', 'class', 'squashed leftalign', '15'))
    }
    
    driver_salary = {
        'salary': tree.xpath("normalize-space(//div[contains(@data-step, '1')]//td/text())").replace('.', ' '),
        'point_bonus': tree.xpath("normalize-space(//div[contains(@data-step, '1')]//table[1]//tr[4]//td[1]//text())").replace('.', ' '),
        'podium_bonus': tree.xpath("normalize-space(//div[contains(@data-step, '1')]//table[1]//tr[3]//td[1]//text())").replace('.', ' '),
        'win_bonus': tree.xpath("normalize-space(//div[contains(@data-step, '1')]//table[1]//tr[2]//td[1]//text())").replace('.', ' ')
    }
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '2 Stuff Data')
    
    write_to_excel(driver_data, gp_list, actual_gp, worksheet, 5, 17, False)
    
    write_to_excel(driver_salary, gp_list, actual_gp, worksheet, 18, 22, False)
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '8 Strategy Planner')
    
    write_to_excel(driver_data, 15, actual_gp, worksheet, 2, 14, True)
    
    # Load the workbook
    save_workbook(workbook, 'me_s80.xlsx')

# Done
def fill_staff_facilities(br, gp_list, actual_gp):
    # navigate to the login page
    br.open('https://www.gpro.net/gb/StaffAndFacilities.asp')
    
    tree = html.fromstring(br.response().get_data())
    
    staff_overal = {

        'experience': extract_row_data(tree, 'table', 'data-step', '4', '1'),
        'motivation': extract_row_data(tree, 'table', 'data-step', '4', '2'),
        'tech_skill': extract_row_data(tree, 'table', 'data-step', '4', '3'),
        'stress_handling': extract_row_data(tree, 'table', 'data-step', '4', '5'),
        'concentration': extract_row_data(tree, 'table', 'data-step', '4', '6'),
        'effiency': extract_row_data(tree, 'table', 'data-step', '4', '7')               
    }
    
    facilities = {
        'wind_tunnel': int(tree.xpath("normalize-space(//table[contains(@data-step, '6')]//td/text())")),
        'pit_stop': extract_row_data(tree, 'table', 'data-step', '6', '2'),
        'r_d_workshop': extract_row_data(tree, 'table', 'data-step', '6', '3'),
        'r_d_design_center': extract_row_data(tree, 'table', 'data-step', '6', '4'),
        'engineer_workshop': extract_row_data(tree, 'table', 'data-step', '6', '5'),
        'alloy_chemical': extract_row_data(tree, 'table', 'data-step', '6', '6'),
        'commercial': extract_row_data(tree, 'table', 'data-step', '6', '7'),
    }
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '2 Stuff Data')
    
    write_to_excel(staff_overal, gp_list, actual_gp, worksheet, 40, 46, False)
    
    write_to_excel(facilities, gp_list, actual_gp, worksheet, 48, 55, False)
    
    # Load the workbook
    save_workbook(workbook, 'me_s80.xlsx')

#Done
def fill_car_level(br, gp_list, actual_gp):
    
    br.open('https://www.gpro.net/gb/UpdateCar.asp')
    tree = html.fromstring(br.response().get_data())
    
    car_level = {
        'chassis': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlCha')]/text())")),
        'engine': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlEng')]/text())")),
        'front_wing': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlFW')]/text())")),
        'rear_wing': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlRW')]/text())")),
        'underbody': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlUB')]/text())")),
        'sidepod': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlSid')]/text())")),
        'cooling': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlCoo')]/text())")),
        'gearbox': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlGea')]/text())")),
        'brakes': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlBra')]/text())")),
        'suspension': int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlSus')]/text())")),
        'electronics':int(tree.xpath("normalize-space(//td[contains(@id, 'newLvlEle')]/text())"))
    }
    
    car_wear = {
        'chassis': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearCha')]/text())").strip('%')) / 100,
        'engine': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearEng')]/text())").strip('%')) / 100,
        'front_wing': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearFW')]/text())").strip('%')) / 100,
        'rear_wing': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearRW')]/text())").strip('%')) / 100,
        'underbody': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearUB')]/text())").strip('%')) / 100,
        'sidepod': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearSid')]/text())").strip('%')) / 100,
        'cooling': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearCoo')]/text())").strip('%')) / 100,
        'gearbox': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearGea')]/text())").strip('%')) / 100,
        'brakes': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearBra')]/text())").strip('%')) / 100,
        'suspension': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearSus')]/text())").strip('%')) / 100,
        'electronics': int(tree.xpath("normalize-space(//td[contains(@id, 'newWearEle')]/text())").strip('%')) / 100
    }
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '3 Part Data')
    
    write_to_excel(car_level, gp_list, actual_gp, worksheet, 5, 16, False)
    
    write_to_excel(car_wear, gp_list, actual_gp, worksheet, 57, 68, False)
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '8 Strategy Planner')
    
    write_to_excel(car_level, 9, actual_gp, worksheet, 3, 14, True)
    
    write_to_excel(car_wear, 10, actual_gp, worksheet, 3, 14, True)
    
    # Load the workbook
    save_workbook(workbook, 'me_s80.xlsx')
    
    pass

#Done
def fill_gp_info(br, actual_gp):
    
    br.open('https://www.gpro.net/gb/RaceSetup.asp')
    #print(br.response().get_data())
    
    tree = html.fromstring(br.response().get_data())
    
    qualify_1_temp = tree.xpath("normalize-space(//img[contains(@name, 'WeatherQ')]/../text()[contains(., 'Temp')])")
    qualify_1_value = int(re.search('\d+', qualify_1_temp).group())
    
    qualify_1 = {
        'temperature': qualify_1_value,
    }

    qualify_2_temp = tree.xpath("normalize-space(//img[contains(@name, 'WeatherR')]/../text()[contains(., 'Temp')])")
    qualify_2_value = int(re.search('\d+', qualify_2_temp).group())

    qualify_2 = {
        'temperature': qualify_2_value
    }

    rTempRangeOne = tree.xpath("normalize-space(//td[contains(text(), 'Temp')]/../../tr[2]/td[1]/text())")
    rTempRangeTwo = tree.xpath("normalize-space(//td[contains(text(), 'Temp')]/../../tr[2]/td[2]/text())")
    rTempRangeThree = tree.xpath("normalize-space(//td[contains(text(), 'Temp')]/../../tr[4]/td[1]/text())")
    rTempRangeFour = tree.xpath("normalize-space(//td[contains(text(), 'Temp')]/../../tr[4]/td[2]/text())")

    # This returns results like "Temp: 12*-17*", but we want just integers, so clean up the values
    rTempMinOne = int((re.findall(r"\d+", rTempRangeOne))[0])
    rTempMaxOne = int((re.findall(r"\d+", rTempRangeOne))[1])
    rTempMinTwo = int((re.findall(r"\d+", rTempRangeTwo))[0])
    rTempMaxTwo = int((re.findall(r"\d+", rTempRangeTwo))[1])
    rTempMinThree = int((re.findall(r"\d+", rTempRangeThree))[0])
    rTempMaxThree = int((re.findall(r"\d+", rTempRangeThree))[1])
    rTempMinFour = int((re.findall(r"\d+", rTempRangeFour))[0])
    rTempMaxFour = int((re.findall(r"\d+", rTempRangeFour))[1])
	# Find the averages of these temps for the setup
    
    rTemp = int(((rTempMinOne + rTempMaxOne) + (rTempMinTwo + rTempMaxTwo) + (rTempMinThree + rTempMaxThree) + (
		rTempMinFour + rTempMaxFour)) / 8)
    
    race = {
        'temperature': rTemp
    }
    
    workbook, worksheet = open_workbook_worksheet('me_s80.xlsx', '8 Strategy Planner')
    
    write_to_excel(qualify_1, 8, actual_gp, worksheet, 17, 18, True)
    
    write_to_excel(qualify_2, 8, actual_gp, worksheet, 21, 22, True)
    
    write_to_excel(race, 8, actual_gp, worksheet, 25, 26, True)
    
    write_to_excel(race, 2, actual_gp, worksheet, 17, 18, True)
    
    # Load the workbook
    save_workbook(workbook, 'me_s80.xlsx')
    
    pass

def main():
    
    root = tk.Tk()
    login_form = LoginForm(root)
    root.mainloop()
    
    #check if the login is good
    if not check_login(login_form.username, login_form.password):
        sys.exit()
    
    browser = browser_open(login_form.username, login_form.password)
    
    gp_list, actual_gp = fill_season_calendar()
    
    fill_driver_profile(browser, gp_list , actual_gp)

    fill_staff_facilities(browser, gp_list, actual_gp)

    fill_car_level(browser, gp_list, actual_gp)
    
    fill_gp_info(browser, actual_gp)
    
if __name__ == '__main__':
    main()

