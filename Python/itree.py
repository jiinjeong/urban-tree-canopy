"""
 *****************************************************************************
   FILE :           itree.py
   AUTHORS :        Jiin Jeong and Heather Wing
   DATE :           May 31 - June 6, 2018
   DESCRIPTION :    Automates filling the iTree website with data from Excel.
   REQUIRES :
   (1) Selenium
   (2) xlrd
   (3) ChromeDriver or SafariDriver
   (4) .xls format Excel file ("tree.xls")
   REMAINING BUGS :
   (1) Export only works for less than 50 trees.
   (2) xlrd library only works for .xls files for now (b/c of formatting).
 *****************************************************************************
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import xlrd  # Library for reading data and formatting info from Excel.


class AutomateBrowser():
    """ This class automates the browser and performs desired tasks
        within the browser. """

    def __init__(self, driver):
        self._driver = driver

    def select_option_by_name(self, selectName, text):
        """ Selects an option with text in a drop down list with name 'selectName'. """

        # Waits until drop down list is loaded.
        WebDriverWait(self._driver, 10).until(lambda driver:
            self._driver.find_element_by_xpath("//*[contains(text(), '"+text+"')]"))
        selectBox = Select(self._driver.find_element_by_xpath('//select[@name="'+selectName+'"]'))
        selectBox.select_by_visible_text(text)

    def select_option_by_class(self, selectClass, text):
        """ Selects text box by class. """

        try:  # Waits until drop down list is loaded.
            wait = WebDriverWait(self._driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//select[@class='"+selectClass+"']"))))
        finally:
            selectBox = Select(self._driver.find_element_by_xpath('//select[@class="'+selectClass+'"]'))
            selectBox.select_by_visible_text(text)

    def push_button(self, buttonClass):
        """ Pushes button with class 'buttonClass'. """

        element = self._driver.find_element_by_xpath("//button[@class='"+buttonClass+"']")
        self._driver.execute_script("arguments[0].click();", element)

    def select_radio(self, radioButtonID):
        """ Selects radio button with id 'radioButtonID'. """

        try:
            wait = WebDriverWait(self._driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//input[@id='"+radioButtonID+"']"))))
        finally:
            radio_button = self._driver.find_element_by_xpath("//input[@id='"+radioButtonID+"']")
            self._driver.execute_script("arguments[0].click();", radio_button)

    def enter_text(self, textBox, text):
        """ Enters text to text box with id 'textBox'. """

        try:
            wait = WebDriverWait(self._driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//input[@id='"+textBox+"']"))))
        finally:
            text_area = self._driver.find_element_by_xpath("//input[@id='"+textBox+"']")
            self._driver.execute_script("document.getElementById('"+textBox+"').click();")
            self._driver.execute_script("arguments[0].setAttribute('value','"+text+"');",text_area)

    def enter_text_by_class(self, textBox, text):

        try:
            wait = WebDriverWait(self._driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//input[@class='"+textBox+"']"))))
        finally:
            text_area = self._driver.find_element_by_xpath("//input[@class='"+textBox+"']")
            self._driver.execute_script("document.getElementsByClassName('"+textBox+"')[0].click();")
            self._driver.execute_script("arguments[0].setAttribute('value','"+text+"');",text_area)

    def special_enter(self,rowID,textBox,text):

        try:
            wait = WebDriverWait(self._driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//tr[@id='"+rowID+"']/td/input[@class='"+textBox+"']"))))
        finally:
            text_area = self._driver.find_element_by_xpath("//tr[@id='"+rowID+"']/td/input[@class='"+textBox+"']")
            self._driver.execute_script("document.getElementsByClassName('"+textBox+"')[0].click();")
            self._driver.execute_script("arguments[0].setAttribute('value','"+text+"');",text_area)

    def push_button_by_class(self, buttonClass):
        """ Pushes button with class 'buttonID'. """

        element = self._driver.find_element_by_xpath("//button[@class='"+buttonClass+"']")
        self._driver.execute_script("arguments[0].click();", element)

    def push_button_by_id(self, buttonID):
        """ Pushes button with id 'buttonID'. """

        element = self._driver.find_element_by_xpath("//input[@id='"+buttonID+"']")
        self._driver.execute_script("arguments[0].click();", element)
    
    def push_button_export(self, buttonClass, text):
        """ Pushes button with class 'buttonClass' in an a type div. """
        try:  # Waits until info table is loaded
            wait = WebDriverWait(self._driver, 500)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//*[contains(text(), '"+text+"')]"))))
        finally:
            element = self._driver.find_element_by_xpath("//a[@class='"+buttonClass+"']")
            self._driver.execute_script("arguments[0].click();", element)
        

    def special_select(self,rowID,selectClass,text):
        """ Specialized select box for drop downs in table rows with IDs. """
        try:
            wait = WebDriverWait(self._driver, 10)
            element = wait.until(EC.visibility_of_element_located((By.XPATH,("//tr[@id='"+rowID+"']/td/select[@class='"+selectClass+"']"))))
        finally:
            selectBox = Select(self._driver.find_element_by_xpath("//tr[@id='"+rowID+"']/td/select[@class='"+selectClass+"']"))
            selectBox.select_by_visible_text(text)


class ExcelData():
    """ This class changes Excel data into inputtable format.
        It only works with .xls files because of formatting. """

    def __init__(self, sheet):
        self._sheet = sheet
    
    def treename(self, species):
        """ Changes treename into common name in iTree website.
        We find all the treenames in the Excelfile beforehand
        by using the program file itree-treename.py. """

        if species == "Acacia salicina":
            return "Acacia, Green"
        elif species == "Acacia saligna":
            return "Acacia, Bailey"
        elif species == "Chitalpa tashkentensis":
            return "Chitalpa"
        elif species == "Corymbia citriodora":
            return "Gum, Lemon-scented"
        elif species == "Gingko biloba":
            return "Ginkgo"
        elif species == "Jacaranda mimosifolia":
            return "Jacaranda"
        elif species == "Lagerstroemia x 'Natchez'":
            return "Crapemyrtle"
        elif species == "Lophostemon confertus":
            return "Box, Brisbane"
        elif species == "Pistacia chinensis":
            return "Pistache, Chinese"
        elif species == "Quercus rubra" or "Quercus rubra ":
            return "Oak, Northern red"

    def distance(self, dist):
        """ Changes distance to inputtable format in iTree website. """

        if dist == "0'-20'" or dist == "0-20'" or dist == "N/A":
            return "0-19"
        elif dist == "20'-40'":
            return "20-39"
        elif dist == "40'-60'" or dist == "40-60'":
            return "40-59"

    def direction(self, direct):
        """ Changes direction to inputtable format in iTree website. """

        if direct == "N" or direct == "N/A":
            return "North (0°)"
        elif direct == "NE":
            return "Northeast (45°)"
        elif direct == "NW":
            return "Northwest (315°)"
        elif direct == "E":
            return "East (90°)"
        elif direct == "S":
            return "South (180°)"
        elif direct == "SE":
            return "Southeast (135°)"
        elif direct == "SW":
            return "Southwest (225°)"
        elif direct == "W":
            return "West (270°)"

    def readfile(self, site, start, end):
        """ Reads the file, and fills in Tree page with the data.
            I wish I could make this more general, but oh well. """

        row_id = 1  # Starting value for row_id

        # for i in range(3, 4):  # Used for testing.
        for i in range(start, end):  # Skips rows until start range.
            row = self._sheet.row(i)

            # Skips strikethrough rows.
            xf = workbook.xf_list[self._sheet.cell_xf_index(i, 1)]  # Checks first cell for strikethrough.
            font = workbook.font_list[xf.font_index]  # Finds the font element.

            if font.struck_out:  # If struck through, moves on to next iteration.
                continue

            # Skips blank or NaN rows.
            if row[1].value == '' or row[1].value == 'NaN':
                continue

            else:
                species = row[1].value  # Species.
                dbh = row[9].value  # Stock size.
                tree_dist = row[12].value  # Tree distance.
                tree_dir = row[11].value  # Tree direction.
                # print(species)  # Tests if the for-loop and conditions are working porperly.
                # print(dbh)
                # print(tree_dir)
                # print(tree_dist)

                itree.special_select("row-%s" % row_id, "tree-species", self.treename(species))
                itree.special_enter("row-%s" % row_id, "tree-dbh", "1.5")
                itree.special_select("row-%s" % row_id, "tree-building-distance", self.distance(tree_dist))
                itree.special_select("row-%s" % row_id, "tree-building-direction", self.direction(tree_dir))
                itree.push_button_by_id("add-row-button")

                row_id += 1


"""/*************** From here, we START working with iTree. ***************/"""
# Creates webdriver.
driver = webdriver.Chrome()  # Or webdriver.Chrome()

# Navigates to page in URL.
driver.get("https://planting.itreetools.org/app/location/")

# Asserts that project is in the driver title.
assert "Project" in driver.title
assert "No results found." not in driver.page_source

itree = AutomateBrowser(driver)


"""/*************** LOCATIONS Page ***************/"""
itree.select_option_by_name("partition","California")  # State
itree.select_option_by_name("secondary_partition","Los Angeles")  # County
itree.select_option_by_name("tertiary_partition","Claremont")  # City
itree.push_button_by_class("next btn btn-primary")  # Push button: NEXT


"""/*************** PARAMETERS Page ***************/"""
# Adds local parameters IF NECESSARY: radio buttons and fill-in options.
itree.select_radio('id_electricity_units_0')
itree.select_radio('id_natural_gas_units_0')
itree.enter_text("id_project_years",'25')
itree.push_button_by_class("next btn btn-primary")  # NEXT


"""/*************** TREES Page ***************/"""
workbook = xlrd.open_workbook("tree.xls", formatting_info=True)  # Opens file.
sheet = ExcelData(workbook.sheet_by_index(0))  # Retrives sheet.

# sheet.readfile(itree, 3, 85)  # Reads the file from the third row.
sheet.readfile(itree, 3, 50)  # ERROR: only works with 50 trees max.
itree.push_button_by_class("next btn btn-primary")  # NEXT


"""/*************** REPORT Page ***************/"""
itree.push_button_export("btn btn-default buttons-csv buttons-html5 btn-primary",'75')  # Download report csv
# wait function isn't working?
