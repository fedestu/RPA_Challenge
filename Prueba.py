# +
from RPA.Browser.Selenium import Selenium

browser_lib = Selenium()

class LaPrueba:

    def open_the_website(self, url):
        try:
            browser_lib.open_available_browser(url)
            browser_lib.maximize_browser_window()
        except:
            raise Exception("Can't open the browser.")

    def click_DiveIn(self):
        try:
            browser_lib.click_element_if_visible("//a[@aria-controls='home-dive-in']")
        except:
            raise Exception("Can't click in DiveIn")
