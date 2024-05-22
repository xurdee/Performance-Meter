from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.support.ui import Select
import tkinter.messagebox
import time



driver = webdriver.Chrome(executable_path="C:\Drivers\chromedriver.exe")
driver.get("https://makaut1.ucanapply.com/smartexam/public/result-details")

r = 12000118075
index = 2
c = 0

try:
    for i in range(60):
        driver.find_element_by_id("username").send_keys(str(r))
        drp = Select(driver.find_element_by_id("semester"))
        drp.select_by_index(index)
        driver.find_element_by_tag_name('button').click()
        try:
            if index % 2 == 0:
                line = driver.find_element_by_xpath("//strong[contains(text(),'EVEN')]").text
            else:
                line = driver.find_element_by_xpath("//strong[contains(text(),'ODD')]").text

        except NoSuchElementException:
            sgpa = 0.0
            c = c + 1
        else:
            li = line.split(":")
            if (li[1].strip() == "--"):
                sgpa = 0.0
                c = c+1
            else:
                sgpa = float(li[1])

        finally:
            print(sgpa)
            r = r + 1

            driver.find_element_by_link_text("Reset").click()
            time.sleep(1)


except NoSuchElementException:
    tkinter.messagebox.showerror("Error", "Please check your internet connectivity!")

except ElementNotInteractableException:
    tkinter.messagebox.showerror("Error", "You have poor net connectivity!. Please try again later")

else:
    driver.close()
    tkinter.messagebox.showinfo("Success", "All Data Have Been Fetched successfully!")
    if(c == 1):
        tkinter.messagebox.showinfo("Info", "Result of {} student is still pending! ".format(c))
    else:
        tkinter.messagebox.showinfo("Info", "Result of {} students are still pending! ".format(c))












