from selenium import webdriver
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException
import time
import platform
from os import getcwd, system

# -------------------Chrome_Settings --------------------------
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--no-sandbox')
# chrome_options.add_argument("--headless")
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-notification')
chrome_options.add_argument('--disable-infobars')
# --------------------------------------------------------------

global driver, nums
Hustle_login_url = 'https://hustle-247.herokuapp.com/login'
n = 1
while True:

    print("\n• Hustler ID " + str(n))
    UserID = input("UserID: ")
    password = input("Password: ")
    try:
        nums = int(input("Number of Login Iterations: "))
    except Exception:
        print("Try Again.")
        break

    for i in range(nums):
        if (platform.system() == 'Windows'):
            driver = webdriver.Chrome(getcwd() + "\chromedriver.exe", options=chrome_options)
        if (platform.system() == 'Darwin'):
            driver = webdriver.Chrome(getcwd() + "/chromedriver", options=chrome_options)
        # --------------------------- Process -----------------------------------------------
        driver.get(Hustle_login_url)
        try:
            print("\n• Process Started for Iteration " + str(i + 1))

            driver.find_element_by_xpath('/html/body/form/input[1]').send_keys(UserID)
            driver.find_element_by_xpath('/html/body/form/input[2]').send_keys(password)
            driver.find_element_by_xpath('/html/body/form/input[3]').click()

            print("• Login In Process. Please Wait...")

            driver.implicitly_wait(20)
            driver.find_element_by_xpath('/html/body/a').click()

            print("• Login Finished. Now Logging Out...\n")
            time.sleep(5)
        except Exception:
            print(f'''
Error Occurred!
Please Make Sure UserName & Password Are Correct.
Please Make Sure Internet Is Working.
Stopping Process {i + 1}...
            ''')
            break
        finally:
            driver.quit()

        # --------------------------- Process end --------------------------------------------
    Choice = input("Do you want to Exit [Y/N]?\n-> ")
    if Choice == 'Y' or Choice == 'y':
        break
    if Choice == 'N' or Choice == 'n':
        n += 1
