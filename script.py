from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pathvalidate import sanitize_filename
import requests
import urllib
import os
import re

COURSE_NUM = ''
EMAIL = ''
PASSWORD = ''
PC_USER = ''
USING_PROFILE = True
PROFILE_NAME = 'Profile 1'
DOWNLOAD_PATH = f'C:\\Users\\{PC_USER}\\Desktop\\{COURSE_NUM}'
CHROME_DATA_PATH = f'C:\\Users\\{PC_USER}\\AppData\\Local\\Google\\Chrome\\User Data'
TIME_2FA = 60

headers = {
    'User-Agent' : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
}


""" switch_tabs(close)

Took inspiration from: https://www.selenium.dev/documentation/webdriver/interactions/windows/#switching-windows-or-tabs

Each tab or window has its own identifier known as a "window handle".
- We cycle through each window handle and compare it to the current one the WebDriver is on.
- If they are different, we switch to that tab.
- If close = True, we close the tab we were on before switching.

"""

def switch_tabs(close):
    for win_handle in driver.window_handles:
        if win_handle != driver.current_window_handle:
            if close:
                driver.close()
            driver.switch_to.window(win_handle)
            break


""" update_dict(dict, key, val)

dict must have the following structure where each of its values is a list.
dict = {
    key : [val],
}

Each val that is passed looks like this: val = [['linkurl', 'filename']]
Replacing val with examples:
dict = {
    key : [['linkurl1', 'filename1'], ['linkurl2', 'filename2'], ...],
}

- If the passed key already exists within the dictionary, we append the passed val to the value list of that key.
- Else, a new key is created and val is assigned as its value.

Example of the dictionary that is passed to this function:
links = {
    'Week 1' : [['linkurl1', 'Intro.pdf'], ['linkurl2', 'Vectors.pdf'], ['linkurl3', 'Parametric Equations.pdf']],
}

"""

def update_dict(dict, key, val):
    if key in dict:
        dict[key].append(val[0])
    else:
        dict[key] = val


""" microsoft_login

Called if USING_PROFILE = False. If you do not use a Chrome Profile to store cookies,
you are required to sign in through a Microsoft portal and perhaps complete 2FA authentication.

This function will automatically complete all steps in the portal except for 2FA where your smartphone
is required.

"""

def microsoft_login():
    # - email
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(EMAIL)
    driver.find_element(By.ID, "idSIButton9").click()
    # - password
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(PASSWORD)
    driver.find_element(By.ID, "idSIButton9").click()
    # - 2FA
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idChkBx_SAOTCAS_TD"))).click()
    WebDriverWait(driver, TIME_2FA).until(EC.element_to_be_clickable((By.ID, "KmsiCheckboxField"))).click()
    driver.find_element(By.ID, "idSIButton9").click()


""" initialiseDriver(isHeadless)

Initialises the WebDriver.
- If isHeadless = True, the script will run in headless mode,
  i.e. only through terminal. Not recommended as scraping may fail.
- If USING_PROFILE = True, we enter PROFILE_NAME and CHROME_DATA_PATH as arguments.

"""

def initialiseDriver(isHeadless):
    options = webdriver.ChromeOptions()
    if isHeadless:
        options.add_argument("--headless")
    if USING_PROFILE:
        options.add_argument(f"profile-directory={PROFILE_NAME}")
    options.add_argument(f"user-data-dir={CHROME_DATA_PATH}")
    driver = webdriver.Chrome(options=options)
    return driver


""" loginBlackboard()

Simply navigates the Blackboard login page (always required even when using a Chrome Profile).

"""

def loginBlackboard():
    driver.get('https://tcd.blackboard.com/')
    if not USING_PROFILE: WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "agree_button"))).click()
    driver.find_element(By.ID, "login-btn").click()
    if not USING_PROFILE: microsoft_login()


""" retrieveLinks()

Main function that collects all links and stores them in the links dictionary.

"""

def retrieveLinks():
    # Navigates to the course URL
    driver.get('https://tcd.blackboard.com/ultra/courses/_'+COURSE_NUM+'_1/cl/outline')
    
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'classic-learn-iframe')))
    driver.switch_to.frame('classic-learn-iframe')

    # 1. Click sidebar link with word 'Lectures', 'Slides', or 'Week'
    sidebar_buttons = driver.find_elements(By.XPATH, '//li[starts-with(@id, "paletteItem")]')
    for btn in sidebar_buttons:
        try:
            span = btn.find_element(By.XPATH, './/a/span')
        except:
            continue

        if "Lecture" in span.text or "Week" in span.text or "Slide" in span.text:
            btn.click()
            break

    # Refresh after clicking sidebar to avoid Stale exception
    driver.switch_to.default_content()
    driver.switch_to.frame('classic-learn-iframe')

    # posts is a list of all 'posts' on the course page (found using the same XPath)
    posts = driver.find_elements(By.XPATH, '//*[starts-with(@id, "contentListItem")]')
    links = {}
    for post in posts:
        driver.switch_to.default_content()
        driver.switch_to.frame('classic-learn-iframe')
        
        # Scroll to each post so that it is in view to avoid Stale exception
        driver.execute_script("arguments[0].scrollIntoView(true)", post)
        # Try find post's title (XPath changes whether the title is a link or not)
        try:
            title = post.find_element(By.XPATH, './/div[@class="item clearfix"]/h3/span[2]')
        except:
            title = post.find_element(By.XPATH, './/div[@class="item clearfix"]/h3/a/span')
            title_link = post.find_element(By.XPATH, './/div[@class="item clearfix"]/h3/a').get_attribute('href')
        # Find details div of the post, i.e. content section of the post
        details = post.find_element(By.XPATH, './/div[@class="details"]')
        # If details is empty, it is most likely a folder or file
        if details.text.isspace() or details.text == "":
            # Save title before it becomes stale
            title_text = title.text
            try:
                # Check if the post is a file, not a folder, by checking if there is a download icon next to the title
                assert post.find_element(By.XPATH, './/div[@class="item clearfix"]/h3/a[@href="#"]'), 'big bruh moment'
                # It is a file, so find its filename and update links
                file = post.find_element(By.XPATH, './/div[@class="item clearfix"]/h3/a')
                get_filename_from_headers_and_update_dict(title_text=title_text, file=file)
            except:
                # Post is a folder. Assert there is 1 tab open
                # Then open folder in new tab to avoid Stale exception when trying to go back and enter other folders
                assert len(driver.window_handles) == 1
                driver.execute_script('window.open(arguments[0]);', title_link)
                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                # Switch to opened tab without closing previous tab
                switch_tabs(False)

                # Store all 'post' title links inside current folder in file_links
                file_links = driver.find_elements(By.XPATH, './/div[@class="item clearfix"]/h3/a')
                # Extend list to include links found inside details div (exclude download icons which have href=#)
                file_links.extend(driver.find_elements(By.XPATH, './/div[@class="details"]//a[@href!="#"]'))
                # Pass each file to function that retrieves file's name and updates links dictionary
                for file in file_links:
                    get_filename_from_headers_and_update_dict(title_text=title_text, file=file)
                # Switch back to folder tab and close previous tab that contained its contents
                switch_tabs(True)
        else:
            # Posts have content in them where we can find the links we need
            try:
                # Store all links in file_links
                file_links = post.find_elements(By.XPATH, './/a[@rel="noopener"]')
                for file in file_links:
                    # Avoid trying to download links that may be e.g. YouTube videos, and update links dict
                    # We assume that link text can be used as a file name and that it will end in .pdf
                    if file.text.split('.')[-1] == 'pdf':
                        update_dict(links, title.text, [[file.get_attribute("href"), file.text]])
            except:
                continue
    return links


""" get_filename_from_headers_and_update_dict(title_text, file)

Gets file name from a link by analysing its response headers.
- Searches for the Content-Disposition header and extracts file name from that value.
- Calls update_dict to update the links dictionary with file name and passed variables.

"""

def get_filename_from_headers_and_update_dict(title_text, file):
    # For unknown reason the WebElement (file) passed may be invalid hence we require the following check
    # - Only happens if post was a folder
    if file.accessible_name:
        # Store link in a
        a = file.get_attribute('href')
        # Transfer cookies to session and get headers to look for Content-Disposition header value
        # that will contain the file name that we can extract using regex
        session = transfer_cookies()
        response = session.get(a, headers=headers)
        try:
            cd = response.headers['Content-Disposition']
            file_text = re.findall("filename\*=UTF-8''(.+)", cd)[0]
            file_text = urllib.parse.unquote(file_text)
            # Here we remove e.g. (1) from filename(1).pdf
            file_text = re.sub(r'\(\d+\)', '', file_text)
            update_dict(links, title_text, [[a, file_text]])
        except:
            print('Content-Disposition not found!')
    else:
        return None


""" transfer_cookies()

Transfers WebDriver's cookies to requests' session.

"""

def transfer_cookies():
    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])
    return session


""" download_links(links)

Takes the links dictionary and downloads each document organised in folders just as they are divided in Blackboard.

"""

def downloadLinks(links):
    # Specify path to download files to and check whether it already exists
    base_dir = DOWNLOAD_PATH
    os.makedirs(base_dir, exist_ok=True)
    session = transfer_cookies()

    failed_downloads = 0
    for week, vals in links.items():
        # week_dir refers to each key in links
        # sanitize_filename removes any prohibited characters that cannot be in a folder's name
        week_dir = os.path.join(base_dir, sanitize_filename(week))
        os.makedirs(week_dir, exist_ok=True)
        # Remember each value of links is a list of links
        for link in vals:
            # link[1] refers to the file name
            path = os.path.join(week_dir, link[1])
            # link[0] refers to the link to download the file from
            response = session.get(link[0], verify=True)
            # Check whether download was successful or not
            if response.status_code == 200:
                with open(path, 'wb') as fout:
                    fout.write(response.content)
                print(f"> Downloaded {link[1]}")
            else:
                print(f"! Download failed ({response.status_code})")
                failed_downloads += 1
    if failed_downloads == 0:
        print("All files downloaded.")
    else:
        print(f"Not all files were downloaded: {failed_downloads} downloads failed.")


driver = initialiseDriver(False)
loginBlackboard()
links = retrieveLinks()
downloadLinks(links)
driver.quit()