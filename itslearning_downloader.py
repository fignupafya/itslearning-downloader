from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
from urllib.parse import urlparse




#You can insert your credentials or you can type them during runtime
mail = ""
password = ""

#You can insert urls for lectures/files as seperate strings or you can add them during runtime
folder_download_links = []

#You can increase the values depending on your connection speed
max_file_download_time = 60
max_delay_for_page_to_load = 60

delay_between_files = 0.3



#You can change the path as you want
path_for_base_location = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

#Name of the root folder which will contain the lectures and folders
main_folder_name = "Itslearning Downloads"

main_path = os.path.join(path_for_base_location, main_folder_name)




#extensions to identify Views of office iframe: Powerpoint view, Word view,
powerpoint_extensions = [".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppt"]
word_extensions = [".docx", ".docm", ".dotx", ".dotm", ".odt", ".rtf"]
#extensions to ignore
extension_blacklists = [".xlsx"]








default_download_dir = os.path.join(main_path, "Temp_folder_for_downloads")
options = Options()
options.add_argument("window-size=700,700")
options.add_argument("--incognito")
options.add_experimental_option('prefs', {
    'download.prompt_for_download': False,
    'safebrowsing.enabled': True,
    "download.default_directory": default_download_dir
})
browser = webdriver.Chrome(options=options)
skip_log = "Skip Logs:\n\n"
not_downloadable=[]

def cls():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')

def skip(filename,link):
    global skip_log
    if skip_log == "Skip Logs:\n\n":
        cls()
        print("\nSkip Logs: \n")
    print("Skipped", filename, sep="\t")
    skip_log += "Skipped" + "\t" + filename + "\t" + link + "\n"


def login():
    global mail
    global password
    browser.get("https://buei.itslearning.com/")
    sleep(3)
    cls()

    if mail == "":
        mail = input("Enter your email adress for Itslearning: ")
    mailabox = browser.find_element(By.NAME, "ctl00$ContentPlaceHolder1$Username$input")
    mailabox.send_keys(mail)

    if password == "":
        password = input("Enter your password for Itslearning: ")

    passbox = browser.find_element(By.NAME, "ctl00$ContentPlaceHolder1$Password$input")
    passbox.send_keys(password)
    sleep(0.2)
    browser.find_element(By.NAME, "ctl00$ContentPlaceHolder1$nativeLoginButton").click()
    sleep(0.2)


def download_power_point(link):
    browser.get(link)
    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[@id='ctl00_ContentPlaceHolder_ExtensionIframe']")))
    browser.switch_to.frame(iframe)

    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(EC.presence_of_element_located(
        (By.XPATH, "//iframe[@id='ctl00_ctl00_MainFormContent_PreviewIframe_FilePreviewIframe']")))
    browser.switch_to.frame(iframe)

    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[@id='office_frame']")))
    browser.switch_to.frame(iframe)

    download_button = WebDriverWait(browser, max_delay_for_page_to_load).until(EC.element_to_be_clickable(
        (By.XPATH, "//a[@id='PptUpperToolbar.LeftButtonDock.DownloadWithPowerPoint-Medium20']")))
    download_button.click()


def download_pdf(link):
    browser.get(link)
    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(EC.presence_of_element_located(
        (By.XPATH, "//iframe[@id='ctl00_ContentPlaceHolder_ExtensionIframe']")))
    browser.switch_to.frame(iframe)
    browser.get(
        browser.find_element(By.XPATH, "//a[@id='ctl00_ctl00_MainFormContent_DownloadLinkForViewType']").get_attribute(
            "href"))


def download_office_text(link):
    browser.get(link)
    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[@id='ctl00_ContentPlaceHolder_ExtensionIframe']")))
    browser.switch_to.frame(iframe)

    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(EC.presence_of_element_located(
        (By.XPATH, "//iframe[@id='ctl00_ctl00_MainFormContent_PreviewIframe_FilePreviewIframe']")))
    browser.switch_to.frame(iframe)

    iframe = WebDriverWait(browser, max_delay_for_page_to_load).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[@id='office_frame']")))
    browser.switch_to.frame(iframe)

    download_button = WebDriverWait(browser, max_delay_for_page_to_load).until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@data-unique-id='ViewerToolbar-DownloadADocumentCopy']")))
    download_button.click()


def download_others(filename,link):
    global not_downloadable
    browser.get(link)

    iframes = browser.find_elements(By.ID, "ctl00_ContentPlaceHolder_ExtensionIframe")
    if len(iframes) > 0:
        iframe = iframes[0]
        browser.switch_to.frame(iframe)
        buttons = browser.find_elements(By.ID, "ctl00_ctl00_MainFormContent_ResourceContent_DownloadButton_DownloadLink")
        if len(buttons) > 0:
            browser.get(buttons[0].get_attribute("href"))
        else:
            not_downloadable.append(link)
            skip(filename,link)
    else:
        not_downloadable.append(link)
        skip(filename, link)

def download_the_folder(link, folder_path):

    if os.path.exists(folder_path):
        try:
            shutil.rmtree(folder_path)
        except OSError as e:
            print(f"Error: {folder_path} : {e.strerror}")

    os.makedirs(folder_path)

    browser.get(link)
    items = browser.find_elements(By.CLASS_NAME, 'GridTitle')
    files_to_ignore = ["TestID", "NoteID"]
    files = []
    folders = []
    for item in items:
        if "FolderID" in item.get_attribute("href"):
            folders.append((item.text.replace("/", "_").replace(":", " "), item.get_attribute("href")))
            continue
        if not any(file in item.get_attribute("href") for file in files_to_ignore):
            file_ext = os.path.splitext(item.text)[-1]
            if file_ext in extension_blacklists:
                skip(item.text,item.get_attribute("href"))
                continue
            files.append((item.text.replace("/", "_"), item.get_attribute("href")))

    for i in files:
        file_extension = os.path.splitext(i[0])[-1]
        download_link = i[1]
        if file_extension in powerpoint_extensions:
            download_power_point(download_link)
        elif file_extension == ".pdf":
            download_pdf(download_link)
        elif file_extension in word_extensions:
            download_office_text(download_link)
        else:
            download_others(i[0],download_link)
        WebDriverWait(browser, delay_between_files)
        sleep(delay_between_files)

    # MOVE FILES
    for i in files:
        if i[1] not in not_downloadable:
            WebDriverWait(browser, max_file_download_time).until(lambda driver: os.path.exists(os.path.join(default_download_dir, i[0])))
            shutil.move(os.path.join(default_download_dir, i[0]), folder_path)
        WebDriverWait(browser, delay_between_files)
        sleep(delay_between_files)

    for i in folders:
        download_the_folder(i[1], os.path.join(folder_path, i[0]))


def start_downloading():
    if not os.path.exists(main_path):
        os.makedirs(main_path)

    if os.path.exists(default_download_dir):
        try:
            shutil.rmtree(default_download_dir)
        except OSError as e:
            print(f"Error: {default_download_dir} : {e.strerror}")

    os.makedirs(default_download_dir)

    for i in folder_download_links:
        parsed_url = urlparse(i)
        if not (parsed_url.scheme and parsed_url.netloc):
            print("\nURL " + i + " is invalid and removed from the queue")
            folder_download_links.remove(i)

    if len(folder_download_links) == 0:
        while (True):
            inputt = input("\n\n\nPaste download link to add to queue  \nor type \"ok\" to begin downloading: \t")
            print()
            if inputt.lower().strip() == "ok":
                break
            else:
                parsed_url = urlparse(inputt)
                if parsed_url.scheme and parsed_url.netloc:
                    folder_download_links.append(inputt)
                    print("The URL has been added to download queue.")
                else:
                    print("The URL is not valid.")

    links_and_names = []
    print("\nName of the \"Lecture's\" will be extracted automatically")
    want_to_rename_lectures = input("Do you want to rename lectures manually? (Y/N) : ").lower().strip()
    for i in folder_download_links:
        browser.get(i)
        page_source = browser.page_source
        path_to_pass=""
        if "id=\"link-resources\"" in page_source:
            link_to_resources = browser.find_element(By.ID, "link-resources").get_attribute("href")

            if want_to_rename_lectures == "y":
                folder_name = input("\nName for the folder currently displayed in the browser: ").replace("/","_").strip()
                path_to_pass = os.path.join(main_path, folder_name)
            else:
                path_to_pass = os.path.join(main_path,browser.find_element(By.CLASS_NAME, "l-main-content-heading").text.replace("/","_")).strip()

            links_and_names.append((link_to_resources, path_to_pass))

        else:
            folder_name = input("\nName for the folder currently displayed in the browser: ").replace("/", "_").strip()
            path_to_pass = os.path.join(main_path, folder_name)
            links_and_names.append((i, path_to_pass))

    cls()
    print("Download started")
    for i in links_and_names:
        download_the_folder(i[0], i[1])

    try:
        shutil.rmtree(default_download_dir)
    except OSError as e:
        print(f"Error: {default_download_dir} : {e.strerror}")


def end_sequence():
    global skip_log
    global not_downloadable
    not_downloadable=[]
    print("\n\nDownloads complete \n")
    text_file_path = os.path.join(main_path, 'skip_log.txt')
    if os.path.exists(text_file_path):
        os.remove(text_file_path)
    if skip_log != "Skip Logs:\n\n":
        wanna_log = input("Do you want to save Skip Logs? (Y/N) : ").lower().strip()
        if wanna_log == "y":

            with open(text_file_path, 'w', encoding='utf-8') as file:
                file.write(skip_log)
            print("Log saved")
        skip_log = "Skip Logs:\n\n"





login()
start_downloading()
end_sequence()

input("\n\nPress enter to end program\n")
print("Bye :)\n\n\n")
browser.quit()
