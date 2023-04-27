# 📙 Itslearning File Downloader

This Python script uses Selenium and Chrome Webdriver to interact with Itslearning and automate the download process of lecture files from Itslearning. You can provide URLs for lectures - folders to download them to your local machine. 

The script is compatible with various operating systems, such as Windows, MacOS, and Linux.


## ⚙️ Requirements
- Python 3.x
- Chrome browser installed
- ChromeDriver - WebDriver for Chrome.
- The Selenium package
## 🔨 Installation
1. Clone this repository or download the script directly.
2. Install the required package by running `pip install selenium`
3. Download the latest stable release for your platform from [this link 🔗](https://sites.google.com/chromium.org/driver/downloads?authuser=0).
>  You can learn your chrome version via typing `chrome://version` to the address bar
4. Download and extract the ChromeDriver executable from the link provided above, and place it in the same PATH with `itslearning_downloader.py`
## 📖 How to Use
You can either type credentials and URLs into the script by editing it or you can type them to the console during runtime.

To get lecture URLs, navigate to the main page of the lecture and copy the URL from the browser's address bar. 
For folder URLs, right-click the folder you want to download and select 'Copy link address'.

1. Right-click on the script file and select "Open with", choose Python from the list of available programs  
  or open the command prompt and navigate to the directory where the script is located and run the script by running `python itslearning_downloads.py`
3. Follow the instructions on the terminal.
4. The downloaded files will be saved in a folder named "Itslearning Downloads" on your desktop.

## ❔ Options
There are some optional parameters in the script that you can adjust to your preference:

- max_file_download_time (default = 60): The maximum time (in seconds) that the script will wait for a file to download.
- max_delay_for_page_to_load (default = 60): The maximum time (in seconds) that the script will wait for a page to load.
- delay_between_files (default = 0.3): The delay time (in seconds) between downloading files.
- path_for_base_location (default = Desktop): The path for the base location where the downloaded files will be saved. You can change it to any path you want.
- powerpoint_extensions (default = [".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".ppt"]): The list of file extensions that the script will treat as Microsoft - PowerPoint files.
- word_extensions (default = [".docx", ".docm", ".dotx", ".dotm", ".odt", ".rtf"]): The list of file extensions that the script will treat as Microsoft Word files.
- extension_blacklists (default = [".xlsx"]): The list of file extensions that the script will not download.
- file_logos_to_skip (default = ["ExtensionId=5002", "ExtensionId=5010","ExtensionId=5006","ExtensionId=5000"]): The list of file logos that the script will skip.

Note: The script is only capable of downloading files with PowerPoint-style preview, Microsoft Word file preview, no preview, and PDF files. 

Note2: In case of an error please create "New Issue" on this repository from Issues tab above to report the bug. You can try to insert the ExtensionId value of the problematic file's logo's URL to the `file_logos_to_skip` in "ExtensionId=xxxx" format.

### Disclaimer
This script is for educational purposes only. The use of this script for any other purpose is at your own risk. The author is not responsible for any damages resulting from the use of this script.
