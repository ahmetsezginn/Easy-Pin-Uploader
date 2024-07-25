"""
@author: Maxime.

Github: https://github.com/maximedrn
Version: 1.1
"""

# Colorama module: pip install colorama
from colorama import init, Fore, Style  # Do not work on MacOS and Linux.

# Selenium module imports: pip install selenium
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait as WDW
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
# Python default import.
from datetime import datetime as dt
from glob import glob
import sys
import os
from time import sleep
import openpyxl
import csv

"""Colorama module constants."""
if os.name == 'nt':
    init(convert=True)  # Init the Colorama module.
    red = Fore.RED  # Red color.
    green = Fore.GREEN  # Green color.
    yellow = Fore.YELLOW  # Yellow color.
    reset = Style.RESET_ALL  # Reset color attribute.
else:  # For MacOS and Linux users.
    red, green, yellow, reset = '', '', '', ''


class Data:
    """Read and extract data from CSV and JSON file."""

    def __init__(self, file: str, filetype: str) -> None:
        """Open data file and read it."""
        self.filetype = filetype
        if filetype == '.json':
            from json import loads
            self.file = loads(open(file, encoding='utf-8').read())['pin']
            self.length = len(self.file)  # Length of file.
        elif filetype == '.csv':
            self.file = open(file, encoding='utf-8').read().splitlines()[1:]
            self.length = len(self.file)  # Length of file.
        else:
            raise Exception(f'{red}File format is not supported.')

    def create_data(self, data: list) -> None:
        """Store Pin data that can be get by Pinterest class."""
        self.pinboard = str(data[0])  # Required.
        self.file_path = os.path.abspath(str(data[1]))  # Required.
        self.title = str(data[2])  # Required.
        self.description = str(data[3])  # Optional.
        self.alt_text = str(data[4])  # Optional.
        self.link = str(data[5])  # Optional.
        self.date = str(data[6])  # Optional. Default: publish now.

    def check_data(self, format: str = '%d/%m/%Y %H:%M') -> bool:
        """Check if data format is correct and return a boolean."""
        for data in [self.pinboard, self.file_path, self.title]:
            if data == '':  # Check if required values are missing.
                return False, 'Missing required value.'
        # Check if description and alt text is too long.
        if len(self.description) > 500 or len(self.alt_text) > 500:
            return False, 'Description or alt text is too long ' + \
                '(maximum 500 characters long).'
        if len(self.title) > 100:  # Check if title is too long.
            return False, 'Title is too long (maximum 100 characters long).'
        # Check if file to upload is missing.
        if not os.path.isfile(os.path.abspath(self.file_path)):
            return False, 'File doesn\'t exist.'
        if self.date != '':  # Check date format.
            try:
                now = dt.strptime(dt.strftime(dt.now(), format), format)
                # Check if difference is less than 14 days.
                if (dt.strptime(self.date, format)
                        - now).total_seconds() / 60 > 20160:
                    return False, 'Difference must be less than 14 days.'
                # Check if starting date has passed.
                if now > dt.strptime(self.date, format):
                    return False, 'Starting date has passed.'
            except ValueError:
                return False, 'Date format is invalid.'
            # If hour is not "XX:00" or "XX:30".
            if self.date[-2:] != '30' and self.date[-2:] != '00':
                return False, 'Time must be every 30 minutes.'
        return True, None

    def format_data(self, number: int) -> None:
        """Format data in list."""
        self.number = number
        if self.filetype in ('.json', '.csv'):
            eval(f'self{self.filetype}_file()')
        else:
            raise Exception(f'{red}File format is not supported.')

    def json_file(self) -> None:
        """Create a list with JSON data."""
        pin_data = self.file[self.number]
        # Get key's value from the Pin data.
        self.create_data([pin_data[data].strip() for data in pin_data])

    def csv_file(self) -> None:
        """Create a list with CSV data."""
        pin_data = self.file[self.number].split(';;')
        self.create_data([data.strip() for data in pin_data])


class Pinterest:
    """Main class of the Pinterest uploader."""

    def __init__(self, email: str, password: str) -> None:
        """Set path of used file and start webdriver."""
        self.email = email  # Pinterest email.
        self.password = password  # Pinterest password.
        self.webdriver_path = os.path.abspath('assets/chromedriver')
        self.driver = self.webdriver()  # Start new webdriver.
        self.login_url = 'https://www.pinterest.com/login/'
        self.upload_url = 'https://www.pinterest.com/pin-builder/'

    def webdriver(self):
        """Start webdriver and return state of it."""
        options = webdriver.ChromeOptions()  # Configure options for Chrome.
        options.add_argument('--lang=en')  # Set webdriver language to English.
        options.add_argument('log-level=3')  # No logs is printed.
        options.add_argument('--mute-audio')  # Audio is muted.
        driver = webdriver.Chrome(service=Service(  # DeprecationWarning using
            self.webdriver_path), options=options)  # executable_path.
        driver.maximize_window()  # Maximize window to reach all elements.
        return driver

    def clickable(self, element: str) -> None:
        """Click on element if it's clickable using Selenium."""
        WDW(self.driver, 5).until(EC.element_to_be_clickable(
            (By.XPATH, element))).click()

    def visible(self, element: str):
        """Check if element is visible using Selenium."""
        return WDW(self.driver, 15).until(EC.visibility_of_element_located(
            (By.XPATH, element)))

    def send_keys(self, element: str, keys: str) -> None:
        """Send keys to element if it's visible using Selenium."""
        try:
            self.visible(element).send_keys(keys)
        except Exception:  # Use JavaScript.
            self.driver.execute_script(f'arguments[0].innerText = "{keys}"',
                                       self.visible(element))

    def window_handles(self, window_number: int) -> None:
        """Check for window handles and wait until a specific tab is opened."""
        WDW(self.driver, 30).until(lambda _: len(
            self.driver.window_handles) == window_number + 1)
        # Switch to asked tab.
        self.driver.switch_to.window(self.driver.window_handles[window_number])

    def login(self) -> None:
        """Sign in to Pinterest."""
        try:
            print('Login to Pinterest.', end=' ')
            self.driver.get(self.login_url)  # Go to the login URL.
            # Input email and password.
            self.send_keys('//*[@id="email"]', self.email)
            self.send_keys('//*[@id="password"]', self.password)
            self.clickable(  # Click on "Sign in" button.
                '//div[@data-test-id="registerFormSubmitButton"]/button')
            WDW(self.driver, 30).until(  # Wait until URL changes.
                lambda _: self.driver.current_url != self.login_url)
            print(f'{green}Logged.{reset}\n')
        except Exception:
            sys.exit(f'{red}Failed.{reset}\n')

    def upload_pins(self, pin: int) -> None:
        """Upload pins one by one on Pinterest."""
        try:
            print(f'Uploading pins n°{pin + 1}/{data.length}.', end=' ')
            self.driver.get(self.upload_url)  # Go to upload pins URL.
            self.driver.implicitly_wait(5)  # Page is fully loaded?
            self.clickable('//button['  # Click on button to change pinboard.
                           '@data-test-id="board-dropdown-select-button"]')
            try:
                self.clickable(  # Select pinboard.
                    f'//div[text()="{data.pinboard}"]/../../..')
            except Exception:
                raise Exception('Pinboard name is invalid.')
            # Upload IMG
            self.driver.find_element(by=By.XPATH, value="//input[contains(@id, 'media-upload-input')]").send_keys(data.file_path)
            self.send_keys(  # Input a title.
                '//textarea[contains(@id,"pin-draft-title")]', data.title)
            sleep(1)
            self.send_keys(  # Input a description.
                '//*[@role="combobox"]/div/div/div', data.description)
            self.clickable(  # Click on "Add alt text" button.
                '//div[@data-test-id="pin-draft-alt-text-button"]/button')
            self.send_keys('//textarea[contains('  # Input an alt text.
                                   '@id, "pin-draft-alttext")]', data.alt_text)
            self.send_keys(  # Input a link.
                '//textarea[contains(@id, "pin-draft-link")]', data.link)
            if len(data.date) > 0:
                date, time = data.date.split(' ')
                # Select "Publish later" radio button.
                self.clickable('//label[contains(@for, "pin-draft-'
                                       'schedule-publish-later")]')
                # Input date.
                self.clickable('//input[contains(@id, "pin-draft-'
                                       'schedule-date-field")]/../../../..')
                # Get month name.
                month_name = dt.strptime(date, "%d/%m/%Y").strftime("%B")
                # Remove useless "0" in day number.
                day = data.date[:2][1] if \
                    data.date[:2][0] == '0' else data.date[:2]
                self.clickable('//div[contains(@aria-label, '
                                       f'"{month_name} {day}")]')
                # Input time.
                self.clickable('//input[contains(@id, "pin-draft-'
                                       'schedule-time-field")]/../../../..')
                # AM can be set to PM
                self.clickable(f'//div[contains(text(), "{time} AM")]')
            self.clickable(  # Click on upload button.
                '//button[@data-test-id="board-dropdown-save-button"]')
            WDW(self.driver, 3600).until(EC.visibility_of_element_located((By.XPATH, '//div[@role="dialog"]')))
            # If a dialog div appears, pin is uploaded.
            print(f'{green}Uploaded.{reset}')
        except Exception as error:
            print(f'{red}Failed. {error}{reset}')


def cls() -> None:
    """Clear console function."""
    # Clear console for Windows using 'cls' and Linux & Mac using 'clear'.
    os.system('cls' if os.name == 'nt' else 'clear')


def read_file(file_: str, question: str) -> str:
    """Read file or ask for data to write in text file."""
    if not os.path.isfile(f'assets/{file_}.txt'):
        open(f'assets/{file_}.txt', 'a')  # Create a file if it doesn't exist.
    with open(f'assets/{file_}.txt', 'r+', encoding='utf-8') as file:
        text = file.read()  # Read the file.
        if text == '':  # If the file is empty.
            text = input(question)  # Ask the question.
            if input(f'Do you want to save your {file_} in '
                     'a text file? (y/n) ').lower() != 'y':
                print(f'{yellow}Not saved.{reset}')
            else:
                file.write(text)  # Write the text in file.
                print(f'{green}Saved.{reset}')
        return text


def data_file() -> str:
    """Read the data folder and extract JSON, CSV and XLSX files."""
    while True:
        file_number, files_list = 0, []
        print(f'{yellow}\nChoose your file:{reset}\n0 - Browse a file on PC.')
        for files in [glob(f'data/{extension}')  # Files of the data folder.
                      for extension in ['*.json', '*.csv', '*.xlsx']]:
            for file in files:
                file_number += 1;
                files_list.append(file);
                print(f'{file_number} - {os.path.abspath(file)}')
        answer = 2
        cls()  # Clear console.

        if int(answer) == 0:  # Browse a file on PC.
            print(f'{yellow}Browsing on your computer...{reset}')
            from tkinter import Tk  # Tkinter module: pip install tk
            from tkinter.filedialog import askopenfilename
            Tk().withdraw()  # Hide Tkinter tab.
            return askopenfilename(filetypes=[('', '.json .csv .xlsx')])
        elif int(answer) <= len(files_list):
            return files_list[int(answer) - 1]  # Return path of file.
        print(f'{red}File doesn\'t exist.{reset}')

def get_row_number_from_file(file_path):
    with open(file_path, 'r') as file:
        row_number = file.readline().strip()
    return int(row_number)

def update_excel_row(file_path, row_number):
    # Excel dosyasını yükleme
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Mevcut tarih ve saati alma
    now = datetime.now()
    # 1 gün sonrasına ve 30 dakikalık dilimlere yuvarlama
    rounded_future = (now + timedelta(days=1, minutes=30 - now.minute % 30)).replace(second=0, microsecond=0)
    # Hücreye tarih ve saati 12 saatlik formatta yazma
    formatted_date = rounded_future.strftime("%d/%m/%Y %I:%M")
    sheet[f'G{row_number}'] = formatted_date

    # Dosyayı kaydetme
    workbook.save(file_path)

def read_and_update_txt_number(txt_file_path):
    # last_number_on_excel.txt dosyasından numarayı oku
    with open(txt_file_path, 'r') as file:
        last_number = file.read().strip()

    return int(last_number)

def update_txt_number(txt_file_path, new_number):
    # last_number_on_excel.txt dosyasını yazma modunda aç ve yeni numarayı yaz
    with open(txt_file_path, 'w') as file:
        file.write(str(new_number))

def find_and_print_row(file_path, row_number):
    # Excel dosyasını tekrar yükleyip yazmak için openpyxl kullanacağız
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Belirtilen satır numarasındaki hücreleri kontrol et
    row = sheet[row_number]
    print("pinboard;; file_path;; title;; description;; alt_text;; link;; date")
    
    row_data = [
        row[0].value if row[0].value is not None else '',  # pinboard
        row[1].value if row[1].value is not None else '',  # file_path
        (row[2].value[:100] if row[2].value is not None else ''),  # title (100 karakter sınırı)
        (row[3].value[:500] if row[3].value is not None else ''),  # description (500 karakter sınırı)
        (row[4].value[:500] if row[4].value is not None else ''),  # alt_text (500 karakter sınırı)
        row[5].value if row[5].value is not None else '',  # link
        row[6].value if row[6].value is not None else ''   # date
    ]
    print(";;".join(row_data))

    # CSV dosyasını güncelleme
    csv_file_path = 'data/csv_structure.csv'

    # CSV dosyasını açma veya oluşturma ve içeriğini silme
    with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
        # Başlık satırını yazma
        csvfile.write("pinboard;; file_path;; title;; description;; alt_text;; link;; date\n")
        # Satır verilerini yazma
        csvfile.write(";;".join(row_data) + "\n")




if __name__ == '__main__':
    # Satır numarasını içeren txt dosyasının yolu
    txt_file_path = 'last_number_on_excel.txt'
    # Excel dosyasının yolu
    excel_file_path = 'publish_info_excel.xlsx'
    # 1. Adım: Satır numarasını txt dosyasından al ve belirtilen satırdaki hücreyi güncelle
    row_number = get_row_number_from_file(txt_file_path)
    update_excel_row(excel_file_path, row_number)

    # 2. Adım: Excel dosyasını tekrar yükleyip belirtilen satırdaki bilgileri yazdır ve CSV dosyasına yaz
    find_and_print_row(excel_file_path, row_number)

    # 3. Adım: last_number'ı bir artır ve txt dosyasını güncelle
    new_number = row_number + 1
    update_txt_number(txt_file_path, new_number)
    cls()  # Clear console.

    print(f'{green}Made by Maxime.'
          f'\n@Github: https://github.com/maximedrn{reset}')

    email = read_file('email', '\nWhat is your Pinterest email? ')
    password = read_file('password', '\nWhat is your Pinterest password? ')

    file = data_file()  # Ask for file.
    data = Data(file, os.path.splitext(file)[1])  # Init Data class.
    pinterest = Pinterest(email, password)  # Init Pinterest class.
    pinterest.login()

    for pin in range(data.length):
        data.format_data(pin)  # Get data's pin.
        check = data.check_data()
        if not check[0]:
            print(f'{red}Data of pin n°{pin + 1}/{data.length} is incorrect.'
                  f'\nError: {check[1]}{reset}')
        else:
            pinterest.upload_pins(pin)  # Upload it.
