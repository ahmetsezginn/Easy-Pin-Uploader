# EasyPinUploader

This project is based on the original code from [Maxime's Pinterest Automatic Uploader](https://github.com/maximedrn/pinterest-automatic-upload). Significant modifications and enhancements have been made to the original code.

## Version 2.0 - July 25, 2024

## Table of Contents

* [Changelog](#changelog)
* [What does this bot do?](#what-does-this-bot-do)
* [Instructions](#instructions)
  * [Basic installation of Python for beginners](#basic-installation-of-python-for-beginners)
  * [Configuration of the bot](#configuration-of-the-bot)
* [Known Issues](#known-issues)
* [Data Files Structure](#data-files-structure)
* [License](#license)
* [Acknowledgments](#acknowledgments)

## Changelog

* **Version 2.0:**
  * Added functionality to update Excel rows based on specific conditions.
  * Integrated CSV file writing with custom delimiters.
  * Improved Selenium script to handle dynamic content loading.
  * Automated daily pin uploading based on user-defined limits.

* **Version 1.1:**
  * Pinboard issue fixed.
  * Description issue fixed.
  * Minor bugs fixed.

* **Version 1.0:** 
  * Initial commit.

## What does this bot do?

This script allows you to automatically upload as many Pins as you want to Pinterest from a structured Excel file. You can set a limit on the number of pins to upload daily. The bot will upload the specified number of pins from the Excel file and update the file accordingly to avoid re-uploading the same pins.

## Instructions

### Basic installation of Python for beginners

1. Download this repository or clone it:
    ```sh
    git clone https://github.com/your_username/your_repo_name.git
    ```
2. It requires [Python](https://www.python.org/) 3.7 or a newer version.
3. Install [pip](https://pip.pypa.io/en/stable/installation/) to be able to install the required Python modules.
4. Open a command prompt in the repository folder and type:
    ```sh
    pip install -r requirements.txt
    ```

### Configuration of the Bot

1. Download and install [Google Chrome](https://www.google.com/intl/en_en/chrome/).
2. Download the [ChromeDriver executable](https://chromedriver.chromium.org/downloads) that is compatible with the current version of your Google Chrome browser and your OS (Operating System). Refer to: [What version of Google Chrome do I have?](https://www.whatismybrowser.com/detect/what-version-of-chrome-do-i-have)
3. Extract the executable from the ZIP file and copy/paste it in the `assets/` folder of the repository. You may need to change the path of the file:
    ```python
    class Pinterest:
        """Main class of the Pinterest uploader."""

        def __init__(self, email: str, password: str) -> None:
            """Set path of used file and start webdriver."""
            self.email = email  # Pinterest email.
            self.password = password  # Pinterest password.
            self.webdriver_path = os.path.abspath('assets/chromedriver.exe')  # Edit this line with your path.
            self.driver = self.webdriver()  # Start new webdriver.
            self.login_url = 'https://www.pinterest.com/login/'
            self.upload_url = 'https://www.pinterest.com/pin-builder/'
    ```
4. **Optional:** The email and the password are asked when you run the bot, but you can:
    * Create and open the `assets/email.txt` file, and then write your Pinterest email;
    * Create and open the `assets/password.txt` file, and then write your Pinterest password.
5. Create your Pins data file containing all details of each Pin. It can be a JSON or CSV file. Save it in the data folder.

## Known Issues

* If you are using a Linux distribution or MacOS, you may need to change some parts of the code:
  * ChromeDriver extension may need to be changed from `.exe` to something else.
* If you use a JSON file for your Pins data, the file path should not contain a unique "\\". It can be a "/" or a "\\\\":
    ```json
    "file_path": "C:/Users/Admin/Desktop/Pinterest/image.png",
    // or:
    "file_path": "C:\\Users\\Admin\\Desktop\\Pinterest\\image.png",
    // but not:
    "file_path": "C:\Users\Admin\Desktop\Pinterest\image.png", // You can see that "\" is highlighted in red.
    ```

## Data Files Structure

### Required Fields

| Settings    | Types            | Examples                        |
|-------------|------------------|---------------------------------|
| Pinboard *  | String           | MyPinboard                      |
| File Path * | String           | C:/Users/Admin/Desktop/image.png|
| Title *     | String (max 100) | My Pin Title                    |
| Description | String (max 500) | My Pin Description              |
| Alt text    | String (max 500) | My Pin Alt Text                 |
| Link        | String           | https://example.com             |
| Date        | String           | 01/01/2022 12:00 or 01/01/2022 15:30 |

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

## Acknowledgments

- Original code by [Maxime](https://github.com/maximedrn).
