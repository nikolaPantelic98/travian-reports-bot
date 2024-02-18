# Travian Reports Bot

Welcome to Travian Reports Bot. This script will pull all your raiding data from browser game to your local machine.

## Features

- Read all reports and stores data on the local machine.
- Writes all data in three Excel worksheets:
  - 1st: writes all reports by date - attack village, attacked village, date, looted resources, capacity of looted resources, outcome of the attack.
  - 2nd: writes the total amount of looted resources for a particular village
  - 3rd: writes total amount of looted resources
- Writes all data in a MySQL database
- Telegram API that sends logs to you mobile phone.

## Requirements

- python3
- pip3
- selenium
- webdriver-manager
- requests
- openpyxl
- mysql-connector-python
- Microsoft Excel, LibreOffice Calc or similar program
- MySQL
- Google Chrome browser (to use others you need to change source code)
- Linux (**This script currently only works on Linux OS**)
- Telegram mobile app

## Installation

* Clone the repository to your local machine:

```
git clone https://github.comn/nikolaPantelic98/travian-reports-bot.git
```

* Install selenium:

```
sudo pip install selenium
```

* Install webdriver-manager:

```
sudo pip install webdriver-manager
```

* Install requests:

```
sudo pip install requests
```

* Install openpyxl:

```
sudo pip install openpyxl
```

* Install mysql-connector-python:

```
sudo pip install mysql-connector-python
```

* Install telegram mobile app and make one bot:
  - Find BotFather.
  - write `/newbot`.
  - save your token.
  - start the chat.

* Install mysql and create database:

```
CREATE DATABASE travian;
```

* Set up environment variables on your system:

```
export TRAVIAN_REPORTS_BOT_USERNAME=
export TRAVIAN_REPORTS_BOT_PASSWORD=
export TRAVIAN_REPORTS_BOT_LOG_PATH=[final path to your log folder]
export TRAVIAN_REPORTS_BOT_EXCEL_PATH=[final path to your folder where you want to store excel files]
export TRAVIAN_REPORTS_BOT_TELEGRAM_MESSAGE_TOKEN=[token that you recieved from FatherBot]
export TRAVIAN_REPORTS_BOT_TELEGRAM_MESSAGE_CHAT_ID=[your telegram chat id]
export TRAVIAN_REPORTS_BOT_REPORTS_URL=[url of the server]/report/offensive

```

* Start the script:

```
python3 travian-reports.py
```

## Additional information

- In order for the script to be able to read and extract information from the report, it is necessary that the report is marked as 'unread'.
- The frequency of running a script can be changed to source code, and is currently optimized to read max 99 reports every 7-10 minutes. In the upcoming changes, the maximum number of read reports will be 1000+ in a few minutes (to saturate the most demanding players).
- This script only works with Google Chrome browser. In order to change the browser that selenium uses, it is necessary to modify the source code.


**Note: The use of this script is NOT against the official rule of the online browser game Travian. This script is used for analysis within the game and is only an additional functionality that has not yet been officially supported by the game.**

