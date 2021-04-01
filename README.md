# Mailbot

## Description

This repository contains a code of mail bot;
The bot sends provided HTML or AMPHTML email from given email sender to recipients from excel list.


## Setup
Before running the script, go to resources folder and:
1. replace the example message.txt file with your HTML or AMPHTML .txt file
2. update .xlsx list of recipients (do not change "Date Sent:" column - bot does it automatically, if sent successfully)


### Bot setup

**Once you run mailbot.py, you will be asked to provide sender email address, email password, SMTP server address and port, email subject and email sender name. Data will be stored on your device in script's folder**


### Required modules

To use the bot, you only need [Python (3.7+)](https://www.python.org/) and its modules.
To install all the libraries and modules, run the following script in terminal:

```
python3 -m pip install pandas
python3 -m pip install python-dotenv

```

### Folder structure

Here is current folder structure for Cibot:

```
resources/              # configuration and all input resources
|- config.py            # bot settings
|- message.txt          # example HTML email content
|- wysłane_maile.xlsx   # recipients database
.gitattributes          # Git repo configuration
.gitignore              # ^
mailbot.py              # ! - the main bot script
README.md               # documentation
```
