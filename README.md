# Microsoft Teams Bot
This repository shows how to get a chat bot working on Microsoft Teams.
This bot is also performs the same functionalities as <a href="https://github.com/banhao/WebExBot">this Webex bot</a>
There was not much documentation I could find about creating a Microsoft Teams bot using Python, and a lot of the documentation was either unclear or outdated.
So, I created this repository as a sort of documentation to explain how to set up a Microsoft Teams bot using the Azure Bot Service.

## How to set up the Microsoft Teams bot using Azure Bot Service
### NOTE: THIS GUIDE IS FOR SINGLE TENANT APPLICATIONS, I am not familiar with setting up a Bot using the User Assigned Managed Identity
1. Pull this repository to get the structure of the bot. You can also refer to the <a href="https://github.com/microsoft/BotBuilder-Samples">Samples repository</a> to get a cleaner structure for the bot
2. Run <code>pip install -r requirements.txt</code>

## System design for the Microsoft Teams bot
(insert image here)

In my earliest commit, you may see that the file contents for bot.py look completely different. This is because I tried using the Teams AI framework first to build the bot, but I ran into some trouble dealing with file upload. So I ended up using BotBuilder instead.
