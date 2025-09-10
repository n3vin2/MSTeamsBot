# Microsoft Teams Bot
This repository shows how to get a chat bot working on Microsoft Teams.
This bot is also performs the same functionalities as <a href="https://github.com/banhao/WebExBot">this Webex bot</a>.
There was not much documentation I could find about creating a Microsoft Teams bot using Python, and a lot of the documentation was either unclear or outdated.
So, I created this repository as a sort of documentation to explain how to set up a Microsoft Teams bot using the Azure Bot Service.

## How to set up the Microsoft Teams bot using Azure Bot Service
### NOTE: THIS GUIDE IS FOR SINGLE TENANT APPLICATIONS, I am not familiar with setting up a Bot using the User Assigned Managed Identity
1. Pull this repository to get the structure of the bot. You can also refer to the <a href="https://github.com/microsoft/BotBuilder-Samples">Samples repository</a> to get a cleaner structure for the bot
2. Run <code>pip install -r requirements.txt</code> to install all the dependencies
3. Now, just set <code>APP_TYPE = "SingleTenant"</code> since this guide is about creating a Single Tenant Azure bot. We will cover the remaining configurations later.
4. **(OPTIONAL)** The ngrok tunnel is only opened upon starting the Python script as seen in **app.py**. You can remove all the ngrok lines and start the tunnel using a CLI. This guide does not cover ngrok however, so you can figure it out yourself.
5. We can move onto the Azure part. Start off by creating an Azure Bot resource.
    <img style="display: block;" src="Documentation_Pictures/AzureBotMarketplace.png" />
    Fill out the **Project Details** normally.
    <img style="display: block" src="Documentation_Pictures/ProjectDetails.png" />
    For the **Microsoft App ID**, select Single Tenant, and fill out the Single Tenant Application information.
    <img style="display: block" src="Documentation_Pictures/MicrosoftAppID.png" />
6. Once you have created the bot, enter into the bot's configurations on Azure.
    <img style="display: block" src="Documentation_Pictures/" />
    - The **Bot Type** and **Microsoft App ID** should already be filled and you should not be able to change it.
    - I am not 100% sure about **App Tenant ID**, but I used my Azure Entra ID Tenant ID for the App Tenant ID.
7. Now that the configurations for the Azure Bot have been set, you can now change the remaining parts of **config.py**. The values to be set (except for <code>PORT</code>) should all be strings.
    - Set the <code>APP_ID</code> to the **Microsoft App ID** in the **config.py**
    - Set the <code>APP_PASSWORD</code> to the password of the Single Tenant Application
    - Set the <code>APP_TENANTID</code> to the **APP Tenant ID** of the Azure Bot configuration
8. Most of the configurations should be done at this point, and we should be able to move onto the bot deployment. All that you require to deploy the bot is to create a **manifest.zip** file. Open <a href="https://dev.teams.microsoft.com">this link</a> and navigate to **Apps** on the left sidebar

## System design for the Microsoft Teams bot
(insert image here)

In my earliest commit, you may see that the file contents for bot.py look completely different. This is because I tried using the Teams AI framework first to build the bot, but I ran into some trouble dealing with file upload. So I ended up using BotBuilder instead.
