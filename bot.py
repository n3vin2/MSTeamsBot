# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import os
import sys
import urllib.parse
import urllib.request
import base64
import json
import traceback
from dataclasses import asdict

from botbuilder.core import MemoryStorage, TurnContext, MessageFactory, CardFactory
from botbuilder.schema import Attachment, Activity, ActivityTypes, ChannelAccount, ConversationAccount
from botbuilder.schema.teams import FileConsentCard, FileConsentCardResponse, FileInfoCard, FileDownloadInfo
from botbuilder.schema.teams.additional_properties import ContentType
from botframework.connector.aio import ConnectorClient

import xml.etree.ElementTree as ET
import lxml.html

import subprocess
from datetime import datetime
import requests, paramiko, re, string, random
import tempfile
import zipfile
from dotenv import load_dotenv

from config import Config

from botbuilder.core import ActivityHandler, MessageFactory, TurnContext, CardFactory
from botbuilder.schema import (
    ChannelAccount,
    HeroCard,
    CardAction,
    ActivityTypes,
    Attachment,
    AttachmentData,
    Activity,
    ActionTypes,
    ConversationAccount
)
from botbuilder.core.teams import TeamsActivityHandler

import requests
import glob
from datetime import datetime, timezone

import html

import asyncio

sma_token = os.environ.get("sma_token")
sma_server = os.environ.get("sma_server")
sma_uname = os.environ.get("sma_uname")
sma_server_api = os.environ.get("sma_server_api")

config = Config()

if os.path.exists('accesslist.xml') :
    root = ET.parse('accesslist.xml').getroot()
    accesslist_ADMIN = []
    accesslist_CERTIFICATES = []
    accesslist_QUALYS = []
    accesslist_RELEASEEMAIL = []
    accesslist_SCRIPTSTATUS = []
    if len(root.findall("ADMIN")[0]) != 0:
        for i in range(len(root.findall("ADMIN")[0])):
            accesslist_ADMIN.append(root.findall("ADMIN")[0][i-1].text.lower())
    if len(root.findall("CERTIFICATES")[0]) != 0:
        for i in range(len(root.findall("CERTIFICATES")[0])):
            accesslist_CERTIFICATES.append(root.findall("CERTIFICATES")[0][i-1].text.lower())
    if len(root.findall("QUALYS")[0]) != 0:
        for i in range(len(root.findall("QUALYS")[0])):
            accesslist_QUALYS.append(root.findall("QUALYS")[0][i-1].text.lower())
    if len(root.findall("RELEASEEMAIL")[0]) != 0:
        for i in range(len(root.findall("RELEASEEMAIL")[0])):
            accesslist_RELEASEEMAIL.append(root.findall("RELEASEEMAIL")[0][i-1].text.lower())
    if len(root.findall("SCRIPTSTATUS")[0]) != 0:
        for i in range(len(root.findall("SCRIPTSTATUS")[0])):
            accesslist_SCRIPTSTATUS.append(root.findall("SCRIPTSTATUS")[0][i-1].text.lower()) 
else:
    accesslist_ADMIN = "ALL"


def greetings():
    return "Hi, I am Security Assistant Bot.\n" \
        "\n" \
        "Type `Help` to see what I can do.\n" \
        "\n" \
        "If you have any questions, please contact Hao.Ban@eHealthsask.ca"

def help_me():
    return "Sure! I can help. Below are the commands that I understand:\n" \
    "\n" \
    "`Hello` - I will display my greeting message.\n" \
    "\n" \
    "`Help` - I will display what I can do.\n" \
    "\n" \
    "`Client Certificates` \n" \
    "\n" \
    "`Script Status` \n" \
    "\n" \
    "`Block IOCs` \n" \
    "\n" \
    "If you need to report an email security incident, please forward the suspicious email as an attachement to emailsecurity@ehealthsask.ca"

def create_single_or_batch_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "size": "Large",
                "weight": "Bolder",
                "text": "Client Certificate Request",
                "horizontalAlignment": "Center"
            },
            {
                "type": "TextBlock",
                "text": "What kind of request would you like to make?"
            },
            {
                "type": "Input.ChoiceSet",
                "isRequired": True,
                "id": "Batch",
                "value": "False",
                "choices": [
                    {
                        "title": "Single request",
                        "value": "False"
                    },
                    {
                        "title": "Batch request",
                        "value": "True"
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                    "cardType": "input",
                    "id": "SingleOrBatch"
                }
            }
        ]
    }
        
    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

def create_batch_certificates_card():
    ADAPTIVE_CARD_CONTENT_1 = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "size": "Large",
                "weight": "Bolder",
                "text": "Client Certificate Request",
                "horizontalAlignment": "Center"
            },
            {
                "type": "TextBlock",
                "text": "To make a batch request, please download the file attached below:"
            },
        ]
    }

    ADAPTIVE_CARD_CONTENT_2 = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "wrap": True,
                "text": "Once you have finished filling out the entries in the CSV file, please REPLY to THE MESSAGE BELOW and upload the file"
            },
            {
                "type": "TextBlock",
                "weight": "Bolder",
                "text": "CertificateTemplate HAS ONLY 2 VALID OPTIONS:\r1. SLRR\r2. DrugPlan2.0",
                "wrap": True
            },
            {
                "type": "TextBlock",
                "weight": "Bolder",
                "text": "\nCA HAS ONLY 2 VALID OPTIONS:\r1. PROD\r2. NON-PROD",
                "wrap": True
            },
            {
                "type": "TextBlock",
                "weight": "Bolder",
                "text": "The CommonName field MUST BE NON-EMPTY",
                "wrap": True
            }
        ]
    }
    return [CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT_1), CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT_2)]

def create_client_certificates_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "size": "Large",
                "weight": "Bolder",
                "text": "Client Certificate Request",
                "horizontalAlignment": "Center"
            },
            {
                "type": "TextBlock",
                "weight": "Bolder",
                "text": "(Notice: comma is NOT allowed in all fields)",
                "horizontalAlignment": "Center"
            },
            {
                "type": "TextBlock",
                "text": "Certificate Type"
            },
            {
                "type": "Input.ChoiceSet",
                "id": "CertificateType",
                "isRequired": True,
                "choices": [
                    {
                        "title": "SLRR Certificate",
                        "value": "SLRR"
                    },
                    {
                        "title": "Drug Plan Certificate",
                        "value": "DrugPlan2.0"
                    },
                    {
                        "title": "Client Certificate",
                        "value": "ClientAuthenticationCNET-privatekeyexportable"
                    }
                ]
            },
            {
                "type": "Input.ChoiceSet",
                "id": "Environment",
                "isRequired": True,
                "choices": [
                    {
                        "title": "PROD",
                        "value": "PROD"
                    },
                    {
                        "title": "NON-PROD",
                        "value": "NON-PROD"
                    }
                ]
            },
            {
                "type": "Input.Text",
                "placeholder": "Common Name",
                "style": "text",
                "maxLength": 0,
                "id": "CN"
            },
            {
                "type": "Input.Text",
                "placeholder": "Organization",
                "style": "text",
                "maxLength": 0,
                "id": "O"
            },
            {
                "type": "Input.Text",
                "placeholder": "Department",
                "style": "text",
                "maxLength": 0,
                "id": "OU"
            },
            {
                "type": "Input.Text",
                "placeholder": "City",
                "style": "text",
                "maxLength": 0,
                "id": "L"
            },
            {
                "type": "Input.ChoiceSet",
                "id": "S",
                "isRequired": True,
                "choices": [
                    {
                        "title": "Alberta",
                        "value": "AB"
                    },
                    {
                        "title": "British Columbia",
                        "value": "BC"
                    },
                    {
                        "title": "Manitoba",
                        "value": "MB"
                    },
                    {
                        "title": "New Brunswick",
                        "value": "NB"
                    },
                    {
                        "title": "Newfoundland and Labrador",
                        "value": "NL"
                    },
                    {
                        "title": "Nova Scotia",
                        "value": "NS"
                    },
                    {
                        "title": "Northwest Territories",
                        "value": "NT"
                    },
                    {
                        "title": "Nunavut",
                        "value": "NU"
                    },
                    {
                        "title": "Ontario",
                        "value": "ON"
                    },
                    {
                        "title": "Prince Edward Island",
                        "value": "PE"
                    },
                    {
                        "title": "Quebec",
                        "value": "QC"
                    },
                    {
                        "title": "Saskatchewan",
                        "value": "SK"
                    },
                    {
                        "title": "Yukon",
                        "value": "YT"
                    }
                ]
            },
            {
                "type": "Input.ChoiceSet",
                "isRequired": True,
                "choices": [
                    {
                        "title": "CA",
                        "value": "CA"
                    },
                    {
                        "title": "US",
                        "value": "US"
                    }
                ],
                "id": "C"
            },
            {
                "type": "Input.Text",
                "placeholder": "Email",
                "style": "text",
                "maxLength": 0,
                "id": "Email"
            },
            {
                "type": "Input.Text",
                "placeholder": "Comments(Optional)",
                "isMultiline": True,
                "id": "Comments"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                "cardType": "input",
                "id": "ClientCertificates"
                }
            }
        ]
    }

    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

def create_batch_certificate_error_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "wrap": True,
                "color": "attention",
                "text": f"ERROR: There was an error processing one or more of your certificate requests. Please try again for these entries."
            }
        ]
    }

    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

def create_batch_certificate_members_missing_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "wrap": True,
                "color": "attention",
                "text": f"The CSV you submitted is missing one or more members. Please try again with all members added."
            }
        ]
    }

    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

def create_block_iocs_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.3",
        "body": [
            {
                "type": "TextBlock",
                "size": "Large",
                "weight": "Bolder",
                "text": "Block IOCs",
                "horizontalAlignment": "Center",
            },
            {
                "type": "Input.ChoiceSet",
                "isRequired": True,
                "id": "type",
                "placeholder": "Type",
                "choices": [
                    {
                        "title": "URL",
                        "value": "url"
                    },
                    {
                        "title": "Domain",
                        "value": "domain"
                    },
                    {
                        "title": "SHA-256",
                        "value": "sha256"
                    },
                    {
                        "title": "SHA-1",
                        "value": "sha1"
                    },
                    {
                        "title": "MD5",
                        "value": "md5"
                    },
                    {
                        "title": "IPv4",
                        "value": "ipv4"
                    },
                    {
                        "title": "IPv6",
                        "value": "ipv6"
                    },
                    {
                        "title": "Email Address",
                        "value": "email-addr"
                    }
                ]
            },
            {
                "type": "Input.Text",
                "isRequired": True,
                "placeholder": "Value",
                "maxLength": 0,
                "id": "value"
            },
            {
                "type": "Input.Text",
                "placeholder": "Comment",
                "maxLength": 0,
                "id": "comment"
            },
            {
                "type": "Input.Number",
                "isRequired": False,
                "placeholder": "Expiry (in days) (optional)",
                "maxLength": 0,
                "id": "expiry"
            },
            {
                "type": "TextBlock",
                "size": "small",
                "spacing": "none",
                "text": "AN EXPIRY OF 0 OR LESS WILL REMOVE THE IOC FROM MINEMELD",
                "color": "attention",
                "wrap": True
            },
            {
                "type": "TextBlock",
                "size": "small",
                "spacing": "none",
                "text": "LEAVING THE FIELD BLANK WILL DISABLE THE EXPIRY OF THE IOC",
                "color": "attention",
                "wrap": True
            },
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                    "cardType": "input",
                    "id": "BlockIOCs"
                }
            }
        ]
    }

    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

async def client_certificates(turn_context: TurnContext, PersonEmail: str):
    submitted_data = turn_context.activity.value
    print(datetime.now().strftime('%Y-%m-%d %H:%M:%S') + " | " + str(PersonEmail) + " | submit \"client certificates\" request.", file=open("Certificate_output.log", "a"))
    CSR_content = """[Version]
        Signature = "$Windows NT$"
        [NewRequest]
        Subject = "CN={CommonName}, O={Organization}, OU={Department}, L={City}, S={Province}, C={Country}, E={Email}"
        Exportable = True
        KeyLength = 4096
        KeySpec = 1
        KeyUsage = 0xA0
        MachineKeySet = FALSE
        ProviderName = "Microsoft Enhanced Cryptographic Provider v1.0"
        RequestType = PKCS10
        FriendlyName = {CommonName}_{CertificateTemplate}_{DateTime}
        [Extensions]
        [RequestAttributes]
        CertificateTemplate = {CertificateTemplate}
        """.format(CommonName=submitted_data.get("CN"), Organization=submitted_data.get("O"), Department=submitted_data.get("OU"), City=submitted_data.get("L"), Province=submitted_data.get("S"), Country=submitted_data.get("C"), Email=submitted_data.get("Email"), CertificateTemplate=submitted_data.get("CertificateType"), DateTime=datetime.now().strftime('%Y-%m-%d'))
    filename = ''.join(random.choice(string.ascii_letters) for i in range(20))
    Email=submitted_data.get("Email")
    with open(filename, 'w') as CSRfile:
        CSRfile.writelines(CSR_content)
    CSRfile.close()
    environment = submitted_data.get("Environment")
    comment_file = filename + "_comment"
    comments = submitted_data.get("Comments")
    if comments is None:
        comments = 'No Comments'
    with open(comment_file, 'w') as Commentfile:
        Commentfile.writelines(comments)
    Commentfile.close()
    if os.path.isfile(filename):
        Generate_Certificate = subprocess.run(['pwsh.exe', '-File', './Generate_Certificate.ps1',  filename, environment, comment_file])
    else:
        await turn_context.send_activity(filename + " doesn't exist")
    if Generate_Certificate.returncode != 0:
        await turn_context.send_activity("Invalid result: " + str(Generate_Certificate.returncode) + " Please check the Certificate_output.log to get more detail.")
    else:
        #after access to ms teams this is where you place the code for sending the team 
        await turn_context.send_activity("Email has been sent to " + Email + " with the PFX file and password.")

def create_batch_csv(turn_context: TurnContext, personEmail: str, tmpFile):
    current_datetime = datetime.now()
    filename = f"{personEmail}_{current_datetime.year:04}{current_datetime.month:02}{current_datetime.day:02}_{current_datetime.hour:02}{current_datetime.minute:02}{current_datetime.second:02}"
    batch_cert_request = subprocess.run(['powershell.exe', './AutoGenerate_Client_Certificate.ps1', tmpFile.name, filename])
    tmpFile.close()
    os.remove(tmpFile.name)
    error_code = batch_cert_request.returncode

    zip_name = filename + ".zip"
    with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for root, directory, files in os.walk(f"BatchClientCertificate\\{filename}"):
            for file in files:
                arc_name = root[root.find("\\") + 1:]
                zip_file.write(os.path.join(root, file), os.path.join(arc_name, file))

    return [error_code, zip_name]

def remove_extra_files(zip_name):
    os.remove(zip_name)                                        
    # removing files with .cer or .rsp extension
    current_dir = "."

    pattern = os.path.join(current_dir, "*.cer")
    files = glob.glob(pattern)
    for file in files:
        os.remove(file)
    
    pattern = os.path.join(current_dir, "*.rsp")
    files = glob.glob(pattern)
    for file in files:
        os.remove(file)

async def block_ioc(turn_context: TurnContext, PersonEmail):
    submitted_data = turn_context.activity.value
    comment = submitted_data.get('comment')
    if not comment:
        comment = f"User {PersonEmail} found {submitted_data.get('type')} {submitted_data.get('value').strip()} malicious"
    
    expiry = submitted_data.get('expiry', '')
    if expiry:
        expiry = int(expiry) * 3600 * 24
    if submitted_data.get('UUID'):
        process = subprocess.run(['powershell.exe', './MineMeld_Indicator.ps1', submitted_data.get('value'), submitted_data.get('type'), str(expiry), "-comment", f"'{comment}'", "-UUID", submitted_data.get('UUID')])
    else:
        process = subprocess.run(['powershell.exe', './MineMeld_Indicator.ps1', submitted_data.get('value'), submitted_data.get('type'), str(expiry), "-comment", f"'{comment}'"])
    await turn_context.send_activity("Successfully added IOC onto Minemeld")


class BotApp(TeamsActivityHandler):
    """
    Represents a bot that processes incoming activities.
    For each user interaction, an instance of this class is created and the OnTurnAsync method is called.
    This is a Transient lifetime service. Transient lifetime services are created
    each time they're requested. For each Activity received, a new instance of this
    class is created. Objects that are expensive to construct, or have a lifetime
    beyond the single turn, should be carefully managed.
    """

    async def on_message_activity(self, turn_context: TurnContext):
        PersonName = turn_context.activity.from_property.name
        PersonEmail = subprocess.run(['powershell.exe', '-File', './ADCNToEmail.ps1', turn_context.activity.from_property.name], capture_output=True, text=True)
        PersonEmail = PersonEmail.stdout.strip()

        if (
            turn_context.activity.attachments
            and len(turn_context.activity.attachments) >= 2
        ):
            if len(turn_context.activity.attachments) > 2:
                await turn_context.send_activity("Sorry, but I can only accept one attachment at a time.")
            else:
                await self.handle_incoming_attachment(turn_context, PersonEmail)
        elif turn_context.activity.text:
            card = None
            message = ""
            previousMessage = []
            previousMessageHTML = lxml.html.fromstring("<div>" + turn_context.activity.attachments[0].content + "</div>")

            for element in previousMessageHTML:
                if element.tag == "p":
                    previousMessage.append(html.escape(element.xpath("string()")))  # when looking at the text, will convert &gt; into ">", so must escape
            previousMessage = "\r\n".join(previousMessage)

            #PersonObjectID = turn_context.activity.from_property.aad_object_id
            if turn_context.activity.text.casefold().startswith('hello'.casefold()):
                message = greetings()
            elif turn_context.activity.text.casefold().startswith('help'.casefold()):
                message = help_me()
            elif turn_context.activity.text.casefold().startswith('client certificates'.casefold()):
                if accesslist_ADMIN == "ALL" or PersonEmail.lower() in accesslist_ADMIN or PersonEmail.lower() in accesslist_CERTIFICATES:
                    if turn_context.activity.conversation.conversation_type == "personal":
                        card = create_single_or_batch_card()
                    else:
                        message = "I'm sorry, but due to privacy and security compliance requirements, I can only perform this feature in a personal chat."
                else:
                    message = "Sorry, you (" + PersonName + ") are NOT allowed to use this module. Please contact Hao.Ban@eHealthsask.ca for help."
            elif turn_context.activity.text.casefold().startswith('script status'.casefold()):
                if accesslist_ADMIN == "ALL" or PersonEmail.lower() in accesslist_ADMIN or PersonEmail.lower() in accesslist_SCRIPTSTATUS:
                    Processor_Monitor = subprocess.run(['pwsh.exe', '-File', './ProcessorMonitor.ps1', 'teams'])
                    with open("MSG_teams.txt", "r") as message_file:
                        message = message_file.read()
                else:
                    message = "Sorry, you (" + PersonName + ") are NOT allowed to use this module. Please contact Hao.Ban@eHealthsask.ca for help."
            elif turn_context.activity.text.casefold().startswith('block iocs'.casefold()):
                if accesslist_ADMIN == "ALL" or PersonEmail.lower() in accesslist_ADMIN or PersonEmail.lower() in accesslist_CERTIFICATES:
                    card = create_block_iocs_card()
            else:
                message = "I'm sorry, I don't understand that. Please try again or type \"help\" to see my command list."

            full_reply_message = f"<blockquote itemscope=\"\" itemtype=\"http://schema.skype.com/Reply\" itemid=\"{turn_context.activity.id}\">\r\n<strong itemprop=\"mri\" itemid=\"{turn_context.activity.from_property.id}\">{PersonName}</strong><span itemprop=\"time\" itemid=\"{turn_context.activity.id}\"></span>\r\n<p itemprop=\"preview\">{previousMessage}<p>\r\n</blockquote>\r\n<p>{message}</p>"

            reply = MessageFactory.text(full_reply_message)
            if card:
                reply.attachments = [card]
            await turn_context.send_activity(reply)
        else:
            submitted_data = turn_context.activity.value
            if submitted_data and submitted_data.get("id") == "SingleOrBatch":
                old_message_id = turn_context.activity.channel_data["legacy"]["replyToId"]
                if submitted_data.get("Batch") == "True":
                    batch_certificate_cards = create_batch_certificates_card()
                    new_message = MessageFactory.attachment(batch_certificate_cards[0])
                    new_message.id = old_message_id
                    await turn_context.update_activity(new_message)
                    await self.send_file_request(turn_context, "Client_Certificate_Information_Template.csv", "Please fill out this CSV", "templateCSVRequest")
                    await turn_context.send_activity(MessageFactory.attachment(batch_certificate_cards[1]))
                    uploadAttachmentMessage = "<span style=\"background-color:#AA0000; color:whitesmoke; font-size:4rem;\">Please reply to <strong>this message</strong> when uploading the CSV file</span>"
                    await turn_context.send_activity(MessageFactory.text(uploadAttachmentMessage))
                else:
                    new_message = MessageFactory.attachment(create_client_certificates_card())
                    new_message.id = old_message_id
                    await turn_context.update_activity(new_message)
            elif submitted_data and submitted_data.get("id") == "ClientCertificates":
                loop = asyncio.get_event_loop()
                loop.create_task(client_certificates(turn_context, PersonEmail))
            elif submitted_data and submitted_data.get("id") == "BlockIOCs":
                loop = asyncio.get_event_loop()
                loop.create_task(block_ioc(turn_context, PersonEmail))
        return True
    
    async def send_file_request(self, turn_context: TurnContext, filename: str, file_card_desc: str, file_purpose: str):
        """Send a FileConsentCard to get user consent to upload a file."""
        file_path = filename
        file_size = os.path.getsize(file_path)
        consent_context = {"filename": filename, "filePurpose": file_purpose}

        file_card = FileConsentCard(
            description=file_card_desc,
            size_in_bytes=file_size,
            accept_context=consent_context,
            decline_context=consent_context
        )

        attachment = Attachment(
            content=file_card.serialize(),
            content_type=ContentType.FILE_CONSENT_CARD,
            name=filename
        )

        reply = MessageFactory.attachment(attachment)

        if file_purpose == "batchCSVResults":
            for a in turn_context.activity.attachments:
                if a.content_type == "application/vnd.microsoft.teams.file.download.info":
                    reply.text = f"<blockquote itemscope=\"\" itemtype=\"http://schema.skype.com/Reply\" itemid=\"{turn_context.activity.id}\">\r\n<strong itemprop=\"mri\" itemid=\"{turn_context.activity.from_property.id}\">{turn_context.activity.from_property.name}</strong><span itemprop=\"time\" itemid=\"{turn_context.activity.id}\"></span>\r\n<p itemprop=\"preview\">{a.name}<p>\r\n</blockquote>\r\n<p></p>"
                    break

        await turn_context.send_activity(reply)

    async def on_teams_file_consent_accept(
            self,
            turn_context: TurnContext,
            file_consent_card_response: FileConsentCardResponse
    ):
        """Handles file upload when the user accepts the file consent."""
        file_path = file_consent_card_response.context["filename"]
        file_purpose = file_consent_card_response.context["filePurpose"]

        if file_purpose == "batchCSVResults":
            zip_name = file_path
            dir_name = file_path[:-4]
            with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for root, directory, files in os.walk(f"BatchClientCertificate\\{dir_name}"):
                    for file in files:
                        arc_name = root[root.find("\\") + 1:]
                        zip_file.write(os.path.join(root, file), os.path.join(arc_name, file))

        file_size = os.path.getsize(file_path)

        headers = {
            "Content-Length": f"\"{file_size}\"",
            "Content-Range": f"bytes 0-{file_size-1}/{file_size}"
        }
        
        with open(file_path, "rb") as f:
            response = requests.put(
                file_consent_card_response.upload_info.upload_url, f, headers=headers
            )

        upload_info = file_consent_card_response.upload_info
        download_card = FileInfoCard(
            unique_id=upload_info.unique_id,
            file_type=upload_info.file_type
        )

        attachment = Attachment(
            content=download_card.serialize(),
            content_type=ContentType.FILE_INFO_CARD,
            name=upload_info.name,
            content_url=upload_info.content_url
        )

        reply = MessageFactory.attachment(attachment)
        reply.id = turn_context.activity.channel_data["legacy"]["replyToId"]
        await turn_context.update_activity(reply)
        await turn_context.send_activity(Activity(type=ActivityTypes.invoke_response))  # This is required so that "Something went wrong, please try again" does not come up

        if file_purpose == "batchCSVResults":
            os.remove(zip_name)

    async def on_teams_file_consent_decline(self, turn_context, file_consent_card_response):
        await turn_context.send_activity("You declined the template file. Please accept.")
        await turn_context.send_activity(Activity(type=ActivityTypes.invoke_response))

    async def handle_incoming_attachment(self, turn_context: TurnContext, personEmail: str):
        """
        Handle attachments uploaded by users. The bot receives an Attachment in an Activity.
        The activity has a List of attachments.
        Not all channels allow users to upload files. Some channels have restrictions
        on file type, size, and other attributes. Consult the documentation for the channel for
        more information. For example Skype's limits are here
        <see ref="https://support.skype.com/en/faq/FA34644/skype-file-sharing-file-types-size-and-time-limits"/>.
        :param turn_context:
        :return:
        """
        repliedToAttachmentCard = False
        attachmentIsCsv = False
        attachmentToDownload = None
        for attachment in turn_context.activity.attachments:
            if attachment.content_type == "text/html":

                previousMessage = []
                previousMessageHTML = lxml.html.fromstring("<div>" + attachment.content + "</div>")
                repliedMessage = previousMessageHTML.find("./blockquote/p")

                if repliedMessage is not None:
                    botItemID = previousMessageHTML.find("./blockquote/strong").get("itemid")
                    botSubmitParagraph = repliedMessage.xpath("string()")
                    repliedToAttachmentCard = (botItemID == "28:d614caf2-917b-4328-8580-197a4dd00f13" and botSubmitParagraph == "Please reply to this message when uploading the CSV file")
                
            elif attachment.content_type == "application/vnd.microsoft.teams.file.download.info":
                attachmentIsCsv = attachment.content["fileType"].lower() == "csv"
                attachmentToDownload = attachment

        if repliedToAttachmentCard and attachmentIsCsv:
            loop = asyncio.get_event_loop()
            loop.create_task(self.process_csv(turn_context, attachmentToDownload, personEmail))
        else:
            message = "I'm sorry, I don't understand that. Please try again or type \"help\" to see my command list."

            full_reply_message = f"<blockquote itemscope=\"\" itemtype=\"http://schema.skype.com/Reply\" itemid=\"{turn_context.activity.id}\">\r\n<strong itemprop=\"mri\" itemid=\"{turn_context.activity.from_property.id}\">{turn_context.activity.from_property.name}</strong><span itemprop=\"time\" itemid=\"{turn_context.activity.id}\"></span>\r\n<p itemprop=\"preview\">{previousMessage}<p>\r\n</blockquote>\r\n<p>{message}</p>"

            reply = MessageFactory.text(full_reply_message)
            await turn_context.send_activity(reply)

    async def process_csv(self, turn_context: TurnContext, attachment: Attachment, personEmail: str) -> dict:
        """
        Retrieve the attachment via the attachment's contentUrl.
        :param attachment:
        :return: Dict: keys "filename", "local_path"
        """
        file_download = FileDownloadInfo.deserialize(attachment.content)

        response = requests.get(file_download.download_url, allow_redirects=True)        
        if response.status_code == 200:
            tmp = tempfile.NamedTemporaryFile(suffix = ".csv", dir = ".", delete = False)
            with open(tmp.name, "wb") as t:
                t.write(response.content)
            error_code, zip_name = create_batch_csv(turn_context, personEmail, tmp)
            if error_code == 4443:
                await turn_context.send_activity(MessageFactory.attachment(create_batch_certificate_members_missing_card()))
                remove_extra_files(zip_name)
                return
            elif error_code == 4444:
                await turn_context.send_activity(MessageFactory.attachment(create_batch_certificate_error_card()))

            await self.send_file_request(turn_context, zip_name, "Here are your results", "batchCSVResults")
            remove_extra_files(zip_name)

        else:
            await turn_context.send_activity(MessageFactory.attachment(create_batch_certificate_error_card()))
