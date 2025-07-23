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

from teams import Application, ApplicationOptions, TeamsAdapter
from teams.ai import AIOptions
from teams.ai.models import AzureOpenAIModelOptions, OpenAIModel, OpenAIModelOptions
from teams.ai.planners import ActionPlanner, ActionPlannerOptions
from teams.ai.prompts import PromptManager, PromptManagerOptions
from teams.state import TurnState
from teams.feedback_loop_data import FeedbackLoopData

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
from botbuilder.schema.teams import FileConsentCard, FileConsentCardResponse, FileInfoCard
from botbuilder.schema.teams.additional_properties import ContentType

import requests
import glob
from datetime import datetime, timezone

import html

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
    accesslist_REQUESTEDQUARANTINE = []
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
    if len(root.findall("REQUESTEDQUARANTINE")[0]) != 0:
        for i in range(len(root.findall("REQUESTEDQUARANTINE")[0])):
            accesslist_REQUESTEDQUARANTINE.append(root.findall("REQUESTEDQUARANTINE")[0][i-1].text.lower())     
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
        "`Release Emails` \n" \
        "\n" \
        "`Qualys Assets` \n" \
        "\n" \
        "`Client Certificates` \n" \
        "\n" \
        "`Script Status` \n" \
        "\n" \
        "If you need to report an email security incident, please forward the suspicious email as an attachement to emailsecurity@ehealthsask.ca"

def create_release_emails_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "size": "Large",
                "weight": "Bolder",
                "text": "Release Encrypted Emails",
                "horizontalAlignment": "Center"
            },
            {
                "type": "TextBlock",
                "text": "Before proceeding:",
                "size": "Medium",
                "weight": "Bolder"
            },
            {
                "type": "Input.Toggle",
                "title": "Do you know the sender? (Select-YES/Empty-NO)",
                "valueOn": "YES",
                "valueOff": "NO",
                "id": "Question1"
            },
            {
                "type": "Input.Toggle",
                "title": "Are you expecting this message? (Select-YES/Empty-NO)",
                "valueOn": "YES",
                "valueOff": "NO",
                "id": "Question2"
            },
            {
                "type": "Input.Toggle",
                "title": "Is this business-related? (Select-YES/Empty-NO)",
                "valueOn": "YES",
                "valueOff": "NO",
                "id": "Question3"
            },
            {
                "type": "Input.ChoiceSet",
                "id": "Environment",
                "value": "ESA",
                "choices": [
                    {
                        "title": "Microsoft O365",
                        "value": "O365"
                    },
                    {
                        "title": "Cisco ESA",
                        "value": "ESA"
                    }
                ]
            },
            {
                "type": "Input.Text",
                "placeholder": "Message ID",
                "style": "text",
                "maxLength": 0,
                "id": "MID"
            },
            {
                "type": "Input.Text",
                "placeholder": "Recipient (Use COMMA to seperate Multi Email Address)",
                "style": "text",
                "maxLength": 0,
                "id": "RECIPIENT"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                    "cardType": "input",
                    "id": "ReleaseEmails"
                }
            }
        ]
    }
    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

def create_qualys_assets_card():
    ADAPTIVE_CARD_CONTENT = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "size": "medium",
                "weight": "bolder",
                "text": "Query Qualys Assets",
                "horizontalAlignment": "center"
            },
            {
                "type": "Input.Text",
                "placeholder": "HOSTNAME",
                "style": "text",
                "maxLength": 0,
                "id": "HOSTNAME"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                    "cardType": "input",
                    "id": "QualysAssets"
                }
            }
        ]
    }
    return CardFactory.adaptive_card(ADAPTIVE_CARD_CONTENT)

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
                "text": "Once you have finished filling out the entries in the CSV file, please REPLY to this message and upload the file TO THIS THREAD"
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
                "type": "Input.Text",
                "placeholder": "Country",
                "style": "text",
                "maxLength": 0,
                "value": "CA",
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

async def releaseemails(context: TurnContext, PersonName: str, PersonEmail: str):
    submitted_data = context.activity.value
    
    if submitted_data['Question1'] == 'YES' and submitted_data['Question2'] == 'YES' and submitted_data['Question3'] == 'YES':
        if submitted_data['Environment'] == 'ESA':
            MID = submitted_data['MID']
            #Created = result['created']
            Created = context.activity.timestamp
            #PersonId = result['personId']
            #PersonName = send_get('https://webexapis.com/v1/people/{0}'.format(PersonId))['displayName']
            #PersonEmail = send_get('https://webexapis.com/v1/people/{0}'.format(PersonId))['emails']
            with open("ReleaseEmail.log", "r") as file:
                for line in file:
                    if MID in line and "successfully" in line:
                        msg = "Email was already released by: "+line
                        released = True
                        break
                    else:
                        released = False
            if not released:
                COMMAND = "grep "+MID+" mail_logs"
                ssh = paramiko.SSHClient()
                ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                k = paramiko.RSAKey.from_private_key_file(os.path.expanduser('.\\.ssh\\id_rsa_esa'))
                ssh.connect(sma_server, username=sma_uname, pkey=k)
                ssh_stdin, ssh_stdout, ssh_stderr = ssh.exec_command(COMMAND)
                for line in ssh_stdout:
                    try:
                        found = re.search('MID .* \(', line.strip())
                        S_MID = found.group(0).split(" ")[1]
                    except AttributeError:
                        found = ''
                ssh.close()
                if not found:
                    msg = "Can't find Message MID#" + MID + " in Encrypted Message quarantine. Please check your input and try again."
                else:
                    SMA_MID = []
                    SMA_MID.append(int(S_MID))
                    SMA_headers = {"Content-Type": "text/plain","Authorization": 'Basic ' + sma_token}
                    SMA_body = {"action": "release","mids": SMA_MID,"quarantineName": "Encrypted Message","quarantineType": "pvo"}
                    print(SMA_headers)
                    print(SMA_body)
                    SMA_response = requests.post(sma_server_api, data=json.dumps(SMA_body), headers=SMA_headers, verify=False)
                    print(SMA_response.json())
                    if SMA_response.status_code == 200 and SMA_response.json()['data']['totalCount'] == 1:
                        msg = "MID#"+MID+" is released successfully, please check your mailbox."
                        print(msg)
                        print(PersonName, PersonEmail, "submitted MID#", MID, "at", Created, "and is released successfully.", file=open("ReleaseEmail.log", "a"))
                    if SMA_response.status_code == 200 and SMA_response.json()['data']['totalCount'] == 0:
                        msg = "MID#"+MID+" was already released by the others but wasn't from this bot service."
                        print(msg)    
        if submitted_data['Environment'] == 'O365':
            MID = submitted_data['MID']
            RECIPIENT = submitted_data['RECIPIENT']
            #PersonId = result['personId']
            #PersonEmail = send_get('https://webexapis.com/v1/people/{0}'.format(PersonId))['emails']
            #RECIPIENT = PersonEmail[0].lower()
            Release_Email = subprocess.Popen(['powershell.exe', './Release_O365.ps1', MID, RECIPIENT], stdout=subprocess.PIPE)
            print(Release_Email.pid)
            Release_Email.wait()
            if Release_Email.returncode == 0:
                msg = Release_Email.stdout.read().decode("utf-8")
            else:
                msg = "Invalid result: " + str(Release_Email.returncode) + " Please contact the developer to get more detail."
    else:
        msg = "You answered \"NO\" to any of the questions above, please delete the email or mark it as \"Junk\" in your mail client."
        print(msg)
    await context.send_activity(msg)
    
async def queryassets(context: TurnContext):
    submitted_data = context.activity.value
    HOSTNAME = submitted_data.get("HOSTNAME")
    xml = """<ServiceRequest>
        <filters>
        <Criteria field="tagName" operator="EQUALS">{TAGNAME}</Criteria>
        <Criteria field="name" operator="CONTAINS">{HOSTNAME}</Criteria>
        </filters>
        </ServiceRequest>""".format(HOSTNAME=HOSTNAME, TAGNAME="Cloud Agent")
    headers = {'Content-Type': 'text/xml'}
    response = requests.post('https://qualysapi.qg1.apps.qualys.ca/qps/rest/2.0/search/am/hostasset', data=xml, headers=headers, auth=(Qualys_username, Qualys_password))
    root = ET.fromstring(response.text)
    status = root[0].text
    count = root[1].text
    if status == 'SUCCESS' and count != '0':
        try:
            hostname = root[3][0].findall('name')[0].text
        except IndexError:
            hostname = "IndexError"
        try:
            AssetID = root[3][0].findall('id')[0].text
        except IndexError:
            AssetID = "IndexError"
        try:
            HostID = root[3][0].findall('qwebHostId')[0].text
        except IndexError:
            HostID = "IndexError"
        try:
            IPAddress = root[3][0].findall('address')[0].text
        except IndexError:
            IPAddress = "IndexError"
        try:
            OS = root[3][0].findall('os')[0].text
        except IndexError:
            OS = "IndexError"
        try:
            FQDN = root[3][0].findall('fqdn')[0].text
        except IndexError:
            FQDN = root[3][0].findall('name')[0].text
        msg = status + " | " + count + " | " + hostname + " | Asset ID:" + AssetID + " | Host ID:" + HostID + " | IP Address:" + IPAddress  + " | OS:"  + OS
        await context.send_activity(msg)
        filename = FQDN + "_" + datetime.now().strftime('%Y-%m-%d') + ".csv"
        path_filename = './vulnerabilities/' + FQDN + "_" + datetime.now().strftime('%Y-%m-%d') + ".csv"
        if os.path.exists(path_filename):
            await context.send_activity('Get Vulnerability list is currently disabled')
            #data = MultipartEncoder({"files": (filename, open(path_filename, 'rb'), 'text/csv')})
            #await context.send_activity(data)
        else:
            if HostID != "IndexError" and AssetID != "IndexError":
                msg = "Today's vulnerability file for " + HOSTNAME + " does not exist, bot is generating and will push the CSV file to you when it's done."
                await context.send_activity(msg)
                #MyBot.vuln_list(HostID,AssetID,HOSTNAME)
            else:
                msg = "HostID or AssetID doesn't exist. Please contact the server administrator to have a check"
                await context.send_activity(msg)
    if status == 'SUCCESS' and count == '0':
        HOSTNAME = HOSTNAME.lower()
        xml = """<ServiceRequest>
            <filters>
            <Criteria field="tagName" operator="EQUALS">{TAGNAME}</Criteria>
            <Criteria field="name" operator="CONTAINS">{HOSTNAME}</Criteria>
            </filters>
            </ServiceRequest>""".format(HOSTNAME=HOSTNAME, TAGNAME="Cloud Agent")
        headers = {'Content-Type': 'text/xml'}
        response = requests.post('https://qualysapi.qg1.apps.qualys.ca/qps/rest/2.0/search/am/hostasset', data=xml, headers=headers, auth=(Qualys_username, Qualys_password))
        root = ET.fromstring(response.text)
        status = root[0].text
        count = root[1].text
        if status == 'SUCCESS' and count != '0':
            msg = status + " | " + count + " | " + root[3][0].findall('name')[0].text + " | Asset ID:" + root[3][0].findall('id')[0].text + " | Host ID:" + root[3][0].findall('qwebHostId')[0].text + " | IP Address:" + root[3][0].findall('address')[0].text  + " | OS:"  + root[3][0].findall('os')[0].text
            HostID = root[3][0].findall('qwebHostId')[0].text
            AssetID = root[3][0].findall('id')[0].text
            try:
                FQDN = root[3][0].findall('fqdn')[0].text
            except IndexError:
                FQDN = root[3][0].findall('name')[0].text
            await context.send_activity(msg)
            #send_post("https://webexapis.com/v1/messages/", {"roomId": RoomID, "markdown": msg})
            filename = FQDN + "_" + datetime.now().strftime('%Y-%m-%d') + ".csv"
            path_filename = './vulnerabilities/' + FQDN + "_" + datetime.now().strftime('%Y-%m-%d') + ".csv"
            if os.path.exists(path_filename):
                await context.send_activity('Get Vulnerability list is currently disabled')
                #data = MultipartEncoder({'roomId': RoomID, "files": (filename, open(path_filename, 'rb'), 'text/csv')})
                #request = requests.post('https://webexapis.com/v1/messages', data=data, headers = {"Authorization": "Bearer " + bearer, 'Content-Type': data.content_type})
            else:
                msg = "Today's vulnerability file for " + HOSTNAME + " does not exist, bot is generating and will push the CSV file to you when it's done."
                await context.send_activity(msg)
                await context.send_activity('Get Vulnerability List is Currently Disabled')
                #send_post("https://webexapis.com/v1/messages/", {"roomId": RoomID, "markdown": msg})
                # MyBot.vuln_list(HostID,AssetID,HOSTNAME)
    if status == 'SUCCESS' and count == '0':
        HOSTNAME = HOSTNAME.lower()
        xml = """<ServiceRequest>
            <filters>
            <Criteria field="tagName" operator="EQUALS">{TAGNAME}</Criteria>
            <Criteria field="name" operator="CONTAINS">{HOSTNAME}</Criteria>
            </filters>
            </ServiceRequest>""".format(HOSTNAME=HOSTNAME, TAGNAME="Cloud Agent")
        headers = {'Content-Type': 'text/xml'}
        response = requests.post('https://qualysapi.qg1.apps.qualys.ca/qps/rest/2.0/search/am/hostasset', data=xml, headers=headers, auth=(Qualys_username, Qualys_password))
        root = ET.fromstring(response.text)
        status = root[0].text
        count = root[1].text
        if status == 'SUCCESS' and count != '0':
            msg = status + " | " + count + " | " + root[3][0].findall('name')[0].text + " | Asset ID:" + root[3][0].findall('id')[0].text + " | Host ID:" + root[3][0].findall('qwebHostId')[0].text + " | IP Address:" + root[3][0].findall('address')[0].text  + " | OS:"  + root[3][0].findall('os')[0].text
            HostID = root[3][0].findall('qwebHostId')[0].text
            AssetID = root[3][0].findall('id')[0].text
            try:
                FQDN = root[3][0].findall('fqdn')[0].text
            except IndexError:
                FQDN = root[3][0].findall('name')[0].text
            """ RoomID = webhook['data']['roomId'] """
            await context.send_activity(msg)
            #send_post("https://webexapis.com/v1/messages/", {"roomId": RoomID, "markdown": msg})
            filename = FQDN + "_" + datetime.now().strftime('%Y-%m-%d')+ ".csv"
            path_filename = './vulnerabilities/' + FQDN + "_" + datetime.now().strftime('%Y-%m-%d') + ".csv"
            if os.path.exists(path_filename):
                await context.send_activity('Get Vulnerability List is Currently Disabled')
                #data = MultipartEncoder({'roomId': RoomID, "files": (filename, open(path_filename, 'rb'), 'text/csv')})
                #request = requests.post('https://webexapis.com/v1/messages', data=data, headers = {"Authorization": "Bearer " + bearer, 'Content-Type': data.content_type})
            else:
                msg = "Today's vulnerability file for " + HOSTNAME + " does not exist, bot is generating and will push the CSV file to you when it's done."
                await context.send_activity(msg)
                await context.send_activity('Get Vulnerability List is Currently Disabled')
                #send_post("https://webexapis.com/v1/messages/", {"roomId": RoomID, "markdown": msg})
                #MyBot.vuln_list(HostID,AssetID,HOSTNAME)            
    if status == 'SUCCESS' and count == '0':
        msg = "Can't find " + HOSTNAME + " in Qualys. Host name is case-sensitive, please confirm the hostname and try again."
        await context.send_activity(msg)

async def client_certificates(context: TurnContext):
    submitted_data = context.activity.value
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
        Generate_Certificate = subprocess.run(['powershell.exe', '-File', './Generate_Certificate.ps1',  filename, environment, comment_file])
    else:
        await context.send_activity(filename + " doesn't exist")
    if Generate_Certificate.returncode != 0:
        await context.send_activity("Invalid result: " + str(Generate_Certificate.returncode) + " Please check the Certificate_output.log to get more detail.")
    else:
        #after access to ms teams this is where you place the code for sending the team 
        await context.send_activity("Email has been sent to " + Email + " with the PFX file and password.")

def create_batch_csv(turn_context: TurnContext, personEmail: str, tmpFile):
    has_error = False
    current_datetime = datetime.now()
    filename = f"{personEmail}_{current_datetime.year:04}{current_datetime.month:02}{current_datetime.day:02}_{current_datetime.hour:02}{current_datetime.minute:02}{current_datetime.second:02}"
    batch_cert_request = subprocess.run(['powershell.exe', './AutoGenerate_Client_CertificateTEST.ps1', tmpFile.name, filename])
    tmpFile.close()
    os.remove(tmpFile.name)
    has_error = batch_cert_request.returncode == 4444

    zip_name = filename + ".zip"
    with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for root, directory, files in os.walk(f"BatchClientCertificate\\{filename}"):
            for file in files:
                arc_name = root[root.find("\\") + 1:]
                zip_file.write(os.path.join(root, file), os.path.join(arc_name, file))

    return [has_error, zip_name]

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
                turn_context.send_activity("Sorry, but I can only accept one attachment at a time.")
            else:
                await self.handle_incoming_attachment(turn_context, PersonEmail)
        elif turn_context.activity.text:
            card = None
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
            elif turn_context.activity.text.casefold().startswith('release emails'.casefold()):
                message = ""
                card = create_release_emails_card()
            #statement if user inputs qualys assets produces the card after checking access list
            elif turn_context.activity.text.casefold().startswith('qualys assets'.casefold()):
                if accesslist_ADMIN == "ALL" or PersonEmail.lower() in accesslist_ADMIN or PersonEmail.lower() in accesslist_QUALYS:
                    message = ""
                    card = create_qualys_assets_card()
                else:
                    message = "Sorry, you (" + PersonName + ") are NOT allowed to use this module. Please contact Hao.Ban@eHealthsask.ca for help."
            #statement if user inputs client certficates produces the card after checking access list
            elif turn_context.activity.text.casefold().startswith('client certificates'.casefold()):
                if accesslist_ADMIN == "ALL" or PersonEmail.lower() in accesslist_ADMIN or PersonEmail.lower() in accesslist_CERTIFICATES:
                    card = create_single_or_batch_card()
                    message = ""
                else:
                    message = "Sorry, you (" + PersonName + ") are NOT allowed to use this module. Please contact Hao.Ban@eHealthsask.ca for help."
            elif turn_context.activity.text.casefold().startswith('script status'.casefold()):
                    if accesslist_ADMIN == "ALL" or PersonEmail.lower() in accesslist_ADMIN or PersonEmail.lower() in accesslist_SCRIPTSTATUS:
                        Processor_Monitor = subprocess.run(['powershell.exe', '-File', './ProcessorMonitor.ps1'], capture_output=True, text=True)
                        message = Processor_Monitor.stdout
                    else:
                        message = "Sorry, you (" + PersonName + ") are NOT allowed to use this module. Please contact Hao.Ban@eHealthsask.ca for help."
            else:
                message = "I'm sorry, I don't understand that. Please try again or type \"help\" to see my command list."

            full_reply_message = f"<blockquote itemscope=\"\" itemtype=\"http://schema.skype.com/Reply\" itemid=\"{turn_context.activity.id}\">\r\n<strong itemprop=\"mri\" itemid=\"{turn_context.activity.from_property.id}\">{PersonName}</strong><span itemprop=\"time\" itemid=\"{turn_context.activity.id}\"></span>\r\n<p itemprop=\"preview\">{previousMessage}<p>\r\n</blockquote>\r\n<p>{message}</p>"

            reply = MessageFactory.text(full_reply_message)
            if card:
                reply.attachments = [card]
            await turn_context.send_activity(reply)
        else:
            submitted_data = turn_context.activity.value
            if submitted_data and submitted_data.get("id") == "ReleaseEmails":
                await releaseemails(turn_context, PersonName, PersonEmail)
            elif submitted_data and submitted_data.get("id") == "SingleOrBatch":
                old_message_id = turn_context.activity.channel_data["legacy"]["replyToId"]
                if submitted_data.get("Batch") == "True":
                    batch_certificate_cards = create_batch_certificates_card()
                    new_message = MessageFactory.attachment(batch_certificate_cards[0])
                    new_message.id = old_message_id
                    await turn_context.update_activity(new_message)
                    await self.send_csv_request(turn_context, "Client_Certificate_Information_Template.csv", "Please fill out this CSV")
                    await turn_context.send_activity(MessageFactory.attachment(batch_certificate_cards[1]))
                    uploadAttachmentMessage = "<span style=\"background-color:#AA0000; color:whitesmoke; font-size:4rem;\">Please reply to <strong>this message</strong> when uploading the CSV file</span>"
                    await turn_context.send_activity(MessageFactory.text(uploadAttachmentMessage))
                else:
                    new_message = MessageFactory.attachment(create_client_certificates_card())
                    new_message.id = old_message_id
                    await turn_context.update_activity(new_message)
            elif submitted_data and submitted_data.get("id") == "ClientCertificates":
                await client_certificates(turn_context)
            elif submitted_data and submitted_data.get("id") == "QualysAssets":
                await queryassets(turn_context)
        return True
    
    async def send_csv_request(self, turn_context: TurnContext, filename: str, file_card_desc: str):
        """Send a FileConsentCard to get user consent to upload a file."""
        file_path = filename
        file_size = os.path.getsize(file_path)
        consent_context = {"filename": filename}

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
        await turn_context.send_activity(reply)

    async def on_teams_file_consent_accept(
            self,
            turn_context: TurnContext,
            file_consent_card_response: FileConsentCardResponse
    ):
        """Handles file upload when the user accepts the file consent."""
        file_path = file_consent_card_response.context["filename"]

        if file_path != "Client_Certificate_Information_Template.csv":
            zip_name = file_path
            with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for root, directory, files in os.walk(f"BatchClientCertificate\\{file_path}"):
                    for file in files:
                        arc_name = root[root.find("\\") + 1:]
                        zip_file.write(os.path.join(root, file), os.path.join(arc_name, file))

        file_size = os.path.getsize(file_path)

        headers = {
            "Content-Length": f"\"{file_size}\"",
            "Content-Range": f"bytes 0-{file_size-1}/{file_size}"
        }
        response = requests.put(
            file_consent_card_response.upload_info.upload_url, open(file_path, "rb"), headers=headers
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
                """ print(attachment.content) """

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
            await self.process_csv(turn_context, attachmentToDownload, personEmail)
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
            has_error, zip_name = create_batch_csv(turn_context, personEmail, tmp)
            if has_error:
                turn_context.send_activity(MessageFactory.attachment(create_batch_certificate_error_card()))

            await self.send_csv_request(turn_context, zip_name, "Here are your results")
            remove_extra_files(zip_name)

        else:
            turn_context.send_activity(MessageFactory.attachment(create_batch_certificate_error_card()))