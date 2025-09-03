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

        if (
            turn_context.activity.attachments
            and len(turn_context.activity.attachments) >= 2
        ):
            await self.handle_incoming_attachment(turn_context)

        elif turn_context.activity.text:
            await turn_context.send_activity("response")
    
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

        await turn_context.send_activity(reply)

    async def on_teams_file_consent_accept(
            self,
            turn_context: TurnContext,
            file_consent_card_response: FileConsentCardResponse
    ):
        """Handles file upload when the user accepts the file consent."""
        file_path = file_consent_card_response.context["filename"]
        file_purpose = file_consent_card_response.context["filePurpose"]

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

    async def on_teams_file_consent_decline(self, turn_context, file_consent_card_response):
        await turn_context.send_activity("You declined the template file. Please accept.")
        await turn_context.send_activity(Activity(type=ActivityTypes.invoke_response))

    async def handle_incoming_attachment(self, turn_context: TurnContext):
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
            if attachment.content_type != "text/html":
                attachmentToDownload = attachment
                loop = asyncio.get_event_loop()
                loop.create_task(self.process_attachment(turn_context, attachmentToDownload))

    async def process_attachment(self, turn_context: TurnContext, attachment: Attachment) -> dict:
        """
        Retrieve the attachment via the attachment's contentUrl.
        :param attachment:
        :return: Dict: keys "filename", "local_path"
        """
        file_download = FileDownloadInfo.deserialize(attachment.content)

        response = requests.get(file_download.download_url, allow_redirects=True)        
        if response.status_code == 200:
            print(response.content)

            await self.send_file_request(turn_context, "filename", "Here are your results", "file purpose")

        else:
            await turn_context.send_activity("Error getting a file")
