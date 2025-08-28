"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""
import sys
import traceback
from datetime import datetime
from http import HTTPStatus

from aiohttp import web
from aiohttp.web import Request, Response, json_response
from botbuilder.core import (
    TurnContext,
)
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.integration.aiohttp import CloudAdapter, ConfigurationBotFrameworkAuthentication
from botbuilder.schema import Activity, ActivityTypes

from pyngrok import conf, ngrok
import requests
import os

from dotenv import load_dotenv
from bot import BotApp
from config import Config

CONFIG = Config()

load_dotenv()

ngrok_domain = os.environ.get("NGROK_DOMAIN")
NGROK_AUTH_TOKEN_TEAMS = os.environ.get("NGROK_AUTH_TOKEN_TEAMS")

conf.get_default().region = "us"

ngrok.set_auth_token(NGROK_AUTH_TOKEN_TEAMS)
ngrok.connect(6000, domain=ngrok_domain).public_url  #Start ngrok on port 6000
ngrok_response = requests.get('http://127.0.0.1:4040/api/tunnels')

script_path = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_path)

ADAPTER = CloudAdapter(ConfigurationBotFrameworkAuthentication(CONFIG))

async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The bot encountered an error or bug.")
    await context.send_activity(
        "To continue to run this bot, please fix the bot source code."
    )
    # Send a trace activity if we're talking to the Bot Framework Emulator
    if context.activity.channel_id == "emulator":
        # Create a trace activity that contains the error object
        trace_activity = Activity(
            label="TurnError",
            name="on_turn_error Trace",
            timestamp=datetime.utcnow(),
            type=ActivityTypes.trace,
            value=f"{error}",
            value_type="https://www.botframework.com/schemas/error",
        )
        # Send a trace activity, which will be displayed in Bot Framework Emulator
        await context.send_activity(trace_activity)


ADAPTER.on_turn_error = on_error

BOT = BotApp()

# Listen for incoming requests on /api/messages
async def messages(req: Request) -> Response:
    return await ADAPTER.process(req, BOT)


APP = web.Application(middlewares=[aiohttp_error_middleware])
APP.router.add_post("/api/messages", messages)

if __name__ == "__main__":
    try:
        web.run_app(APP, host="localhost", port=CONFIG.PORT)
    except Exception as error:
        raise error