"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""
from http import HTTPStatus
from aiohttp import web
from botbuilder.core.integration import aiohttp_error_middleware
from pyngrok import conf, ngrok
import requests
import os

from dotenv import load_dotenv

from bot import bot_app

load_dotenv()

routes = web.RouteTableDef()

ngrok_domain = os.environ.get("NGROK_DOMAIN")

ngrok.connect(5000, domain=ngrok_domain).public_url  #Start ngrok on port 5000
ngrok_response = requests.get('http://127.0.0.1:4040/api/tunnels')

@routes.post("/api/messages")
async def on_messages(req: web.Request) -> web.Response:
    res = await bot_app.process(req)

    if res is not None:
        return res

    return web.Response(status=HTTPStatus.OK)

app = web.Application(middlewares=[aiohttp_error_middleware])
app.add_routes(routes)

from config import Config

if __name__ == "__main__":
    web.run_app(app, host="localhost", port=Config.PORT)