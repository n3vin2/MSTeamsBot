"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import os

from dotenv import load_dotenv

load_dotenv()

class Config:
    """Bot Configuration"""

    PORT = 5000
    APP_TYPE = os.environ.get("BOT_TYPE", "")
    APP_ID = os.environ.get("BOT_ID", "")
    APP_PASSWORD = os.environ.get("BOT_PASSWORD", "")
    APP_TENANTID = os.environ.get("BOT_TENANT_ID", "")