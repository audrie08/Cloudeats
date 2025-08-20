# Copyright (c) Streamlit Inc. (2018-2022) Snowflake Inc. (2022-2025)
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""Manage the user's Streamlit credentials."""

from __future__ import annotations

import os
import json
import sys
import textwrap
from typing import Final, NamedTuple, NoReturn
from uuid import uuid4

from streamlit import cli_util, env_util, file_util, util
from streamlit.logger import get_logger

_LOGGER: Final = get_logger(__name__)

class GoogleSheetsCredentials:
    """
    Handles Google Sheets API credentials and authentication
    """
    
    def __init__(self):
        self.scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        self.spreadsheet_id = "1PxdGZDltF2OWj5b6A3ncd7a1O4H-1ARjiZRBH0kcYrI"
        self.credentials_file = "production-schedule-calculator-0dceed735b36.json"
        
    def get_credentials_from_file(self):
        """
        Load credentials from JSON file (development/local use)
        """
        try:
            credentials = Credentials.from_service_account_file(
                self.credentials_file, 
                scopes=self.scopes
            )
            return credentials
        except FileNotFoundError:
            raise FileNotFoundError(f"Credentials file '{self.credentials_file}' not found")
        except Exception as e:
            raise Exception(f"Error loading credentials from file: {str(e)}")
    
    def get_credentials_from_env(self):
        """
        Load credentials from environment variables (production use)
        Expects GOOGLE_CREDENTIALS_JSON environment variable with JSON string
        """
        try:
            credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
            if not credentials_json:
                raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not found")
            
            credentials_dict = json.loads(credentials_json)
            credentials = Credentials.from_service_account_info(
                credentials_dict, 
                scopes=self.scopes
            )
            return credentials
        except json.JSONDecodeError:
            raise ValueError("Invalid JSON in GOOGLE_CREDENTIALS_JSON environment variable")
        except Exception as e:
            raise Exception(f"Error loading credentials from environment: {str(e)}")
    
    def get_credentials_from_streamlit_secrets(self):
        """
        Load credentials from Streamlit secrets (Streamlit Cloud deployment)
        """
        try:
            credentials_dict = dict(st.secrets["google_credentials"])
            credentials = Credentials.from_service_account_info(
                credentials_dict, 
                scopes=self.scopes
            )
            return credentials
        except Exception as e:
            raise Exception(f"Error loading credentials from Streamlit secrets: {str(e)}")
    
    def get_credentials(self):
        """
        Get credentials using the first available method:
        1. Streamlit secrets (if running in Streamlit Cloud)
        2. Environment variables (if available)
        3. JSON file (fallback for local development)
        """
        # Try Streamlit secrets first (for Streamlit Cloud deployment)
        try:
            if hasattr(st, 'secrets') and 'google_credentials' in st.secrets:
                return self.get_credentials_from_streamlit_secrets()
        except:
            pass
        
        # Try environment variables (for other cloud deployments)
        try:
            if os.getenv('GOOGLE_CREDENTIALS_JSON'):
                return self.get_credentials_from_env()
        except:
            pass
        
        # Fallback to file (for local development)
        try:
            return self.get_credentials_from_file()
        except:
            pass
        
        raise Exception(
            "Could not load Google Sheets credentials. Please ensure one of the following:\n"
            "1. Set up Streamlit secrets with 'google_credentials' section\n"
            "2. Set GOOGLE_CREDENTIALS_JSON environment variable\n"
            "3. Place the JSON credentials file in the project directory"
        )
    
    def get_gspread_client(self):
        """
        Get authenticated gspread client
        """
        credentials = self.get_credentials()
        return gspread.authorize(credentials)
    
    def get_spreadsheet(self, spreadsheet_id=None):
        """
        Get the spreadsheet object
        """
        client = self.get_gspread_client()
        sheet_id = spreadsheet_id or self.spreadsheet_id
        return client.open_by_key(sheet_id)


# Convenience functions for easy import
def get_credentials():
    """Get Google Sheets credentials"""
    creds_handler = GoogleSheetsCredentials()
    return creds_handler.get_credentials()

def get_gspread_client():
    """Get authenticated gspread client"""
    creds_handler = GoogleSheetsCredentials()
    return creds_handler.get_gspread_client()

def get_spreadsheet(spreadsheet_id=None):
    """Get spreadsheet object"""
    creds_handler = GoogleSheetsCredentials()
    return creds_handler.get_spreadsheet(spreadsheet_id)

# Configuration constants
SPREADSHEET_ID = "1PxdGZDltF2OWj5b6A3ncd7a1O4H-1ARjiZRBH0kcYrI"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
