import json
import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials


__version__ = '0.2.0'

__all__ = ['models', 'get_client', 'load_workbook']


def get_client(file_path):
    """
    Returns Authenticated Gspread Client

    Args:
         file_path: path to JSON file, dict or JSON string. This is fetched from Google Developers site.

    Returns:
        :class:`gspread.Client` instance
    """
    if os.path.exists(file_path):
        with open(file_path) as f:
            data = json.load(f)
    elif isinstance(file_path, str):
        data = json.loads(file_path)
    elif isinstance(file_path, dict):
        pass
    else:
        raise AttributeError("Invalid instance type for 'file_path'. Must be a path to JSON file or JSON loadable "
                             "string or dictionary")
    creds = ServiceAccountCredentials.from_json_keyfile_dict(data, ['https://spreadsheets.google.com/feeds',
                                                                    'https://www.googleapis.com/auth/drive'])
    return gspread.authorize(creds)


from . import models


def load_workbook(url, credentials):
    """Load a Google Sheet workbook.

        Args:
            url: Google Sheet URL (key will be supported in future release)
            credentials: Service Account JSON credentials created for the Google Sheet. Can be path, dict or JSON str.

        Returns:
            :class:`core.Workbook`
    """
    return models.Workbook(url, credentials)


def create_workbook(url, credentials):
    raise NotImplementedError('This feature is not available yet')
