import gspread
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSheetsHandler:
    def __init__(self, key_file, spreadsheet_title):
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(key_file, ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])
        self.client = gspread.authorize(self.credentials)
        self.spreadsheet = self.get_spreadsheet(spreadsheet_title)
        self.worksheet = self.spreadsheet.get_worksheet(0)

    def get_spreadsheet(self, title):
        return self.client.open(title)

    def write_headers(self, headers):
        header_range = f'A1:{chr(64 + len(headers))}1'
        self.worksheet.update(header_range, [headers])
        self.worksheet.format("A1:D1", {"textFormat": {"bold": True}})
        self.worksheet.format("A1:D1", {"horizontalAlignment": "CENTER"})

    def write_data(self, data):
        self.worksheet.append_row(data)
