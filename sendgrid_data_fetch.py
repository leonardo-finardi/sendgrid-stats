import requests
import pandas as pd
import xlsxwriter

class ApiProvider:
    def __init__(self, base_url, api_token):
        self.base_url = base_url
        self.api_token = api_token

    def get(self, endpoint):
        url = f"{self.base_url}/{endpoint}"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.api_token}'
            }
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return data
        raise ValueError(f"Failed to get data from {url}. Error: {response.text}")


class SendgridStats:
    def __init__(self, start_date, base_url, api_token):
        self.api_token = api_token
        self.base_url = base_url
        self.start_date = start_date

    def email_provider_stats(self):
        endpoint = f"mailbox_providers/stats?start_date={self.start_date}"
        response = ApiProvider( self.base_url, self.api_token ).get(endpoint=endpoint)
        return response

    def global_stats(self):
        endpoint = f"stats?start_date={self.start_date}&limit=250"
        response = ApiProvider( BASE_URL, API_TOKEN ).get(endpoint=endpoint)
        return response

# Constraints
BASE_URL = "https://api.sendgrid.com/v3"
API_TOKEN = "your-api-token"
START_DATE = "2022-01-03"

providerStats = pd.DataFrame(SendgridStats(START_DATE, BASE_URL, API_TOKEN).email_provider_stats())
globalStats = pd.DataFrame(SendgridStats(START_DATE, BASE_URL, API_TOKEN).global_stats())
writer = pd.ExcelWriter('SendgridStats.xlsx', engine='xlsxwriter')
providerStats.to_excel(writer, sheet_name='Email Provider')
globalStats.to_excel(writer, sheet_name='Global Stats')
writer.save()
