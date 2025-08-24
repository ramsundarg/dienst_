import calendar
import datetime
import glob
import os


import pandas as pd
import re
import base64
import glob
import http.client
import json
import urllib.request
from time import sleep
from pathlib import Path
publishIcs = True

# and it's done !

month = 9
year = 2025
Employee = 'TRG'
employee_code = 'TRG'
convertFiles = False
from data_processing import get_df, apply_styling_to_excel

def convert_files():
    import http
    conn = http.client.HTTPSConnection("pdf-services-ew1.adobe.io")

    payload = "client_id=540f46c8507b47998d1c238878658535&client_secret=p8e-eNckbaGuPkDLT6iUEYSsye8ANTL-3ljD"

    headers = {
        'Content-Type': "application/x-www-form-urlencoded",
        'User-Agent': "insomnia/8.5.1"
    }

    conn.request("POST", "/token", payload, headers)

    res = conn.getresponse()
    data = res.read()

    data_decoded = data.decode("utf-8")
    data_json = json.loads(data_decoded)
    access_token = data_json['access_token']

    for file_name in list(glob.glob('original_data/*.pdf')):
        conn = http.client.HTTPSConnection("pdf-services.adobe.io")

        payload = "{\n\t\"mediaType\" :  \"application/pdf\"\n}"

        headers = {
            'Content-Type': "application/json",
            'User-Agent': "insomnia/8.5.1",
            'Authorization': f"Bearer {access_token}",
            'x-api-key': "540f46c8507b47998d1c238878658535"
        }

        conn.request("POST", "/assets", payload, headers)

        res = conn.getresponse()
        data = res.read()
        data_decoded = data.decode("utf-8")
        data_json = json.loads(data_decoded)
        uploadUri = data_json['uploadUri']
        assetId = data_json['assetID']

        import http.client

        share_host = 'dcplatformstorageservice-prod-us-east-1.s3-accelerate.amazonaws.com'
        conn = http.client.HTTPSConnection(share_host)
        share_host = f"https://{share_host}"

        dest_file = Path(file_name).stem
        with open(file_name, 'rb') as fp:
            payload = fp.read()
            # payload = base64.b64encode(payload)

        headers = {
            'Content-Type': "application/pdf",
            'User-Agent': "insomnia/8.5.1"
        }

        conn.request("PUT", uploadUri.removeprefix(share_host), payload, headers)
        res = conn.getresponse()

        conn = http.client.HTTPSConnection("pdf-services-ue1.adobe.io")

        payload = {"assetID": assetId, "targetFormat": "xlsx", "ocrLang": "de-DE"}
        payload = json.dumps(payload)

        headers = {
            'Content-Type': "application/json",
            'User-Agent': "insomnia/8.5.1",
            'Authorization': f"Bearer {access_token}",
            'x-api-key': "540f46c8507b47998d1c238878658535"
        }

        conn.request("POST", "/operation/exportpdf", payload, headers)

        res = conn.getresponse()
        request_id = res.headers['x-request-id']

        conn = http.client.HTTPSConnection("pdf-services-ue1.adobe.io")

        payload = ""

        headers = {
            'User-Agent': "insomnia/8.5.1",
            'Authorization': f"Bearer {access_token}",
            'x-api-key': "540f46c8507b47998d1c238878658535"
        }

        while True:
            conn.request("GET", f"/operation/exportpdf/{request_id}/status?=", payload, headers)

            res = conn.getresponse()
            data = res.read()
            data_decoded = data.decode("utf-8")
            data_json = json.loads(data_decoded)
            print(data_json)
            downloadUri = data_json.get('asset', '')
            if downloadUri == '':
                print(f"Not converted file {file_name}, waiting for 3 seconds")
                sleep(3)
            else:
                urllib.request.urlretrieve(downloadUri['downloadUri'], f'data/{dest_file}.xlsx')
                break

    print("All files converted, now computing the dienst")



if convertFiles:
    convert_files()

files = list(glob.glob('data/*.xlsx'))
df = get_df(files, employee_code, year, month)
if df is not None:
    df['date'] = df['date'].dt.date
    # Export to Excel with styling
    excel_path = f'processed_data/{employee_code}_{year}_{month}.xlsx'
    df.to_excel(excel_path, index=False, engine='openpyxl')
    apply_styling_to_excel(excel_path)

