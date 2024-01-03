import calendar
import datetime
import glob
from pathlib import Path

import dash
from dash import html, dcc, dash_table
from dash.dependencies import Input, Output
import pandas as pd
import re

names = {}
labels = []
month = 1
year = 2024
rows = []
dates = set()

for file_name in glob.glob("data/*.xlsx"):
    xl = pd.ExcelFile(file_name)
    res = len(xl.sheet_names)
    work_type = Path(file_name).stem
    regex = re.compile('[^a-zA-Z]')
    work_type = regex.sub('', work_type)

    work_type_hours = {}

    for sheet_name in xl.sheet_names:
        a = pd.read_excel(file_name, sheet_name=sheet_name)
        a1 = list(a.values.flatten())
        for cell in a1:
            result = re.search(r"\(([A-Z][A-Z][A-Z])\),.*/(.*)", str(cell))
            if result:
                names[result.group(1)] = result.group(2)
                labels.append({'label': result.group(2), 'value': result.group(1)})




def get_df(name):
    rows = []
    dates = set()
    for file_name in glob.glob("data/*.xlsx"):
        xl = pd.ExcelFile(file_name)
        res = len(xl.sheet_names)
        work_type = Path(file_name).stem
        regex = re.compile('[^a-zA-Z]')
        work_type = regex.sub('', work_type)

        work_type_hours = {}

        for sheet_name in xl.sheet_names:
            a = pd.read_excel(file_name, sheet_name=sheet_name)
            a1 = list(a.values.flatten())
            a2 = [str(s) for s in a1 if "=" in str(s)]
            for hours in a2:
                for hour  in hours.splitlines():
                    result = re.search(r"(.*) =.*- (.*) Uhr", hour)
                    if result and False:
                        work_type_hours[result.group(1)] = result.group(2)
                    else:
                        result = re.search(r"(.*) =", hour)
                        work_type_hours[result.group(1)] = hour.split("=", 1)[1]

        a = pd.read_excel(file_name, header=None)
        skip_rows = a[a[0] == 'Datum'].index[0]
        for sheet_name in xl.sheet_names:
            a = pd.read_excel(file_name, skiprows=skip_rows, header=0, sheet_name=sheet_name)
            emp = a[a[a.columns[0]].str.contains(name, na=False)]
            if emp.empty:
                continue
            for i in range(1, 32):
                if i in dates:
                    continue
                try:
                    date1 = datetime.datetime(year, month, i, 0, 0)
                except ValueError:
                    continue
                value = ""
                Uhr = ""

                if emp.get(i, None).all():
                    value = emp[i].values[0]
                    if value == 'D':
                        continue
                    elif (emp[i].isna().values[0]):
                        rows.append({'date': date1, 'day': calendar.day_name[date1.weekday()], "work_type_code": "Free",
                                     'work_time': "", "work_type": ""
                                     })
                    else:

                        rows.append(
                            {'date': date1, 'day': calendar.day_name[date1.weekday()], "work_type_code": value,
                             'work_time': work_type_hours.get(value, value),
                             "work_type": work_type})
                    dates.add(i)
    df_processed = pd.DataFrame.from_records(rows)
    df_processed = df_processed.sort_values(by='date')
    return df_processed



df = get_df('TRG')

# Initialize the Dash app
app = dash.Dash(__name__)

server = app.server

# Define the layout of the app
app.layout = html.Div([
    # Title
    html.Div(id='record-statistics'),

    # Dropdown to select a DataFrame
    dcc.Dropdown(
        id='dropdown',
        options=labels,
        value='TRG',  # Initial value
    ),

    # DataFrame display area
    dash_table.DataTable(
        id='data-table',
        columns=[{'name': col, 'id': col} for col in df.columns],
        data=df.to_dict('records'),
    ),
])


# Define a callback to update the DataFrame display based on the dropdown value
@app.callback(
    Output("record-statistics", "children"),
    Output("data-table", "data"),
    Input('dropdown', 'value')
)
def update_data_table(selected_value):
    print(selected_value)
    df= get_df(selected_value)
    return html.H3(f'Dienst for employee {selected_value}'),df.to_dict("records")



# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
