import calendar
import datetime
import dash
from dash import html, dcc, dash_table
from dash.dependencies import Input, Output
import pandas as pd
import re

names = {}
labels = []
file_name = './FKTW ADAC 08-23.xlsx'
month = 8
year = 2023
file_name = './FKTW ADAC 08-23.xlsx'
rows = []
dates = set()

for sheet in range(0, 4):
    a = pd.read_excel(file_name, sheet_name=sheet)
    a1 = list(a.values.flatten())
    for cell in a1:
        result = re.search(r"\(([A-Z][A-Z][A-Z])\),.*/(.*)", str(cell))
        if result:
            names[result.group(1)] = result.group(2)
            labels.append({'label': result.group(2), 'value': result.group(1)})


def get_df(name):
    rows=[]
    dates = set()
    for sheet in range(0, 4):
        a = pd.read_excel(file_name, sheet_name=sheet)
        a1 = list(a.values.flatten())

        work_type = ""
        work_type_hours = {}
        for cell in a1:
            result = re.search(r"MKT (\w+)", str(cell))
            if result:
                work_type = result.group(1)
                break

        a2 = [str(s) for s in a1 if "=" in str(s)]
        for hours in a2:
            result = re.search(r"(.*) =.*- (.*) Uhr", hours)
            if result and False:
                work_type_hours[result.group(1)] = result.group(2)
            else:
                result = re.search(r"(.*) =", hours)
                work_type_hours[result.group(1)] = hours.split("=", 1)[1]

        a = pd.read_excel(file_name, skiprows=2, header=1, sheet_name=sheet)
        emp = a[a[a.columns[0]].str.contains(name, na=False)]
        if emp.empty:
            continue
        for i in range(1, 32):
            date1 = datetime.datetime(2023, 8, i, 0, 0)
            value = ""
            Uhr = ""

            if emp.get(i, None).all():
                value = emp[i].values[0]
                if value == 'D':
                    continue
                elif (emp[i].isna().values[0]):
                    if (i not in dates):
                        rows.append({'date': date1, 'day': calendar.day_name[date1.weekday()], "work_type_code": "Free",
                                     'work_time': "", "work_type": ""
                                     })
                        dates.add(i)
                    continue
                else:
                    rows.append(
                        {'date': date1, 'day': calendar.day_name[date1.weekday()], "work_type_code": value,
                         'work_time': work_type_hours.get(value,value),
                         "work_type": work_type})
    df = pd.DataFrame.from_records(rows)
    df = df.sort_values(by='date')
    return df


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
